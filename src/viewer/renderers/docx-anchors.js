/*
 * Floating ("anchored") objects in Word documents: shapes and arrows, text boxes and pictures
 * with text wrapping. docx-preview does not draw them, so before rendering each one is swapped
 * for an invisible marker in its paragraph, and afterwards drawn at its position:
 *   - in front of / behind text (wrapNone): placed over the page at its offsets
 *   - square / tight wrap: a float in its paragraph, so the text runs around it
 *   - top and bottom: a block above the paragraph's text
 */
import { el } from '../sanitize.js';
import { PRESET_SHAPES, svgEl } from './shapes.js';

const MARK = '⁣dvA';
const EMU = 9525; // EMU per CSS px
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

const local = (node, name) => (node ? [...node.getElementsByTagName('*')].filter((n) => n.localName === name) : []);
const first = (node, name) => local(node, name)[0] || null;
const kids = (node, name) => (node ? [...node.children].filter((c) => c.localName === name) : []);
const attr = (node, name) => node?.getAttribute(name) ?? node?.getAttribute(`w:${name}`) ?? null;

/**
 * Replaces each anchored drawing in word/document.xml with a marker run.
 * Returns the anchors (XML elements) by marker number, or null when there are none.
 */
export async function extractAnchors(zip) {
  const part = zip.file('word/document.xml');
  if (!part) return null;
  const xml = await part.async('string');
  if (!xml.includes('<wp:anchor')) return null;
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return null;
  const anchors = [];
  for (const drawing of local(doc, 'drawing')) {
    const anchor = kids(drawing, 'anchor')[0];
    if (!anchor) continue;
    const graphicData = first(anchor, 'graphicData');
    const uri = graphicData?.getAttribute('uri') || '';
    if (!/wordprocessingShape|picture/.test(uri)) continue; // charts, groups, canvases: left to docx-preview
    const t = doc.createElementNS(W_NS, 'w:t');
    t.textContent = `${MARK}${anchors.length}⁣`;
    anchors.push(anchor);
    drawing.replaceWith(t);
  }
  if (!anchors.length) return null;
  zip.file('word/document.xml', new XMLSerializer().serializeToString(doc));
  return anchors;
}

/** Theme colors of the document (for schemeClr). */
export async function readThemeColors(zip) {
  const xml = await zip.file('word/theme/theme1.xml')?.async('string');
  const out = {};
  if (!xml) return out;
  for (const name of ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
    const m = new RegExp(`<a:${name}>[\\s\\S]*?(?:val|lastClr)="([0-9A-Fa-f]{6})"`).exec(xml);
    if (m) out[name] = m[1];
  }
  return out;
}

function colorOf(fill, theme) {
  const c = fill?.children?.[0];
  if (!c) return null;
  let hex = c.localName === 'srgbClr' ? c.getAttribute('val')
    : c.localName === 'schemeClr' ? theme[{ tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[c.getAttribute('val')] || c.getAttribute('val')]
      : c.localName === 'sysClr' ? c.getAttribute('lastClr') : null;
  if (!hex) return null;
  let rgb = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16));
  let alpha = 1;
  for (const m of c.children) {
    const v = Number(m.getAttribute('val')) / 100000;
    if (m.localName === 'lumMod') rgb = rgb.map((x) => x * v);
    else if (m.localName === 'lumOff') rgb = rgb.map((x) => x + 255 * v);
    else if (m.localName === 'tint') rgb = rgb.map((x) => x + (255 - x) * (1 - v));
    else if (m.localName === 'shade') rgb = rgb.map((x) => x * v);
    else if (m.localName === 'alpha') alpha = v;
  }
  hex = rgb.map((x) => Math.round(Math.min(255, Math.max(0, x))));
  return `rgba(${hex.join(',')},${alpha})`;
}

/** Position of an anchor relative to its page (px), and its wrapping. */
function placement(anchor) {
  const axis = (name) => {
    const p = kids(anchor, name)[0];
    const off = kids(p, 'posOffset')[0];
    const align = kids(p, 'align')[0];
    return { from: p?.getAttribute('relativeFrom') || 'column', offset: off ? Number(off.textContent) / EMU : null, align: align?.textContent || null };
  };
  const ext = kids(anchor, 'extent')[0];
  const wrap = ['wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom', 'wrapNone'].find((w) => kids(anchor, w).length) || 'wrapNone';
  const dist = (k) => Number(anchor.getAttribute(k) || 0) / EMU;
  return {
    h: axis('positionH'), v: axis('positionV'),
    w: Number(ext?.getAttribute('cx') || 0) / EMU, ht: Number(ext?.getAttribute('cy') || 0) / EMU,
    wrap, behind: anchor.getAttribute('behindDoc') === '1',
    dist: { l: dist('distL'), r: dist('distR'), t: dist('distT'), b: dist('distB') },
  };
}

/** The drawing itself: a picture, or a shape (outline + fill) with an optional text box. */
async function drawObject(anchor, ctx, w, h) {
  const box = el('div', 'dv-docx-float');
  box.style.width = `${w}px`;
  box.style.height = `${h}px`;
  const pic = first(anchor, 'pic');
  if (pic) {
    const rid = first(pic, 'blip')?.getAttribute('r:embed');
    const url = rid && await ctx.image(rid);
    if (url) {
      const img = el('img', 'dv-docx-float-pic');
      img.src = url;
      img.alt = first(anchor, 'docPr')?.getAttribute('descr') || '';
      box.append(img);
    }
    return box;
  }
  const wsp = first(anchor, 'wsp');
  if (!wsp) return box;
  const spPr = kids(wsp, 'spPr')[0];
  const style = kids(wsp, 'style')[0];
  const geom = kids(spPr, 'prstGeom')[0]?.getAttribute('prst') || 'rect';
  const xfrm = kids(spPr, 'xfrm')[0];
  const rot = Number(xfrm?.getAttribute('rot') || 0) / 60000;
  const flipH = xfrm?.getAttribute('flipH') === '1';
  const flipV = xfrm?.getAttribute('flipV') === '1';
  const tf = [rot && `rotate(${rot}deg)`, flipH && 'scaleX(-1)', flipV && 'scaleY(-1)'].filter(Boolean);
  if (tf.length) box.style.transform = tf.join(' ');
  const fill = kids(spPr, 'noFill').length ? null
    : colorOf(kids(spPr, 'solidFill')[0], ctx.theme) || (style && colorOf(kids(style, 'fillRef')[0], ctx.theme));
  const ln = kids(spPr, 'ln')[0];
  const stroke = ln && kids(ln, 'noFill').length ? null
    : colorOf(kids(ln, 'solidFill')[0], ctx.theme) || (style && colorOf(kids(style, 'lnRef')[0], ctx.theme));
  const strokeW = Math.max(1, Number(ln?.getAttribute('w') || 9525) / EMU);
  const svg = svgEl('svg', { class: 'dv-docx-geom', viewBox: `0 0 ${Math.max(w, 1)} ${Math.max(h, 1)}`, preserveAspectRatio: 'none', 'aria-hidden': 'true' });
  if (geom === 'line' || geom.startsWith('straightConnector') || geom.startsWith('bentConnector')) {
    const line = svgEl('line', { x1: 0, y1: 0, x2: w, y2: h, stroke: stroke || '#000', 'stroke-width': strokeW });
    const defs = svgEl('defs');
    for (const [end, which] of [['headEnd', 'marker-start'], ['tailEnd', 'marker-end']]) {
      const type = kids(ln, end)[0]?.getAttribute('type');
      if (!type || type === 'none') continue;
      const id = `dvw${Math.random().toString(36).slice(2, 9)}`;
      const marker = svgEl('marker', { id, viewBox: '0 0 10 10', refX: type === 'oval' || type === 'diamond' ? 5 : 9, refY: 5, markerWidth: 3, markerHeight: 3, orient: 'auto-start-reverse', markerUnits: 'strokeWidth' });
      marker.append(type === 'oval' ? svgEl('circle', { cx: 5, cy: 5, r: 4.5, fill: stroke || '#000' })
        : svgEl('path', { d: type === 'diamond' ? 'M0 5L5 0L10 5L5 10Z' : type === 'stealth' ? 'M0 0L10 5L0 10L3 5Z' : 'M0 0L10 5L0 10Z', fill: stroke || '#000' }));
      defs.append(marker);
      line.setAttribute(which, `url(#${id})`);
    }
    svg.append(defs, line);
  } else {
    const gd = new Map(local(kids(spPr, 'prstGeom')[0], 'gd').map((g) => [g.getAttribute('name'), Number(/val\s+(-?\d+)/.exec(g.getAttribute('fmla') || '')?.[1])]));
    const adj = (name, def) => (Number.isFinite(gd.get(name)) ? gd.get(name) : def);
    const draw = PRESET_SHAPES[geom];
    const shape = draw ? svgEl('path', { d: draw(w, h, adj) })
      : geom === 'ellipse' ? svgEl('ellipse', { cx: w / 2, cy: h / 2, rx: w / 2, ry: h / 2 })
        : geom === 'roundRect' ? svgEl('rect', { x: 0, y: 0, width: w, height: h, rx: Math.min(w, h) * adj('adj', 16667) / 100000 })
          : svgEl('rect', { x: 0, y: 0, width: w, height: h });
    shape.setAttribute('fill', fill || 'none');
    if (stroke) { shape.setAttribute('stroke', stroke); shape.setAttribute('stroke-width', strokeW); shape.setAttribute('vector-effect', 'non-scaling-stroke'); }
    svg.append(shape);
  }
  box.append(svg);
  // text inside the shape / text box
  const content = first(wsp, 'txbxContent');
  if (content) {
    const bodyPr = kids(wsp, 'bodyPr')[0];
    const ins = (k, d) => Number(bodyPr?.getAttribute(k) ?? d) / EMU;
    const text = el('div', 'dv-docx-float-text');
    text.style.padding = `${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px`;
    text.style.justifyContent = { ctr: 'center', b: 'flex-end' }[bodyPr?.getAttribute('anchor')] || 'flex-start';
    for (const p of kids(content, 'p')) {
      const para = el('p');
      const jc = attr(first(p, 'jc'), 'val');
      if (jc) para.style.textAlign = { both: 'justify', center: 'center', right: 'right', end: 'right' }[jc] || 'left';
      for (const r of kids(p, 'r')) {
        const span = el('span', null, local(r, 't').map((t) => t.textContent).join(''));
        const rPr = kids(r, 'rPr')[0];
        if (kids(rPr, 'b').length) span.style.fontWeight = '700';
        if (kids(rPr, 'i').length) span.style.fontStyle = 'italic';
        const sz = attr(kids(rPr, 'sz')[0], 'val');
        if (sz) span.style.fontSize = `${Number(sz) / 2}pt`;
        const color = attr(kids(rPr, 'color')[0], 'val');
        if (color && color !== 'auto') span.style.color = `#${color}`;
        para.append(span);
      }
      if (!para.childNodes.length) para.append(el('br'));
      text.append(para);
    }
    box.append(text);
  }
  return box;
}

/**
 * Draws the anchors at their markers after docx-preview has rendered the pages.
 * @param {HTMLElement[]} pages the rendered section.docx pages
 * @param {Element[]} anchors from extractAnchors
 * @param {{ theme: object, image: (rid: string) => Promise<string|null> }} ctx
 */
export async function placeAnchors(pages, anchors, ctx) {
  for (const page of pages) {
    const walker = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
    const hits = [];
    for (let n = walker.nextNode(); n; n = walker.nextNode()) if (n.data.includes(MARK)) hits.push(n);
    if (!hits.length) continue;
    if (getComputedStyle(page).position === 'static') page.style.position = 'relative';
    const cs = getComputedStyle(page);
    const margin = { l: parseFloat(cs.paddingLeft) || 0, r: parseFloat(cs.paddingRight) || 0, t: parseFloat(cs.paddingTop) || 0 };
    const pageW = page.offsetWidth;
    const content = pageW - margin.l - margin.r;
    for (const node of hits) {
      const m = new RegExp(`${MARK}(\\d+)⁣`).exec(node.data);
      node.data = node.data.replace(/⁣dvA\d+⁣/g, '');
      const anchor = m && anchors[Number(m[1])];
      if (!anchor) continue;
      const para = node.parentElement?.closest('p') || node.parentElement;
      const pos = placement(anchor);
      const obj = await drawObject(anchor, ctx, pos.w, pos.ht);
      // horizontal: from the column (text area) or the page, by offset or alignment
      const base = { page: 0, margin: margin.l, column: margin.l, leftMargin: 0, character: margin.l }[pos.h.from] ?? margin.l;
      const span = pos.h.from === 'page' ? pageW : content;
      let x = pos.h.offset != null ? base + pos.h.offset
        : base + ({ left: 0, center: (span - pos.w) / 2, right: span - pos.w, inside: 0, outside: span - pos.w }[pos.h.align] ?? 0);
      // vertical: from the paragraph / line, the margin or the page
      const paraTop = para ? para.offsetTop : margin.t;
      const vBase = { paragraph: paraTop, line: paraTop, page: 0, margin: margin.t, topMargin: 0 }[pos.v.from] ?? paraTop;
      const y = pos.v.offset != null ? vBase + pos.v.offset : vBase;
      if ((pos.wrap === 'wrapSquare' || pos.wrap === 'wrapTight' || pos.wrap === 'wrapThrough') && para && pos.v.from === 'paragraph' && (pos.v.offset ?? 0) >= 0) {
        // text wraps around it: float on the side it sits on, inside its paragraph
        const right = x - margin.l + pos.w / 2 > content / 2;
        Object.assign(obj.style, {
          float: right ? 'right' : 'left', marginTop: `${pos.v.offset || 0}px`,
          marginLeft: right ? `${pos.dist.l}px` : `${Math.max(0, x - margin.l)}px`,
          marginRight: right ? `${Math.max(0, margin.l + content - x - pos.w)}px` : `${pos.dist.r}px`,
          marginBottom: `${pos.dist.b}px`, position: 'relative',
        });
        para.prepend(obj);
      } else if (pos.wrap === 'wrapTopAndBottom' && para) {
        Object.assign(obj.style, { display: 'block', marginLeft: `${Math.max(0, x - margin.l)}px`, marginTop: `${pos.v.offset || 0}px`, marginBottom: `${pos.dist.b}px`, position: 'relative' });
        para.prepend(obj);
      } else {
        // in front of (or behind) the text, at its exact place on the page
        x = Math.max(-pos.w, Math.min(x, pageW));
        Object.assign(obj.style, { position: 'absolute', left: `${x}px`, top: `${y}px`, zIndex: pos.behind ? '-1' : '2' });
        if (pos.behind) page.style.isolation = 'isolate';
        page.append(obj);
      }
    }
  }
}
