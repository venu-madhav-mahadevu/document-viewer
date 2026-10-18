import JSZip from 'jszip';
import { el, sanitizeLinks } from '../sanitize.js';

/*
 * Lightweight, dependency-free PPTX renderer (JSZip + DOMParser only).
 *
 * Covers what most business decks use: text boxes and placeholders (with size,
 * position, alignment and font inheritance from layout/master), pictures,
 * filled/outlined shapes, groups, connector lines, tables and theme colors.
 * Not covered: charts, SmartArt, animations, gradients, EMF/WMF images - those
 * show a labelled placeholder so nothing silently disappears.
 */

const EMU_PER_PX = 9525; // 914400 EMU per inch / 96 px per inch
const PT_TO_PX = 96 / 72;
const px = (emu) => Number(emu || 0) / EMU_PER_PX;

// ---------- XML helpers (namespace-agnostic, matched by localName) ----------
const kids = (node, name) => (node ? [...node.children].filter((c) => !name || c.localName === name) : []);
const child = (node, ...path) => {
  let cur = node;
  for (const name of path) {
    cur = cur && [...cur.children].find((c) => c.localName === name);
    if (!cur) return null;
  }
  return cur;
};
const attr = (node, name) => (node ? node.getAttribute(name) : null);
const rid = (node, name) => node && (node.getAttribute(`r:${name}`) || node.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', name));

function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

class Package {
  constructor(zip) { this.zip = zip; this.xmlCache = new Map(); }

  async xml(path) {
    if (!this.xmlCache.has(path)) {
      const file = this.zip.file(path);
      const text = file ? await file.async('string') : null;
      this.xmlCache.set(path, text ? new DOMParser().parseFromString(text, 'application/xml') : null);
    }
    return this.xmlCache.get(path);
  }

  async rels(path) {
    const relsPath = path.replace(/([^/]+)$/, '_rels/$1.rels');
    const doc = await this.xml(relsPath);
    const map = new Map();
    for (const r of doc ? [...doc.getElementsByTagName('Relationship')] : []) {
      map.set(r.getAttribute('Id'), {
        type: r.getAttribute('Type').split('/').pop(),
        target: r.getAttribute('TargetMode') === 'External' ? r.getAttribute('Target') : resolvePath(path, r.getAttribute('Target')),
        external: r.getAttribute('TargetMode') === 'External',
      });
    }
    return map;
  }
}

// ---------- Colors ----------
function hexToRgb(hex) {
  const n = parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}
function hslToRgb([h, s, l]) {
  if (s === 0) return [l, l, l].map((v) => Math.round(v * 255));
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1; if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3), f(h), f(h - 1 / 3)].map((v) => Math.round(v * 255));
}

const PRESET = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' };
const CLR_MAP = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };

function colorFrom(node, theme) {
  if (!node) return null;
  const c = node.children[0];
  if (!c) return null;
  let hex = null;
  switch (c.localName) {
    case 'srgbClr': hex = attr(c, 'val'); break;
    case 'sysClr': hex = attr(c, 'lastClr') || '000000'; break;
    case 'prstClr': hex = PRESET[attr(c, 'val')] || '000000'; break;
    case 'schemeClr': { const v = attr(c, 'val'); hex = theme[CLR_MAP[v] || v] || null; break; }
    default: return null;
  }
  if (!hex) return null;
  let rgb = hexToRgb(hex);
  let alpha = 1;
  for (const mod of c.children) {
    const v = Number(attr(mod, 'val')) / 100000;
    if (mod.localName === 'lumMod' || mod.localName === 'lumOff') {
      const hsl = rgbToHsl(rgb);
      hsl[2] = mod.localName === 'lumMod' ? hsl[2] * v : Math.min(1, hsl[2] + v);
      rgb = hslToRgb(hsl);
    } else if (mod.localName === 'tint') {
      rgb = rgb.map((ch) => Math.round(ch + (255 - ch) * (1 - v)));
    } else if (mod.localName === 'shade') {
      rgb = rgb.map((ch) => Math.round(ch * v));
    } else if (mod.localName === 'alpha') {
      alpha = v;
    }
  }
  return `rgba(${rgb.join(',')},${alpha})`;
}

async function loadTheme(pkg, masterPath) {
  const rels = await pkg.rels(masterPath);
  const themeRel = [...rels.values()].find((r) => r.type === 'theme');
  const doc = themeRel && await pkg.xml(themeRel.target);
  const scheme = doc && doc.getElementsByTagNameNS('*', 'clrScheme')[0];
  const theme = { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6' };
  for (const c of scheme ? scheme.children : []) {
    const inner = c.children[0];
    if (!inner) continue;
    // sysClr's `val` is a name like "windowText"; the hex is in `lastClr`.
    theme[c.localName] = inner.localName === 'sysClr' ? (attr(inner, 'lastClr') || '000000') : attr(inner, 'val');
  }
  const fonts = doc && doc.getElementsByTagNameNS('*', 'fontScheme')[0];
  theme.majorFont = attr(child(fonts, 'majorFont', 'latin'), 'typeface') || 'Calibri Light';
  theme.minorFont = attr(child(fonts, 'minorFont', 'latin'), 'typeface') || 'Calibri';
  return theme;
}

// ---------- Placeholder inheritance ----------
function placeholderOf(sp) {
  const ph = child(sp, 'nvSpPr', 'nvPr', 'ph') || child(sp, 'nvPicPr', 'nvPr', 'ph') || child(sp, 'nvGraphicFramePr', 'nvPr', 'ph');
  if (!ph) return null;
  return { type: attr(ph, 'type') || 'body', idx: attr(ph, 'idx') };
}

function indexPlaceholders(doc) {
  const byIdx = new Map(), byType = new Map();
  if (!doc) return { byIdx, byType };
  for (const sp of doc.getElementsByTagNameNS('*', 'sp')) {
    const ph = placeholderOf(sp);
    if (!ph) continue;
    if (ph.idx != null && !byIdx.has(ph.idx)) byIdx.set(ph.idx, sp);
    if (!byType.has(ph.type)) byType.set(ph.type, sp);
  }
  return { byIdx, byType };
}

const TYPE_ALIASES = { ctrTitle: 'title', subTitle: 'body', obj: 'body' };

function findPlaceholder(index, ph) {
  if (!ph) return null;
  return (ph.idx != null && index.byIdx.get(ph.idx)) || index.byType.get(ph.type) ||
    index.byType.get(TYPE_ALIASES[ph.type]) || null;
}

// ---------- Rendering ----------
class SlideContext {
  constructor({ pkg, theme, rels, slidePath, layoutIdx, masterIdx, masterDoc, blobUrls }) {
    Object.assign(this, { pkg, theme, rels, slidePath, layoutIdx, masterIdx, masterDoc, blobUrls });
  }

  inherited(sp) {
    const ph = placeholderOf(sp);
    if (!ph) return [];
    return [findPlaceholder(this.layoutIdx, ph), findPlaceholder(this.masterIdx, ph)].filter(Boolean);
  }

  async image(embedId) {
    const rel = this.rels.get(embedId);
    if (!rel || rel.external) return null;
    if (/\.(emf|wmf|tiff?)$/i.test(rel.target)) return { unsupported: rel.target.split('.').pop().toUpperCase() };
    const file = this.pkg.zip.file(rel.target);
    if (!file) return null;
    const ext = rel.target.split('.').pop().toLowerCase();
    const mime = { svg: 'image/svg+xml', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' }[ext] || 'application/octet-stream';
    const url = URL.createObjectURL(new Blob([await file.async('arraybuffer')], { type: mime }));
    this.blobUrls.push(url);
    return { url };
  }
}

function readXfrm(xfrm) {
  if (!xfrm) return null;
  const off = child(xfrm, 'off'), ext = child(xfrm, 'ext');
  if (!off || !ext) return null;
  return {
    x: Number(attr(off, 'x')), y: Number(attr(off, 'y')),
    cx: Number(attr(ext, 'cx')), cy: Number(attr(ext, 'cy')),
    rot: Number(attr(xfrm, 'rot') || 0) / 60000,
    flipH: attr(xfrm, 'flipH') === '1', flipV: attr(xfrm, 'flipV') === '1',
  };
}

function place(node, box, map) {
  const p = map({ x: box.x, y: box.y, cx: box.cx, cy: box.cy });
  Object.assign(node.style, {
    position: 'absolute', left: `${px(p.x)}px`, top: `${px(p.y)}px`,
    width: `${px(p.cx)}px`, height: `${px(p.cy)}px`,
  });
  const t = [];
  if (box.rot) t.push(`rotate(${box.rot}deg)`);
  if (box.flipH) t.push('scaleX(-1)');
  if (box.flipV) t.push('scaleY(-1)');
  if (t.length) node.style.transform = t.join(' ');
}

function applyShapeStyle(node, spPr, ctx, styleNode) {
  const fill = child(spPr, 'solidFill');
  if (fill) node.style.background = colorFrom(fill, ctx.theme);
  else if (!child(spPr, 'noFill') && styleNode) {
    const ref = child(styleNode, 'fillRef');
    if (ref && attr(ref, 'idx') !== '0') node.style.background = colorFrom(ref, ctx.theme);
  }
  const ln = child(spPr, 'ln');
  if (ln && !child(ln, 'noFill')) {
    const color = colorFrom(child(ln, 'solidFill'), ctx.theme);
    if (color) node.style.border = `${Math.max(1, px(attr(ln, 'w') || 9525))}px solid ${color}`;
  } else if (!ln && styleNode) {
    const ref = child(styleNode, 'lnRef');
    const color = ref && attr(ref, 'idx') !== '0' && colorFrom(ref, ctx.theme);
    if (color) node.style.border = `1px solid ${color}`;
  }
  const geom = attr(child(spPr, 'prstGeom'), 'prst');
  if (geom === 'ellipse') node.style.borderRadius = '50%';
  else if (geom === 'roundRect') node.style.borderRadius = '10%';
}

const DEFAULT_SIZES = { title: 44, ctrTitle: 54, subTitle: 24, body: [28, 24, 20, 18, 18], other: 18 };

function inheritedRunProps(ctx, sp, lvl) {
  const ph = placeholderOf(sp);
  const chain = [child(sp, 'txBody', 'lstStyle'), ...ctx.inherited(sp).map((s) => child(s, 'txBody', 'lstStyle'))];
  const txStyles = child(ctx.masterDoc?.documentElement, 'txStyles');
  const styleName = !ph ? 'otherStyle' : (ph.type === 'title' || ph.type === 'ctrTitle') ? 'titleStyle' : 'bodyStyle';
  chain.push(child(txStyles, styleName));
  const props = {};
  for (const lst of chain) {
    const lvlPr = child(lst, `lvl${lvl + 1}pPr`);
    const def = child(lvlPr, 'defRPr');
    if (props.sz == null && attr(def, 'sz')) props.sz = Number(attr(def, 'sz')) / 100;
    if (props.color == null && child(def, 'solidFill')) props.color = colorFrom(child(def, 'solidFill'), ctx.theme);
    if (props.algn == null && attr(lvlPr, 'algn')) props.algn = attr(lvlPr, 'algn');
    if (props.b == null && attr(def, 'b')) props.b = attr(def, 'b') === '1';
    // `in`, not `== null`: an explicit buNone (null) must stop inheritance.
    if (!('bullet' in props) && lvlPr) {
      if (child(lvlPr, 'buNone')) props.bullet = null;
      else if (child(lvlPr, 'buChar')) props.bullet = attr(child(lvlPr, 'buChar'), 'char');
    }
  }
  if (props.sz == null) {
    const d = DEFAULT_SIZES[ph?.type] ?? (ph ? DEFAULT_SIZES.body : DEFAULT_SIZES.other);
    props.sz = Array.isArray(d) ? d[Math.min(lvl, d.length - 1)] : d;
  }
  return { ...props, isTitle: styleName === 'titleStyle' };
}

function renderText(txBody, sp, ctx) {
  const bodyPr = child(txBody, 'bodyPr');
  const inheritedBodyPr = ctx.inherited(sp).map((s) => child(s, 'txBody', 'bodyPr')).filter(Boolean);
  const pick = (name) => [bodyPr, ...inheritedBodyPr].map((b) => attr(b, name)).find((v) => v != null);

  const box = el('div', 'dv-pptx-text');
  const anchor = pick('anchor') || 't';
  box.style.justifyContent = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
  const ins = (name, def) => px(pick(name) ?? def);
  box.style.padding = `${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px`;
  if (pick('vert') === 'vert' || pick('vert') === 'eaVert') box.style.writingMode = 'vertical-rl';
  if (pick('wrap') === 'none') box.style.whiteSpace = 'nowrap';
  const fontScale = Number(attr(child(bodyPr, 'normAutofit'), 'fontScale') || 100000) / 100000;
  // Shapes (not text boxes) take their default text color from the shape style.
  const styleColor = colorFrom(child(sp, 'style', 'fontRef'), ctx.theme);

  for (const p of kids(txBody, 'p')) {
    const pPr = child(p, 'pPr');
    const lvl = Number(attr(pPr, 'lvl') || 0);
    const inh = inheritedRunProps(ctx, sp, lvl);
    const para = el('p');
    para.style.textAlign = { ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[attr(pPr, 'algn') || inh.algn] || 'left';
    if (lvl) para.style.marginLeft = `${lvl * 24}px`;

    let bullet = inh.bullet;
    if (child(pPr, 'buNone')) bullet = null;
    else if (child(pPr, 'buChar')) bullet = attr(child(pPr, 'buChar'), 'char');
    else if (child(pPr, 'buAutoNum')) bullet = '•';
    const hasText = kids(p).some((r) => (r.localName === 'r' || r.localName === 'fld') && child(r, 't')?.textContent);
    if (bullet && hasText) para.append(el('span', 'dv-pptx-bullet', `${bullet} `));

    const endSz = Number(attr(child(p, 'endParaRPr'), 'sz') || 0) / 100;
    para.style.fontSize = `${(endSz || inh.sz) * fontScale * PT_TO_PX}px`;
    para.style.minHeight = '1em';

    for (const r of kids(p)) {
      if (r.localName === 'br') { para.append(el('br')); continue; }
      if (r.localName !== 'r' && r.localName !== 'fld') continue;
      const rPr = child(r, 'rPr');
      const span = el('span', null, child(r, 't')?.textContent ?? '');
      const sz = Number(attr(rPr, 'sz') || 0) / 100 || inh.sz;
      span.style.fontSize = `${sz * fontScale * PT_TO_PX}px`;
      const bold = attr(rPr, 'b') != null ? attr(rPr, 'b') === '1' : inh.b;
      if (bold) span.style.fontWeight = '700';
      if (attr(rPr, 'i') === '1') span.style.fontStyle = 'italic';
      const u = attr(rPr, 'u');
      if (u && u !== 'none') span.style.textDecoration = 'underline';
      if (attr(rPr, 'strike') && attr(rPr, 'strike') !== 'noStrike') span.style.textDecoration = 'line-through';
      const color = colorFrom(child(rPr, 'solidFill'), ctx.theme) ||
        (placeholderOf(sp) ? inh.color : styleColor || inh.color) || `#${ctx.theme.dk1}`;
      span.style.color = color;
      const face = attr(child(rPr, 'latin'), 'typeface');
      const font = !face || face === '+mn-lt' ? ctx.theme.minorFont : face === '+mj-lt' ? ctx.theme.majorFont : face;
      span.style.fontFamily = `"${(inh.isTitle && !face) ? ctx.theme.majorFont : font}", "Segoe UI", Arial, sans-serif`;
      const link = child(rPr, 'hlinkClick');
      const rel = link && ctx.rels.get(rid(link, 'id'));
      if (rel?.external) {
        const a = el('a');
        a.href = rel.target;
        a.append(span);
        para.append(a);
      } else {
        para.append(span);
      }
    }
    box.append(para);
  }
  return box;
}

async function renderShape(sp, ctx, map) {
  const spPr = child(sp, 'spPr');
  let box = readXfrm(child(spPr, 'xfrm'));
  if (!box) {
    for (const parent of ctx.inherited(sp)) {
      box = readXfrm(child(parent, 'spPr', 'xfrm'));
      if (box) break;
    }
  }
  if (!box) return null;
  const node = el('div', 'dv-pptx-shape');
  place(node, box, map);
  applyShapeStyle(node, spPr, ctx, child(sp, 'style'));
  const blip = child(spPr, 'blipFill', 'blip');
  if (blip) {
    const img = await ctx.image(rid(blip, 'embed'));
    if (img?.url) Object.assign(node.style, { backgroundImage: `url("${img.url}")`, backgroundSize: 'cover' });
  }
  const txBody = child(sp, 'txBody');
  if (txBody) node.append(renderText(txBody, sp, ctx));
  return node;
}

async function renderPicture(pic, ctx, map) {
  const box = readXfrm(child(pic, 'spPr', 'xfrm'));
  if (!box) return null;
  const blip = child(pic, 'blipFill', 'blip');
  const img = blip && await ctx.image(rid(blip, 'embed'));
  if (!img) return null;
  if (img.unsupported) return placeholderBox(box, map, `${img.unsupported} image`);
  const node = el('img', 'dv-pptx-pic');
  node.alt = attr(child(pic, 'nvPicPr', 'cNvPr'), 'descr') || '';
  node.src = img.url;
  const src = child(pic, 'blipFill', 'srcRect');
  if (src) {
    // Cropped pictures: show the cropped region by scaling an inner image.
    const [l, t, r, b] = ['l', 't', 'r', 'b'].map((k) => Number(attr(src, k) || 0) / 100000);
    const wrap = el('div', 'dv-pptx-crop');
    place(wrap, box, map);
    Object.assign(node.style, {
      position: 'absolute', width: `${100 / (1 - l - r)}%`, height: `${100 / (1 - t - b)}%`,
      left: `${(-l / (1 - l - r)) * 100}%`, top: `${(-t / (1 - t - b)) * 100}%`,
    });
    wrap.append(node);
    return wrap;
  }
  place(node, box, map);
  return node;
}

function placeholderBox(box, map, label) {
  const node = el('div', 'dv-pptx-unsupported', label);
  place(node, box, map);
  return node;
}

function renderConnector(cxn, ctx, map) {
  const spPr = child(cxn, 'spPr');
  const box = readXfrm(child(spPr, 'xfrm'));
  if (!box) return null;
  const ln = child(spPr, 'ln');
  const color = colorFrom(child(ln, 'solidFill'), ctx.theme) || colorFrom(child(cxn, 'style', 'lnRef'), ctx.theme) || '#000';
  const width = Math.max(1, px(attr(ln, 'w') || 12700));
  const p = map(box);
  const w = Math.max(px(p.cx), width), h = Math.max(px(p.cy), width);
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'dv-pptx-line');
  Object.assign(svg.style, { position: 'absolute', left: `${px(p.x)}px`, top: `${px(p.y)}px`, width: `${w}px`, height: `${h}px`, overflow: 'visible' });
  if (box.rot) svg.style.transform = `rotate(${box.rot}deg)`;
  const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  const [x1, x2] = box.flipH ? [w, 0] : [0, w];
  const [y1, y2] = box.flipV ? [h, 0] : [0, h];
  Object.entries({ x1, y1, x2, y2, stroke: color, 'stroke-width': width }).forEach(([k, v]) => line.setAttribute(k, v));
  svg.append(line);
  return svg;
}

function renderGraphicFrame(gf, ctx, map) {
  const box = readXfrm(child(gf, 'xfrm'));
  if (!box) return null;
  const data = child(gf, 'graphic', 'graphicData');
  const tbl = child(data, 'tbl');
  if (!tbl) {
    const uri = attr(data, 'uri') || '';
    const label = uri.includes('chart') ? 'Chart' : uri.includes('diagram') ? 'SmartArt diagram' : 'Embedded object';
    return placeholderBox(box, map, `${label} (not shown in preview)`);
  }
  const wrap = el('div', 'dv-pptx-table');
  place(wrap, box, map);
  wrap.style.height = 'auto';
  const table = el('table');
  const colgroup = el('colgroup');
  for (const gc of kids(child(tbl, 'tblGrid'), 'gridCol')) {
    const col = el('col');
    col.style.width = `${px(attr(gc, 'w'))}px`;
    colgroup.append(col);
  }
  table.append(colgroup);
  for (const tr of kids(tbl, 'tr')) {
    const row = el('tr');
    row.style.height = `${px(attr(tr, 'h'))}px`;
    for (const tc of kids(tr, 'tc')) {
      if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') continue;
      const cell = el('td');
      if (attr(tc, 'gridSpan')) cell.colSpan = Number(attr(tc, 'gridSpan'));
      if (attr(tc, 'rowSpan')) cell.rowSpan = Number(attr(tc, 'rowSpan'));
      const fill = colorFrom(child(tc, 'tcPr', 'solidFill'), ctx.theme);
      if (fill) cell.style.background = fill;
      const txBody = child(tc, 'txBody');
      if (txBody) cell.append(renderText(txBody, tc, ctx));
      row.append(cell);
    }
    table.append(row);
  }
  wrap.append(table);
  return wrap;
}

async function renderTree(tree, ctx, map, out) {
  for (const node of kids(tree)) {
    let rendered = null;
    switch (node.localName) {
      case 'sp': rendered = await renderShape(node, ctx, map); break;
      case 'pic': rendered = await renderPicture(node, ctx, map); break;
      case 'cxnSp': rendered = renderConnector(node, ctx, map); break;
      case 'graphicFrame': rendered = renderGraphicFrame(node, ctx, map); break;
      case 'grpSp': {
        const x = child(node, 'grpSpPr', 'xfrm');
        const off = readXfrm(x);
        const chOff = child(x, 'chOff'), chExt = child(x, 'chExt');
        if (off && chOff && chExt) {
          const sx = off.cx / (Number(attr(chExt, 'cx')) || off.cx || 1);
          const sy = off.cy / (Number(attr(chExt, 'cy')) || off.cy || 1);
          const cx0 = Number(attr(chOff, 'x')), cy0 = Number(attr(chOff, 'y'));
          const inner = (b) => map({ x: off.x + (b.x - cx0) * sx, y: off.y + (b.y - cy0) * sy, cx: b.cx * sx, cy: b.cy * sy });
          await renderTree(node, ctx, inner, out);
        } else {
          await renderTree(node, ctx, map, out);
        }
        break;
      }
      case 'AlternateContent': {
        const fallback = child(node, 'Fallback') || child(node, 'Choice');
        if (fallback) await renderTree(fallback, ctx, map, out);
        break;
      }
      default: break;
    }
    if (rendered) out.append(rendered);
  }
}

async function background(docs, ctx) {
  for (const doc of docs) {
    const bg = doc && child(doc.documentElement, 'cSld', 'bg');
    if (!bg) continue;
    const bgPr = child(bg, 'bgPr');
    if (bgPr) {
      const solid = child(bgPr, 'solidFill');
      if (solid) return { color: colorFrom(solid, ctx.theme) };
      const blip = child(bgPr, 'blipFill', 'blip');
      if (blip) {
        // Resolve against the part that owns the background (slide, layout or master).
        const owner = doc === docs[0] ? ctx.rels : doc.__rels;
        const saved = ctx.rels;
        ctx.rels = owner || saved;
        const img = await ctx.image(rid(blip, 'embed'));
        ctx.rels = saved;
        if (img?.url) return { image: img.url };
      }
      const grad = child(bgPr, 'gradFill', 'gsLst');
      if (grad) {
        const stops = kids(grad, 'gs').map((gs) => `${colorFrom(gs, ctx.theme)} ${Number(attr(gs, 'pos')) / 1000}%`);
        return { image: null, gradient: `linear-gradient(180deg, ${stops.join(', ')})` };
      }
    }
    const ref = child(bg, 'bgRef');
    if (ref) return { color: colorFrom(ref, ctx.theme) };
  }
  return { color: `#${ctx.theme.lt1 || 'FFFFFF'}` };
}

export async function renderPptx({ buffer, mount }) {
  const zip = await JSZip.loadAsync(buffer);
  const pkg = new Package(zip);
  const presPath = 'ppt/presentation.xml';
  const pres = await pkg.xml(presPath);
  if (!pres) throw new Error('This file does not look like a PowerPoint presentation.');
  const presRels = await pkg.rels(presPath);
  const sldSz = pres.getElementsByTagNameNS('*', 'sldSz')[0];
  const W = px(attr(sldSz, 'cx') || 12192000), H = px(attr(sldSz, 'cy') || 6858000);

  const root = el('div', 'dv-pptx');
  mount.append(root);
  const blobUrls = [];
  const themes = new Map();

  const slideIds = [...pres.getElementsByTagNameNS('*', 'sldId')];
  let index = 0;
  for (const sldId of slideIds) {
    index++;
    const rel = presRels.get(rid(sldId, 'id'));
    if (!rel) continue;
    const slidePath = rel.target;
    const slideDoc = await pkg.xml(slidePath);
    if (!slideDoc || attr(slideDoc.documentElement, 'show') === '0') continue;
    const rels = await pkg.rels(slidePath);
    const layoutPath = [...rels.values()].find((r) => r.type === 'slideLayout')?.target;
    const layoutDoc = layoutPath && await pkg.xml(layoutPath);
    const layoutRels = layoutPath ? await pkg.rels(layoutPath) : new Map();
    const masterPath = [...layoutRels.values()].find((r) => r.type === 'slideMaster')?.target;
    const masterDoc = masterPath && await pkg.xml(masterPath);
    if (layoutDoc) layoutDoc.__rels = layoutRels;
    if (masterDoc) masterDoc.__rels = masterPath ? await pkg.rels(masterPath) : new Map();
    if (masterPath && !themes.has(masterPath)) themes.set(masterPath, await loadTheme(pkg, masterPath));
    const theme = themes.get(masterPath) || { dk1: '000000', lt1: 'FFFFFF', minorFont: 'Calibri', majorFont: 'Calibri Light' };

    const ctx = new SlideContext({
      pkg, theme, rels, slidePath, masterDoc,
      layoutIdx: indexPlaceholders(layoutDoc), masterIdx: indexPlaceholders(masterDoc), blobUrls,
    });

    const slide = el('div', 'dv-page dv-pptx-slide');
    slide.style.width = `${W}px`;
    slide.style.height = `${H}px`;
    slide.setAttribute('aria-label', `Slide ${index}`);
    const bg = await background([slideDoc, layoutDoc, masterDoc], ctx);
    if (bg.color) slide.style.background = bg.color;
    if (bg.gradient) slide.style.background = bg.gradient;
    if (bg.image) Object.assign(slide.style, { backgroundImage: `url("${bg.image}")`, backgroundSize: 'cover' });

    const identity = (b) => b;
    // Layout/master decorations (logos, bars) appear unless the slide hides them.
    if (attr(slideDoc.documentElement, 'showMasterSp') !== '0') {
      for (const [doc, docRels] of [[masterDoc, masterDoc?.__rels], [layoutDoc, layoutRels]]) {
        const tree = doc && child(doc.documentElement, 'cSld', 'spTree');
        if (!tree) continue;
        const decorations = el('div', 'dv-pptx-layer');
        const deco = new SlideContext({ ...ctx, rels: docRels, layoutIdx: { byIdx: new Map(), byType: new Map() }, masterIdx: { byIdx: new Map(), byType: new Map() } });
        const nonPlaceholders = tree.cloneNode(true);
        for (const n of [...nonPlaceholders.children]) if (placeholderOf(n)) n.remove();
        await renderTree(nonPlaceholders, deco, identity, decorations);
        slide.append(decorations);
      }
    }
    const content = el('div', 'dv-pptx-layer');
    await renderTree(child(slideDoc.documentElement, 'cSld', 'spTree'), ctx, identity, content);
    slide.append(content);
    slide.append(el('div', 'dv-pptx-num', String(index)));
    root.append(slide);
  }
  sanitizeLinks(root);

  return {
    naturalWidth: () => W,
    destroy() { blobUrls.forEach((u) => URL.revokeObjectURL(u)); },
  };
}
