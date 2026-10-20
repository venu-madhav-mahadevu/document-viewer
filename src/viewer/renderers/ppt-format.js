/*
 * PowerPoint 97-2003 (.ppt) slide reader, following Microsoft's published
 * [MS-PPT] and [MS-ODRAW] specifications. Reads each slide as a drawing:
 *   - shapes: position, size, rotation, flips, fill, outline, rectangles,
 *     rounded rectangles, ellipses, lines, pictures and groups (tables)
 *   - text: paragraphs (alignment, indents, spacing, bullets) and runs
 *     (bold, italic, underline, size, color, font, super/subscript)
 *   - the slide master: background, decorations and default text styles
 *   - color schemes and the font list
 * Coordinates are in "master units" (576 per inch).
 */

const cp1252 = new TextDecoder('windows-1252');
const utf16 = new TextDecoder('utf-16le');

const R = {
  DOCUMENT: 0x03e8, DOCUMENT_ATOM: 0x03e9, SLIDE: 0x03ee, SLIDE_ATOM: 0x03ef, NOTES: 0x03f0,
  SLIDE_PERSIST: 0x03f3, MAIN_MASTER: 0x03f8, ENVIRONMENT: 0x03f2, PPDRAWING: 0x040c, DRAWING_GROUP: 0x040b,
  COLOR_SCHEME: 0x07f0, FONT_COLLECTION: 0x07d5, FONT_ENTITY: 0x0fb7, PLACEHOLDER: 0x0bc3,
  OUTLINE_TEXT_REF: 0x0f9e, TEXT_HEADER: 0x0f9f, TEXT_CHARS: 0x0fa0, STYLE_TEXT_PROP: 0x0fa1,
  TX_MASTER_STYLE: 0x0fa3, TEXT_BYTES: 0x0fa8, SLIDE_LIST: 0x0ff0, PERSIST_DIR: 0x1772,
  DGG: 0xf000, BSTORE: 0xf001, DG: 0xf002, SPGR: 0xf003, SP: 0xf004, FBSE: 0xf007, FSPGR: 0xf009,
  FSP: 0xf00a, FOPT: 0xf00b, CLIENT_TEXTBOX: 0xf00d, CHILD_ANCHOR: 0xf00f, CLIENT_ANCHOR: 0xf010, CLIENT_DATA: 0xf011,
};

const DEFAULT_SCHEME = ['#ffffff', '#000000', '#808080', '#000000', '#bbe0e3', '#333399', '#009999', '#99cc00'];
const EMU_PER_MU = 914400 / 576;

const u16 = (b, p) => b[p] | (b[p + 1] << 8);
const i16 = (b, p) => (u16(b, p) << 16) >> 16;
const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
const i32 = (b, p) => b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24);
const hex = (r, g, b) => `#${[r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('')}`;

/** Child records of a container: [{ type, inst, body, end, container }]. */
function records(b, start, end) {
  const out = [];
  let p = start;
  while (p + 8 <= end) {
    const verInst = u16(b, p);
    const len = u32(b, p + 4);
    const body = p + 8;
    if (body + len > end) break;
    out.push({ type: u16(b, p + 2), inst: verInst >> 4, container: (verInst & 0xf) === 0xf, body, end: body + len });
    p = body + len;
  }
  return out;
}
const child = (b, rec, type) => records(b, rec.body, rec.end).find((r) => r.type === type);
const children = (b, rec, type) => records(b, rec.body, rec.end).filter((r) => r.type === type);

// ------------------------------------------------------------ text styles

function colorIndex(b, p, scheme) {
  const index = b[p + 3];
  if (index === 0xfe) return hex(b[p], b[p + 1], b[p + 2]);
  if (index < 8) return scheme?.[index] ?? DEFAULT_SCHEME[index];
  return undefined;
}

/** TextPFException (paragraph properties); returns [props, nextOffset]. */
function readPF(b, p, scheme) {
  const masks = u32(b, p); p += 4;
  const pf = {};
  if (masks & 0x000f) {
    const flags = u16(b, p); p += 2;
    if (masks & 0x1) pf.hasBullet = !!(flags & 0x1);
    if (masks & 0x4) pf.bulletHasColor = !!(flags & 0x4);
    if (masks & 0x8) pf.bulletHasSize = !!(flags & 0x8);
  }
  if (masks & 0x0080) { pf.bulletChar = u16(b, p); p += 2; }
  if (masks & 0x0010) { pf.bulletFont = u16(b, p); p += 2; }
  if (masks & 0x0040) { pf.bulletSize = i16(b, p); p += 2; }
  if (masks & 0x0020) { pf.bulletColor = colorIndex(b, p, scheme); p += 4; }
  if (masks & 0x0800) { pf.align = u16(b, p); p += 2; }
  if (masks & 0x1000) { pf.lineSpacing = i16(b, p); p += 2; }
  if (masks & 0x2000) { pf.spaceBefore = i16(b, p); p += 2; }
  if (masks & 0x4000) { pf.spaceAfter = i16(b, p); p += 2; }
  if (masks & 0x0100) { pf.leftMargin = i16(b, p); p += 2; }
  if (masks & 0x0400) { pf.indent = i16(b, p); p += 2; }
  if (masks & 0x8000) p += 2;                                  // defaultTabSize
  if (masks & 0x100000) { const n = u16(b, p); p += 2 + n * 4; } // tabStops
  if (masks & 0x10000) p += 2;                                  // fontAlign
  if (masks & 0xe0000) p += 2;                                  // wrapFlags
  if (masks & 0x200000) p += 2;                                 // textDirection
  return [pf, p];
}

/** TextCFException (character properties); returns [props, nextOffset]. */
function readCF(b, p, scheme) {
  const masks = u32(b, p); p += 4;
  const cf = {};
  if (masks & 0xffff) {
    const style = u16(b, p); p += 2;
    if (masks & 0x1) cf.bold = !!(style & 0x1);
    if (masks & 0x2) cf.italic = !!(style & 0x2);
    if (masks & 0x4) cf.underline = !!(style & 0x4);
  }
  if (masks & 0x10000) { cf.font = u16(b, p); p += 2; }
  if (masks & 0x200000) p += 2; // oldEAFontRef
  if (masks & 0x400000) p += 2; // ansiFontRef
  if (masks & 0x800000) p += 2; // symbolFontRef
  if (masks & 0x20000) { cf.size = u16(b, p); p += 2; }
  if (masks & 0x40000) { cf.color = colorIndex(b, p, scheme); p += 4; }
  if (masks & 0x80000) { cf.position = i16(b, p); p += 2; }
  if (masks & 0x100000) p += 4;   // pp10ext
  if (masks & 0x1000000) p += 2;  // newEAFontRef
  if (masks & 0x2000000) p += 2;  // csFontRef
  if (masks & 0x4000000) p += 4;  // pp11ext
  return [cf, p];
}

/** StyleTextPropAtom: paragraph runs and character runs over `length` characters. */
function readStyleRuns(b, rec, length, scheme) {
  const pf = [];
  const cf = [];
  let p = rec.body;
  try {
    for (let covered = 0; covered < length && p + 6 <= rec.end;) {
      const count = u32(b, p);
      const level = u16(b, p + 4);
      const [props, next] = readPF(b, p + 6, scheme);
      pf.push({ count, level, props });
      covered += count; p = next;
    }
    for (let covered = 0; covered < length && p + 4 <= rec.end;) {
      const count = u32(b, p);
      const [props, next] = readCF(b, p + 4, scheme);
      cf.push({ count, props });
      covered += count; p = next;
    }
  } catch { /* keep what was read */ }
  return { pf, cf };
}

/** TxMasterStyleAtom: per-level default paragraph and character properties. */
function readMasterStyle(b, rec, scheme) {
  const levels = [];
  let p = rec.body;
  const n = u16(b, p); p += 2;
  for (let i = 0; i < n && p < rec.end; i++) {
    if (rec.inst >= 5) p += 2;
    const [pf, q] = readPF(b, p, scheme);
    const [cf, r] = readCF(b, q, scheme);
    levels.push({ pf, cf });
    p = r;
  }
  return levels;
}

// Text types: 0 title, 1 body, 2 notes, 4 other, 5 centered body (subtitle), 6 centered title, 7 half body, 8 quarter body.
const BASE_TYPE = { 5: 1, 6: 0, 7: 1, 8: 1 };

function styleFor(masterStyles, type, level) {
  const merged = { pf: {}, cf: {} };
  const chain = [4, BASE_TYPE[type], type].filter((t, i, a) => t !== undefined && a.indexOf(t) === i);
  for (const t of chain) {
    const levels = masterStyles[t];
    if (!levels) continue;
    for (const l of [levels[0], levels[Math.min(level, levels.length - 1)]]) {
      if (!l) continue;
      Object.assign(merged.pf, l.pf);
      Object.assign(merged.cf, l.cf);
    }
  }
  return merged;
}

// ------------------------------------------------------------------ shapes

function readOpt(b, rec) {
  const props = {};
  if (!rec) return props;
  for (let i = 0; i < rec.inst; i++) {
    const at = rec.body + i * 6;
    if (at + 6 > rec.end) break;
    props[u16(b, at) & 0x3fff] = u32(b, at + 2);
  }
  return props;
}

function optColor(v, scheme) {
  if (v === undefined) return undefined;
  const r = v & 0xff; const g = (v >> 8) & 0xff; const bl = (v >> 16) & 0xff; const flags = v >>> 24;
  if (flags & 0x08) return scheme?.[r] ?? DEFAULT_SCHEME[r] ?? '#000000';
  if (flags & 0x10) return undefined; // system colors: use the default
  return hex(r, g, bl);
}

const boolProp = (v, bit, useBit) => (v === undefined || !(v & useBit) ? undefined : !!(v & bit));

function readAnchor(b, sp) {
  const ca = child(b, sp, R.CLIENT_ANCHOR);
  if (ca) {
    const len = ca.end - ca.body;
    if (len >= 16) return { t: i32(b, ca.body), l: i32(b, ca.body + 4), r: i32(b, ca.body + 8), bt: i32(b, ca.body + 12) };
    if (len >= 8) return { t: i16(b, ca.body), l: i16(b, ca.body + 2), r: i16(b, ca.body + 4), bt: i16(b, ca.body + 6) };
  }
  const ch = child(b, sp, R.CHILD_ANCHOR);
  if (ch) return { l: i32(b, ch.body), t: i32(b, ch.body + 4), r: i32(b, ch.body + 8), bt: i32(b, ch.body + 12), child: true };
  return null;
}

// --------------------------------------------------------------- document

/**
 * Reads a PowerPoint 97-2003 file.
 * @returns {{ size: {w,h}, slides: Array<{ background?: string, shapes: Shape[] }> }}
 */
export function readPresentation(doc, pictureBlobs) {
  const top = records(doc, 0, doc.length);

  // Persist directory: persist id -> byte offset of the record (later edits win).
  const persist = new Map();
  for (const rec of top.filter((r) => r.type === R.PERSIST_DIR)) {
    let p = rec.body;
    while (p + 4 <= rec.end) {
      const v = u32(doc, p); p += 4;
      const id = v & 0xfffff; const n = v >>> 20;
      for (let k = 0; k < n && p + 4 <= rec.end; k++, p += 4) persist.set(id + k, u32(doc, p));
    }
  }
  const recordAt = (offset) => records(doc, offset, doc.length)[0];

  const docRec = top.find((r) => r.type === R.DOCUMENT);
  if (!docRec) throw new Error('No document record');
  const docAtom = child(doc, docRec, R.DOCUMENT_ATOM);
  const size = docAtom ? { w: i32(doc, docAtom.body), h: i32(doc, docAtom.body + 4) } : { w: 5760, h: 4320 };

  // Fonts and document-level "other" text style.
  const fonts = [];
  const env = child(doc, docRec, R.ENVIRONMENT);
  const docStyles = {};
  if (env) {
    const fc = child(doc, env, R.FONT_COLLECTION);
    if (fc) {
      for (const fe of children(doc, fc, R.FONT_ENTITY)) {
        const name = utf16.decode(doc.subarray(fe.body, fe.body + 64));
        fonts.push(name.replace(/\0.*$/s, ''));
      }
    }
    for (const ts of children(doc, env, R.TX_MASTER_STYLE)) docStyles[ts.inst] = readMasterStyle(doc, ts, DEFAULT_SCHEME);
  }

  // Pictures: BStore entry (1-based pib) -> blob, matched by offset in the Pictures stream.
  const blips = [];
  const dg = child(doc, docRec, R.DRAWING_GROUP);
  const dgg = dg && child(doc, dg, R.DGG);
  const bstore = dgg && child(doc, dgg, R.BSTORE);
  if (bstore) {
    for (const fbse of children(doc, bstore, R.FBSE)) blips.push(pictureBlobs.get(u32(doc, fbse.body + 28)) || null);
  }

  // Slide lists: inst 0 = slides, 1 = masters. Each entry: persist ref + its outline texts.
  const lists = { 0: [], 1: [] };
  for (const list of children(doc, docRec, R.SLIDE_LIST)) {
    const out = lists[list.inst] ?? [];
    let entry = null;
    let header = 4;
    for (const r of records(doc, list.body, list.end)) {
      if (r.type === R.SLIDE_PERSIST) {
        entry = { persist: u32(doc, r.body), slideId: u32(doc, r.body + 12), texts: [] };
        out.push(entry);
      } else if (!entry) continue;
      else if (r.type === R.TEXT_HEADER) { header = u32(doc, r.body); entry.texts.push({ type: header, text: '', style: null }); }
      else if (r.type === R.TEXT_CHARS || r.type === R.TEXT_BYTES) {
        const t = entry.texts.at(-1) || (entry.texts.push({ type: header, text: '', style: null }), entry.texts.at(-1));
        t.text = r.type === R.TEXT_CHARS ? utf16.decode(doc.subarray(r.body, r.end)) : cp1252.decode(doc.subarray(r.body, r.end));
      } else if (r.type === R.STYLE_TEXT_PROP && entry.texts.length) entry.texts.at(-1).style = r;
    }
  }

  const schemeOf = (container) => {
    const cs = children(doc, container, R.COLOR_SCHEME).find((r) => r.inst === 1);
    if (!cs) return null;
    return Array.from({ length: 8 }, (_, i) => hex(doc[cs.body + i * 4], doc[cs.body + i * 4 + 1], doc[cs.body + i * 4 + 2]));
  };

  // Masters, by slide id.
  const masters = new Map();
  for (const entry of lists[1]) {
    const off = persist.get(entry.persist);
    const rec = off !== undefined ? recordAt(off) : null;
    if (!rec || rec.type !== R.MAIN_MASTER) continue;
    const scheme = schemeOf(rec) || DEFAULT_SCHEME;
    const styles = { ...docStyles };
    for (const ts of children(doc, rec, R.TX_MASTER_STYLE)) styles[ts.inst] = readMasterStyle(doc, ts, scheme);
    masters.set(entry.slideId, { rec, scheme, styles, texts: entry.texts });
  }
  const firstMaster = masters.values().next().value;

  // --- shapes of one drawing (slide or master)
  function readDrawing(container, ctx) {
    const shapes = [];
    let background;
    const drawing = child(doc, container, R.PPDRAWING);
    const dgc = drawing && child(doc, drawing, R.DG);
    if (!dgc) return { shapes, background };

    const shapeFrom = (sp, tf) => {
      const fspRec = child(doc, sp, R.FSP);
      if (!fspRec) return null;
      const spt = fspRec.inst;
      const flags = u32(doc, fspRec.body + 4);
      const opt = readOpt(doc, child(doc, sp, R.FOPT));
      if (flags & 0x400) { // background shape
        const filled = boolProp(opt[0x1bf], 0x10, 0x100000);
        if (filled !== false) background = optColor(opt[0x181], ctx.scheme) ?? '#ffffff';
        return null;
      }
      if (flags & 0x5) return null; // group header / patriarch
      const a = readAnchor(doc, sp);
      if (!a) return null;
      let rect = tf({ x: a.l, y: a.t, w: a.r - a.l, h: a.bt - a.t });

      const cd = child(doc, sp, R.CLIENT_DATA);
      const ph = cd && child(doc, cd, R.PLACEHOLDER);
      const placeholder = ph ? doc[ph.body + 4] : 0;
      if (ctx.master && placeholder) return null; // master placeholders are templates ("Click to edit…")

      const rot = opt[0x004] !== undefined ? (opt[0x004] | 0) / 65536 : 0; // 16.16 fixed-point degrees
      const norm = ((rot % 360) + 360) % 360;
      if ((norm >= 45 && norm < 135) || (norm >= 225 && norm < 315)) { // anchor stored rotated
        const cx = rect.x + rect.w / 2; const cy = rect.y + rect.h / 2;
        rect = { x: cx - rect.h / 2, y: cy - rect.w / 2, w: rect.h, h: rect.w };
      }

      const isText = spt === 202 || placeholder;
      const filled = boolProp(opt[0x1bf], 0x10, 0x100000) ?? (isText || spt === 75 ? false : true);
      const lined = boolProp(opt[0x1ff], 0x08, 0x80000) ?? (isText || spt === 75 ? false : true);
      const shape = {
        ...rect, rot, flipH: !!(flags & 0x40), flipV: !!(flags & 0x80),
        kind: spt === 3 ? 'ellipse' : spt === 2 ? 'roundRect' : spt === 20 || spt === 32 || (flags & 0x100) ? 'line' : 'rect',
        fill: filled ? optColor(opt[0x181], ctx.scheme) ?? ctx.scheme[4] : null,
        fillOpacity: opt[0x182] !== undefined ? opt[0x182] / 65536 : 1,
        line: lined ? optColor(opt[0x1c0], ctx.scheme) ?? ctx.scheme[1] : null,
        lineWidth: (opt[0x1cb] ?? 9525) / EMU_PER_MU,
        adjust: opt[0x147],
      };
      const pib = opt[0x104] ?? (opt[0x180] === 3 ? opt[0x186] : undefined);
      if (pib && blips[pib - 1]) { shape.picture = blips[pib - 1]; shape.fill = null; }

      // Text: inline in the shape, or a reference into the slide list's outline texts.
      const tb = child(doc, sp, R.CLIENT_TEXTBOX);
      if (tb) {
        let text = null; let type = 4; let style = null;
        for (const r of records(doc, tb.body, tb.end)) {
          if (r.type === R.TEXT_HEADER) type = u32(doc, r.body);
          else if (r.type === R.TEXT_CHARS) text = utf16.decode(doc.subarray(r.body, r.end));
          else if (r.type === R.TEXT_BYTES) text = cp1252.decode(doc.subarray(r.body, r.end));
          else if (r.type === R.STYLE_TEXT_PROP) style = r;
          else if (r.type === R.OUTLINE_TEXT_REF) {
            const ref = ctx.texts?.[u32(doc, r.body)];
            if (ref) { text = ref.text; type = ref.type; style = ref.style; }
          }
        }
        if (text) {
          shape.text = {
            type, anchor: opt[0x087] ?? (type === 0 || type === 6 ? 1 : 0), wrap: opt[0x085] !== 2,
            insets: [opt[0x082] ?? 45720, opt[0x083] ?? 91440, opt[0x084] ?? 45720, opt[0x081] ?? 91440].map((v) => v / EMU_PER_MU),
            paragraphs: buildParagraphs(text, type, style, ctx),
          };
        }
      }
      if (placeholder && !shape.text && !shape.picture) return null; // empty placeholder
      return shape;
    };

    const walkGroup = (spgr, tf) => {
      const items = records(doc, spgr.body, spgr.end);
      const head = items[0]?.type === R.SP ? items[0] : null;
      let inner = tf;
      if (head) {
        const fspgr = child(doc, head, R.FSPGR);
        const a = readAnchor(doc, head);
        const fsp = child(doc, head, R.FSP);
        const isPatriarch = fsp && (u32(doc, fsp.body + 4) & 0x4);
        if (fspgr && a && !isPatriarch) {
          const c = { x: i32(doc, fspgr.body), y: i32(doc, fspgr.body + 4), w: i32(doc, fspgr.body + 8) - i32(doc, fspgr.body), h: i32(doc, fspgr.body + 12) - i32(doc, fspgr.body + 4) };
          const g = tf({ x: a.l, y: a.t, w: a.r - a.l, h: a.bt - a.t });
          const sx = c.w ? g.w / c.w : 1; const sy = c.h ? g.h / c.h : 1;
          inner = (r) => ({ x: g.x + (r.x - c.x) * sx, y: g.y + (r.y - c.y) * sy, w: r.w * sx, h: r.h * sy });
        }
      }
      for (const it of items) {
        if (it === head) continue;
        if (it.type === R.SPGR) walkGroup(it, inner);
        else if (it.type === R.SP) { const s = shapeFrom(it, inner); if (s) shapes.push(s); }
      }
    };

    for (const it of records(doc, dgc.body, dgc.end)) {
      if (it.type === R.SPGR) walkGroup(it, (r) => r);
      else if (it.type === R.SP) { const s = shapeFrom(it, (r) => r); if (s) shapes.push(s); }
    }
    return { shapes, background };
  }

  function buildParagraphs(text, type, styleRec, ctx) {
    const raw = text.replace(/\r$/, '');
    const runs = styleRec ? readStyleRuns(doc, styleRec, raw.length + 1, ctx.scheme) : { pf: [], cf: [] };
    const at = (list, i) => { let n = 0; for (const r of list) { n += r.count; if (i < n) return r; } return list.at(-1); };
    const paragraphs = [];
    let i = 0;
    for (const ptext of raw.split('\r')) {
      const pr = at(runs.pf, i);
      const level = pr?.level ?? 0;
      const master = styleFor(ctx.styles, type, level);
      const pf = { ...master.pf, ...(pr?.props || {}) };
      const segs = [];
      for (let k = 0; k < ptext.length; k++) {
        const cr = at(runs.cf, i + k);
        const cf = { ...master.cf, ...(cr?.props || {}) };
        const last = segs.at(-1);
        if (last && last.run === cr) last.text += ptext[k];
        else segs.push({ run: cr, text: ptext[k], cf });
      }
      const firstCf = segs[0]?.cf || { ...master.cf, ...(at(runs.cf, i)?.props || {}) };
      paragraphs.push({
        level, pf, empty: !ptext.length, emptyCf: firstCf,
        runs: segs.map(({ text: t, cf }) => ({ text: t, ...cf, fontName: fonts[cf.font] })),
        bulletFontName: fonts[pf.bulletFont],
      });
      i += ptext.length + 1;
    }
    return paragraphs;
  }

  // --- slides in presentation order
  const slides = [];
  for (const entry of lists[0]) {
    const off = persist.get(entry.persist);
    const rec = off !== undefined ? recordAt(off) : null;
    if (!rec || rec.type !== R.SLIDE) continue;
    const sa = child(doc, rec, R.SLIDE_ATOM);
    const masterId = sa ? u32(doc, sa.body + 12) : 0;
    const slideFlags = sa ? u16(doc, sa.body + 20) : 0x7;
    const master = masters.get(masterId) || firstMaster;
    const scheme = (!(slideFlags & 0x2) && schemeOf(rec)) || master?.scheme || DEFAULT_SCHEME;
    const styles = master?.styles || docStyles;

    const own = readDrawing(rec, { scheme, styles, texts: entry.texts });
    let background = own.background;
    let masterShapes = [];
    if (master) {
      const m = readDrawing(master.rec, { scheme, styles, texts: master.texts, master: true });
      if ((slideFlags & 0x4) || !background) background = m.background ?? background;
      if (slideFlags & 0x1) masterShapes = m.shapes;
    }
    slides.push({ background: background || scheme[0], shapes: [...masterShapes, ...own.shapes] });
  }
  if (!slides.length) throw new Error('No slides found');
  return { size, slides };
}
