/**
 * Floating objects in Word 97-2003 documents: shapes, text boxes and pictures drawn with
 * OfficeArt ([MS-ODRAW]) and anchored to a character of the text ([MS-DOC] FSPA).
 */

const u16 = (b, p) => b[p] | (b[p + 1] << 8);
const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
const i32 = (b, p) => b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24);
const hex = (r, g, b) => `#${[r, g, b].map((x) => x.toString(16).padStart(2, '0')).join('')}`;

const BLIP_TYPES = { 0xf01d: 'image/jpeg', 0xf02a: 'image/jpeg', 0xf01e: 'image/png', 0xf01f: 'image/bmp', 0xf029: 'image/tiff' };
const HEADS = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];

/** OfficeArt records in bytes[start, end): { type, inst, ver, body, end }. */
function records(b, start, end) {
  const out = [];
  for (let q = start; q + 8 <= end;) {
    const len = u32(b, q + 4);
    out.push({ ver: b[q] & 0x0f, inst: u16(b, q) >> 4, type: u16(b, q + 2), body: q + 8, end: Math.min(end, q + 8 + len) });
    q += 8 + len;
  }
  return out;
}
const child = (b, rec, type) => records(b, rec.body, rec.end).find((r) => r.type === type);

/**
 * Shape properties (OPT and tertiary OPT): id -> value, plus the data of complex properties
 * (arrays). Arrays may give their size without their 6-byte header; both are accepted.
 */
function readProps(b, recs) {
  const props = {};
  const complex = {};
  for (const rec of recs) {
    if (!rec) continue;
    let cx = rec.body + rec.inst * 6;
    for (let k = 0; k < rec.inst; k++) {
      const at = rec.body + k * 6;
      const id = u16(b, at);
      const v = u32(b, at + 2);
      props[id & 0x3fff] = v;
      if (!(id & 0x8000)) continue;
      let take = v;
      if (cx + 6 <= rec.end) {
        const raw = u16(b, cx + 4);
        const full = 6 + u16(b, cx) * (raw === 0xfff0 ? 4 : raw);
        if (v + 6 === full) take = full;
      }
      complex[id & 0x3fff] = b.subarray(cx, Math.min(rec.end, cx + take));
      cx += take;
    }
  }
  return { props, complex };
}

const color = (v) => {
  if (v === undefined) return undefined;
  if (v >>> 24 & 0x18) return undefined; // scheme / system colors: the default
  return hex(v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff);
};
const flag = (v, bit, useBit) => (v === undefined || !(v & useBit) ? undefined : !!(v & bit));

/** An array property: its elements as byte offsets into `data`, and the element size. */
function arrayOf(data) {
  if (!data || data.length < 6) return { n: 0, size: 0 };
  const raw = u16(data, 4);
  const size = raw === 0xfff0 ? 4 : raw;
  return { n: Math.min(u16(data, 0), Math.floor((data.length - 6) / (size || 1))), size };
}

/**
 * Custom geometry (vertices, path segments and guide formulas) as an SVG path in the shape's
 * own coordinates, or null. Arcs written as "line to" commands followed by a radius and a
 * start / sweep angle pair (as LibreOffice saves them) are drawn as arcs.
 */
function customPath(props, complex) {
  const verts = complex[0x145];
  if (!verts) return null;
  const adjust = (k) => { const v = props[0x147 + k]; return v === undefined ? 0 : v | 0; };

  // guides: [MS-ODRAW] MSOSG formulas, evaluated in order (a guide may use earlier ones)
  const guides = [];
  const constantGuide = [];
  const g = arrayOf(complex[0x156]);
  const gd = complex[0x156];
  const ref = (v) => {
    if (v >= 0x400 && v < 0x400 + guides.length) return guides[v - 0x400];
    if (v >= 0x147 && v <= 0x150) return adjust(v - 0x147);
    if (v >= 0x140 && v <= 0x143) return props[v] ?? (v >= 0x142 ? 21600 : 0);
    return 0;
  };
  const bigAdjust = [0, 1, 2, 3].some((k) => adjust(k) > 0x7fff);
  const deg = (v) => ((v / 65536) * Math.PI) / 180; // 16.16 fixed-point degrees
  for (let k = 0; k < g.n; k++) {
    const o = 6 + k * g.size;
    const sgf = u16(gd, o);
    const op = sgf & 0x1fff;
    const [a, b2, c] = [0, 1, 2].map((i) => {
      const raw = u16(gd, o + 2 + i * 2);
      if (sgf & (0x2000 << i)) return ref(raw);
      // 100% in preset formulas (100000) does not fit in 16 bits: LibreOffice writes 32767
      return raw === 0x7fff && bigAdjust ? 100000 : raw;
    });
    let r;
    switch (op) {
      case 0: r = a + b2 - c; break;
      case 1: r = c ? (a * b2) / c : 0; break;
      case 2: r = (a + b2) / 2; break;
      case 3: r = Math.abs(a); break;
      case 4: r = Math.min(a, b2); break;
      case 5: r = Math.max(a, b2); break;
      case 6: r = a > 0 ? b2 : c; break;
      case 7: r = Math.sqrt(a * a + b2 * b2 + c * c); break;
      case 8: r = (Math.atan2(b2, a) * 180 / Math.PI) * 65536; break;
      case 9: r = a * Math.sin(deg(b2)); break;
      case 10: r = a * Math.cos(deg(b2)); break;
      case 11: r = a * Math.cos(Math.atan2(c, b2)); break;
      case 12: r = a * Math.sin(Math.atan2(c, b2)); break;
      case 13: r = Math.sqrt(Math.max(0, a)); break;
      case 14: r = a + b2 * 65536 - c * 65536; break;
      case 15: r = b2 ? c * Math.sqrt(Math.max(0, 1 - (a / b2) ** 2)) : 0; break;
      case 16: r = a * Math.tan(deg(b2)); break;
      default: r = 0;
    }
    guides.push(r);
    constantGuide.push(op === 0 && !(sgf & 0xe000));
  }

  // vertices: 32-bit pairs (0x80000000 + n = guide n) or 16-bit pairs
  const va = arrayOf(verts);
  const pts = [];
  const isConst = [];
  for (let k = 0; k < va.n; k++) {
    const o = 6 + k * va.size;
    if (va.size >= 8) {
      const xy = [i32(verts, o), i32(verts, o + 4)].map((v) => (v & 0x80000000 && (v & 0x7fffffff) < guides.length ? { g: v & 0x7fffffff } : { v }));
      pts.push(xy.map((e) => (e.g !== undefined ? guides[e.g] : e.v)));
      isConst.push(xy.every((e) => e.g !== undefined && constantGuide[e.g]));
    } else {
      pts.push([(u16(verts, o) << 16) >> 16, (u16(verts, o + 2) << 16) >> 16]);
      isConst.push(false);
    }
  }

  // segments
  const sa = arrayOf(complex[0x146]);
  const segs = sa.n ? Array.from({ length: sa.n }, (_, k) => u16(complex[0x146], 6 + k * 2)) : [0x4000, 0x0000 | (pts.length - 1), 0x6001, 0x8000];
  const used = segs.reduce((n, s) => n + ({ 0: s & 0x1fff, 1: 3 * (s & 0x1fff), 2: 1 }[s >> 13] ?? (s >> 13 === 5 ? s & 0xff : 0)), 0);
  let arcsLeft = Math.max(0, pts.length - used); // vertices beyond the plain reading: one per arc
  const n = (v) => +v.toFixed(2);
  let d = '';
  let i = 0;
  let cur = [0, 0];
  let start = [0, 0];
  for (const s of segs) {
    const type = s >> 13;
    const count = s & 0x1fff;
    if (type === 2 && pts[i]) { cur = start = pts[i++]; d += `M${n(cur[0])} ${n(cur[1])}`; }
    else if (type === 0) {
      for (let c = 0; c < count && pts[i]; c++) {
        if (arcsLeft > 0 && isConst[i + 1] && pts[i + 1]) {
          const [wR, hR] = pts[i]; const [st, sw] = pts[i + 1];
          i += 2; arcsLeft--;
          const a0 = (st * Math.PI) / 180; const a1 = ((st + sw) * Math.PI) / 180;
          const cx = cur[0] - wR * Math.cos(a0); const cy = cur[1] - hR * Math.sin(a0);
          cur = [cx + wR * Math.cos(a1), cy + hR * Math.sin(a1)];
          d += `A${n(wR)} ${n(hR)} 0 ${Math.abs(sw) > 180 ? 1 : 0} ${sw > 0 ? 1 : 0} ${n(cur[0])} ${n(cur[1])}`;
        } else { cur = pts[i++]; d += `L${n(cur[0])} ${n(cur[1])}`; }
      }
    } else if (type === 1) {
      for (let c = 0; c < count && pts[i + 2]; c++) {
        d += `C${[pts[i], pts[i + 1], pts[i + 2]].map((p) => `${n(p[0])} ${n(p[1])}`).join(' ')}`;
        cur = pts[i + 2]; i += 3;
      }
    } else if (type === 3) { d += 'Z'; cur = start; }
    else if (type === 5) i += s & 0xff; // escapes (fill / line switches, other arcs): skipped
  }
  return d || null;
}

/** A blip (picture) record at `at`: its bytes and type. */
function readBlip(b, at, end) {
  if (at + 8 > b.length) return null;
  const inst = u16(b, at) >> 4;
  const type = u16(b, at + 2);
  const len = u32(b, at + 4);
  if (!BLIP_TYPES[type]) return null;
  const uids = [0x46b, 0x6e1, 0x7a9, 0x6e3, 0x6e5].includes(inst) ? 32 : 16;
  return { mime: BLIP_TYPES[type], bytes: b.subarray(at + 8 + uids + 1, Math.min(end ?? b.length, at + 8 + len)) };
}

/**
 * The document's floating objects: per object its anchor character, its box in twips
 * (relative to what `relH` / `relV` name), how text wraps around it and what to draw.
 * `storyText(n)` returns the paragraphs of text box story n (1-based).
 */
export function readDrawings(wd, table, fcOf, lcbOf, storyText) {
  // FSPA anchors in the main document: CPs (n + 1) then 26-byte FSPA entries
  const fcSpa = fcOf(40); const lcbSpa = lcbOf(40);
  if (!lcbSpa || !lcbOf(50)) return [];
  const count = (lcbSpa - 4) / 30;
  const anchors = [];
  for (let k = 0; k < count; k++) {
    const q = fcSpa + (count + 1) * 4 + k * 26;
    const f = u16(table, q + 20);
    anchors.push({
      cp: i32(table, fcSpa + k * 4), spid: u32(table, q),
      rect: { x: i32(table, q + 4), y: i32(table, q + 8), w: i32(table, q + 12) - i32(table, q + 4), h: i32(table, q + 16) - i32(table, q + 8) },
      relH: (f >> 1) & 3, relV: (f >> 3) & 3, wrap: (f >> 5) & 0xf, below: !!(f & 0x4000),
    });
  }

  // OfficeArtContent: the drawing group (pictures), then per drawing a dgglbl byte and a DgContainer
  let p = fcOf(50);
  const end = p + lcbOf(50);
  const dgg = records(table, p, Math.min(end, p + 8 + u32(table, p + 4)))[0];
  const blips = [];
  const bstore = dgg && child(table, dgg, 0xf001);
  for (const fbse of bstore ? records(table, bstore.body, bstore.end).filter((r) => r.type === 0xf007) : []) {
    const inner = fbse.body + 36 + table[fbse.body + 33];
    blips.push(inner + 8 <= fbse.end ? readBlip(table, inner, fbse.end) : readBlip(wd, u32(table, fbse.body + 28)));
  }
  p += 8 + u32(table, p + 4);
  const shapes = new Map();
  while (p + 9 <= end) {
    const label = table[p++]; // 0 = main document, 1 = headers
    const len = u32(table, p + 4);
    const visit = (s, e) => {
      for (const r of records(table, s, e)) {
        if (r.type === 0xf004) {
          const fsp = child(table, r, 0xf00a);
          if (!fsp) continue;
          const kids = records(table, r.body, r.end);
          const { props, complex } = readProps(table, [kids.find((x) => x.type === 0xf00b), kids.find((x) => x.type === 0xf122)]);
          shapes.set(u32(table, fsp.body), { spt: fsp.inst, flags: u32(table, fsp.body + 4), props, complex });
        } else if (r.ver === 0x0f) visit(r.body, r.end);
      }
    };
    if (label === 0) visit(p, p + 8 + len);
    p += 8 + len;
  }

  const out = [];
  for (const a of anchors) {
    const s = shapes.get(a.spid);
    if (!s) continue;
    const { props, complex, spt, flags } = s;
    const isText = spt === 202;
    const filled = flag(props[0x1bf], 0x10, 0x100000) ?? !isText;
    const lined = flag(props[0x1ff], 0x08, 0x80000) ?? !isText;
    const shape = {
      spt,
      kind: spt === 3 ? 'ellipse' : spt === 2 ? 'roundRect' : spt === 20 || spt === 32 ? 'line' : spt === 1 || isText || spt === 75 ? 'rect' : 'custom',
      fill: filled ? color(props[0x181]) ?? '#ffffff' : null,
      fillOpacity: props[0x182] !== undefined ? props[0x182] / 65536 : 1,
      line: lined ? color(props[0x1c0]) ?? '#000000' : null,
      lineWidth: (props[0x1cb] ?? 9525) / 12700, // EMU -> pt
      flipH: !!(flags & 0x40), flipV: !!(flags & 0x80),
      rot: props[0x004] !== undefined ? (props[0x004] | 0) / 65536 : 0,
      adjust: props[0x147],
      anchorText: props[0x087] ?? 0,
      insets: [0x082, 0x083, 0x084, 0x081].map((k) => (props[k] ?? (k === 0x081 || k === 0x083 ? 91440 : 45720)) / 9525), // top right bottom left, px
    };
    if (props[0x1d0] || props[0x1d1]) shape.arrows = { start: HEADS[props[0x1d0] || 0], end: HEADS[props[0x1d1] || 0] };
    if (shape.kind === 'custom') {
      shape.path = customPath(props, complex);
      if (!shape.path) shape.kind = 'rect';
    }
    const pib = props[0x104];
    if (pib && blips[pib - 1]) { shape.picture = blips[pib - 1]; shape.fill = null; shape.line = lined && props[0x1ff] !== undefined ? shape.line : null; }
    const txid = props[0x080];
    if (txid) shape.text = storyText(txid >>> 16);
    out.push({ ...a, shape });
  }

  // LibreOffice saves a shape's text as a separate, unfilled text box positioned on the page:
  // move that text into the shape it belongs to (the next one, horizontally inside it).
  for (let k = 0; k < out.length; k++) {
    const o = out[k];
    const next = out[k + 1];
    const empty = !o.shape.text || !o.shape.text.some((pp) => pp.segments.length);
    if (!empty || !next || next.shape.spt !== 202 || !next.shape.text) continue;
    if (next.rect.x >= o.rect.x - 60 && next.rect.x + next.rect.w <= o.rect.x + o.rect.w + 60) {
      o.shape.text = next.shape.text;
      out.splice(k + 1, 1);
    }
  }
  return out;
}
