/*
 * Word 97-2003 (.doc) formatting reader, following Microsoft's published
 * [MS-DOC] specification. Reads, besides the text:
 *   - character formatting (CHPX): bold, italic, underline, strike, color,
 *     highlight, size, font, super/subscript, caps, hidden text
 *   - paragraph formatting (PAPX): alignment, indents, spacing, line height
 *   - styles (STSH): Normal, Heading 1-9 ... with their inheritance
 *   - lists (PlfLst / PlfLfo): bullets and numbering with level indents
 *   - the font table (SttbfFfn)
 *   - sections: page size, margins, orientation and columns (PlcfSed)
 *   - tables (TDefTable): column widths, merged cells, borders, cell shading
 *   - fields: hyperlinks, page numbers; headers and footers (PlcfHdd)
 *   - inline pictures (PICF in the Data stream) and embedded equations (MTEF)
 * Floating shapes, text boxes and pictures are read in doc-drawing.js. Line breaking follows
 * the browser, so a page may end a line earlier or later than in Word.
 */

import { bulletCode } from '../bullets.js';

/** Page-number fields in headers / footers, replaced with the page's numbers when laid out. */
export const PAGE_MARK = '\u2063dvPAGE\u2063';
export const PAGES_MARK = '\u2063dvPAGES\u2063';
const OBJ = '\ufffc'; // an inline picture or object (equation) in paragraph text

import { readDrawings } from './doc-drawing.js';
import { readMtef } from './mtef.js';

const cp1252 = new TextDecoder('windows-1252');
const utf16 = new TextDecoder('utf-16le');

// FibRgFcLcb97 starts at byte 154 of the FIB; each entry is an (fc, lcb) pair.
const FIB_RGFCLCB = 154;
const FC = { stshf: 1, plcfSed: 6, plcfBteChpx: 12, plcfBtePapx: 13, sttbfFfn: 15, clx: 33, plfLst: 73, plfLfo: 74 };

// Word's 16 "ico" colors.
const ICO = [null, '#000000', '#0000ff', '#00ffff', '#00ff00', '#ff00ff', '#ff0000', '#ffff00', '#ffffff',
  '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0'];

const u16 = (b, p) => b[p] | (b[p + 1] << 8);
const i16 = (b, p) => (u16(b, p) << 16) >> 16;
const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
const i32 = (b, p) => b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24);

/** Iterates the SPRMs (property modifiers) in bytes[start, end). */
function forEachSprm(b, start, end, fn) {
  let p = start;
  while (p + 2 <= end) {
    const op = u16(b, p);
    let q = p + 2;
    let size;
    switch (op >> 13) {
      case 0: case 1: size = 1; break;
      case 2: case 4: case 5: size = 2; break;
      case 3: size = 4; break;
      case 7: size = 3; break;
      default: // variable length
        if (op === 0xd608 || op === 0xd606) { size = u16(b, q) - 1; q += 2; } // table definitions
        else { size = b[q]; q += 1; }
    }
    if (q + size > end) break;
    fn(op, q, size);
    p = q + size;
  }
}

// --------------------------------------------------------------- properties

const toggle = (v, base) => (v === 0 ? false : v === 1 ? true : v === 0x80 ? !!base : v === 0x81 ? !base : !!v);

function applyChp(chp, b, start, end, base = {}) {
  forEachSprm(b, start, end, (op, at) => {
    switch (op) {
      case 0x0835: chp.bold = toggle(b[at], base.bold); break;
      case 0x0836: chp.italic = toggle(b[at], base.italic); break;
      case 0x0837: chp.strike = toggle(b[at], base.strike); break;
      case 0x2a53: chp.strike = toggle(b[at], base.strike); break; // double strike
      case 0x083a: chp.smallCaps = toggle(b[at], base.smallCaps); break;
      case 0x083b: chp.caps = toggle(b[at], base.caps); break;
      case 0x083c: chp.hidden = toggle(b[at], base.hidden); break;
      case 0x2a3e: chp.underline = b[at]; break;                           // kul: 1 single, 3 double, 0 none
      case 0x4a43: chp.size = u16(b, at) / 2; break;                      // half-points -> pt
      case 0x2a42: chp.color = ICO[b[at]] ?? chp.color; break;
      case 0x6870: chp.color = b[at + 3] ? null                            // fAuto
        : `#${[b[at], b[at + 1], b[at + 2]].map((x) => x.toString(16).padStart(2, '0')).join('')}`; break;
      case 0x2a0c: chp.highlight = ICO[b[at]] ?? null; break;
      case 0x4a4f: chp.font = u16(b, at); break;                            // ASCII font index
      case 0x2a48: chp.vert = b[at]; break;                                 // 1 super, 2 sub
      case 0x6a03: chp.picLocation = u32(b, at); break;                    // picture: PICF offset in the Data stream
      case 0x0855: chp.special = b[at] !== 0; break;                        // fSpec: the character is an object
      case 0x080a: chp.ole = b[at] !== 0; break;                            // fOle2: an embedded object (equation…)
      default:
    }
  });
  return chp;
}

const hex3 = (b, q) => `#${[b[q], b[q + 1], b[q + 2]].map((x) => x.toString(16).padStart(2, '0')).join('')}`;
/** Shd80 (2 bytes): pattern 1 = solid foreground, 0 = clear (background color). */
const shd80 = (v) => { const ipat = v >> 10; return ipat === 1 ? ICO[v & 0x1f] : ipat === 0 ? ICO[(v >> 5) & 0x1f] : null; };
/** Shd (10 bytes): foreground and background COLORREF (byte 4 = 0xff means automatic) and pattern. */
const shdRgb = (b, q) => {
  const ipat = u16(b, q + 8);
  const fore = b[q + 3] === 0xff ? null : hex3(b, q);
  const back = b[q + 7] === 0xff ? null : hex3(b, q + 4);
  return ipat === 1 ? fore : ipat === 0 ? back : back || fore;
};
/** Brc80 (4 bytes): width in eighths of a point, line type, ico color. */
const brc80 = (b, q) => (b[q + 1] && b[q + 1] !== 0xff ? { width: b[q] / 8, type: b[q + 1], color: ICO[b[q + 2]] || '#000000' } : null);

/**
 * sprmTDefTable: the row's column edges (twips) and, per cell, its merge flags and borders
 * ([MS-DOC] TDefTableOperand / TC80).
 */
function readTableDef(b, at) {
  const n = b[at];
  const centers = Array.from({ length: n + 1 }, (_, i) => i16(b, at + 1 + i * 2));
  const tc0 = at + 1 + (n + 1) * 2;
  const cells = Array.from({ length: n }, (_, i) => {
    const q = tc0 + i * 20;
    const grf = u16(b, q);
    return {
      width: centers[i + 1] - centers[i],
      merged: !!(grf & 0x2), vertMerge: !!(grf & 0x20), vertRestart: !!(grf & 0x40), vAlign: (grf >> 7) & 3,
      borders: { top: brc80(b, q + 4), left: brc80(b, q + 8), bottom: brc80(b, q + 12), right: brc80(b, q + 16) },
    };
  });
  return { centers, cells };
}

function applyPap(pap, b, start, end) {
  forEachSprm(b, start, end, (op, at) => {
    switch (op) {
      case 0x2403: case 0x2461: pap.align = b[at]; break;                   // 0 left 1 center 2 right 3 justify
      case 0x840f: case 0x845e: pap.left = i16(b, at); break;               // twips
      case 0x840e: case 0x845d: pap.right = i16(b, at); break;
      case 0x8411: case 0x8460: pap.first = i16(b, at); break;
      case 0xa413: pap.before = u16(b, at); break;
      case 0xa414: pap.after = u16(b, at); break;
      case 0x6412: pap.line = { value: i16(b, at), mult: i16(b, at + 2) }; break;
      case 0x460b: pap.ilfo = i16(b, at); break;
      case 0x260a: pap.ilvl = b[at]; break;
      case 0x2640: pap.outline = b[at]; break;
      case 0x2407: pap.pageBreakBefore = b[at] !== 0; break;
      case 0x246d: pap.contextual = b[at] !== 0; break;                    // no space between same-style paragraphs
      case 0x442d: {                                                        // Shd80: paragraph shading (ico colors)
        const v = u16(b, at); const ipat = v >> 10;
        pap.shading = ipat === 1 ? ICO[v & 0x1f] : ICO[(v >> 5) & 0x1f] || null;
        break;
      }
      case 0xd608: pap.tap = { ...pap.tap, ...readTableDef(b, at) }; break; // table row: column edges and cells
      case 0xd609: {                                                        // cell shading, Word 97 colors
        const n = b[at - 1] / 2;
        pap.tap = { ...pap.tap, shd80: Array.from({ length: n }, (_, i) => shd80(u16(b, at + i * 2))) };
        break;
      }
      case 0xd612: {                                                        // cell shading, RGB (cells 0-21)
        const n = Math.floor(b[at - 1] / 10);
        pap.tap = { ...pap.tap, shd: Array.from({ length: n }, (_, i) => shdRgb(b, at + i * 10)) };
        break;
      }
      case 0xc64d: {                                                        // Shd: paragraph shading (RGB)
        const rgb = (q) => (b[q + 3] === 0xff ? null : `#${[b[q], b[q + 1], b[q + 2]].map((x) => x.toString(16).padStart(2, '0')).join('')}`);
        const ipat = u16(b, at + 8);
        pap.shading = ipat === 1 ? rgb(at) : rgb(at + 4);
        break;
      }
      case 0x6424: case 0x6425: case 0x6426: case 0x6427: {                 // Brc80 top/left/bottom/right
        const side = ['top', 'left', 'bottom', 'right'][op - 0x6424];
        pap.borders = { ...pap.borders, [side]: b[at + 1] && b[at + 1] !== 0xff ? { width: b[at] / 8, color: ICO[b[at + 2]] || '#000000' } : null };
        break;
      }
      case 0xc64e: case 0xc64f: case 0xc650: case 0xc651: {                 // Brc top/left/bottom/right
        const side = ['top', 'left', 'bottom', 'right'][op - 0xc64e];
        const auto = b[at + 3] === 0xff;
        const color = auto ? '#000000' : `#${[b[at], b[at + 1], b[at + 2]].map((x) => x.toString(16).padStart(2, '0')).join('')}`;
        pap.borders = { ...pap.borders, [side]: b[at + 5] && b[at + 5] !== 0xff ? { width: b[at + 4] / 8, color } : null };
        break;
      }
      default:
    }
  });
  return pap;
}

// ------------------------------------------------------------------- styles

function readStyles(t, fc, lcb) {
  const styles = [];
  if (!lcb) return { styles, defaultFont: 0 };
  const cbStshi = u16(t, fc);
  const stshi = fc + 2;
  const cstd = u16(t, stshi);
  const cbBase = u16(t, stshi + 2);
  const defaultFont = u16(t, stshi + 12); // rgftcStandardChpStsh[0]
  let p = stshi + cbStshi;
  const raw = [];
  for (let i = 0; i < cstd && p + 2 <= fc + lcb; i++) {
    const cb = u16(t, p);
    raw.push(cb ? { at: p + 2, cb } : null);
    p += 2 + cb;
  }
  for (const r of raw) {
    if (!r) { styles.push(null); continue; }
    const { at, cb } = r;
    const sti = u16(t, at) & 0x0fff;
    const stk = u16(t, at + 2) & 0x000f;
    const base = u16(t, at + 2) >> 4;
    const cupx = u16(t, at + 4) & 0x000f;
    let q = at + cbBase;
    const cch = u16(t, q);
    const name = utf16.decode(t.subarray(q + 2, q + 2 + cch * 2));
    q += 2 + cch * 2 + 2;
    const upx = [];
    for (let k = 0; k < cupx && q + 2 <= at + cb; k++) {
      const len = u16(t, q);
      upx.push({ at: q + 2, len });
      q += 2 + len + (len & 1);
    }
    styles.push({ sti, stk, base: base === 0x0fff ? -1 : base, name, upx });
  }
  return { styles, defaultFont };
}

/** Effective paragraph and character properties of style `istd`, with inheritance. */
function makeStyleResolver(t, styles) {
  const cache = new Map();
  const resolve = (istd, depth = 0) => {
    if (cache.has(istd)) return cache.get(istd);
    const s = styles[istd];
    let out = { pap: {}, chp: {}, sti: s?.sti ?? 0 };
    if (s && depth < 20) {
      const parent = s.base >= 0 && s.base !== istd ? resolve(s.base, depth + 1) : { pap: {}, chp: {} };
      const pap = { ...parent.pap };
      const chp = { ...parent.chp };
      if (s.stk === 1) { // paragraph style: PAPX (istd + sprms), then CHPX
        if (s.upx[0] && s.upx[0].len >= 2) applyPap(pap, t, s.upx[0].at + 2, s.upx[0].at + s.upx[0].len);
        if (s.upx[1]) applyChp(chp, t, s.upx[1].at, s.upx[1].at + s.upx[1].len, parent.chp);
      } else if (s.stk === 2 && s.upx[0]) {
        applyChp(chp, t, s.upx[0].at, s.upx[0].at + s.upx[0].len, parent.chp);
      }
      out = { pap, chp, sti: s.sti };
    }
    cache.set(istd, out);
    return out;
  };
  return resolve;
}

// -------------------------------------------------------------------- fonts

function readFonts(t, fc, lcb) {
  const fonts = [];
  if (!lcb) return fonts;
  let p = fc;
  const extended = u16(t, p) === 0xffff;
  if (extended) p += 2;
  const count = u16(t, p);
  p += 4; // cData + cbExtra
  for (let i = 0; i < count && p < fc + lcb; i++) {
    const cb = t[p];
    const start = p + 1 + 39; // FFN fixed part
    let end = start;
    while (end + 1 < p + 1 + cb && u16(t, end) !== 0) end += 2;
    fonts.push(utf16.decode(t.subarray(start, end)));
    p += 1 + cb;
  }
  return fonts;
}

// -------------------------------------------------------------------- lists

function readLists(t, fcLst, lcbLst, fcLfo, lcbLfo) {
  const lists = new Map(); // lsid -> levels[]
  const lfos = [];         // ilfo - 1 -> lsid
  if (lcbLst) {
    const n = i16(t, fcLst);
    const lstf = [];
    for (let i = 0; i < n; i++) {
      const at = fcLst + 2 + i * 28;
      lstf.push({ lsid: i32(t, at), simple: t[at + 26] & 1 });
    }
    let p = fcLst + 2 + n * 28;
    for (const l of lstf) {
      const levels = [];
      for (let k = 0; k < (l.simple ? 1 : 9); k++) {
        const start = i32(t, p);
        const nfc = t[p + 4];
        const cbChpx = t[p + 24];
        const cbPapx = t[p + 25];
        const papAt = p + 28;
        const chpAt = papAt + cbPapx;
        const pap = applyPap({}, t, papAt, chpAt);
        const chp = applyChp({}, t, chpAt, chpAt + cbChpx);
        const xstAt = chpAt + cbChpx;
        const cch = u16(t, xstAt);
        const text = Array.from({ length: cch }, (_, j) => u16(t, xstAt + 2 + j * 2));
        levels.push({ start, nfc, pap, chp, text });
        p = xstAt + 2 + cch * 2;
      }
      lists.set(l.lsid, levels);
    }
  }
  if (lcbLfo) {
    const n = i32(t, fcLfo);
    for (let i = 0; i < n; i++) lfos.push(i32(t, fcLfo + 4 + i * 16));
  }
  return { lists, lfos };
}

const ROMAN = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']];
function formatNumber(n, nfc) {
  switch (nfc) {
    case 1: case 2: {
      let s = ''; let v = n;
      for (const [k, r] of ROMAN) while (v >= k) { s += r; v -= k; }
      return nfc === 1 ? s.toUpperCase() : s;
    }
    case 3: case 4: {
      let s = ''; let v = n;
      while (v > 0) { v--; s = String.fromCharCode(97 + (v % 26)) + s; v = Math.floor(v / 26); }
      return nfc === 3 ? s.toUpperCase() : s;
    }
    case 22: return String(n).padStart(2, '0');
    default: return String(n);
  }
}


// ----------------------------------------------------------------- FKP runs

/** Runs [fcStart, fcEnd) -> property bytes, from a CHPX or PAPX bin table. */
function readRuns(wd, t, fc, lcb, kind) {
  const runs = [];
  if (!lcb) return runs;
  const n = (lcb - 4) / 8;
  for (let i = 0; i < n; i++) {
    const pn = u32(t, fc + (n + 1) * 4 + i * 4) & 0x3fffff;
    const page = pn * 512;
    if (page + 512 > wd.length) continue;
    const count = wd[page + 511];
    for (let k = 0; k < count; k++) {
      const start = u32(wd, page + k * 4);
      const end = u32(wd, page + (k + 1) * 4);
      if (kind === 'chp') {
        const off = wd[page + (count + 1) * 4 + k] * 2;
        runs.push({ start, end, at: off ? page + off + 1 : 0, len: off ? wd[page + off] : 0 });
      } else {
        const off = wd[page + (count + 1) * 4 + k * 13] * 2;
        if (!off) { runs.push({ start, end, at: 0, len: 0 }); continue; }
        let cb = wd[page + off];
        let at = page + off + 1;
        let len;
        if (cb) len = cb * 2 - 1;
        else { cb = wd[at]; at += 1; len = cb * 2; }
        runs.push({ start, end, at, len });
      }
    }
  }
  runs.sort((a, b) => a.start - b.start);
  return runs;
}

function findRun(runs, fc) {
  let lo = 0; let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (fc < runs[mid].start) hi = mid - 1;
    else if (fc >= runs[mid].end) lo = mid + 1;
    else return mid;
  }
  return -1;
}

// ------------------------------------------------------------------ section

/**
 * Every section of the document: where it ends (CP), page size, margins, orientation and
 * columns (twips). Word 97 defaults where a value is not stored.
 */
export function readSections(wd, table) {
  const fc = u32(wd, FIB_RGFCLCB + FC.plcfSed * 8);
  const lcb = u32(wd, FIB_RGFCLCB + FC.plcfSed * 8 + 4);
  const n = lcb ? (lcb - 4) / 16 : 0;
  const out = [];
  let prev = { width: 12240, height: 15840, left: 1800, right: 1800, top: 1440, bottom: 1440, cols: 1, colSpace: 720, landscape: false, headerTop: 720, footerBottom: 720 };
  for (let i = 0; i < n; i++) {
    // columns start from one in each section; the page setup carries over unless stored
    const sec = { ...prev, cols: 1, colSpace: 720, cpEnd: i32(table, fc + (i + 1) * 4) };
    const fcSepx = u32(table, fc + (n + 1) * 4 + i * 12 + 2);
    if (fcSepx !== 0xffffffff && fcSepx + 2 <= wd.length) {
      const cb = u16(wd, fcSepx);
      forEachSprm(wd, fcSepx + 2, Math.min(wd.length, fcSepx + 2 + cb), (op, at) => {
        switch (op) {
          case 0xb01f: sec.width = u16(wd, at); break;
          case 0xb020: sec.height = u16(wd, at); break;
          case 0xb021: sec.left = u16(wd, at); break;
          case 0xb022: sec.right = u16(wd, at); break;
          case 0x9023: sec.top = Math.abs(i16(wd, at)); break;
          case 0x9024: sec.bottom = Math.abs(i16(wd, at)); break;
          case 0x301d: sec.landscape = wd[at] === 2; break;     // bOrientation
          case 0x500b: sec.cols = u16(wd, at) + 1; break;       // ccolM1
          case 0x900c: sec.colSpace = u16(wd, at); break;       // dxaColumns
          case 0xb017: sec.headerTop = u16(wd, at); break;      // dyaHdrTop
          case 0xb018: sec.footerBottom = u16(wd, at); break;   // dyaHdrBottom
          default:
        }
      });
    }
    out.push(sec);
    prev = sec;
  }
  return out;
}

/** Page size and margins (twips) of the first section; Word 97 defaults when not stored. */
export function readPageSetup(wd, table) {
  const page = { width: 12240, height: 15840, left: 1800, right: 1800, top: 1440, bottom: 1440 };
  const fc = u32(wd, FIB_RGFCLCB + FC.plcfSed * 8);
  const lcb = u32(wd, FIB_RGFCLCB + FC.plcfSed * 8 + 4);
  if (!lcb) return page;
  const n = (lcb - 4) / 16;
  if (n < 1) return page;
  const fcSepx = u32(table, fc + (n + 1) * 4 + 2);
  if (fcSepx === 0xffffffff || fcSepx + 2 > wd.length) return page;
  const cb = u16(wd, fcSepx);
  forEachSprm(wd, fcSepx + 2, Math.min(wd.length, fcSepx + 2 + cb), (op, at) => {
    switch (op) {
      case 0xb01f: page.width = u16(wd, at); break;
      case 0xb020: page.height = u16(wd, at); break;
      case 0xb021: page.left = u16(wd, at); break;
      case 0xb022: page.right = u16(wd, at); break;
      case 0x9023: page.top = Math.abs(i16(wd, at)); break;
      case 0x9024: page.bottom = Math.abs(i16(wd, at)); break;
      default:
    }
  });
  return page;
}

// --------------------------------------------------------------- document

/**
 * Reads a Word 97-2003 document into paragraphs:
 *   [{ end: '\r' | '\x07', pap, segments: [{ text, chp }] }]
 * `pap` and `chp` are fully resolved (style + direct formatting).
 */
const BLIP_TYPES = { 0xf01d: 'image/jpeg', 0xf02a: 'image/jpeg', 0xf01e: 'image/png', 0xf01f: 'image/bmp', 0xf029: 'image/tiff' };

/**
 * An inline picture: the PICF at `at` in the Data stream gives the display size (twips, scaled
 * by mx / my per mille) and is followed by OfficeArt records holding the picture's blip.
 */
export function readInlinePicture(data, at) {
  if (!data || at + 68 > data.length) return null;
  const lcb = u32(data, at);
  const cbHeader = u16(data, at + 4);
  const end = Math.min(data.length, at + lcb);
  const mm = u16(data, at + 6);
  const width = (u16(data, at + 28) * u16(data, at + 32)) / 1000;
  const height = (u16(data, at + 30) * u16(data, at + 34)) / 1000;
  let p = at + cbHeader;
  if (mm === 0x66) p += 1 + data[p]; // picture name
  // walk records (descending into containers) to the first bitmap blip
  const find = (from, to) => {
    for (let q = from; q + 8 <= to;) {
      const ver = data[q] & 0x0f;
      const inst = u16(data, q) >> 4;
      const type = u16(data, q + 2);
      const len = u32(data, q + 4);
      const body = q + 8;
      if (ver === 0x0f) { const hit = find(body, Math.min(to, body + len)); if (hit) return hit; }
      else if (type === 0xf007) { const hit = find(body + 36, Math.min(to, body + len)); if (hit) return hit; } // FBSE then its blip
      else if (BLIP_TYPES[type]) {
        const uids = [0x46b, 0x6e1, 0x7a9, 0x6e3, 0x6e5].includes(inst) ? 32 : 16; // second UID present
        const start = body + uids + 1;
        return { mime: BLIP_TYPES[type], bytes: data.subarray(start, Math.min(to, body + len)) };
      }
      q = body + len;
    }
    return null;
  };
  const blip = find(p, end);
  if (!blip || !width || !height) return null;
  return { ...blip, width, height };
}

export function readWordDocument(wd, table, data = null, objectData = null) {
  const fcOf = (i) => u32(wd, FIB_RGFCLCB + i * 8);
  const lcbOf = (i) => u32(wd, FIB_RGFCLCB + i * 8 + 4);
  const ccpText = i32(wd, 76);

  // Piece table: CP ranges -> byte offsets (8-bit or UTF-16 text).
  const fcClx = fcOf(FC.clx);
  const lcbClx = lcbOf(FC.clx);
  let p = fcClx;
  while (p < fcClx + lcbClx && table[p] === 0x01) p += 3 + u16(table, p + 1);
  if (table[p] !== 0x02) throw new Error('This Word document is damaged.');
  const plc = p + 5;
  const nPieces = (u32(table, p + 1) - 4) / 12;
  const pieces = [];
  for (let i = 0; i < nPieces; i++) {
    const cpStart = i32(table, plc + i * 4);
    const cpEnd = i32(table, plc + (i + 1) * 4);
    const raw = u32(table, plc + (nPieces + 1) * 4 + i * 8 + 2);
    const compressed = !!(raw & 0x40000000);
    pieces.push({ cpStart, cpEnd, compressed, fc: compressed ? (raw & 0x3fffffff) / 2 : raw });
  }

  const { styles, defaultFont } = readStyles(table, fcOf(FC.stshf), lcbOf(FC.stshf));
  const style = makeStyleResolver(table, styles);
  const fonts = readFonts(table, fcOf(FC.sttbfFfn), lcbOf(FC.sttbfFfn));
  const { lists, lfos } = readLists(table, fcOf(FC.plfLst), lcbOf(FC.plfLst), fcOf(FC.plfLfo), lcbOf(FC.plfLfo));
  const chpRuns = readRuns(wd, table, fcOf(FC.plcfBteChpx), lcbOf(FC.plcfBteChpx), 'chp');
  const papRuns = readRuns(wd, table, fcOf(FC.plcfBtePapx), lcbOf(FC.plcfBtePapx), 'pap');

  const papCache = new Map();
  const papFor = (fc) => {
    const r = findRun(papRuns, fc);
    if (papCache.has(r)) return papCache.get(r);
    const run = papRuns[r];
    const istd = run && run.len >= 2 ? u16(wd, run.at) : 0;
    const st = style(istd);
    const pap = { ...st.pap, istd, sti: st.sti };
    if (run && run.len > 2) applyPap(pap, wd, run.at + 2, run.at + run.len);
    const out = { pap, styleChp: st.chp };
    papCache.set(r, out);
    return out;
  };
  const chpFor = (r, styleChp) => {
    const chp = { ...styleChp };
    const run = chpRuns[r];
    if (run && run.len) applyChp(chp, wd, run.at, run.at + run.len, styleChp);
    if (chp.font === undefined) chp.font = defaultFont;
    chp.fontName = fonts[chp.font] || '';
    return chp;
  };

  // Section ends (CP): a section break (\x0c there) also ends its paragraph.
  let sections = [];
  try { sections = readSections(wd, table); } catch { sections = []; }
  const sectionEnds = sections.map((sc) => sc.cpEnd);

  /**
   * Reads the text between two CPs into paragraphs. The main document is CP 0..ccpText;
   * headers and footers are stories after it. PAGE / NUMPAGES fields become PAGE_MARK /
   * PAGES_MARK, filled in with real numbers when pages are laid out.
   */
  let equations = 0;
  const walk = (cpFrom, cpTo) => {
    const paras = [];
    let section = 0;
    let segs = [];
    let seg = null;
    const fields = []; // open fields: { code, inCode, link, page }
    let pending = [];  // characters of the current paragraph: [char, fc, link]
    let at = cpFrom;   // CP after the character being read
    let paraStart = cpFrom;
    const flush = (end, fcEnd) => {
      const { pap, styleChp } = papFor(fcEnd);
      segs = [];
      seg = null;
      let lastRun = -2;
      for (const [ch, fc, link, object] of pending) {
        const r = findRun(chpRuns, fc);
        if (r !== lastRun || !seg || seg.link !== link || ch === OBJ || seg.text.endsWith(OBJ)) {
          seg = { text: '', chp: chpFor(r, styleChp), link };
          if (object) seg.object = object;
          segs.push(seg);
          lastRun = r;
        }
        seg.text += ch;
      }
      for (const x of segs) {
        if (x.text !== OBJ) continue;
        // an equation: drawn from its own data (MTEF), its preview picture is not needed
        if (x.object?.equation) {
          const id = x.object.objFc != null ? chpFor(findRun(chpRuns, x.object.objFc), styleChp).picLocation : undefined;
          try { x.equation = readMtef(objectData?.(id, x.object.index)); } catch { x.equation = null; }
          if (x.equation) continue;
        }
        x.picture = x.chp.picLocation != null ? readInlinePicture(data, x.chp.picLocation) : null;
        if (!x.picture) x.text = '';
      }
      paras.push({ end, pap, styleChp, section, cpStart: paraStart, cpEnd: at, segments: segs.filter((x) => x.text) });
      paraStart = at;
      pending = [];
    };
    const linkOf = () => [...fields].reverse().find((f) => f.link)?.link || null;
    let lastFc = 0;
    for (const pc of pieces) {
      const from = Math.max(pc.cpStart, cpFrom);
      const to = Math.min(pc.cpEnd, cpTo);
      if (to <= from) continue;
      const off = from - pc.cpStart;
      const fc0 = pc.fc + (pc.compressed ? off : off * 2);
      const text = pc.compressed
        ? cp1252.decode(wd.subarray(fc0, fc0 + (to - from)))
        : utf16.decode(wd.subarray(fc0, fc0 + (to - from) * 2));
      let cp = from;
      for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        const fc = fc0 + (pc.compressed ? i : i * 2);
        lastFc = fc;
        cp++;
        at = cp;
        // fields: \x13 code \x14 result \x15; HYPERLINK makes the result a link, PAGE a page number
        if (ch === '\x13') { fields.push({ code: '', inCode: true, link: null, page: null }); continue; }
        if (ch === '\x14') {
          const f = fields.at(-1);
          if (f) {
            f.inCode = false;
            const m = /^\s*HYPERLINK\s+(\\l\s+)?"([^"]+)"/i.exec(f.code);
            if (m) f.link = m[1] ? `#${m[2]}` : m[2];
            const pg = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(f.code)?.[1].toUpperCase();
            if (pg) { f.page = pg; pending.push([pg === 'PAGE' ? PAGE_MARK : PAGES_MARK, fc, linkOf()]); }
          }
          continue;
        }
        if (ch === '\x15') {
          const f = fields.pop();
          // a PAGE field written without a result part (\x13 PAGE \x15)
          const pg = f && f.inCode && /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(f.code)?.[1].toUpperCase();
          if (pg) pending.push([pg === 'PAGE' ? PAGE_MARK : PAGES_MARK, fc, linkOf()]);
          continue;
        }
        if (fields.some((f) => f.inCode)) {
          const f = fields.at(-1);
          if (f.inCode) { f.code += ch; if (ch === '\x01') f.objFc = fc; } // the object's storage id is on this character
          continue;
        }
        if (fields.some((f) => f.page)) continue; // the saved page number is replaced by the marker
        if (ch === '\r' || ch === '\x07') { flush(ch, fc); continue; }
        if (ch === '\x0c' && sectionEnds.includes(cp)) { flush('\r', fc); section++; continue; } // section break
        if (ch === '\x01') { // inline picture, or the result of an EMBED field (equation…)
          const embed = [...fields].reverse().find((f) => /^\s*EMBED\b/i.test(f.code));
          const object = embed ? { equation: true, objFc: embed.objFc, index: equations++ } : null; // equation if its data is one
          pending.push([OBJ, fc, linkOf(), object]);
          continue;
        }
        if (/[\x02\x05\x08\x1f]/.test(ch)) continue; // notes, drawn objects, optional hyphen
        pending.push([ch === '\x1e' ? '-' : ch === '\xa0' ? '\u00a0' : ch, fc, linkOf()]);
      }
    }
    if (pending.length) flush('\r', lastFc);
    return paras;
  };

  const paras = walk(0, ccpText);

  // Headers and footers per section (PlcfHdd): 6 separator stories, then per section
  // even header, header, even footer, footer, first-page header, first-page footer.
  try {
    const fcHdd = fcOf(11);
    const lcbHdd = lcbOf(11);
    const ccpFtn = i32(wd, 80);
    if (lcbHdd > 8) {
      const base = ccpText + ccpFtn;
      const n = lcbHdd / 4;
      const cps = Array.from({ length: n }, (_, i) => i32(table, fcHdd + i * 4));
      const story = (k) => {
        const a = cps[k]; const b2 = cps[k + 1];
        if (a == null || b2 == null || b2 - a <= 1) return null; // empty: the section uses the previous one
        return walk(base + a, base + b2).filter((pp, i, all) => i < all.length - 1 || pp.segments.length); // drop the closing mark
      };
      paras.headers = sections.map((_, si) => ({ header: story(6 + si * 6 + 1), footer: story(6 + si * 6 + 3) }));
    }
  } catch { /* headers and footers are optional */ }

  // Floating shapes, text boxes and pictures, attached to the paragraph holding their anchor.
  // Text box stories follow the main text, footnotes, headers, comments and endnotes.
  try {
    const base = ccpText + i32(wd, 80) + i32(wd, 84) + i32(wd, 92) + i32(wd, 96);
    const fcTx = fcOf(56);
    const n = lcbOf(56) >= 4 ? (lcbOf(56) - 4) / 26 : 0;
    const cps = Array.from({ length: n + 1 }, (_, k) => i32(table, fcTx + k * 4));
    const storyText = (k) => {
      if (!(k >= 1 && k <= n)) return null;
      const story = walk(base + cps[k - 1], base + cps[k]);
      while (story.length && !story.at(-1).segments.length) story.pop();
      return story;
    };
    for (const d of readDrawings(wd, table, fcOf, lcbOf, storyText)) {
      const para = paras.find((pp) => d.cp >= pp.cpStart && d.cp < pp.cpEnd);
      if (para) (para.drawings ||= []).push(d);
    }
  } catch (err) { console.warn('Drawings could not be read', err); }

  // List markers ("•", "1.", "a)") with per-list counters.
  const counters = new Map();
  for (const para of paras) {
    const { ilfo, ilvl = 0 } = para.pap;
    if (!ilfo || ilfo > lfos.length) continue;
    const levels = lists.get(lfos[ilfo - 1]);
    const lvl = levels?.[Math.min(ilvl, levels.length - 1)];
    if (!lvl) continue;
    const key = lfos[ilfo - 1];
    const c = counters.get(key) || [];
    for (let k = 0; k < 9; k++) {
      if (k === ilvl) c[k] = (c[k] ?? (levels[k]?.start ?? 1) - 1) + 1;
      else if (k > ilvl) c[k] = undefined; // a deeper level restarts
    }
    counters.set(key, c);
    let marker = '';
    for (const code of lvl.text) {
      if (code < 9) marker += formatNumber(c[code] ?? levels[code]?.start ?? 1, levels[code]?.nfc ?? 0);
      else marker += bulletCode(code, fonts[lvl.chp.font] || ''); // Symbol / Wingdings -> Unicode
    }
    para.list = { marker, pap: lvl.pap, chp: lvl.chp };
  }
  return paras;
}

// ------------------------------------------------------------------ to CSS

const tw = (v) => `${(v / 15).toFixed(1)}px`; // twips -> CSS px (1440 tw = 96 px)

export function paragraphStyle(pap, list) {
  const left = pap.left ?? list?.pap.left;
  const first = pap.first ?? list?.pap.first;
  const css = { margin: `${tw(pap.before || 0)} ${tw(pap.right || 0)} ${tw(pap.after || 0)} ${tw(left || 0)}` };
  if (first) css.textIndent = tw(first);
  for (const [side, brc] of Object.entries(pap.borders || {})) {
    if (!brc) continue;
    const name = side[0].toUpperCase() + side.slice(1);
    css[`border${name}`] = `${Math.max(1, brc.width * (4 / 3)).toFixed(1)}px solid ${brc.color}`;
    css[`padding${name}`] = '2px';
  }
  if (pap.align) css.textAlign = ['left', 'center', 'right', 'justify'][pap.align] || 'left';
  if (pap.shading) css.backgroundColor = pap.shading;
  if (pap.line) {
    const { value, mult } = pap.line;
    if (mult) css.lineHeight = String(Math.max(0.8, value / 240) * 1.15);
    else if (value) css.lineHeight = tw(Math.abs(value));
  }
  return css;
}

const MONO = /mono|courier|consol/i;
const SERIF = /times|roman|georgia|cambria|garamond|serif/i;
export function charStyle(chp) {
  const css = {};
  if (chp.bold) css.fontWeight = '700';
  if (chp.italic) css.fontStyle = 'italic';
  const deco = [chp.underline && 'underline', chp.strike && 'line-through'].filter(Boolean);
  if (deco.length) css.textDecoration = deco.join(' ');
  if (chp.underline === 3) css.textDecorationStyle = 'double';
  if (chp.color) css.color = chp.color;
  if (chp.highlight) css.backgroundColor = chp.highlight;
  if (chp.size) css.fontSize = `${chp.size}pt`;
  if (chp.fontName) {
    const generic = MONO.test(chp.fontName) ? 'monospace' : SERIF.test(chp.fontName) ? 'serif' : 'sans-serif';
    css.fontFamily = `"${chp.fontName.replace(/"/g, '')}", ${generic}`;
  }
  if (chp.vert === 1) { css.verticalAlign = 'super'; css.fontSize = '0.7em'; }
  if (chp.vert === 2) { css.verticalAlign = 'sub'; css.fontSize = '0.7em'; }
  if (chp.caps) css.textTransform = 'uppercase';
  if (chp.smallCaps) css.fontVariant = 'small-caps';
  if (chp.hidden) css.display = 'none';
  return css;
}
