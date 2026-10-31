/*
 * Excel 97-2003 (.xls, BIFF8) cell formatting that SheetJS CE does not hand over, read from the
 * Workbook stream following Microsoft's [MS-XLS] specification:
 *   - fonts (bold, italic, underline / double, strikethrough, color, size, name, super/subscript)
 *   - alignment, wrap, indent, rotation, borders, pattern fills (XF records + PALETTE)
 *   - rich text inside a cell (SST formatting runs), frozen panes (PANE)
 *   - conditional formatting "cell value" and "formula" rules (CONDFMT / CF)
 *   - pictures placed on sheets (MSODRAWING anchors + MSODRAWINGGROUP pictures)
 * The result has the same shape as readCellStyles() for .xlsx, so the sheet view draws both alike.
 */

const DEFAULT_PALETTE = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', // 0-7
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000',
  '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99',
  '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];
const BORDER = [null, ['thin', 1], ['medium', 2], ['dashed', 1], ['dotted', 1], ['thick', 3], ['double', 3], ['hair', 1],
  ['mediumDashed', 2], ['dashDot', 1], ['mediumDashDot', 2], ['dashDotDot', 1], ['mediumDashDotDot', 2], ['slantDashDot', 2]];
const BORDER_CSS = { dotted: 'dotted', hair: 'dotted', dashed: 'dashed', mediumDashed: 'dashed', dashDot: 'dashed', mediumDashDot: 'dashed', dashDotDot: 'dashed', mediumDashDotDot: 'dashed', slantDashDot: 'dashed', double: 'double' };
const FILLS = [null, 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp',
  'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];
const ALIGN_H = [null, 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
const ALIGN_V = ['top', 'center', null, 'justify', 'distributed'];

const u16 = (b, p) => b[p] | (b[p + 1] << 8);
const u32 = (b, p) => (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0;
const col = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const addr = (r, c) => `${col(c)}${r + 1}`;
const latin1 = new TextDecoder('latin1');
const utf16 = new TextDecoder('utf-16le');

/** All BIFF records of the stream, with CONTINUE records kept next to the one they extend. */
function readRecords(b) {
  const out = [];
  for (let p = 0; p + 4 <= b.length;) {
    const type = u16(b, p); const len = u16(b, p + 2);
    const rec = { type, body: p + 4, end: p + 4 + len };
    if (type === 0x003c && out.length) (out.at(-1).cont ||= []).push(rec);
    else out.push(rec);
    p += 4 + len;
  }
  return out;
}

/** XLUnicodeString with a 1- or 2-byte character count. */
function readString(b, p, cchBytes = 2) {
  const cch = cchBytes === 1 ? b[p] : u16(b, p);
  const flags = b[p + cchBytes];
  const start = p + cchBytes + 1;
  const wide = flags & 1;
  const text = wide ? utf16.decode(b.subarray(start, start + cch * 2)) : latin1.decode(b.subarray(start, start + cch));
  return { text, next: start + cch * (wide ? 2 : 1) };
}

/** Shared strings with their formatting runs; strings may continue across CONTINUE records. */
function readSst(b, rec) {
  const segs = [{ body: rec.body + 8, end: rec.end }, ...(rec.cont || []).map((c) => ({ body: c.body, end: c.end }))];
  let s = 0; let p = segs[0].body;
  const need = () => { if (p >= segs[s].end && s + 1 < segs.length) { s++; p = segs[s].body; } };
  const out = [];
  const total = u32(b, rec.body + 4);
  for (let i = 0; i < total && s < segs.length; i++) {
    need();
    const cch = u16(b, p); p += 2;
    need();
    let flags = b[p]; p += 1;
    let runs = 0; let ext = 0;
    if (flags & 0x8) { runs = u16(b, p); p += 2; }
    if (flags & 0x4) { ext = u32(b, p); p += 4; }
    let text = '';
    let left = cch;
    while (left > 0) {
      need();
      if (text.length && p === segs[s].body) { flags = (flags & ~1) | (b[p] & 1); p += 1; } // a continued string restarts with its flags
      const wide = flags & 1;
      const fit = Math.min(left, Math.floor((segs[s].end - p) / (wide ? 2 : 1)));
      text += wide ? utf16.decode(b.subarray(p, p + fit * 2)) : latin1.decode(b.subarray(p, p + fit));
      p += fit * (wide ? 2 : 1);
      left -= fit;
      if (left > 0 && s + 1 >= segs.length) break;
      if (left > 0) { s++; p = segs[s].body; }
    }
    const fmt = [];
    for (let k = 0; k < runs; k++) { need(); fmt.push({ ich: u16(b, p), ifnt: u16(b, p + 2) }); p += 4; }
    p += ext;
    out.push({ text, runs: fmt });
  }
  return out;
}

/**
 * Formula tokens ([MS-XLS] Rgce) back to formula text, for the token types conditional formats use.
 * Relative refs (ptgRefN / ptgAreaN) are relative to the rule's first cell.
 */
const FUNCS = { 0: 'COUNT', 1: 'IF', 4: 'SUM', 5: 'AVERAGE', 6: 'MIN', 7: 'MAX', 36: 'AND', 37: 'OR', 38: 'NOT', 346: 'COUNTIF', 345: 'SUMIF', 336: 'CONCATENATE' };
const FIXED_ARGS = { 38: 1, 346: 2 }; // arguments of fixed-argument functions (NOT, COUNTIF)
function decodeFormula(b, p, end, origin) {
  const stack = [];
  const bin = { 0x03: '+', 0x04: '-', 0x05: '*', 0x06: '/', 0x07: '^', 0x08: '&', 0x09: '<', 0x0a: '<=', 0x0b: '=', 0x0c: '>=', 0x0d: '>', 0x0e: '<>' };
  const ref = (rw, cl, relN) => {
    const rowRel = cl & 0x8000; const colRel = cl & 0x4000;
    let r = rw; let c = cl & 0x3fff;
    if (relN) { if (rowRel) r = origin.r + ((rw << 16) >> 16); if (colRel) c = origin.c + ((c << 18) >> 18); }
    return `${colRel ? '' : '$'}${col(c)}${rowRel ? '' : '$'}${r + 1}`;
  };
  while (p < end) {
    const t = b[p++];
    const base = t & 0x1f | (t >= 0x20 ? 0x20 : 0);
    if (bin[t]) { const y = stack.pop(); const x = stack.pop(); stack.push(`${x}${bin[t]}${y}`); continue; }
    switch (t) {
      case 0x1e: stack.push(String(u16(b, p))); p += 2; break; // ptgInt
      case 0x1f: stack.push(String(new DataView(b.buffer, b.byteOffset + p, 8).getFloat64(0, true))); p += 8; break; // ptgNum
      case 0x17: { const s = readString(b, p, 1); stack.push(`"${s.text}"`); p = s.next; break; } // ptgStr
      case 0x1d: stack.push(b[p] ? 'TRUE' : 'FALSE'); p += 1; break; // ptgBool
      case 0x15: stack.push(`(${stack.pop()})`); break; // ptgParen
      default:
        if ([0x24, 0x44, 0x64].includes(t)) { stack.push(ref(u16(b, p), u16(b, p + 2), false)); p += 4; } // ptgRef
        else if ([0x2c, 0x4c, 0x6c].includes(t)) { stack.push(ref(u16(b, p), u16(b, p + 2), true)); p += 4; } // ptgRefN
        else if ([0x25, 0x45, 0x65].includes(t)) { stack.push(`${ref(u16(b, p), u16(b, p + 4), false)}:${ref(u16(b, p + 2), u16(b, p + 6), false)}`); p += 8; } // ptgArea
        else if ([0x2d, 0x4d, 0x6d].includes(t)) { stack.push(`${ref(u16(b, p), u16(b, p + 4), true)}:${ref(u16(b, p + 2), u16(b, p + 6), true)}`); p += 8; } // ptgAreaN
        else if ([0x21, 0x41, 0x61].includes(t)) { // ptgFunc: a function with a fixed number of arguments
          const id = u16(b, p); p += 2;
          const n = FIXED_ARGS[id] ?? 1;
          const args = stack.splice(stack.length - n, n);
          stack.push(`${FUNCS[id] || 'FN'}(${args.join(',')})`);
        }
        else if ([0x22, 0x42, 0x62].includes(t)) { const n = b[p] & 0x7f; const f = FUNCS[u16(b, p + 1) & 0x7fff]; p += 3; const args = stack.splice(stack.length - n, n); stack.push(`${f || 'FN'}(${args.join(',')})`); } // ptgFuncVar
        else return null; // something this reader does not decode
        void base;
    }
  }
  return stack.length === 1 ? stack[0] : null;
}

/** DXF-style formatting stored in a CF record: font color / bold / italic and the fill. */
function readCfFormat(b, p, palette) {
  const flags = u32(b, p); p += 6; // flags (4) + reserved (2)
  const out = {};
  if (flags & 0x02000000) p += 2; // fHasNumFmt: ifmt
  if (flags & 0x04000000) { // fHasFont: 118-byte font block
    const color = u32(b, p + 80);
    const sts = u32(b, p + 68);
    const weight = u16(b, p + 72);
    if (color !== 0xffffffff && palette[color]) out.color = `#${palette[color]}`;
    if (weight >= 700) out.bold = true;
    if (sts & 0x2) out.italic = true;
    p += 118;
  }
  if (flags & 0x08000000) p += 8; // alignment
  if (flags & 0x10000000) p += 8; // border
  if (flags & 0x20000000) { // pattern
    const fls = (u16(b, p) >>> 10) & 0x3f; const icv = u16(b, p + 2); // DXFPat: fill pattern, then colors
    const fore = icv & 0x7f; const back = (icv >>> 7) & 0x7f;
    // solid patterns show the background color in conditional formats
    const c = fls <= 1 ? palette[back] || palette[fore] : palette[fore];
    if (c) out.fill = `#${c}`;
  }
  return out;
}

/** Anchored pictures in a sheet's drawing: OfficeArtClientAnchorSheet + the picture's BStore index. */
function readSheetPictures(b, rec, blips) {
  const bytes = [b.subarray(rec.body, rec.end), ...(rec.cont || []).map((c) => b.subarray(c.body, c.end))];
  const data = new Uint8Array(bytes.reduce((n, x) => n + x.length, 0));
  let o = 0;
  for (const x of bytes) { data.set(x, o); o += x.length; }
  const pics = [];
  const walk = (start, end) => {
    for (let p = start; p + 8 <= end;) {
      const ver = u16(data, p); const type = u16(data, p + 2); const len = u32(data, p + 4);
      const body = p + 8;
      if ((ver & 0xf) === 0xf) {
        if (type === 0xf004) { // SpContainer: look for a picture reference and an anchor
          let pib = 0; let anchor = null;
          for (let q = body; q + 8 <= body + len;) {
            const t = u16(data, q + 2); const l = u32(data, q + 4);
            if (t === 0xf00b) { // OPT: property 0x0104 = pib
              const n = u16(data, q) >> 4;
              for (let k = 0; k < n; k++) { const id = u16(data, q + 8 + k * 6) & 0x3fff; if (id === 0x0104) pib = u32(data, q + 8 + k * 6 + 2); }
            } else if (t === 0xf010 && l >= 18) {
              anchor = { c1: u16(data, q + 10), dx1: u16(data, q + 12), r1: u16(data, q + 14), dy1: u16(data, q + 16), c2: u16(data, q + 18), dx2: u16(data, q + 20), r2: u16(data, q + 22), dy2: u16(data, q + 24) };
            }
            q += 8 + l;
          }
          if (pib && anchor && blips[pib - 1]) pics.push({ anchor, ...blips[pib - 1] });
        } else walk(body, body + len);
      }
      p = body + len;
    }
  };
  walk(0, data.length);
  return pics;
}

/** BStore pictures of the workbook's drawing group, as data URLs (PNG / JPEG). */
function readBlips(b, rec) {
  const bytes = [b.subarray(rec.body, rec.end), ...(rec.cont || []).map((c) => b.subarray(c.body, c.end))];
  const data = new Uint8Array(bytes.reduce((n, x) => n + x.length, 0));
  let o = 0;
  for (const x of bytes) { data.set(x, o); o += x.length; }
  const blips = [];
  const walk = (start, end) => {
    for (let p = start; p + 8 <= end;) {
      const ver = u16(data, p); const type = u16(data, p + 2); const len = u32(data, p + 4);
      const body = p + 8;
      if (type === 0xf007) { // FBSE, with its embedded BLIP after the 36-byte header
        const nameLen = data[body + 33];
        const bp = body + 36 + nameLen;
        const btype = u16(data, bp + 2); const blen = u32(data, bp + 4);
        const inst = u16(data, bp) >> 4;
        const mime = { 0xf01d: 'image/jpeg', 0xf02a: 'image/jpeg', 0xf01e: 'image/png' }[btype];
        if (mime && bp + 8 + blen <= end) {
          // header: 16-byte UID (+16 more for some instances) + 1-byte tag
          const skip = 16 + ([0x46b, 0x6e1, 0x6e3, 0x6e5, 0x7a9].includes(inst) ? 16 : 0) + 1;
          const img = data.subarray(bp + 8 + skip, bp + 8 + blen);
          let s = '';
          for (let i = 0; i < img.length; i += 0x8000) s += String.fromCharCode(...img.subarray(i, i + 0x8000));
          blips.push({ src: `data:${mime};base64,${btoa(s)}` });
        } else blips.push(null);
      } else if ((ver & 0xf) === 0xf) walk(body, body + len);
      p = body + len;
    }
  };
  walk(0, data.length);
  return blips;
}

/**
 * @param {Uint8Array} b the Workbook stream
 * @param {number} maxRows only cells in the first rows are read
 * @returns {Record<string, object>} sheet name -> cell address -> style (and '!freeze', '!cf', '!pics', '!rich')
 */
export function readXlsStyles(b, maxRows) {
  const recs = readRecords(b);
  const fonts = []; const xfs = []; const sheets = []; let sst = []; let blips = [];
  const palette = DEFAULT_PALETTE.slice();
  let globals = true;
  for (const r of recs) {
    if (r.type === 0x000a && globals) { globals = false; continue; }
    if (!globals) break;
    if (r.type === 0x0031) { // FONT
      const nameRec = readString(b, r.body + 14, 1);
      fonts.push({ size: u16(b, r.body) / 20, italic: !!(u16(b, r.body + 2) & 0x2), strike: !!(u16(b, r.body + 2) & 0x8), icv: u16(b, r.body + 4), bold: u16(b, r.body + 6) >= 700, sss: u16(b, r.body + 8), uls: b[r.body + 10], name: nameRec.text });
      if (fonts.length === 4) fonts.push(null); // font index 4 is never used in BIFF8
    } else if (r.type === 0x0092) { // PALETTE: colors 8.. replaced
      const n = u16(b, r.body);
      for (let i = 0; i < n; i++) palette[8 + i] = [0, 1, 2].map((k) => b[r.body + 2 + i * 4 + k].toString(16).padStart(2, '0')).join('').toUpperCase();
    } else if (r.type === 0x00e0) xfs.push(r); // XF
    else if (r.type === 0x0085) sheets.push({ pos: u32(b, r.body), name: readString(b, r.body + 6, 1).text });
    else if (r.type === 0x00fc) sst = readSst(b, r);
    else if (r.type === 0x00eb) { try { blips = readBlips(b, r); } catch { blips = []; } }
  }
  const color = (icv) => (icv === 0x7fff || icv === 64 || icv === 0x41 ? null : palette[icv] ? `#${palette[icv]}` : null);
  const defFont = fonts[0] || { size: 10, name: '' };

  const fontStyle = (f, s) => {
    if (!f) return;
    if (f.bold) s.b = true;
    if (f.italic) s.i = true;
    if (f.strike) s.strike = true;
    if (f.uls === 1 || f.uls === 0x21) s.u = true;
    if (f.uls === 2 || f.uls === 0x22) s.u = 'double';
    const c = color(f.icv);
    if (c && c !== '#000000') s.color = c;
    if (f.size && f.size !== defFont.size) s.size = f.size / defFont.size;
    if (f.name && f.name !== defFont.name) s.font = f.name;
  };
  const cache = new Map();
  const xfStyle = (ixfe) => {
    if (cache.has(ixfe)) return cache.get(ixfe);
    const r = xfs[ixfe];
    let out = null;
    if (r) {
      const p = r.body;
      const s = {};
      fontStyle(fonts[u16(b, p)], s);
      const al = b[p + 6];
      if (ALIGN_H[al & 7]) s.align = ALIGN_H[al & 7];
      if (al & 0x8) s.wrap = true;
      const v = ALIGN_V[(al >> 4) & 7];
      if (v) s.valign = v;
      if (b[p + 7]) s.rotate = b[p + 7];
      if (b[p + 8] & 0xf) s.indent = b[p + 8] & 0xf;
      const b1 = u32(b, p + 10); const b2 = u32(b, p + 14);
      const sides = { left: [b1 & 0xf, (b1 >>> 16) & 0x7f], right: [(b1 >>> 4) & 0xf, (b1 >>> 23) & 0x7f], top: [(b1 >>> 8) & 0xf, b2 & 0x7f], bottom: [(b1 >>> 12) & 0xf, (b2 >>> 7) & 0x7f] };
      const border = {};
      for (const [side, [dg, icv]] of Object.entries(sides)) {
        const st = BORDER[dg];
        if (st) border[side] = `${st[1]}px ${BORDER_CSS[st[0]] || 'solid'} ${color(icv) || '#000000'}`;
      }
      if (Object.keys(border).length) s.border = border;
      const fls = (b2 >>> 26) & 0x3f;
      const icvs = u16(b, p + 18);
      if (FILLS[fls] && fls > 1) s.pattern = { pattern: FILLS[fls], fg: color(icvs & 0x7f) || '#000000', bg: color((icvs >>> 7) & 0x7f) || '#ffffff' };
      if (Object.keys(s).length) out = s;
    }
    cache.set(ixfe, out);
    return out;
  };

  // rich text as OOXML-style runs, so the sheet view's richText() draws .xls and .xlsx alike
  const esc = (t) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const richXml = (str) => {
    if (!str?.runs?.length) return null;
    const cuts = [{ ich: 0, ifnt: -1 }, ...str.runs].sort((x, y) => x.ich - y.ich);
    let xml = '';
    cuts.forEach((run, k) => {
      const text = str.text.slice(run.ich, cuts[k + 1]?.ich ?? str.text.length);
      if (!text) return;
      const f = run.ifnt >= 0 ? fonts[run.ifnt] : null;
      let pr = '';
      if (f) {
        if (f.bold) pr += '<b/>';
        if (f.italic) pr += '<i/>';
        if (f.strike) pr += '<strike/>';
        if (f.uls) pr += f.uls === 2 ? '<u val="double"/>' : '<u/>';
        if (f.sss === 1) pr += '<vertAlign val="superscript"/>';
        if (f.sss === 2) pr += '<vertAlign val="subscript"/>';
        const c = color(f.icv);
        if (c && c !== '#000000') pr += `<color rgb="FF${c.slice(1)}"/>`;
      }
      xml += `<r>${pr ? `<rPr>${pr}</rPr>` : ''}<t xml:space="preserve">${esc(text)}</t></r>`;
    });
    return xml;
  };

  const result = {};
  for (const sh of sheets) {
    const map = {};
    const rich = {};
    const start = recs.findIndex((r) => r.body - 4 === sh.pos);
    if (start < 0) continue;
    let cfRanges = null;
    const rules = [];
    for (let i = start + 1; i < recs.length && recs[i].type !== 0x000a; i++) {
      const r = recs[i];
      const t = r.type;
      const cell = (rw, cl, ixfe) => { if (rw < maxRows) { const s = xfStyle(ixfe); if (s) map[addr(rw, cl)] = s; } };
      if ([0x00fd, 0x0203, 0x027e, 0x0201, 0x0205, 0x0006, 0x0204].includes(t)) {
        cell(u16(b, r.body), u16(b, r.body + 2), u16(b, r.body + 4));
        if (t === 0x00fd) { // LABELSST: rich text runs
          const x = richXml(sst[u32(b, r.body + 6)]);
          if (x) rich[addr(u16(b, r.body), u16(b, r.body + 2))] = x;
        }
      } else if (t === 0x00bd || t === 0x00be) { // MULRK / MULBLANK
        const rw = u16(b, r.body); const first = u16(b, r.body + 2); const step = t === 0x00bd ? 6 : 2;
        const lastCol = u16(b, r.end - 2);
        for (let c = first, p = r.body + 4; c <= lastCol; c++, p += step) cell(rw, c, u16(b, p));
      } else if (t === 0x0041) { // PANE: frozen rows / columns
        map['!freeze'] = { cols: u16(b, r.body), rows: u16(b, r.body + 2) };
      } else if (t === 0x01b0) { // CONDFMT: ranges of the following CF records
        const n = u16(b, r.body + 12);
        cfRanges = [];
        for (let k = 0; k < n; k++) {
          const q = r.body + 14 + k * 8;
          cfRanges.push(`${addr(u16(b, q), u16(b, q + 4))}:${addr(u16(b, q + 2), u16(b, q + 6))}`);
        }
      } else if (t === 0x01b1 && cfRanges) { // CF: cell-value or formula rule with its formatting
        const ct = b[r.body]; const cp = b[r.body + 1];
        const cce1 = u16(b, r.body + 2); const cce2 = u16(b, r.body + 4);
        const fmt = readCfFormat(b, r.body + 6, palette);
        const flags = u32(b, r.body + 6);
        let fmtLen = 6 + (flags & 0x02000000 ? 2 : 0) + (flags & 0x04000000 ? 118 : 0) + (flags & 0x08000000 ? 8 : 0) + (flags & 0x10000000 ? 8 : 0) + (flags & 0x20000000 ? 4 : 0) + (flags & 0x40000000 ? 2 : 0);
        const f1Start = r.body + 6 + fmtLen;
        const m = /^([A-Z]+)(\d+)/.exec(cfRanges[0] || 'A1');
        const origin = { r: Number(m[2]) - 1, c: m[1].split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1 };
        const f1 = cce1 ? decodeFormula(b, f1Start, f1Start + cce1, origin) : null;
        const f2 = cce2 ? decodeFormula(b, f1Start + cce1, f1Start + cce1 + cce2, origin) : null;
        const ops = [null, 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual'];
        if (f1 != null) {
          rules.push(ct === 1
            ? { ranges: cfRanges, type: 'cellIs', operator: ops[cp], formulas: [f1, f2].filter((x) => x != null), dxf: fmt, priority: rules.length, cfvo: [], colors: [] }
            : { ranges: cfRanges, type: 'expression', formulas: [f1], dxf: fmt, priority: rules.length, cfvo: [], colors: [] });
        }
      } else if (t === 0x00ec) { // MSODRAWING: pictures placed on the sheet
        try {
          for (const pic of readSheetPictures(b, r, blips)) {
            const a = pic.anchor;
            (map['!pics'] ||= []).push({
              kind: 'twoCellAnchor', src: pic.src,
              // .xls offsets are in 1/1024 of a cell's width and 1/256 of its height: converted later
              from: { col: a.c1, row: a.r1, colFrac: a.dx1 / 1024, rowFrac: a.dy1 / 256 },
              to: { col: a.c2, row: a.r2, colFrac: a.dx2 / 1024, rowFrac: a.dy2 / 256 },
            });
          }
        } catch { /* pictures are optional */ }
      }
    }
    if (rules.length) map['!cf'] = rules;
    if (Object.keys(rich).length) map['!rich'] = rich;
    result[sh.name] = map;
  }
  return result;
}
