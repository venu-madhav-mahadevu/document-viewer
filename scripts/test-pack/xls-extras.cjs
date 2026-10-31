// Adds to the .xls test file the Excel features LibreOffice leaves out when it saves .xls, written
// as Excel's own BIFF8 records ([MS-XLS]): the stripe pattern fill (XF), the gradient fill
// (XFCRC + XFExt), the color scale, data bars and icon set (CondFmt12 + CF12) and the filter
// buttons (_FilterDatabase name + AutoFilterInfo).
// usage: node xls-extras.cjs out/test-excel.xls   (the file is changed in place)
const fs = require('fs');
const XLSX = require('xlsx');

const file = process.argv[2];
const cfb = XLSX.CFB.read(fs.readFileSync(file), { type: 'buffer' });
const entry = XLSX.CFB.find(cfb, 'Workbook');
const src = Buffer.from(entry.content);

// ---------- records
let recs = [];
for (let p = 0; p + 4 <= src.length;) {
  const type = src.readUInt16LE(p); const len = src.readUInt16LE(p + 2);
  recs.push({ type, data: Buffer.from(src.subarray(p + 4, p + 4 + len)) });
  p += 4 + len;
}
const rec = (type, data) => ({ type, data });
const BOF = 0x0809; const EOF = 0x000a;

// sheet substreams: [first record index, EOF index] by sheet name
const sheetNames = recs.filter((r) => r.type === 0x0085).map((r) => r.data.toString('latin1', 8, 8 + r.data[6]));
const sheets = () => {
  const out = []; let start = -1;
  recs.forEach((r, i) => {
    if (r.type === BOF && i > 0) start = i;
    if (r.type === EOF && start >= 0) { out.push([start, i]); start = -1; }
  });
  return out;
};
const sheetIndex = (name) => { const i = sheetNames.indexOf(name); if (i < 0) throw new Error(`no sheet ${name}`); return i; };
const findIn = (name, test) => { const [a, b] = sheets()[sheetIndex(name)]; for (let i = a; i <= b; i++) if (test(recs[i])) return i; return -1; };

// small writers
const u8 = (v) => Buffer.from([v & 0xff]);
const u16 = (v) => { const b = Buffer.alloc(2); b.writeUInt16LE(v); return b; };
const u32 = (v) => { const b = Buffer.alloc(4); b.writeUInt32LE(v >>> 0); return b; };
const f64 = (v) => { const b = Buffer.alloc(8); b.writeDoubleLE(v); return b; };
const cat = (...parts) => Buffer.concat(parts);
const rgba = (hex) => Buffer.from([parseInt(hex.slice(0, 2), 16), parseInt(hex.slice(2, 4), 16), parseInt(hex.slice(4, 6), 16), 0xff]); // LongRGBA
const ref8 = (r1, r2, c1, c2) => cat(u16(r1), u16(r2), u16(c1), u16(c2));

// ---------- 1. fills: two new XFs (copies of the cells' XFs), the cells pointed at them
const xfIdx = recs.map((r, i) => (r.type === 0x00e0 ? i : -1)).filter((i) => i >= 0);
const lastXf = xfIdx.at(-1);
const xfCount = xfIdx.length;
const cellXf = (sheet, row, col) => {
  const i = findIn(sheet, (r) => [0x00fd, 0x0203, 0x027e, 0x0201, 0x0205].includes(r.type) && r.data.readUInt16LE(0) === row && r.data.readUInt16LE(2) === col);
  if (i < 0) throw new Error(`no cell ${sheet} R${row + 1}C${col + 1}`);
  return i;
};
const patternCell = cellXf('Colors & fonts', 10, 1); // B11
const gradientCell = cellXf('Colors & fonts', 11, 1); // B12
const xfOf = (i) => Buffer.from(recs[xfIdx[recs[i].data.readUInt16LE(4)]].data);

// stripe pattern: fls 8 = darkUp, foreground = the light blue already used, background white (icv 9)
const patternXf = xfOf(patternCell);
const b2 = patternXf.readUInt32LE(14);
patternXf.writeUInt32LE(((b2 & 0x03ffffff) | (8 << 26)) >>> 0, 14);
const icvs = patternXf.readUInt16LE(18);
patternXf.writeUInt16LE((icvs & 0xc07f) | (9 << 7), 18);
// gradient: the XF keeps its solid blue (what programs without gradients show); the XFExt adds the gradient
const gradientXf = xfOf(gradientCell);
recs[patternCell].data.writeUInt16LE(xfCount, 4);
recs[gradientCell].data.writeUInt16LE(xfCount + 1, 4);

// XFCRC: [MS-OSHARED] 2.4.3 CRC (polynomial 0xAF) over the data of every XF record
const CRC_TABLE = Array.from({ length: 256 }, (_, i) => {
  let v = i << 24;
  for (let k = 0; k < 8; k++) v = v & 0x80000000 ? ((v << 1) ^ 0xaf) >>> 0 : (v << 1) >>> 0;
  return v >>> 0;
});
const crc = (value, bytes) => { for (const byte of bytes) value = ((value << 8) ^ CRC_TABLE[((value >>> 24) ^ byte) & 0xff]) >>> 0; return value; };
const allXfs = [...xfIdx.map((i) => recs[i].data), patternXf, gradientXf];
const xfcrc = cat(u16(0x087c), u16(0), Buffer.alloc(8), u16(0), u16(allXfs.length), u32(allXfs.reduce(crc, 0)));
// XFExt with one ExtProp (type 6, XFExtGradient): linear, 0 degrees (left to right), white -> 0070C0
const stop = (hex, pos) => cat(u16(2), rgba(hex), f64(pos), f64(0)); // GradStop: RGB
const gradient = cat(u32(0), f64(0), f64(0), f64(0), f64(0), f64(0), u32(2), stop('FFFFFF', 0), stop('0070C0', 1));
const xfext = cat(u16(0x087d), u16(0), Buffer.alloc(8), u16(0), u16(xfCount + 1), u16(0), u16(1), u16(6), u16(4 + gradient.length), gradient);
recs.splice(lastXf + 1, 0, rec(0x00e0, patternXf), rec(0x00e0, gradientXf), rec(0x087c, xfcrc), rec(0x087d, xfext));

// ---------- 2. color scale, data bars, icon set: CondFmt12 + CF12 on the "Conditional formatting" sheet
const cfvo = (type, value) => cat(u8(type), u16(0), type === 2 || type === 3 ? Buffer.alloc(0) : f64(value)); // CFVO, no formula
const cfColor = (hex) => cat(u32(2), rgba(hex), f64(0)); // CFColor: RGB
let priority = 100;
const cf12 = (ct, template, rgbCT) => cat(
  u16(0x087a), u16(0), Buffer.alloc(8), // frtRefHeader (ref8 zero)
  u8(ct), u8(0), u16(0), u16(0),        // ct, cp, cce1, cce2
  u32(0), u16(0),                       // dxf: cbDxf 0 + reserved
  u16(0),                               // fmlaActive: no formula
  u8(0), u16(priority++), u16(template), u8(16), Buffer.alloc(16), // flags, ipriority, icfTemplate, template parameters
  rgbCT,
);
let nID = 10;
const condFmt12 = (row) => cat(
  u16(0x0879), u16(1), ref8(row, row, 1, 3), // frtRefHeaderU (fFrtRef = 1, ref8 = refBound)
  u16(1), u16(nID++ << 1), ref8(row, row, 1, 3), // ccf, fToughRecalc / nID, refBound
  u16(1), ref8(row, row, 1, 3),              // sqref: B:D of the row
);
const colorScale = cat(u16(0), u8(0), u8(3), u8(3), u8(0x03), // fClamp + fBackground
  cfvo(2), f64(0), cfvo(5, 50), f64(0.5), cfvo(3), f64(1),
  f64(0), cfColor('F8696B'), f64(0.5), cfColor('FFEB84'), f64(1), cfColor('63BE7B'));
const dataBar = cat(u16(0), u8(0), u8(0x02), u8(10), u8(90), cfColor('638EC6'), cfvo(2), cfvo(3)); // fShowValue, 10% - 90% of the cell
const iconSet = cat(u16(0), u8(0), u8(3), u8(0x00), u8(0), // 3 states, icon set 0 = 3 colored arrows
  cfvo(2), u8(1), Buffer.alloc(4), cfvo(4, 33), u8(1), Buffer.alloc(4), cfvo(4, 67), u8(1), Buffer.alloc(4));
const cfSheet = 'Conditional formatting';
const cfEof = sheets()[sheetIndex(cfSheet)][1];
recs.splice(cfEof, 0,
  rec(0x0879, condFmt12(2)), rec(0x087a, cf12(3, 2, colorScale)),
  rec(0x0879, condFmt12(3)), rec(0x087a, cf12(4, 3, dataBar)),
  rec(0x0879, condFmt12(4)), rec(0x087a, cf12(6, 4, iconSet)));

// ---------- 3. filter buttons on "Notes & links" B22:D25
const fSheet = 'Notes & links';
const itab = sheetIndex(fSheet);
const ext = recs.findIndex((r) => r.type === 0x0017);
if (ext < 0) throw new Error('no EXTERNSHEET');
const xtiCount = recs[ext].data.readUInt16LE(0);
recs[ext].data = cat(u16(xtiCount + 1), recs[ext].data.subarray(2), u16(0), u16(itab), u16(itab)); // a reference to the sheet
const ixti = xtiCount;
const area3d = cat(u8(0x3b), u16(ixti), u16(21), u16(24), u16(1), u16(3)); // PtgArea3d $B$22:$D$25
const lbl = cat(u16(0x0021), u8(0), u8(1), u16(area3d.length), u16(0), u16(itab + 1), u8(0), u8(0), u8(0), u8(0), u8(0), u8(0x0d), area3d); // hidden built-in _FilterDatabase
recs.splice(ext + 1, 0, rec(0x0018, lbl));
const dims = findIn(fSheet, (r) => r.type === 0x0200);
recs.splice(dims, 0, rec(0x009d, u16(3))); // AutoFilterInfo: 3 columns

// ---------- write back, sheet offsets in BOUNDSHEET updated
const write = () => Buffer.concat(recs.flatMap((r) => [u16(r.type), u16(r.data.length), r.data]));
let pos = 0; const bofPos = [];
recs.forEach((r, i) => { if (r.type === BOF && i > 0) bofPos.push(pos); pos += 4 + r.data.length; });
recs.filter((r) => r.type === 0x0085).forEach((r, i) => r.data.writeUInt32LE(bofPos[i], 0));
const out = write();
entry.content = out;
entry.size = out.length;
fs.writeFileSync(file, XLSX.CFB.write(cfb, { type: 'buffer' }));
console.log(`ok ${file}: ${src.length} -> ${out.length} bytes`);
