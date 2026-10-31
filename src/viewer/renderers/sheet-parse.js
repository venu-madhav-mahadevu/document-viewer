// Spreadsheet parsing, shared by the Web Worker (normal case) and the page itself
// (fallback when a worker cannot run, e.g. blocked by a security policy).
import * as XLSX from 'xlsx';
import { readCellStyles } from './xlsx-styles.js';
import { readXlsStyles } from './xls-format.js';
import { evaluate } from './formula.js';

const isCompoundFile = (buffer) => {
  const b = new Uint8Array(buffer, 0, Math.min(8, buffer.byteLength));
  return b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0; // .xls (OLE compound file)
};

export function parseWorkbook(buffer, isCsv, maxRows) {
  const xls = !isCsv && isCompoundFile(buffer);
  const options = {
    sheetRows: maxRows, // stop reading after the rows we show: far faster on huge sheets
    cellDates: true,
    cellFormula: xls, // .xls: formulas whose text results were not saved are worked out below
    cellHTML: false,
    cellStyles: true, // cell background colors (fills)
  };
  // Raw workbook files (for fonts, alignment and borders) only for normal-size files.
  const withFiles = !isCsv && buffer.byteLength < 25 * 1024 * 1024;
  const wb = isCsv
    ? XLSX.read(new TextDecoder().decode(buffer), { ...options, type: 'string', raw: true })
    : XLSX.read(buffer, { ...options, type: 'array', bookFiles: withFiles });
  let styles = null;
  try {
    if (xls && withFiles) {
      const stream = XLSX.CFB.find(XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' }), 'Workbook')
        || XLSX.CFB.find(XLSX.CFB.read(new Uint8Array(buffer), { type: 'array' }), 'Book');
      if (stream) styles = readXlsStyles(Uint8Array.from(stream.content), maxRows);
    } else if (withFiles) styles = readCellStyles(wb, maxRows);
  } catch { styles = null; }
  if (xls) fillMissingTextResults(wb);
  return { SheetNames: wb.SheetNames, Sheets: wb.Sheets, Workbook: wb.Workbook, CellStyles: styles };
}

/**
 * Some programs save .xls files without the text result of formulas (Excel recalculates on
 * open). Those cells read as 0: work the formula out, and show it when the answer is text.
 */
function fillMissingTextResults(wb) {
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    for (const key of Object.keys(ws)) {
      if (key[0] === '!') continue;
      const cell = ws[key];
      if (!cell?.f || !(cell.t === 'n' && cell.v === 0)) continue;
      try {
        const v = evaluate(cell.f, name, wb);
        if (typeof v === 'string') { cell.t = 's'; cell.v = v; cell.w = v; }
        else if (typeof v === 'boolean') { cell.t = 'b'; cell.v = v; cell.w = v ? 'TRUE' : 'FALSE'; }
      } catch { /* leave the saved value */ }
    }
  }
}
