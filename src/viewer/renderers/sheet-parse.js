// Spreadsheet parsing, shared by the Web Worker (normal case) and the page itself
// (fallback when a worker cannot run, e.g. blocked by a security policy).
import * as XLSX from 'xlsx';
import { readCellStyles } from './xlsx-styles.js';

export function parseWorkbook(buffer, isCsv, maxRows) {
  const options = {
    sheetRows: maxRows, // stop reading after the rows we show: far faster on huge sheets
    cellDates: true,
    cellFormula: false,
    cellHTML: false,
    cellStyles: true, // cell background colors (fills)
  };
  // Raw workbook files (for fonts, alignment and borders) only for normal-size files.
  const withFiles = !isCsv && buffer.byteLength < 25 * 1024 * 1024;
  const wb = isCsv
    ? XLSX.read(new TextDecoder().decode(buffer), { ...options, type: 'string', raw: true })
    : XLSX.read(buffer, { ...options, type: 'array', bookFiles: withFiles });
  let styles = null;
  try { if (withFiles) styles = readCellStyles(wb, maxRows); } catch { styles = null; }
  return { SheetNames: wb.SheetNames, Sheets: wb.Sheets, Workbook: wb.Workbook, CellStyles: styles };
}
