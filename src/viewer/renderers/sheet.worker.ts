// @ts-nocheck -- plain JavaScript kept as .ts so Angular bundles it as a Web Worker.
// Parses spreadsheets off the main thread so large files never freeze the page.
import * as XLSX from 'xlsx';
import { readCellStyles } from './xlsx-styles.js';

self.onmessage = ({ data }) => {
  const { buffer, isCsv, maxRows } = data;
  try {
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
    self.postMessage({ ok: true, wb: { SheetNames: wb.SheetNames, Sheets: wb.Sheets, Workbook: wb.Workbook, CellStyles: styles } });
  } catch (err) {
    self.postMessage({ ok: false, message: err?.message || 'This spreadsheet could not be read.' });
  }
};
