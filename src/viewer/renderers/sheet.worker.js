// Parses spreadsheets off the main thread so large files never freeze the page.
import * as XLSX from 'xlsx';

self.onmessage = ({ data }) => {
  const { buffer, isCsv, maxRows } = data;
  try {
    const options = {
      sheetRows: maxRows, // stop reading after the rows we show: far faster on huge sheets
      cellDates: true,
      cellFormula: false,
      cellHTML: false,
    };
    const wb = isCsv
      ? XLSX.read(new TextDecoder().decode(buffer), { ...options, type: 'string', raw: true })
      : XLSX.read(buffer, { ...options, type: 'array' });
    self.postMessage({ ok: true, wb: { SheetNames: wb.SheetNames, Sheets: wb.Sheets, Workbook: wb.Workbook } });
  } catch (err) {
    self.postMessage({ ok: false, message: err?.message || 'This spreadsheet could not be read.' });
  }
};
