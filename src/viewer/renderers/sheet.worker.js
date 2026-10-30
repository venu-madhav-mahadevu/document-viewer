// Parses spreadsheets off the main thread so large files never freeze the page.
import { parseWorkbook } from './sheet-parse.js';

self.onmessage = ({ data }) => {
  const { buffer, isCsv, maxRows } = data;
  try {
    self.postMessage({ ok: true, wb: parseWorkbook(buffer, isCsv, maxRows) });
  } catch (err) {
    self.postMessage({ ok: false, message: err?.message || 'This spreadsheet could not be read.' });
  }
};
