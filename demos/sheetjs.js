import { setup } from './shared.js';
import source from './sheetjs.js?raw';

// #region demo
import * as XLSX from 'xlsx';

async function render(buffer, name, out) {
  // One call reads .xlsx .xlsm .xlsb .xls .ods .csv .tsv
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  const tabs = document.createElement('div');
  tabs.className = 'tabs';
  const sheetArea = document.createElement('div');
  out.append(tabs, sheetArea);

  function showSheet(sheetName) {
    const html = XLSX.utils.sheet_to_html(workbook.Sheets[sheetName], { header: '', footer: '' });
    sheetArea.innerHTML = html; // SheetJS HTML-escapes every cell value
    sheetArea.querySelector('table').className = 'data';
    tabs.querySelectorAll('button').forEach((b) => b.setAttribute('aria-pressed', String(b.textContent === sheetName)));
  }

  for (const sheetName of workbook.SheetNames) {
    const button = document.createElement('button');
    button.textContent = sheetName;
    button.onclick = () => showSheet(sheetName);
    tabs.append(button);
  }
  showSheet(workbook.SheetNames[0]);
}
// #endregion

setup({ sample: 'regional-sales.xlsx', accept: '.xlsx,.xlsm,.xlsb,.xls,.ods,.csv,.tsv', render, source });
