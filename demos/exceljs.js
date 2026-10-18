import { setup } from './shared.js';
import source from './exceljs.js?raw';

// #region demo
import ExcelJS from 'exceljs';

async function render(buffer, name, out) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer); // .xlsx only, but gives styles, fonts, fills

  workbook.eachSheet((sheet) => {
    const title = document.createElement('h3');
    title.textContent = `${sheet.name} (${sheet.rowCount} rows)`;

    const table = document.createElement('table');
    table.className = 'data';
    sheet.eachRow({ includeEmpty: false }, (row) => {
      const tr = table.insertRow();
      row.eachCell({ includeEmpty: true }, (cell) => {
        const td = tr.insertCell();
        td.textContent = cell.text;                       // formatted text, never HTML
        if (cell.font?.bold) td.style.fontWeight = '700'; // styles are available
      });
    });
    out.append(title, table);
  });
}
// #endregion

setup({ sample: 'regional-sales.xlsx', accept: '.xlsx', render, source });
