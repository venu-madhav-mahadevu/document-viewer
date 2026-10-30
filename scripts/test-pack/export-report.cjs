// Exports TEST-CHECKLIST.xlsx to public/test-report/data/test-results.json for the test report page.
// Run after the checklist results change: node scripts/test-pack/export-report.cjs
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const ROOT = path.join(__dirname, '..', '..');
const CHECKLIST = path.join(ROOT, 'public', 'samples', 'test-pack', 'TEST-CHECKLIST.xlsx');
const OUT = path.join(ROOT, 'public', 'test-report', 'data', 'test-results.json');

const FORMAT = { jpg: 'Image', jpeg: 'Image', tif: 'TIFF', tiff: 'TIFF' };
const FAMILY = { DOCX: 'Word', DOCM: 'Word', DOC: 'Word', XLSX: 'Excel', XLSM: 'Excel', XLS: 'Excel', CSV: 'CSV',
  PPTX: 'PowerPoint', PPT: 'PowerPoint', PDF: 'PDF', Image: 'Images', TIFF: 'Images' };

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(CHECKLIST);
  const ws = wb.getWorksheet('All tests');
  const rows = [];
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const v = (i) => (row.getCell(i).value == null ? '' : String(row.getCell(i).value));
    const ext = v(3).split('.').pop().toLowerCase();
    const format = FORMAT[ext] || ext.toUpperCase();
    rows.push({ id: v(2), file: v(3), format, family: FAMILY[format], where: v(4), feature: v(5), expected: v(6), result: v(8), seen: v(9) });
  });
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT, JSON.stringify(rows));
  console.log(`${rows.length} tests -> ${path.relative(ROOT, OUT)}`);
})();
