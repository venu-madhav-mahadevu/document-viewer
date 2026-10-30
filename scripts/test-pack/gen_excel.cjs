// Excel test workbook. usage: node gen_excel.cjs <PREFIX> <output.xlsx>   PREFIX = XLSX | XLSM | XLS
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const [PREFIX, OUTFILE] = process.argv.slice(2);
const FILE = { XLSX: 'test-excel.xlsx', XLSM: 'test-excel.xlsm', XLS: 'test-excel.xls' }[PREFIX];
const tests = [];
let n = 0;
const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
const thin = { style: 'thin', color: { argb: 'FFBFBFBF' } };
const box = (b) => ({ top: b, left: b, bottom: b, right: b });

/** One test row: A = ID + what, B..D = the test cells (filled by `fill`), E = expected. */
function testRow(ws, sheetName, what, expected, fill, status = '') {
  n += 1;
  const id = `${PREFIX}-${String(n).padStart(2, '0')}`;
  const row = ws.addRow([`${id} · ${what}`]);
  tests.push({ id, file: FILE, where: `Sheet "${sheetName}", row ${row.number}`, what, expected, status });
  row.getCell(1).font = { bold: true, color: { argb: 'FF1F4E79' } };
  row.getCell(1).alignment = { vertical: 'middle', wrapText: true };
  row.getCell(5).value = `Expected: ${expected}`;
  row.getCell(5).font = { italic: true, color: { argb: 'FF595959' }, size: 9 };
  row.getCell(5).alignment = { wrapText: true, vertical: 'middle' };
  fill(row);
  return row;
}
function sheet(wb, name, title) {
  const ws = wb.addWorksheet(name);
  ws.columns = [{ width: 38 }, { width: 30 }, { width: 22 }, { width: 22 }, { width: 60 }];
  ws.addRow([title]).getCell(1).font = { bold: true, size: 14, color: { argb: 'FF1F4E79' } };
  const h = ws.addRow(['Test (ID · what is tested)', 'Test cell', 'Test cell', 'Test cell', 'Expected result']);
  h.eachCell((c) => { c.font = { bold: true }; c.fill = solid('FFD9E1F2'); c.border = box(thin); });
  return ws;
}

(async () => {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Docs Viewer test pack';

  // ---------------- 1. Values and number formats
  let ws = sheet(wb, 'Values & formats', `${FILE} — values and number formats`);
  const fmt = (what, expected, value, numFmt, status) => testRow(ws, 'Values & formats', what, expected, (r) => {
    r.getCell(2).value = value; if (numFmt) r.getCell(2).numFmt = numFmt;
  }, status);
  if (PREFIX === 'XLSM') testRow(ws, 'Values & formats', 'Macro-enabled workbook (.xlsm)', 'Opens normally; no macro runs', () => {});
  if (PREFIX === 'XLS') testRow(ws, 'Values & formats', 'Excel 97-2003 workbook (.xls)', 'Opens normally with the values below', () => {});
  fmt('Plain text', 'Shows: Hello world', 'Hello world');
  fmt('Whole number', 'Shows: 42 (right-aligned)', 42);
  fmt('Decimal number', 'Shows: 1234.5678', 1234.5678);
  fmt('Thousands separator', 'Shows: 1,234,567.89', 1234567.891, '#,##0.00');
  fmt('Currency', 'Shows: $1,234.50', 1234.5, '$#,##0.00');
  fmt('Negative currency in red', 'Shows: -$1,234.50 in RED', -1234.5, '$#,##0.00;[Red]-$#,##0.00');
  fmt('Negative in brackets', 'Shows: (1,234.50)', -1234.5, '#,##0.00;(#,##0.00)');
  fmt('Accounting', 'Shows: $ 1,234.50 with the $ at the left edge', 1234.5, '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)');
  fmt('Percent', 'Shows: 45.67%', 0.4567, '0.00%');
  fmt('Fraction', 'Shows: 1 3/4', 1.75, '# ?/?');
  fmt('Scientific', 'Shows: 1.23E+08', 123456789, '0.00E+00');
  fmt('Date (m/d/yyyy)', 'Shows: 10/21/2026', new Date(Date.UTC(2026, 9, 21)), 'm/d/yyyy');
  fmt('Date (long)', 'Shows: Wednesday, October 21, 2026', new Date(Date.UTC(2026, 9, 21)), 'dddd, mmmm d, yyyy');
  fmt('Date and time', 'Shows: 2026-10-21 14:30', new Date(Date.UTC(2026, 9, 21, 14, 30)), 'yyyy-mm-dd hh:mm');
  fmt('Time', 'Shows: 9:45:30 AM', new Date(Date.UTC(1899, 11, 30, 9, 45, 30)), 'h:mm:ss AM/PM');
  fmt('Leading zeros (ID)', 'Shows: 000042 (zeros kept)', 42, '000000');
  fmt('Long number as text (account / card)', 'Shows all 16 digits: 1234567812345678 (not 1.23E+15)', '1234567812345678', '@');
  fmt('Phone number format', 'Shows: (408) 555-1234', 4085551234, '(###) ###-####');
  fmt('TRUE / FALSE', 'Shows: TRUE', true);
  fmt('Empty cell', 'Test cell is empty', null);
  testRow(ws, 'Values & formats', 'Long text overflow', 'The long sentence in the test cell stays on one line and runs into the empty cells to the right', (r) => {
    r.getCell(2).value = 'This is a long sentence that overflows into the next empty cells.';
  }, '⚠️');
  testRow(ws, 'Values & formats', 'Accented and Asian text', 'Shows: café  Zürich  日本語  中文  한국어', (r) => { r.getCell(2).value = 'café  Zürich  日本語  中文  한국어'; });

  // ---------------- 2. Colors and fonts
  ws = sheet(wb, 'Colors & fonts', `${FILE} — colors and fonts`);
  const S2 = 'Colors & fonts';
  for (const [name, argb, text] of [['Red', 'FFFF0000', 'FFFFFFFF'], ['Orange', 'FFFFC000', 'FF000000'], ['Yellow', 'FFFFFF00', 'FF000000'], ['Green', 'FF00B050', 'FFFFFFFF'], ['Blue', 'FF0070C0', 'FFFFFFFF'], ['Purple', 'FF7030A0', 'FFFFFFFF'], ['Black', 'FF000000', 'FFFFFFFF']]) {
    testRow(ws, S2, `Background color ${name}`, `Test cell background is ${name.toLowerCase()} with ${text === 'FFFFFFFF' ? 'white' : 'black'} text "${name}"`, (r) => {
      r.getCell(2).value = name; r.getCell(2).fill = solid(argb); r.getCell(2).font = { color: { argb: text } };
    });
  }
  testRow(ws, S2, 'Theme color with tint', 'Light blue (theme accent 1, 60% lighter) background', (r) => { r.getCell(2).value = 'Theme tint'; r.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 4, tint: 0.6 } }; });
  testRow(ws, S2, 'Pattern fill (stripes)', 'Diagonal stripe pattern in the cell background', (r) => { r.getCell(2).value = 'Pattern'; r.getCell(2).fill = { type: 'pattern', pattern: 'darkUp', fgColor: { argb: 'FF9BC2E6' }, bgColor: { argb: 'FFFFFFFF' } }; }, '⚠️');
  testRow(ws, S2, 'Gradient fill', 'Cell fades from white to blue', (r) => { r.getCell(2).value = 'Gradient'; r.getCell(2).fill = { type: 'gradient', gradient: 'angle', degree: 0, stops: [{ position: 0, color: { argb: 'FFFFFFFF' } }, { position: 1, color: { argb: 'FF0070C0' } }] }; }, '❌');
  for (const [what, font, value] of [['Text color red', { color: { argb: 'FFC00000' } }, 'Red text'], ['Text color green', { color: { argb: 'FF00B050' } }, 'Green text'],
    ['Bold', { bold: true }, 'Bold'], ['Italic', { italic: true }, 'Italic'], ['Underline', { underline: true }, 'Underline'], ['Double underline', { underline: 'double' }, 'Double underline'],
    ['Strikethrough', { strike: true }, 'Strikethrough'], ['Font size 16', { size: 16 }, 'Big 16 pt'], ['Font size 8', { size: 8 }, 'Small 8 pt'],
    ['Font Times New Roman', { name: 'Times New Roman' }, 'Times New Roman'], ['Font Courier New', { name: 'Courier New' }, 'Courier New'], ['Bold + italic + color', { bold: true, italic: true, color: { argb: 'FF7030A0' } }, 'Purple bold italic']]) {
    testRow(ws, S2, what, `Test cell text: "${value}" looks ${what.toLowerCase()}`, (r) => { r.getCell(2).value = value; r.getCell(2).font = font; });
  }
  testRow(ws, S2, 'Superscript / subscript', 'Shows m² and H₂O with small raised / lowered digits', (r) => {
    r.getCell(2).value = { richText: [{ text: 'm' }, { text: '2', font: { vertAlign: 'superscript' } }, { text: '  H' }, { text: '2', font: { vertAlign: 'subscript' } }, { text: 'O' }] };
  }, '❌');
  testRow(ws, S2, 'Mixed formats in one cell (rich text)', 'One cell with "Red" in red, "bold" in bold and "blue" in blue', (r) => {
    r.getCell(2).value = { richText: [{ text: 'Red ', font: { color: { argb: 'FFFF0000' } } }, { text: 'bold ', font: { bold: true } }, { text: 'blue', font: { color: { argb: 'FF0070C0' } } }] };
  }, '⚠️');

  // ---------------- 3. Alignment, borders, layout
  ws = sheet(wb, 'Alignment & borders', `${FILE} — alignment, borders, merged cells`);
  const S3 = 'Alignment & borders';
  for (const [what, al] of [['Align left', { horizontal: 'left' }], ['Align center', { horizontal: 'center' }], ['Align right', { horizontal: 'right' }]]) {
    testRow(ws, S3, what, `"${what.split(' ')[1]}" sits at the ${what.split(' ')[1].toLowerCase()} of the cell`, (r) => { r.getCell(2).value = what.split(' ')[1]; r.getCell(2).alignment = al; r.getCell(2).border = box(thin); });
  }
  for (const [what, v] of [['Vertical top', 'top'], ['Vertical middle', 'middle'], ['Vertical bottom', 'bottom']]) {
    testRow(ws, S3, what, `In this tall row, the text sits at the ${v}`, (r) => { r.height = 45; r.getCell(2).value = v; r.getCell(2).alignment = { vertical: v }; r.getCell(2).border = box(thin); });
  }
  testRow(ws, S3, 'Wrap text', 'The long text wraps onto several lines inside the cell; the row is taller', (r) => { r.height = 60; r.getCell(2).value = 'This long text wraps inside the cell instead of overflowing.'; r.getCell(2).alignment = { wrapText: true, vertical: 'top' }; r.getCell(2).border = box(thin); });
  testRow(ws, S3, 'Indent', 'Text starts further from the left edge', (r) => { r.getCell(2).value = 'Indented'; r.getCell(2).alignment = { indent: 3 }; r.getCell(2).border = box(thin); });
  testRow(ws, S3, 'Rotated text (45°)', 'Text is written diagonally', (r) => { r.height = 50; r.getCell(2).value = 'Rotated'; r.getCell(2).alignment = { textRotation: 45 }; }, '❌');
  for (const [what, b] of [['Border thin', { style: 'thin' }], ['Border medium', { style: 'medium' }], ['Border thick', { style: 'thick' }], ['Border dashed', { style: 'dashed' }], ['Border dotted', { style: 'dotted' }], ['Border double', { style: 'double' }], ['Border red', { style: 'medium', color: { argb: 'FFFF0000' } }]]) {
    testRow(ws, S3, what, `Test cell has a ${what.split(' ')[1]} border on all four sides`, (r) => { r.getCell(2).value = what.split(' ')[1]; r.getCell(2).border = box(b); r.getCell(2).alignment = { horizontal: 'center' }; });
  }
  testRow(ws, S3, 'Bottom border only', 'Only a blue line under the test cell', (r) => { r.getCell(2).value = 'Bottom'; r.getCell(2).border = { bottom: { style: 'medium', color: { argb: 'FF0070C0' } } }; });
  testRow(ws, S3, 'Merged cells (across)', 'Cells B to D are one merged cell with centered text', (r) => {
    r.getCell(2).value = 'Merged across B:D'; ws.mergeCells(r.number, 2, r.number, 4); r.getCell(2).alignment = { horizontal: 'center' }; r.getCell(2).fill = solid('FFFFF2CC'); r.getCell(2).border = box(thin);
  });
  const mr = testRow(ws, S3, 'Merged cells (down)', 'The test cell spans this row and the next one', (r) => { r.getCell(2).value = 'Merged down 2 rows'; r.getCell(2).alignment = { vertical: 'middle' }; r.getCell(2).fill = solid('FFE2EFDA'); });
  ws.addRow([]); ws.mergeCells(mr.number, 2, mr.number + 1, 2);
  testRow(ws, S3, 'Column widths', 'Column A is wide (tests), E is widest (expected), B-D medium', () => {});
  const hid = testRow(ws, S3, 'Hidden row', 'The next row (with "HIDDEN ROW") is not shown', () => {});
  const hr = ws.addRow(['HIDDEN ROW - you should not see this']); hr.hidden = true;
  testRow(ws, S3, 'Hidden column', 'Column F ("HIDDEN COLUMN") is not shown', (r) => { r.getCell(6).value = 'HIDDEN COLUMN'; });
  ws.getColumn(6).hidden = true;

  // ---------------- 4. Formulas
  ws = sheet(wb, 'Formulas', `${FILE} — formulas (results saved in the file)`);
  const S4 = 'Formulas';
  const base = ws.addRow(['Input numbers', 10, 20, 30]);
  const br = base.number;
  for (const [what, formula, result, show] of [['SUM', `SUM(B${br}:D${br})`, 60, '60'], ['AVERAGE', `AVERAGE(B${br}:D${br})`, 20, '20'], ['MAX', `MAX(B${br}:D${br})`, 30, '30'],
    ['IF', `IF(B${br}>5,"Big","Small")`, 'Big', 'Big'], ['Text join', `CONCATENATE("Total: ",SUM(B${br}:D${br}))`, 'Total: 60', 'Total: 60'],
    ['Percentage', `B${br}/D${br}`, 1 / 3, '33.3%'], ['Cross-sheet reference', "'Values & formats'!B4", 'Hello world', 'Hello world']]) {
    testRow(ws, S4, `Formula ${what}`, `Shows the result ${show} (not the formula text)`, (r) => {
      r.getCell(2).value = { formula, result }; if (what === 'Percentage') r.getCell(2).numFmt = '0.0%';
      r.getCell(3).value = `=${formula}`; r.getCell(3).font = { name: 'Courier New', size: 9, color: { argb: 'FF808080' } };
    });
  }
  testRow(ws, S4, 'Error value', 'Shows #DIV/0!', (r) => { r.getCell(2).value = { formula: `B${br}/0`, result: { error: '#DIV/0!' } }; });

  // ---------------- 5. Conditional formatting
  ws = sheet(wb, 'Conditional formatting', `${FILE} — conditional formatting (colors from rules)`);
  const S5 = 'Conditional formatting';
  const cf = (what, expected, rule, values, status = '❌') => {
    const r = testRow(ws, S5, what, expected, (row) => values.forEach((v, i) => { row.getCell(2 + i).value = v; }), status);
    ws.addConditionalFormatting({ ref: `B${r.number}:D${r.number}`, rules: [rule] });
  };
  cf('Color scale (red - yellow - green)', 'B is red (low), C yellow (middle), D green (high)', { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }] }, [10, 50, 90]);
  cf('Data bars', 'Blue bars inside the cells, longest in D', { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }, [20, 50, 100]);
  cf('Icon set (arrows)', 'Red down arrow in B, yellow side arrow in C, green up arrow in D', { type: 'iconSet', iconSet: '3Arrows', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }] }, [10, 50, 90]);
  cf('Highlight values greater than 50', 'Only D (90) has a light red background with dark red text', { type: 'cellIs', operator: 'greaterThan', formulae: [50], style: { fill: solid('FFFFC7CE'), font: { color: { argb: 'FF9C0006' } } } }, [10, 50, 90]);
  cf('Highlight duplicate values', 'B and C (both "Apple") are highlighted', { type: 'duplicateValues', style: { fill: solid('FFFFEB9C') } }, ['Apple', 'Apple', 'Pear']);

  // ---------------- 6. Notes, links, picture, filter, frozen rows
  ws = sheet(wb, 'Notes & links', `${FILE} — notes, links, picture, filter, frozen rows`);
  const S6 = 'Notes & links';
  ws.views = [{ state: 'frozen', ySplit: 2 }];
  testRow(ws, S6, 'Frozen top rows', 'Scroll down: the title and the header row stay at the top', () => {}, '❌');
  testRow(ws, S6, 'Cell note / comment', 'Test cell has a red corner mark; hovering shows "This is a cell note"', (r) => { r.getCell(2).value = 'Has a note'; r.getCell(2).note = 'This is a cell note'; }, '❌');
  testRow(ws, S6, 'Hyperlink to a website', 'Blue underlined "example.com"; click opens https://example.com', (r) => { r.getCell(2).value = { text: 'example.com', hyperlink: 'https://example.com' }; r.getCell(2).font = { color: { argb: 'FF0563C1' }, underline: true }; }, '⚠️');
  testRow(ws, S6, 'Link to another sheet', 'Click "Go to Formulas" jumps to the Formulas sheet', (r) => { r.getCell(2).value = { text: 'Go to Formulas', hyperlink: "#'Formulas'!A1" }; r.getCell(2).font = { color: { argb: 'FF0563C1' }, underline: true }; }, '❌');
  const pr = testRow(ws, S6, 'Picture in the sheet', 'The "Embedded picture" test chart appears over columns B-D below this row', (r) => { r.height = 20; }, '❌');
  const imgId = wb.addImage({ filename: path.join(__dirname, 'assets', 'picture.png'), extension: 'png' });
  ws.addImage(imgId, { tl: { col: 1, row: pr.number }, ext: { width: 360, height: 240 } });
  for (let i = 0; i < 12; i++) ws.addRow([]);
  const fr = testRow(ws, S6, 'Filter buttons', 'The header row of the small table below has filter drop-down buttons', () => {}, '❌');
  const th = ws.addRow(['', 'Name', 'City', 'Amount']);
  [['', 'Alpha', 'Fremont', 120], ['', 'Beta', 'San Jose', 80], ['', 'Gamma', 'Oakland', 150]].forEach((v) => ws.addRow(v));
  ws.autoFilter = { from: { row: th.number, column: 2 }, to: { row: th.number + 3, column: 4 } };
  testRow(ws, S6, 'Drop-down list (data validation)', 'Test cell shows "Open"; in Excel it has a drop-down with Open / Closed', (r) => { r.getCell(2).value = 'Open'; r.getCell(2).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Open,Closed"'] }; }, '⚠️');
  for (let i = 0; i < 40; i++) ws.addRow([`Filler row ${i + 1} (scroll to test frozen rows)`]);

  // ---------------- 7. Hidden sheet + sheet tab color
  const hidden = wb.addWorksheet('Hidden sheet', { state: 'hidden' });
  hidden.addRow(['HIDDEN SHEET - you should not see this sheet as a tab']);
  n += 1; tests.push({ id: `${PREFIX}-${String(n).padStart(2, '0')}`, file: FILE, where: 'Sheet tabs', what: 'Hidden sheet', expected: 'There are 6 sheet tabs; "Hidden sheet" is not one of them', status: '' });
  wb.getWorksheet('Values & formats').properties.tabColor = { argb: 'FF1F4E79' };
  n += 1; tests.push({ id: `${PREFIX}-${String(n).padStart(2, '0')}`, file: FILE, where: 'Sheet tabs', what: 'Sheet tabs and switching', expected: 'Tabs: Values & formats, Colors & fonts, Alignment & borders, Formulas, Conditional formatting, Notes & links; clicking a tab shows that sheet', status: '' });

  await wb.xlsx.writeFile(OUTFILE);
  fs.writeFileSync(path.join(__dirname, `tests-excel-${PREFIX.toLowerCase()}.json`), JSON.stringify(tests, null, 1));
  console.log(`excel-${PREFIX.toLowerCase()}: ${tests.length} tests`);
})();
