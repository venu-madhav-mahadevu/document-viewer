// Finishes the test pack: .docm/.xlsm content types, clean-up, TEST-CHECKLIST.xlsx.
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');
const OUT = path.join(__dirname, 'out');

async function setMainType(src, dest, from, to) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(OUT, src)));
  const types = await zip.file('[Content_Types].xml').async('string');
  if (!types.includes(from)) throw new Error(`${src}: content type not found`);
  zip.file('[Content_Types].xml', types.replace(from, to));
  fs.writeFileSync(path.join(OUT, dest), await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  fs.unlinkSync(path.join(OUT, src));
}

// Viewer status for the binary formats differs from their modern versions.
function adjust(t) {
  if (t.id.startsWith('DOC-')) {
    if (/picture|Shapes|Text box|Floating|Header and footer|Two-column|Equation/i.test(t.what)) t.status = '❌';
    else if (/Hyperlink|Merged cells|Landscape/i.test(t.what)) t.status = '⚠️';
  }
  if (t.id.startsWith('XLS-')) {
    if (/Text color|Bold|Italic|Underline|Strikethrough|Font size|Font Times|Font Courier|Bold \+|Superscript|Mixed formats|Border|Align|Vertical|Wrap text|Indent|Rotated|Bottom border/i.test(t.what)) t.status = '❌';
  }
  if (t.id.startsWith('PPT-') && /Gradient|Transparency|Cropped|border/i.test(t.what)) t.status = t.status || '⚠️';
  return t;
}

(async () => {
  await setMainType('test-word-docm.docx', 'test-word.docm', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', 'application/vnd.ms-word.document.macroEnabled.main+xml');
  await setMainType('test-excel-xlsm.xlsx', 'test-excel.xlsm', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
  for (const f of fs.readdirSync(OUT)) {
    if (/-for-(doc|xls|ppt)\.|\.(docx|xlsx|pptx)\.pdf$/.test(f)) fs.unlinkSync(path.join(OUT, f));
  }

  // collect tests in format order
  const order = ['word-docx', 'word-docm', 'word-doc', 'excel-xlsx', 'excel-xlsm', 'excel-xls', 'csv', 'slides-pptx', 'slides-ppt', 'pdf', 'images'];
  const all = order.flatMap((name) => JSON.parse(fs.readFileSync(path.join(__dirname, `tests-${name}.json`), 'utf8')).map(adjust));

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Docs Viewer test pack';
  const solid = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

  // How to use
  const how = wb.addWorksheet('How to use', { properties: { tabColor: { argb: 'FF1F4E79' } } });
  how.columns = [{ width: 120 }];
  [
    ['Docs Viewer — manual test checklist', { bold: true, size: 18, color: { argb: 'FF1F4E79' } }],
    [''],
    ['1. Open a test file from the "Test pack" section in the viewer (or the files in public/samples/test-pack/).'],
    ['2. Every test inside the file is labeled with its ID (e.g. DOCX-07), what is tested, and the expected result.'],
    ['3. Find the same ID in the "All tests" sheet, compare what you see with "Expected result", and pick Pass / Fail / N/A.'],
    ['4. Write what you saw in "Tester notes" for every Fail (a screenshot name helps), plus the device and browser.'],
    ['5. If unsure what is correct, open the same file in Microsoft Word / Excel / PowerPoint / Adobe Reader and compare.'],
    [''],
    ['"Viewer status" = what the current viewer is known to do:  ✅ should match the original   ⚠️ partly supported   ❌ not supported yet.'],
    ['Report ❌ items too: they show what is missing compared with the original file.'],
    [''],
    [`Total tests: ${all.length}.  Files: ${[...new Set(all.map((t) => t.file))].length}.`],
  ].forEach(([text, font]) => { const r = how.addRow([text]); if (font) r.getCell(1).font = font; r.getCell(1).alignment = { wrapText: true }; });

  // All tests
  const ws = wb.addWorksheet('All tests', { views: [{ state: 'frozen', ySplit: 1 }], properties: { tabColor: { argb: 'FF70AD47' } } });
  ws.columns = [
    { header: '#', width: 5 }, { header: 'Test ID', width: 11 }, { header: 'File', width: 30 }, { header: 'Where', width: 26 },
    { header: 'What is tested', width: 40 }, { header: 'Expected result', width: 70 }, { header: 'Viewer status', width: 13 },
    { header: 'Result (Pass / Fail / N/A)', width: 16 }, { header: 'Tester notes (what you saw, device, browser)', width: 45 },
  ];
  ws.getRow(1).eachCell((c) => { c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; c.fill = solid('FF1F4E79'); c.alignment = { vertical: 'middle', wrapText: true }; });
  ws.getRow(1).height = 32;
  all.forEach((t, i) => {
    const r = ws.addRow([i + 1, t.id, t.file, t.where, t.what, t.expected, t.status || '✅', '', '']);
    r.alignment = { vertical: 'top', wrapText: true };
    r.getCell(7).alignment = { horizontal: 'center', vertical: 'top' };
    r.getCell(8).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Pass,Fail,N/A"'] };
    r.getCell(8).fill = solid('FFFFFDE7');
  });
  const last = all.length + 1;
  ws.autoFilter = { from: 'A1', to: `I${last}` };
  ws.addConditionalFormatting({ ref: `H2:H${last}`, rules: [
    { type: 'cellIs', operator: 'equal', formulae: ['"Pass"'], style: { fill: solid('FFC6EFCE'), font: { color: { argb: 'FF006100' } } } },
    { type: 'cellIs', operator: 'equal', formulae: ['"Fail"'], style: { fill: solid('FFFFC7CE'), font: { color: { argb: 'FF9C0006' } } } },
  ] });

  // Summary per file
  const sum = wb.addWorksheet('Summary', { properties: { tabColor: { argb: 'FFFFC000' } } });
  sum.columns = [{ header: 'File', width: 32 }, { header: 'Tests', width: 8 }, { header: 'Pass', width: 8 }, { header: 'Fail', width: 8 }, { header: 'Not tested yet', width: 14 }, { header: 'Known ❌ (not supported)', width: 22 }];
  sum.getRow(1).eachCell((c) => { c.font = { bold: true }; c.fill = solid('FFD9E1F2'); });
  for (const file of [...new Set(all.map((t) => t.file))]) {
    const k = all.filter((t) => t.file === file).length;
    const rng = `'All tests'!$C$2:$C$${last}`;
    sum.addRow([file, k,
      { formula: `COUNTIFS(${rng},A${sum.rowCount + 1},'All tests'!$H$2:$H$${last},"Pass")`, result: 0 },
      { formula: `COUNTIFS(${rng},A${sum.rowCount + 1},'All tests'!$H$2:$H$${last},"Fail")`, result: 0 },
      { formula: `B${sum.rowCount + 1}-C${sum.rowCount + 1}-D${sum.rowCount + 1}`, result: k },
      all.filter((t) => t.file === file && t.status === '❌').length]);
  }
  await wb.xlsx.writeFile(path.join(OUT, 'TEST-CHECKLIST.xlsx'));
  fs.writeFileSync(path.join(__dirname, 'all-tests.json'), JSON.stringify(all, null, 1));
  console.log(`checklist: ${all.length} tests, ${[...new Set(all.map((t) => t.file))].length} files`);
  console.log(fs.readdirSync(OUT).join('\n'));
})();
