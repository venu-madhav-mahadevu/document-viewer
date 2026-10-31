// Writes the test results (Pass / Fail + what was seen) into TEST-CHECKLIST.xlsx.
// usage: node fill-checklist.cjs <TEST-CHECKLIST.xlsx>
const ExcelJS = require('exceljs');
const file = process.argv[2];
const ENV = 'Stage + Angular, Chrome 1400 px (spot checks Firefox, Safari iPhone), ' + (process.env.RUN_DATE || '30 Oct 2026');

// Word (.docx / .docm share the same tests; key = what is tested)
const word = {
  'Font families': ['Pass', 'Each line in its font (Calibri/Cambria shown with look-alike fonts Carlito/Caladea when not installed).'],
  'Font sizes': ['Pass', '8 to 28 pt, growing.'],
  'Text styles': ['Pass', 'All 9 styles correct, incl. double underline, x², H₂O, small caps.'],
  'Text colors and highlight': ['Pass', 'Colors, yellow / green highlight and light-blue paragraph shading shown.'],
  'Alignment': ['Pass', 'Left, centered, right; justified paragraph has a straight right edge.'],
  'Indents and line spacing': ['Pass', '1 inch indent, first-line and hanging indent, 1.0 / 1.5 / 2.0 spacing.'],
  'Headings and styles': ['Pass', 'Title with underline rule, Heading 1-3 smaller each time, theme blue.'],
  'Round bullets (Symbol font)': ['Pass', '• bullets.'],
  'Square bullets (Wingdings §)': ['Pass', '▪ bullets (not §).'],
  'Arrow bullets (Wingdings Ø)': ['Pass', '➢ bullets (not Ø).'],
  'Check-mark bullets (Wingdings ü)': ['Pass', '✓ bullets (not ü).'],
  'Hollow circle bullets (Courier New o)': ['Pass', 'o bullets.'],
  'Dash bullets': ['Pass', '– bullets.'],
  'Numbered list': ['Pass', '1. 2. 3.'],
  'Multi-level list (1 / a / i / bullet)': ['Pass', '1. / a) / i. / ▪, indented more each level.'],
  'Table borders, header shading, column widths': ['Pass', 'Dark blue header with white bold text, borders, numbers right-aligned.'],
  'Merged cells and cell colors': ['Pass', 'Merged top row, "Merged down" spans 2 rows, green/yellow/red cells.'],
  'Inline picture': ['Pass', 'Picture shown about 4 inches wide.'],
  'Shapes and arrows over a picture': ['Pass', 'Red arrow at the gray ramp, yellow "Look here" box top right, green circle at the ramp end, blue arrow line below the fine lines (fixed 31 Oct 2026).'],
  'Text box': ['Pass', 'Light gray box with a dark border and two lines of text, right of the paragraph (fixed 31 Oct 2026).'],
  'Floating picture with text wrap': ['Pass', 'Picture 2 on the right; the text wraps on its left (fixed 31 Oct 2026).'],
  'Hyperlink': ['Pass', 'Blue underlined; opens https://example.com in a new tab.'],
  'Equation (math)': ['Pass', 'Shown as math: c² raised, a over b stacked (slightly smaller than in Word).'],
  'Page break': ['Pass', 'Section 5 starts on page 5. Note: page 1 is taller than a real page (long pages are not split in two like Word does).'],
  'Header and footer': ['Pass', 'Header on all 7 pages (pages 6-7 take it over from page 1, as in Word); footer shows the right number on each page: Page 1 … Page 7. Fixed 30 Oct 2026.'],
  'Two-column layout': ['Pass', 'Text flows in two columns.'],
  'Landscape page in a portrait document': ['Pass', 'Page 7 is landscape (wider than tall), pages 1-6 portrait.'],
  'Macro-enabled document (.docm)': ['Pass', 'Opens like the .docx; no macro runs, no warning.'],
};
const doc = {
  'Word 97-2003 binary document (.doc)': ['Pass', 'Opens, 11 pages (the same pages as LibreOffice shows).'],
  'Text styles': ['Pass', 'All styles correct, incl. double underline (fixed 31 Oct 2026).'],
  'Text colors and highlight': ['Pass', 'Text colors, highlights and the light-blue paragraph background (fixed 31 Oct 2026).'],
  'Table borders, header shading, column widths': ['Pass', 'Dark blue header row with white bold Region/Q1/Q2/Total, borders, numbers right-aligned (fixed 31 Oct 2026).'],
  'Merged cells and cell colors': ['Pass', 'Top row merged across, "Merged down" spans 2 rows, green/yellow/red cells (fixed 31 Oct 2026).'],
  'Inline picture': ['Pass', 'Test chart shown 4 inches wide (fixed 31 Oct 2026).'],
  'Shapes and arrows over a picture': ['Pass', 'Red right-arrow at the gray ramp, yellow rounded "Look here" box top right, green circle outline at the ramp end, blue line with arrow head below (fixed 31 Oct 2026).'],
  'Text box': ['Pass', 'Light gray box with a dark border and two lines of text, right of the paragraph (fixed 31 Oct 2026).'],
  'Floating picture with text wrap': ['Pass', 'Purple Picture 2 on the right; the text wraps on its left (fixed 31 Oct 2026).'],
  'Hyperlink': ['Pass', 'Blue underlined "Open example.com"; clicking opens https://example.com in a new tab (fixed 31 Oct 2026).'],
  'Equation (math)': ['Pass', 'Shown as math: E = mc² with the 2 raised, a over b stacked (fixed 31 Oct 2026).'],
  'Page break': ['Pass', 'Each section starts on a new page, page for page the same as LibreOffice (DOC-25 on page 9 of 11). Note: the test text says "page 5", which assumes every section fits on one page; sections 1, 2 and 4 are longer (fixed 31 Oct 2026).'],
  'Header and footer': ['Pass', 'Header "Docs Viewer test pack — header" on every page; footer "Page N" with the right number on each of the 11 pages (fixed 31 Oct 2026).'],
  'Two-column layout': ['Pass', 'Text flows in two columns side by side (fixed 31 Oct 2026).'],
  'Landscape page in a portrait document': ['Pass', 'Last page is landscape, the pages before it portrait (fixed 31 Oct 2026).'],
};

// Excel (.xlsx / .xlsm / .xls share tests; key = what is tested)
const excel = {
  'Plain text': ['Pass', 'Hello world'], 'Whole number': ['Pass', '42, right-aligned'], 'Decimal number': ['Pass', '1234.5678'],
  'Thousands separator': ['Pass', '1,234,567.89'], 'Currency': ['Pass', '$1,234.50'], 'Negative currency in red': ['Pass', '-$1,234.50 in red'],
  'Negative in brackets': ['Pass', '(1,234.50)'], 'Accounting': ['Pass', '$ at the left edge, 1,234.50 at the right (fixed 30 Oct 2026).'],
  'Percent': ['Pass', '45.67%'], 'Fraction': ['Pass', '1 3/4'], 'Scientific': ['Pass', '1.23E+08'],
  'Date (m/d/yyyy)': ['Pass', '10/21/2026 (left-aligned; Excel right-aligns dates)'], 'Date (long)': ['Pass', 'Wednesday, October 21, 2026 (left-aligned)'],
  'Date and time': ['Pass', '2026-10-21 14:30 (left-aligned)'], 'Time': ['Pass', '9:45:30 AM (left-aligned)'],
  'Leading zeros (ID)': ['Pass', '000042'], 'Long number as text (account / card)': ['Pass', 'All 16 digits'],
  'Phone number format': ['Pass', '(408) 555-1234'], 'TRUE / FALSE': ['Pass', 'TRUE (left-aligned; Excel centers it)'], 'Empty cell': ['Pass', 'Empty'],
  'Long text overflow': ['Pass', 'One line, runs on into the empty cells to the right (fixed in this round).'],
  'Accented and Asian text': ['Pass', 'café Zürich 日本語 中文 한국어'],
  'Background color Red': ['Pass', 'Red, white text'], 'Background color Orange': ['Pass', 'Orange, black text'], 'Background color Yellow': ['Pass', 'Yellow, black text'],
  'Background color Green': ['Pass', 'Green, white text'], 'Background color Blue': ['Pass', 'Blue, white text'], 'Background color Purple': ['Pass', 'Purple, white text'],
  'Background color Black': ['Pass', 'Black, white text'], 'Theme color with tint': ['Pass', 'Light blue'],
  'Pattern fill (stripes)': ['Pass', 'Diagonal light-blue stripes (fixed 31 Oct 2026).'], 'Gradient fill': ['Pass', 'Fades from white to blue (fixed 31 Oct 2026).'],
  'Text color red': ['Pass', 'Red'], 'Text color green': ['Pass', 'Green'], 'Bold': ['Pass', 'Bold'], 'Italic': ['Pass', 'Italic'], 'Underline': ['Pass', 'Underlined'],
  'Double underline': ['Pass', 'Double underline (fixed 30 Oct 2026).'], 'Strikethrough': ['Pass', 'Struck through'], 'Font size 16': ['Pass', 'Big'], 'Font size 8': ['Pass', 'Small'],
  'Font Times New Roman': ['Pass', 'Times'], 'Font Courier New': ['Pass', 'Courier'], 'Bold + italic + color': ['Pass', 'Purple bold italic'],
  'Superscript / subscript': ['Pass', 'm² and H₂O with raised / lowered digits (fixed 30 Oct 2026).'], 'Mixed formats in one cell (rich text)': ['Pass', '"Red" in red, "bold" in bold, "blue" in blue (fixed 30 Oct 2026).'],
  'Align left': ['Pass', 'Left'], 'Align center': ['Pass', 'Centered'], 'Align right': ['Pass', 'Right'],
  'Vertical top': ['Pass', 'Top'], 'Vertical middle': ['Pass', 'Middle'], 'Vertical bottom': ['Pass', 'Bottom'],
  'Wrap text': ['Pass', 'Wraps onto 3 lines, row taller'], 'Indent': ['Pass', 'Indented'], 'Rotated text (45°)': ['Pass', 'Written diagonally at 45° (fixed 30 Oct 2026).'],
  'Border thin': ['Pass', 'Thin'], 'Border medium': ['Pass', 'Medium'], 'Border thick': ['Pass', 'Thick'], 'Border dashed': ['Pass', 'Dashed'],
  'Border dotted': ['Pass', 'Dotted'], 'Border double': ['Pass', 'Double'], 'Border red': ['Pass', 'Red'], 'Bottom border only': ['Pass', 'Blue line under the cell only'],
  'Merged cells (across)': ['Pass', 'B-D merged, centered'], 'Merged cells (down)': ['Pass', 'Spans 2 rows'],
  'Column widths': ['Pass', 'A wide, E widest, B-D medium (Excel column widths now kept exactly).'],
  'Hidden row': ['Pass', 'Not shown (fixed in this round; it was visible before).'], 'Hidden column': ['Pass', 'Column F not shown (fixed in this round; it was visible before).'],
  'Formula SUM': ['Pass', '60'], 'Formula AVERAGE': ['Pass', '20'], 'Formula MAX': ['Pass', '30'], 'Formula IF': ['Pass', 'Big'],
  'Formula Text join': ['Pass', 'Total: 60'], 'Formula Percentage': ['Pass', '33.3%'], 'Formula Cross-sheet reference': ['Pass', 'Hello world'], 'Error value': ['Pass', '#DIV/0!'],
  'Color scale (red - yellow - green)': ['Pass', 'B red, C yellow, D green (fixed 31 Oct 2026).'],
  'Data bars': ['Pass', 'Blue bars, longest in D (fixed 31 Oct 2026).'],
  'Icon set (arrows)': ['Pass', 'Red down arrow in B, yellow side arrow in C, green up arrow in D (fixed 31 Oct 2026).'],
  'Highlight values greater than 50': ['Pass', 'Only D (90) light red with dark red text (fixed 31 Oct 2026).'],
  'Highlight duplicate values': ['Pass', 'B and C ("Apple") highlighted (fixed 31 Oct 2026; the test file did not contain this rule before, see the 31 Oct run).'],
  'Frozen top rows': ['Pass', 'Title and header row stay on top while scrolling (fixed 30 Oct 2026).'],
  'Cell note / comment': ['Pass', 'Red corner mark; hovering shows "This is a cell note" (fixed 30 Oct 2026).'],
  'Hyperlink to a website': ['Pass', 'Blue underlined; opens https://example.com in a new tab (fixed 30 Oct 2026).'], 'Link to another sheet': ['Pass', 'Clicking "Go to Formulas" opens the Formulas sheet (fixed 30 Oct 2026).'],
  'Picture in the sheet': ['Pass', 'The test chart shows over columns B-D below row 7 (fixed 31 Oct 2026).'], 'Filter buttons': ['Pass', 'Name, City and Amount have filter buttons (shown only: the viewer is read-only) (fixed 31 Oct 2026).'],
  'Drop-down list (data validation)': ['Pass', 'Shows "Open" (view only, no drop-down - expected for a viewer).'],
  'Hidden sheet': ['Pass', '6 tabs; "Hidden sheet" not listed.'], 'Sheet tabs and switching': ['Pass', 'All 6 tabs in order; clicking shows that sheet.'],
  'Macro-enabled workbook (.xlsm)': ['Pass', 'Opens like the .xlsx; no macro runs, no warning.'],
  'Excel 97-2003 workbook (.xls)': ['Pass', 'Opens, 6 sheets.'],
};
const XLS_NA = 'This .xls file does not contain it: the test file was saved as .xls by LibreOffice, which left it out (Excel 97-2003 files can only hold it through later extensions). Checked in the records of the file, 31 Oct 2026.';
const xls = {
  'Pattern fill (stripes)': ['N/A', 'Plain light-blue fill: the file stores a solid color for this cell. ' + XLS_NA],
  'Gradient fill': ['N/A', 'Plain blue fill: the file stores a solid color for this cell. ' + XLS_NA],
  'Color scale (red - yellow - green)': ['N/A', 'No colors: the rule is missing. ' + XLS_NA],
  'Data bars': ['N/A', 'No bars: the rule is missing. ' + XLS_NA],
  'Icon set (arrows)': ['N/A', 'No arrows: the rule is missing. ' + XLS_NA],
  'Filter buttons': ['N/A', 'No filter buttons: the filter range is missing. ' + XLS_NA],
  'Formula IF': ['Pass', 'Big (the file has no saved text result; worked out from the formula, as Excel does when it opens the file) (fixed 31 Oct 2026).'],
  'Formula Text join': ['Pass', 'Total: 60 (worked out from the formula) (fixed 31 Oct 2026).'],
  'Formula Cross-sheet reference': ['Pass', 'Hello world (worked out from the formula) (fixed 31 Oct 2026).'],
  'Merged cells (down)': ['Pass', 'Spans 2 rows (shaded green from the file).'],
  'Merged cells (across)': ['Pass', 'B-D merged, text centered (fixed 31 Oct 2026).'],
  'Rotated text (45°)': ['Pass', 'Written diagonally at 45° (fixed 31 Oct 2026).'],
};

const csv = {
  'Plain text': 'Hello world', 'Comma inside quotes': 'Smith, John in one cell', 'Quotes inside quotes': 'He said "hello"',
  'Line break inside a cell': 'Two lines in one cell (fixed in this round; it was one line before).', 'Accented characters (UTF-8)': 'café naïve Zürich São Paulo, no garbling',
  'Asian text': '日本語 中文 한국어', 'Emoji': 'Done ✅ Thanks 🙂', 'Leading zeros': '00042 (zeros kept)', 'Long number (account / card)': 'All 16 digits',
  'Decimal number': '1234.56', 'Negative number': '-75.25', 'Currency text': '$1,234.50', 'Percent text': '45%', 'ISO date': '2026-10-21',
  'US date': '10/21/2026', 'Empty value': 'Empty', 'Spaces at start and end': 'padded (spaces kept in the cell)',
  'Very long text': 'Column capped; text ends with … and the full text shows on hover (fixed in this round; it stretched the table before).',
  'Formula-looking text (safety)': '=1+1 shown as text', 'UTF-8 file with byte-order mark (Excel "CSV UTF-8")': 'First cell reads "Test ID"',
  'Table view': '4 columns, 19 data rows + header (file has 19 tests; the expected "20 rows" counts the header)',
};

const slidesCommon = {
  'Title and subtitle': ['Pass', 'Large title, smaller gray subtitle below it (fixed 31 Oct 2026: the test file saved the titles with zero height).'],
  'Slide count and order': ['Pass', 'Page bar shows 10; 10 thumbnails in order.'],
  '16:9 slide size': ['Pass', 'Wide 16:9 slides.'],
  'Text colors': ['Pass', 'Red, blue, green.'], 'Font sizes': ['Pass', 'Small 12 pt small, Large 32 pt big.'], 'Font families': ['Pass', 'Times, Courier, Georgia.'],
  'Round bullet': ['Pass', '•'], 'Wingdings §': ['Pass', '▪'], 'Wingdings q': ['Pass', '❑'], 'Wingdings n': ['Pass', '■'],
  'Wingdings Ø': ['Pass', '➢'], 'Wingdings ü': ['Pass', '✓'], 'Wingdings v': ['Pass', '❖'],
  'Multi-level indents': ['Pass', 'Levels 1-4 indented more and smaller.'],
  'Shrink text on overflow': ['Pass', 'All 14 lines above the red line.'],
  'Basic shapes': ['Pass', 'Rectangle, rounded box, circle and triangle with their colors and labels (fixed 31 Oct 2026).'],
  'Arrows': ['Pass', 'Right, left-right and up arrows have arrow shapes (fixed 31 Oct 2026).'],
  'Other shapes': ['Pass', 'Star, callout with its pointer, chevron, pentagon and hexagon have their real outlines (fixed 31 Oct 2026).'],
  'Rotation': ['Pass', 'Tilted 30°.'], 'Outline only': ['Pass', 'Red circle outline, no fill.'],
  'Connector with arrow head': ['Pass', 'Blue line with an arrow head on the right (fixed 31 Oct 2026).'],
  'Picture': ['Pass', 'Whole chart, top-left.'],
  'Grouped shapes': ['Pass', 'Group 1-3 overlapping, stepping down.'],
  'Transparency': ['Pass', 'Red circle is see-through; TEXT BEHIND visible.'],
  'Column chart': ['Pass', 'Columns for Q1-Q4 with gridlines and a West / East legend, in the deck theme colors (fixed 31 Oct 2026).'],
  'Pie chart': ['Pass', 'Pie with Deposits 55%, Loans 35%, Other 10% and a legend (fixed 31 Oct 2026).'],
  'Slide background color': ['Pass', 'Light yellow background.'],
  'Speaker notes': ['Pass', '"Speaker notes · slide 10" under the slide shows the notes (fixed 31 Oct 2026).'],
  'Table header': ['Pass', 'Blue header row, white text.'], 'Banded rows': ['Pass', 'Alternating shading.'],
  'Cell color': ['Pass', '2,700 light green.'], 'Merged cells': ['Pass', 'Last row is one cell.'],
};
const pptx = {
  'Bold, italic, underline, strikethrough': ['Pass', 'All four correct.'],
  'Center alignment and superscript': ['Pass', 'Centered; the 2 is small and raised (fixed 30 Oct 2026).'],
  'Numbered and lettered': ['Pass', '1. 2. then a), numbers the same size as the text (fixed 30 Oct 2026).'],
  'Cropped picture': ['Pass', 'Only the top-left part of the chart, enlarged.'],
  'Picture with border': ['Pass', 'Thick red border (fixed 30 Oct 2026).'],
  'Table header': ['Pass', 'Blue header row with white bold text (fixed 31 Oct 2026).'],
  'Banded rows': ['Pass', 'Rows alternate light / darker blue (fixed 31 Oct 2026).'],
  'Gradient fill': ['Pass', 'Box fades from white (left) to dark blue (right) (fixed 31 Oct 2026).'],
};
const ppt = {
  'PowerPoint 97-2003 file (.ppt)': ['Pass', 'Opens with all 10 slides.'],
  'Bold, italic, underline, strikethrough': ['N/A', 'Bold, italic, underline OK. The .ppt format has no strikethrough setting (it was lost when this test file was saved as .ppt; LibreOffice shows none either), so this line cannot be checked.'],
  'Center alignment and superscript': ['Pass', 'Centered; x² raised.'],
  'Numbered and lettered': ['Pass', '1. 2. then a) (fixed 31 Oct 2026).'],
  'Cropped picture': ['Pass', 'Only the top-left part of the chart, enlarged (fixed 31 Oct 2026).'],
  'Picture with border': ['Pass', 'Thick red border.'],
  'Gradient fill': ['Pass', 'Box fades from white (left) to dark blue (right) (fixed 31 Oct 2026).'],
};

const pdf = {
  'Standard fonts': ['Pass', 'Helvetica, Times, Courier.'], 'Embedded fonts': ['Pass', 'Arial and Calibri.'], 'Sizes and colors': ['Pass', '8-28 pt; red, blue, green.'],
  'Japanese, Chinese, Korean': ['Pass', 'All three lines, no empty boxes.'],
  'Search (Ctrl+F)': ['Pass', 'Finds 3 on page 1: the 2 test words + the word in the test description (the file\'s "2 matches" wording is off by one).'],
  'Select and copy text': ['Pass', 'Text selectable.'], 'Lines': ['Pass', 'Solid and dashed lines.'], 'Filled shapes': ['Pass', 'Blue rectangle, orange circle, yellow triangle.'],
  'Outline shape': ['Pass', 'Green rounded outline.'],
  'Gradient': ['Pass', 'Box fades white to dark blue (test file fixed in this round: the gradient was not limited to its box and covered half the page).'],
  'JPEG image': ['Pass', 'Chart, colors match.'], 'PNG image': ['Pass', 'Purple Picture 2.'], 'Zoom': ['Pass', 'Page re-drawn at the new zoom; stays sharp.'],
  'External link': ['Pass', 'Clicking opens https://example.com in a new tab (fixed 30 Oct 2026).'],
  'Internal link': ['Pass', 'Clicking "go to page 1" jumps to page 1 (fixed 30 Oct 2026).'],
  'Bookmarks / outline': ['Pass', '9 bookmarks above the thumbnails; clicking one goes to its page (fixed 31 Oct 2026; the test file had lost its bookmarks, see the 31 Oct (3) run).'],
  'Text field': ['Pass', 'Box shows "Jane Doe" (view only).'], 'Checkbox': ['Pass', 'Checked.'], 'Drop-down': ['Pass', 'Shows "Checking".'],
  'Radio buttons': ['Pass', 'Personal selected, Business not.'], 'A4 page': ['Pass', 'Narrower and taller than Letter.'], 'Landscape page': ['Pass', 'Wider than tall.'],
  'Legal page': ['Pass', 'Taller than Letter.'], 'Scanned page without text': ['Pass', 'Picture shown; words inside the picture are not found by search.'],
  'Page stored with 90° rotation': ['Pass', 'Page 3 shown turned 90° clockwise.'], 'Page count and thumbnails': ['Pass', '9 pages; thumbnails show the different sizes.'],
  'Password prompt': ['Pass', 'Asks for the password; nothing shown before.'], 'Wrong password': ['Pass', '"Incorrect password. Try again."'],
  'Correct password': ['Pass', 'test1234 opens the 9 pages.'],
};

const images = {
  'Image opens': 'Shows fully.', 'Colors': 'Each bar matches its name.', 'Gray levels': 'Black to white steps, no tint.', 'Sharpness': 'Fine lines sharp at 100% zoom.',
  'Orientation': 'BOTTOM-LEFT / BOTTOM-RIGHT at the bottom.', '.jpeg extension': 'Same as test-image.jpg.', 'Zoom in and out': 'Zoom works; lines sharp.',
  'Progressive JPEG (web-optimised)': 'Same as normal JPEG.', 'Rotation from camera data (EXIF orientation 6)': 'Upright, reads normally.',
  'CMYK color JPEG': 'Colors normal.', 'Very large image (40 megapixels)': 'Opens in about 2 s; zoom works; page stays responsive.',
  'Multi-page TIFF (5 pages)': '5 pages, PAGE n / 5 on each.', 'Page thumbnails': '5 thumbnails in order.', 'Go to page 3 from the page bar': 'Page 3 shown.',
  'No compression': 'Opens, colors correct.', 'LZW compression': 'Opens, colors correct.', 'PackBits compression': 'Opens, colors correct.',
  'Deflate (ZIP) compression': 'Opens, colors correct.', 'JPEG compression (color scan)': 'Opens, colors correct.',
  'Black and white fax (CCITT Group 4)': 'Opens, readable, black and white only.', 'Black and white fax (CCITT Group 3)': 'Opens, readable, black and white only.',
  'Grayscale 8-bit': 'Opens, gray levels.', 'CMYK color': 'Opens, colors correct.',
  '16-bit grayscale (medical / archive scans)': 'Opens, readable gray image (test file fixed in this round: it held 8-bit values in a 16-bit file, so it looked black).',
};

function resultFor(t) {
  const w = t.what;
  const id = t.id;
  if (/^DOC-/.test(id)) return doc[w] || word[w];
  if (/^DOC[XM]-/.test(id)) return word[w];
  if (/^XLS-/.test(id)) return xls[w] || excel[w];
  if (/^XLS[XM]-/.test(id)) return excel[w];
  if (/^CSV-/.test(id)) return csv[w] && ['Pass', csv[w]];
  if (/^PPTX-/.test(id)) return pptx[w] || slidesCommon[w];
  if (/^PPT-/.test(id)) return ppt[w] || slidesCommon[w];
  if (/^PDF-/.test(id)) return pdf[w];
  if (/^(JPE?G|TIFF?)-/.test(id)) return images[w] && ['Pass', images[w]];
  return null;
}

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet('All tests');
  const head = ws.getRow(1);
  head.getCell(10).value = 'Tested on';
  head.getCell(10).font = head.getCell(9).font; head.getCell(10).fill = head.getCell(9).fill; head.getCell(10).alignment = head.getCell(9).alignment;
  ws.getColumn(10).width = 30;
  ws.getColumn(9).width = 60;
  const missing = [];
  const counts = {};
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const t = { id: row.getCell(2).value, file: row.getCell(3).value, what: row.getCell(5).value };
    const r = resultFor(t);
    if (!r) { missing.push(`${t.id} ${t.what}`); return; }
    row.getCell(8).value = r[0];
    row.getCell(9).value = r[1];
    row.getCell(10).value = ENV;
    row.getCell(10).alignment = { vertical: 'top', wrapText: true };
    // viewer status column follows the test result
    const status = r[0] === 'Pass' ? '✅' : r[0] === 'N/A' ? '➖' : /OK|correct except|but |only|Merged, but/i.test(r[1]) ? '⚠️' : '❌';
    row.getCell(7).value = status;
    counts[t.file] = counts[t.file] || { pass: 0, fail: 0, na: 0, total: 0 };
    counts[t.file].total++;
    counts[t.file][{ Pass: 'pass', Fail: 'fail', 'N/A': 'na' }[r[0]]]++;
  });
  ws.autoFilter = { from: 'A1', to: `J${ws.rowCount}` };

  // Summary: write the computed values (Excel recalculates the formulas when the file opens)
  const sum = wb.getWorksheet('Summary');
  const last = sum.getRow(sum.rowCount);
  if (last.getCell(1).value === 'TOTAL') sum.spliceRows(sum.rowCount, 1); // re-run: drop the old total
  sum.getRow(1).getCell(5).value = 'N/A';
  sum.getRow(1).getCell(6).value = 'Pass rate';
  let tp = 0, tf = 0, tn = 0, tt = 0;
  sum.eachRow((row, n) => {
    if (n === 1) return;
    const c = counts[row.getCell(1).value];
    if (!c) return;
    const range = `'All tests'!$C$2:$C$${ws.rowCount}`;
    const res = `'All tests'!$H$2:$H$${ws.rowCount}`;
    row.getCell(3).value = { formula: `COUNTIFS(${range},A${n},${res},"Pass")`, result: c.pass };
    row.getCell(4).value = { formula: `COUNTIFS(${range},A${n},${res},"Fail")`, result: c.fail };
    row.getCell(5).value = { formula: `COUNTIFS(${range},A${n},${res},"N/A")`, result: c.na };
    row.getCell(6).value = { formula: `IF(B${n}-E${n}=0,"",C${n}/(B${n}-E${n}))`, result: c.pass / (c.total - c.na) };
    row.getCell(6).numFmt = '0%';
    tp += c.pass; tf += c.fail; tn += c.na; tt += c.total;
  });
  const tr = sum.addRow(['TOTAL', tt, tp, tf, tn, tp / (tt - tn)]);
  tr.font = { bold: true }; tr.getCell(6).numFmt = '0%';

  // How to use: the history of test runs (rewritten each time, so a re-run never duplicates lines)
  const how = wb.getWorksheet('How to use');
  for (let n = 13; n <= Math.max(how.rowCount, 40); n++) { const c = how.getRow(n).getCell(1); c.value = null; c.font = {}; }
  const RUNS = [
    ['22 Oct 2026', '372 Pass, 134 Fail (74%)', 'First full run. Fixed: Excel hidden rows/columns, column widths and text overflow, CSV line breaks and long text; test files: Word page 7 columns, PDF gradient, 16-bit TIFF, Excel column B width.'],
    ['30 Oct 2026', '401 Pass, 104 Fail, 1 N/A (79%)', 'Fixed: Word page numbers and header on every page; Excel frozen rows, clickable links, cell notes, double underline, superscript/subscript, mixed formats in a cell, rotated text, accounting $; PDF clickable links; PowerPoint superscript, a) numbering, number size, picture border.'],
    ['31 Oct 2026', '417 Pass, 88 Fail, 1 N/A (83%)', 'Fixed: PowerPoint shape outlines (triangle, arrows, star, callouts, chevron, pentagon, hexagon…) in .pptx and .ppt, arrow heads on lines, gradient fills, table styles (header, banded rows), .ppt numbered lists and picture cropping; test files: slide titles were saved with zero height.'],
    ['31 Oct 2026 (2)', '435 Pass, 70 Fail, 1 N/A (86%)', 'Fixed: Excel conditional formatting (color scales, data bars, icon sets, highlight rules, duplicates, top/bottom, above average, text rules), pictures in sheets, filter buttons, gradient and pattern fills; test file: the duplicate-values rule was missing (ExcelJS cannot write it; now a formula rule).'],
    ['31 Oct 2026 (3)', '442 Pass, 63 Fail, 1 N/A (88%)', 'Fixed: PDF bookmarks panel; PowerPoint charts (column, bar, line, area, pie, doughnut) in .pptx and .ppt, speaker notes in .pptx and .ppt; test file: the PDF lost its bookmarks when page 3 was rotated.'],
    ['31 Oct 2026 (4)', '448 Pass, 57 Fail, 1 N/A (89%)', 'Fixed: Word floating objects (shapes and arrows, text boxes, pictures with square / top-and-bottom wrap, in front of or behind text) in .docx and .docm.'],
    ['31 Oct 2026 (5)', '485 Pass, 14 Fail, 7 N/A (97%)', 'Fixed: Excel 97-2003 (.xls) fonts, colors, underline, size, font names, alignment, wrap, indent, rotation, borders, rich text in a cell, frozen panes, conditional rules (cell value, formula), pictures; formulas whose text results were not saved are worked out. N/A: 6 .xls tests whose features the test file does not contain.'],
  ];
  if (process.env.RUN_DATE && !RUNS.some((r) => r[0] === process.env.RUN_DATE)) {
    RUNS.push([process.env.RUN_DATE, `${tp} Pass, ${tf} Fail${tn ? `, ${tn} N/A` : ''} (${Math.round((tp / (tt - tn)) * 100)}%)`, `Fixed: ${process.env.FIXED || ''}`]);
  }
  let row = 14;
  const put = (text, font) => { const c = how.getRow(row++).getCell(1); c.value = text; c.alignment = { wrapText: true }; if (font) c.font = font; };
  put('Test runs (Stage and Angular, Chrome 1400 px with spot checks in Firefox and Safari iPhone):', { bold: true, color: { argb: 'FF1F4E79' } });
  const latest = `${tp} Pass, ${tf} Fail${tn ? `, ${tn} N/A` : ''} (${Math.round((tp / (tt - tn)) * 100)}%)`;
  for (const [date, result, fixed] of RUNS) put(`${date}: ${result === 'RESULT' ? latest : result}. ${fixed}`);
  row++;
  put('Results are in the Result / Tester notes / Tested on columns of "All tests".');
  put('"Viewer status" shows the tested result: ✅ matches the original, ⚠️ partly, ❌ missing, ➖ cannot be checked with this file (N/A).');
  await wb.xlsx.writeFile(file);
  console.log(`${tp} pass, ${tf} fail, ${tn} n/a, ${tt} total; missing: ${missing.length}`);
  missing.forEach((m) => console.log('  ', m));
  for (const [f, c] of Object.entries(counts)) console.log(f.padEnd(28), c.pass, '/', c.total);
})();
