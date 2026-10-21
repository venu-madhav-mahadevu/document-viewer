# Manual test pack

`public/samples/test-pack/` holds one or more test files for every supported format, plus
`TEST-CHECKLIST.xlsx`. Every test inside a file is labeled with:

- an **ID** (e.g. `DOCX-07`, `XLSX-23`, `PPTX-10`),
- **what is tested**,
- the **expected result** (what the original application shows).

The checklist lists all tests (ID, file, where, what, expected, known viewer status) with a
Pass / Fail / N/A column and notes, plus a summary per file.

| Format | Test file(s) |
|---|---|
| `.docx` `.docm` `.doc` | `test-word.*`: fonts, sizes, styles, colors, highlight, alignment, indents, spacing, headings, 6 bullet types (incl. Wingdings), numbered and multi-level lists, tables (shading, merged cells), inline / floating pictures, shapes and arrows, text box, hyperlink, equation, header/footer with page number, two columns, landscape page |
| `.xlsx` `.xlsm` `.xls` | `test-excel.*`: 21 number formats, text/Asian text, fills (solid, theme, pattern, gradient), font styles and colors, rich text, alignment, wrap, indent, rotation, 8 border styles, merged cells, hidden row/column, formulas with results, conditional formatting, notes, links, picture, filter, frozen rows, drop-down, hidden sheet |
| `.csv` | `test-data.csv`: quotes, commas, line breaks, UTF-8 with BOM, accents, Asian text, emoji, leading zeros, long numbers, dates, formula-looking text |
| `.pptx` `.ppt` | `test-slides.*` (16:9, 10 slides): text formatting, 9 bullet types incl. Wingdings, multi-level, numbering, shrink to fit, 12 shapes incl. arrows, rotation, connector, pictures (cropped, border), table, group, gradient, transparency, charts, background, notes |
| `.pdf` | `test-pdf.pdf` (9 pages): standard / embedded fonts, Japanese / Chinese / Korean, search word, drawings, gradient, images, links, bookmarks, form fields, A4 / landscape / Legal, rotated page, scan-only page; `test-pdf-protected.pdf` (password `test1234`) |
| `.jpg` `.jpeg` | test chart, `.jpeg`, progressive, phone rotation (EXIF), CMYK, 8000 x 5000 |
| `.tif` `.tiff` | 5-page scan, no compression, LZW, PackBits, Deflate, JPEG, fax G3 / G4, grayscale, CMYK, 16-bit |

## Rebuilding the pack

Requirements: Python 3 with `python-docx python-pptx pillow reportlab pypdf lxml`, Node.js,
and (for `.doc` / `.xls` / `.ppt`) the `staging` branch's `@matbee/libreoffice-converter`.

```bash
cd scripts/test-pack
pip install --target pylibs python-docx python-pptx pillow reportlab pypdf lxml
python gen_images.py                       # JPEG / TIFF + pictures used by the other files
python gen_word.py DOCX out/test-word.docx
python gen_word.py DOCM out/test-word-docm.docx
python gen_word.py DOC  out/test-word-for-doc.docx
node gen_excel.cjs XLSX out/test-excel.xlsx
node gen_excel.cjs XLSM out/test-excel-xlsm.xlsx
node gen_excel.cjs XLS  out/test-excel-for-xls.xlsx
python gen_csv.py
python gen_pptx.py PPTX out/test-slides.pptx
python gen_pptx.py PPT  out/test-slides-for-ppt.pptx
python gen_pdf.py
node lo-convert.mjs out/test-word-for-doc.docx=doc out/test-excel-for-xls.xlsx=xls out/test-slides-for-ppt.pptx=ppt
mv out/test-word-for-doc.doc out/test-word.doc
mv out/test-excel-for-xls.xls out/test-excel.xls
mv out/test-slides-for-ppt.ppt out/test-slides.ppt
node finalize.cjs                          # .docm / .xlsm types, clean-up, TEST-CHECKLIST.xlsx
cp out/* ../../public/samples/test-pack/
```
