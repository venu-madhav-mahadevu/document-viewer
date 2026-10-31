# docs-viewer

A document viewer that runs entirely in the browser. Documents are opened, drawn and searched
on the page itself; nothing is uploaded and no outside service is called.

| Format | Extensions | Read by |
|---|---|---|
| PDF | `.pdf` (incl. password-protected) | [pdf.js](https://github.com/mozilla/pdf.js) |
| Word | `.docx` `.docm` | [docx-preview](https://github.com/VolodymyrBaydalka/docxjs), plus our own floating shapes, text boxes, pictures, page numbers |
| Word 97-2003 | `.doc` | Our own reader (`doc-format.js`, `doc-drawing.js`, `mtef.js`) on [cfb](https://github.com/SheetJS/js-cfb) |
| Excel | `.xlsx` `.xlsm` `.csv` | [SheetJS CE](https://sheetjs.com), plus our own styles, conditional formatting, charts, pictures |
| Excel 97-2003 | `.xls` | SheetJS CE, plus our own formatting reader (`xls-format.js`) |
| PowerPoint | `.pptx` | Our own renderer (`pptx.js`, `shapes.js`, `charts.js`) on [JSZip](https://stuk.github.io/jszip/) |
| PowerPoint 97-2003 | `.ppt` | Our own reader (`ppt-format.js`) on cfb |
| Images | `.jpg` `.jpeg` | Browser |
| TIFF | `.tif` `.tiff` (multi-page, CCITT G3/G4, LZW, JPEG, Deflate, 16-bit) | [UTIF.js](https://github.com/photopea/UTIF.js) |
| Text | `.txt` | Browser |

Viewing features: page thumbnails, PDF bookmarks, search (Ctrl+F), zoom and fit to width,
page-by-page printing, download, keyboard navigation, phone layout, dark mode.

Tested with 506 test cases (the files are in `test/files`): 505 pass, 1 cannot be checked
(the PowerPoint 97-2003 format has no strikethrough).

## Use

```bash
npm install
npm run build        # dist/: docs-viewer.js, docs-viewer.css, fonts/, pdfjs/, format chunks
```

Copy `dist/` into your application (keep the folder together) and:

```html
<link rel="stylesheet" href="docs-viewer/docs-viewer.css">
<div id="viewer" style="height: 100vh"></div>
<script type="module">
  import { createDocViewer, ACCEPT } from './docs-viewer/docs-viewer.js';

  const viewer = createDocViewer(document.getElementById('viewer'));
  await viewer.open(file);                           // File from <input> or drag and drop
  await viewer.open(blob, 'statement.pdf');          // Blob or ArrayBuffer + file name
  await viewer.open('/api/docs/123', 'report.xlsx'); // same-origin URL
</script>
```

From a bundler (Vite, webpack, Angular): `import { createDocViewer } from 'docs-viewer'` and
`import 'docs-viewer/style.css'`, then serve `dist/fonts/` and `dist/pdfjs/` and pass their
location as `assetsUrl`.

### Options

`createDocViewer(container, options)`

| Option | Default | |
|---|---|---|
| `assetsUrl` | folder of `docs-viewer.js` | where the `fonts/` and `pdfjs/` folders are served |
| `maxBytes` | 200 MB | larger files are refused with a message |
| `download` | `true` | show the Download button |
| `print` | `true` | show the Print button |
| `fetchInit` | – | `fetch()` options for URLs (headers, credentials) |
| `onBack` | – | shows a Back button that calls this function |
| `badge` | `false` | show the format label next to the file name |

### Viewer methods

`open(source, name?)`, `loading(message, name?)`, `zoomIn()`, `zoomOut()`, `fitWidth()`,
`setZoom(z)`, `goToPage(n)`, `find(text)`, `print()`, `destroy()`.

Also exported: `ACCEPT` (for `<input type="file" accept>`), `SUPPORTED_EXTENSIONS`, `FORMATS`,
`LIBRARIES`, `formatBytes()`, `formatDuration()`.

## Develop and test

```bash
npm run dev          # demo page with the library source: http://localhost:5185/demo/
npm run build
npm test             # opens all 28 test files with the built library (Playwright, Chrome)
node test/library.test.cjs firefox
node test/library.test.cjs webkit
```

The demo page also runs the built library: `/demo/?dist`.

## Licenses

Copyright (c) 2026 Venu Madhav Mahadevu. All rights reserved; see [LICENSE](LICENSE). Use,
copying or distribution needs written permission. The open-source libraries and fonts it uses
keep their own licenses, which must travel with any copy; see
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
