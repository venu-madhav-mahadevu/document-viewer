# Docs Viewer

A browser-based document viewer for PDF, Word, Excel/CSV, PowerPoint, images and text, built
**entirely from open-source libraries**. It runs 100% client-side: documents are never uploaded,
and no Google, Microsoft or other third-party service is involved.

| Format | Extensions | Engine | License |
|---|---|---|---|
| PDF | `.pdf` (incl. password-protected) | [pdf.js](https://github.com/mozilla/pdf.js) | Apache-2.0 |
| Word | `.docx` `.docm` `.dotx` | [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) | Apache-2.0 |
| Spreadsheets | `.xlsx` `.xlsm` `.xlsb` `.xls` `.ods` `.csv` `.tsv` | [SheetJS CE](https://sheetjs.com) 0.20.3 | Apache-2.0 |
| PowerPoint | `.pptx` `.pptm` `.ppsx` `.potx` | In-house renderer (`src/viewer/renderers/pptx.js`) + [JSZip](https://stuk.github.io/jszip/) | MIT |
| Images | `.png` `.jpg` `.gif` `.webp` `.bmp` `.svg` `.avif` | Browser | - |
| Text | `.txt` `.json` `.xml` `.md` `.log` `.yaml` `.sql` ... | Browser | - |

Files without an extension (e.g. streamed from an API) are detected from their content.

## Run it

```bash
npm install
npm run dev        # http://localhost:5180
npm run build      # static output in dist/
npm run preview    # serve dist/ locally
```

`dist/` is plain static files. Host it on any internal web server (IIS, nginx, Apache, S3/CloudFront
behind SSO, etc.). No backend is required.

## Library demos with code

`demos/` has one page per open-source library. Each page renders a sample file and shows the code
that did it, next to the result. Open `http://localhost:5180/demos/` (or **Library demos with code**
in the app sidebar).

| Demo | Library (npm) | License | Format |
|---|---|---|---|
| `full-viewer` | This project's `createDocViewer()` | Apache-2.0 / MIT | All |
| `pdfjs` | `pdfjs-dist` | Apache-2.0 | PDF |
| `docx-preview` | `docx-preview` | Apache-2.0 | Word (layout) |
| `mammoth` | `mammoth` + `dompurify` | BSD-2 / Apache-2.0 | Word (clean HTML) |
| `sheetjs` | `xlsx` (SheetJS CE) | Apache-2.0 | Excel / CSV / ODS |
| `exceljs` | `exceljs` | MIT | Excel with styles |
| `papaparse` | `papaparse` | MIT | CSV / TSV |
| `pptx-jszip` | `jszip` | MIT | PowerPoint text + images |
| `viewerjs` | `viewerjs` | MIT | Images (zoom/rotate) |
| `markdown` | `marked` + `dompurify` + `highlight.js` | MIT / Apache-2.0 / BSD-3 | Markdown and code |

The alternative libraries (mammoth, exceljs, papaparse, marked, dompurify, highlight.js, viewerjs)
are dev dependencies used only by the demos, so the production viewer still ships just four
libraries. `npm audit --omit=dev` reports 0 vulnerabilities. The full audit flags only `sprintf-js`,
which belongs to mammoth's command-line tool and never runs in the browser.

## Sample files

`public/samples/` has 46 example files, and they are listed on the home screen so you can open each
one with a click (the **Docs Viewer** title returns to that list). You can also open one directly, e.g.
`http://localhost:5180/?src=/samples/strategy-deck.pptx`.

| Group | Files |
|---|---|
| PDF | `quarterly-report.pdf` (12 pages), `mixed-page-sizes.pdf` (Letter/A4, portrait/landscape), `protected-statement.pdf` (password `demo1234`) |
| Word | `.docx`, `.dotx`, `.docm` |
| Spreadsheets | `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.ods`, `.csv`, `.tsv` |
| PowerPoint | `.pptx`, `.ppsx`, `.potx`, `.pptm` |
| Images | `.png`, `.jpg`, `.gif` (animated), `.webp`, `.avif`, `.bmp`, `.ico`, `.svg` |
| Text & code | `.txt`, `.log`, `.md`, `.json`, `.xml`, `.yaml`, `.yml`, `.ini`, `.cfg`, `.html`, `.htm`, `.css`, `.js`, `.ts`, `.sql` |
| Detection & unsupported | a PDF with no extension, `.doc`, `.ppt`, `.rtf`, `.zip` |

Some samples deliberately contain script payloads (`logo.svg`, `customers.json`, `page.html`,
`accounts.csv`) to show that document content is never executed.

## Embedding in another app

The viewer is a self-contained module (`src/viewer/`) with no framework dependency, so it drops into
React, Angular, Vue or plain HTML.

```js
import { createDocViewer } from './viewer/index.js';

const viewer = createDocViewer(document.getElementById('viewer'), {
  maxBytes: 200 * 1024 * 1024, // reject larger files
  download: true,              // show the Download button
  fetchInit: { headers: { Authorization: `Bearer ${token}` } }, // for URL sources
});

await viewer.open(file);                              // File from <input type=file> or drag-drop
await viewer.open(blob, 'statement.pdf');             // Blob / ArrayBuffer + file name
await viewer.open('/api/documents/123', 'report.xlsx'); // same-origin URL

viewer.zoomIn(); viewer.zoomOut(); viewer.fitWidth(); viewer.setZoom(1.5);
viewer.destroy();
```

React example:

```jsx
function DocumentPreview({ file }) {
  const ref = useRef(null);
  const viewer = useRef(null);
  useEffect(() => { viewer.current = createDocViewer(ref.current); return () => viewer.current.destroy(); }, []);
  useEffect(() => { if (file) viewer.current.open(file); }, [file]);
  return <div ref={ref} style={{ height: '80vh' }} />;
}
```

The container needs a height; the viewer fills it.

## Security

- **No data leaves the browser.** Rendering is local, and the CSP in `index.html` allows no
  third-party origins (`default-src 'self'`). Mirror that CSP as an HTTP header in production.
- **No CDNs.** All libraries are bundled at build time, and pdf.js fonts/CMaps/wasm are served from
  your own origin (`scripts/copy-pdfjs-assets.mjs`).
- **Document content is never executed.** Spreadsheet cells and text files are written with
  `textContent`. Hyperlinks inside documents are restricted to `http`/`https`/`mailto` and open with
  `noopener noreferrer`. pdf.js runs with `isEvalSupported: false`, and SVGs are shown via `<img>`
  (no script execution).
- **URL loading is same-origin only**, so `?src=` can't be abused to fetch arbitrary hosts.
- `npm audit`: 0 vulnerabilities. SheetJS is installed from the official SheetJS registry because
  the npm `xlsx@0.18.5` package has unpatched CVEs (CVE-2023-30533, CVE-2024-22363).

## Known limitations

- **Legacy binary `.doc` / `.ppt`**: no mature open-source browser renderer exists. Users get a
  clear message and a Download button. `.xls` is supported.
- **PowerPoint**: text, placeholders, pictures, shapes, groups, lines, tables, backgrounds and theme
  colors render. Charts, SmartArt, gradients on shapes, table styles, animations and EMF/WMF images
  are shown as labelled placeholders or simplified.
- **Word**: rendering is HTML-based, so page breaks and complex floating layouts may differ slightly
  from Word.
- **Spreadsheets**: formulas show the value cached by Excel at save time (they aren't recalculated).
  Very large sheets display the first 5,000 rows × 256 columns.

If full-fidelity Office rendering or editing becomes a requirement, the open-source route is a
self-hosted **ONLYOFFICE Docs** or **Collabora Online** server (both AGPL; commercial licenses
available), embedded with the same "nothing leaves our network" model.
