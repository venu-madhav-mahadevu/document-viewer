# Docs Viewer

A browser-based document viewer for PDF, Word, Excel/CSV, PowerPoint, images and text, built
**entirely from open-source libraries**. It runs 100% client-side: documents are never uploaded,
and no Google, Microsoft or other third-party service is involved.

**Supported formats (exactly these 15):**
`.pdf` `.doc` `.docx` `.docm` `.xls` `.xlsx` `.xlsm` `.csv` `.ppt` `.pptx` `.jpg` `.jpeg` `.tif` `.tiff` `.txt`

| Format | Extensions | Engine | License |
|---|---|---|---|
| PDF | `.pdf` (incl. password-protected) | [pdf.js](https://github.com/mozilla/pdf.js) | Apache-2.0 |
| Word | `.docx` `.docm` | [docx-preview](https://github.com/VolodymyrBaydalka/docxjs) | Apache-2.0 |
| Word 97-2003 | `.doc` | Text, formatting, styles, lists, tables and page breaks read from the binary file with [cfb](https://github.com/SheetJS/js-cfb) (`doc-format.js`) | Apache-2.0 |
| Excel | `.xlsx` `.xlsm` `.xls` `.csv` | [SheetJS CE](https://sheetjs.com) 0.20.3 | Apache-2.0 |
| PowerPoint | `.pptx` | In-house renderer (`src/viewer/renderers/pptx.js`) + [JSZip](https://stuk.github.io/jszip/) | MIT |
| PowerPoint 97-2003 | `.ppt` | Slides drawn from their shapes, pictures and formatted text, read with [cfb](https://github.com/SheetJS/js-cfb) (`ppt-format.js`) | Apache-2.0 |
| Images | `.jpg` `.jpeg` | Browser | - |
| TIFF | `.tif` `.tiff` (multi-page scans, CCITT G3/G4, LZW, JPEG, Deflate) | [UTIF.js](https://github.com/photopea/UTIF.js) | MIT |
| Text | `.txt` | Browser | - |

Any other file type is rejected with a "not supported" message, and the file picker only offers
these formats. Files without an extension (e.g. streamed from an API) are identified from their
content. `.doc` and `.ppt` are binary formats with no open-source browser renderer; this project reads
them itself (formatting, styles, shapes and pictures). Rare features (embedded objects, text
effects, exact line breaking) are simplified.

**Google Docs-style viewing features** (no extra libraries; built on browser APIs):

| Feature | Details |
|---|---|
| Page thumbnails | Sidebar of page/slide previews (PDF, Word, PowerPoint), drawn lazily; click to jump |
| Find in document | `Ctrl+F`, every match highlighted, `3 of 12`, Enter / Shift+Enter. Searches all PDF pages, even ones not drawn yet (CSS Custom Highlight API) |
| Print | `Ctrl+P` prints only the document, one page per sheet; slides print landscape; spreadsheets print the open sheet |
| Full screen | For reading and presenting slides |
| Page / zoom pill | Floating `Page 3 / 12 · − 100% +` bar, like Google Drive's preview |
| Keyboard | `←` `→` previous/next page, `Ctrl` `+`/`−`/`0` zoom, `Esc` closes search |

## Angular version (this branch)

The app is an **Angular 21** application built with current Angular features:

| Feature | Where | Benefit |
|---|---|---|
| Zoneless change detection | `app.config.ts` | No zone.js; change detection only when signals change |
| Signals, `input()` / `output()`, `viewChild()` | all components | Fine-grained updates |
| Lazy routes (`loadComponent`) + `PreloadAllModules` | `app.routes.ts` | Home loads without the viewer; the view page is preloaded when idle |
| `@defer (on viewport; on idle)` | home gallery | Lower groups render when needed |
| `resource()` (experimental in Angular 21) | view page | Loading/error states built in; switching documents cancels the running download |
| `afterNextRender()` | `<doc-viewer>` | Viewer created once the component is in the page |
| Angular service worker (`ngsw-config.json`) | production build | Instant repeat visits and offline use; 36-50 MB test files are not cached |
| `OnPush`, standalone components, built-in `@for`/`@if` | all | Less work per change, smaller bundles |

First load is about 96 KB (compressed); each format's library loads only when such a file is opened.

**Deployment path:** the production build uses `baseHref: /document-viewer/angular/` (GitHub Pages).
For another server, change `baseHref` in `angular.json` (the service worker needs the real path). The viewer itself (`src/viewer/`) is unchanged, framework-free code wrapped in an
Angular component:

```html
<doc-viewer [src]="file" [download]="true" [showBack]="true" (back)="goHome()" />
```

`src` can be a `File`, `Blob`, `ArrayBuffer` or a same-origin URL. Types: `src/viewer/index.d.ts`.

| Angular file | What it does |
|---|---|
| `src/app/doc-viewer/doc-viewer.ts` | `<doc-viewer>` component (wraps `createDocViewer`) |
| `src/app/documents.service.ts` | Sample list, download with progress, documents opened in this session |
| `src/app/home/home.ts` | Home page: document gallery (`#/`) |
| `src/app/view/view.ts` | View page: `#/view?file=<sample>` or `#/view?src=<same-origin URL>` |
| `src/app/app.ts` | Shell: sidebar, Open file button, drag and drop, routed page, phone layout |

The demo pages in `demos/` are not moved to Angular yet; the app links to the staging demos.

## Run it

```bash
npm install
npm start          # http://localhost:4200
npm run build      # static output in dist/docs-viewer/browser/
```

Angular 21 works with Node.js 20.19+, 22.12+ or 24+. (Angular 22 needs Node.js 24.15+.)

The build output is plain static files. Host it on any internal web server (IIS, nginx, Apache, S3/CloudFront
behind SSO, etc.). No backend is required.

## Library demos with code

**Compatibility test** (`demos/compare.html`): runs all 11 libraries on one sample of each of the 15
formats, live in the browser, and shows a pass / unreadable / fail matrix with a preview of every
result. Result on the samples: the
Docs Viewer component opens 15 / 15; each single library opens only its own formats.

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
| `papaparse` | `papaparse` | MIT | CSV |
| `pptx-jszip` | `jszip` | MIT | PowerPoint text + images |
| `utif` | `utif` | MIT | TIFF scans (multi-page, fax, LZW) |
| `viewerjs` | `viewerjs` | MIT | Images (zoom/rotate) |

The alternative libraries (mammoth, exceljs, papaparse, dompurify, highlight.js, viewerjs) are dev
dependencies used only by the demos, so the production viewer ships just its own six libraries
(pdfjs-dist, docx-preview, xlsx, jszip, cfb, utif). `npm audit --omit=dev` reports 0
vulnerabilities. The full audit flags only `sprintf-js`, which belongs to mammoth's command-line
tool and never runs in the browser.

## Opening documents

The demo app opens only the platform's documents: visitors open the documents listed on the home screen (or a same-origin
link such as `?src=/samples/report.pdf`). There is no way to open files from the visitor's computer,
dropped files are blocked. The Download button saves the original file. Developers embedding the component can
still call `viewer.open(file)` with any File, Blob or URL their application provides.

## Sample files

`public/samples/` has 17 example files covering all 15 supported formats. They are listed on the
home screen so you can open each one with a click (the **Docs Viewer** title returns to that list).
You can also open one directly, e.g. `http://localhost:5180/?src=/samples/strategy-deck.pptx`.

| Group | Files |
|---|---|
| PDF | `quarterly-report.pdf` (12 pages), `mixed-page-sizes.pdf` (Letter/A4, portrait/landscape), `protected-statement.pdf` (password `demo1234`) |
| Word | `onboarding-guide.docx`, `onboarding-macro.docm`, `onboarding-guide.doc` |
| Excel | `regional-sales.xlsx`, `branches-macro.xlsm`, `branches-legacy.xls`, `accounts.csv` |
| PowerPoint | `strategy-deck.pptx`, `strategy-deck.ppt` |
| Images | `photo.jpg`, `photo-chart.jpeg`, `scanned-contract.tif` (3-page fax-compressed scan), `color-chart.tiff`, `color-scan-jpeg.tif` (JPEG-compressed) |
| Text | `readme.txt` |

`accounts.csv` deliberately contains a `<script>` payload to show that document content is never
executed.

## Where the libraries are declared

All libraries and formats are declared in **one file, `src/viewer/formats.js`**: for each kind of
document it lists the file extensions, the toolbar label, the library that renders it and how that
renderer is loaded. File-type detection, the file picker (`ACCEPT`) and the viewer all read from it,
so adding, removing or swapping a library means editing only that file (plus `package.json` for the
npm package).

## Embedding in another app

The viewer is a self-contained module (`src/viewer/`) with no framework dependency, so it drops into
React, Angular, Vue or plain HTML.

```js
import { createDocViewer } from './viewer/index.js';

const viewer = createDocViewer(document.getElementById('viewer'), {
  maxBytes: 200 * 1024 * 1024, // reject larger files (the message states the file's size)
  download: true,              // Download button (set false for view-only)
  print: true,                 // Print button and Ctrl+P
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
  third-party origins (`default-src 'self'`). `netlify.toml` also sends it as an HTTP header, with
  `frame-ancestors 'none'` so other sites can't embed the viewer.
- **No CDNs.** All libraries are bundled at build time, and pdf.js fonts/CMaps/wasm are served from
  your own origin (`scripts/copy-pdfjs-assets.mjs`).
- **Document content is never executed.** Spreadsheet cells and text files are written with
  `textContent`. Hyperlinks inside documents are restricted to `http`/`https`/`mailto` and open with
  `noopener noreferrer`. pdf.js runs with `isEvalSupported: false`. Macros in `.docm` / `.xlsm`
  files are never run.
- **URL loading is same-origin only**, so `?src=` can't be abused to fetch arbitrary hosts.
- `npm audit --omit=dev` (the shipped code): 0 vulnerabilities. The full audit lists 3 moderate
  findings in `sprintf-js`, used only by mammoth's command-line tool for one demo page, never in the
  browser. SheetJS is installed from the official SheetJS registry because
  the npm `xlsx@0.18.5` package has unpatched CVEs (CVE-2023-30533, CVE-2024-22363).

## Known limitations

- **Legacy binary `.doc` / `.ppt`**: no mature open-source browser renderer exists, so they open as a
  labelled simplified preview (text, tables and slide text; `.ppt` pictures in a gallery) without
  the original fonts or layout. `.xls` is fully supported.
- **PowerPoint**: text, placeholders, pictures, shapes, groups, lines, tables, backgrounds and theme
  colors render. Charts, SmartArt, gradients on shapes, table styles, animations and EMF/WMF images
  are shown as labelled placeholders or simplified.
- **Word**: rendering is HTML-based, so page breaks and complex floating layouts may differ slightly
  from Word.
- **Spreadsheets**: formulas show the value cached by Excel at save time (they aren't recalculated).
  Very large sheets display the first 5,000 rows × 256 columns. A 50 MB, 370,000-row workbook
  opens in about 4 to 13 seconds (parsed in a background worker, 77 MB of memory).
- **Large TIFF scans**: pages are decoded only as they come near the screen and released when far
  away, so the 60-page, 36 MB sample opens in under 0.1 s and holds about 5 pages in memory. A page
  takes a moment to appear the first time it scrolls into view.
- **Large PDFs**: a 49 MB, 100-page PDF opens in under 1 second; pages are drawn as they scroll in.
- **View-only is not copy protection**: hiding the Download button stops casual downloads, but a
  browser that can display a file can also save it (developer tools, screenshots, print to PDF).
  Truly restricted documents need server-side controls such as watermarking or rendering to images.

If full-fidelity Office rendering or editing becomes a requirement, the open-source route is a
self-hosted **ONLYOFFICE Docs** or **Collabora Online** server (both AGPL; commercial licenses
available), embedded with the same "nothing leaves our network" model.
