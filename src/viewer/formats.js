// ONE place that declares every library and every supported format.
// To add, remove or swap a library: edit this file only. detect.js and the
// viewer (index.js) read everything from here.

/** The open-source libraries the viewer uses (npm packages are in package.json). */
export const LIBRARIES = {
  pdfjs: { name: 'pdf.js', npm: 'pdfjs-dist', license: 'Apache-2.0', url: 'https://github.com/mozilla/pdf.js' },
  docxPreview: { name: 'docx-preview', npm: 'docx-preview', license: 'Apache-2.0', url: 'https://github.com/VolodymyrBaydalka/docxjs' },
  sheetjs: { name: 'SheetJS CE', npm: 'xlsx', license: 'Apache-2.0', url: 'https://sheetjs.com/' },
  jszip: { name: 'JSZip + project PPTX renderer', npm: 'jszip', license: 'MIT', url: 'https://github.com/Stuk/jszip' },
  cfb: { name: 'cfb + project reader', npm: 'cfb', license: 'Apache-2.0', url: 'https://github.com/SheetJS/js-cfb' },
  utif: { name: 'UTIF.js', npm: 'utif', license: 'MIT', url: 'https://github.com/photopea/UTIF.js' },
  browser: { name: 'Browser built-in', npm: null, license: null, url: null },
};

/**
 * Each kind of document: its file extensions, toolbar label, the library that
 * renders it, and how that renderer is loaded. Renderers are imported only when
 * their format is first opened, so a PDF never downloads the spreadsheet engine.
 * The order here is the order of SUPPORTED_EXTENSIONS (and the file picker).
 *
 *   paged: has pages/slides (thumbnails, page numbers, page-by-page printing)
 *   badge: reuse another kind's badge color
 */
export const FORMATS = {
  pdf: {
    extensions: ['pdf'], label: 'PDF', library: 'pdfjs', paged: true,
    load: () => import('./renderers/pdf.js').then((m) => m.renderPdf),
  },
  doc: {
    extensions: ['doc'], label: 'Word 97-2003', library: 'cfb', paged: true, badge: 'docx',
    load: () => import('./renderers/legacy.js').then((m) => m.renderDoc),
  },
  docx: {
    extensions: ['docx', 'docm'], label: 'Word', library: 'docxPreview', paged: true,
    load: () => import('./renderers/docx.js').then((m) => m.renderDocx),
  },
  sheet: {
    extensions: ['xls', 'xlsx', 'xlsm', 'csv'], label: 'Spreadsheet', library: 'sheetjs', paged: false,
    load: () => import('./renderers/sheet.js').then((m) => m.renderSheet),
  },
  ppt: {
    extensions: ['ppt'], label: 'PowerPoint 97-2003', library: 'cfb', paged: true, badge: 'pptx',
    load: () => import('./renderers/legacy.js').then((m) => m.renderPpt),
  },
  pptx: {
    extensions: ['pptx'], label: 'Presentation', library: 'jszip', paged: true,
    load: () => import('./renderers/pptx.js').then((m) => m.renderPptx),
  },
  image: {
    extensions: ['jpg', 'jpeg'], label: 'Image', library: 'browser', paged: false,
    load: () => import('./renderers/basic.js').then((m) => m.renderImage),
  },
  tiff: {
    extensions: ['tif', 'tiff'], label: 'TIFF', library: 'utif', paged: true,
    load: () => import('./renderers/tiff.js').then((m) => m.renderTiff),
  },
  text: {
    extensions: ['txt'], label: 'Text', library: 'browser', paged: false,
    load: () => import('./renderers/basic.js').then((m) => m.renderText),
  },
};

/** ['pdf', 'doc', 'docx', ...] — the only extensions the viewer accepts. */
export const SUPPORTED_EXTENSIONS = Object.values(FORMATS).flatMap((f) => f.extensions);

/** 'docm' -> 'docx', 'csv' -> 'sheet', unknown -> undefined. */
export const kindForExtension = (ext) =>
  Object.keys(FORMATS).find((kind) => FORMATS[kind].extensions.includes(ext));

/** Value for <input type="file" accept="…">. */
export const ACCEPT = SUPPORTED_EXTENSIONS.map((e) => `.${e}`).join(',');
