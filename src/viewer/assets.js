/**
 * Where the viewer's static files are served from: fonts/ (look-alike Office fonts) and
 * pdfjs/ (pdf.js character maps, fonts and decoders). By default they sit next to the built
 * library file; `createDocViewer(el, { assetsUrl })` points somewhere else.
 */
let base = import.meta.url.replace(/[^/]*$/, ''); // folder of this file (written so bundlers keep it as is)

export function setAssetsUrl(url) {
  if (url) base = new URL(url.endsWith('/') ? url : `${url}/`, document.baseURI).href;
}

export const assetUrl = (path) => new URL(path, base).href;
