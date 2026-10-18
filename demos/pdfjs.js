import { setup } from './shared.js';
import source from './pdfjs.js?raw';

// #region demo
import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';

// pdf.js parses in a Web Worker so the page never freezes.
pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

async function render(buffer, name, out) {
  const pdf = await pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    isEvalSupported: false, // security: never eval code from the PDF
    standardFontDataUrl: new URL('../pdfjs/standard_fonts/', location.href).href,
  }).promise;

  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const viewport = page.getViewport({ scale: 1.25 });

    const canvas = document.createElement('canvas');
    canvas.className = 'page';
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    out.append(canvas);

    await page.render({ canvas, viewport }).promise;
  }
}
// #endregion

setup({ sample: 'quarterly-report.pdf', accept: '.pdf', render, source });
