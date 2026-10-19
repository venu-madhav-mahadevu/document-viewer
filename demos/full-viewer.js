import { setup } from './shared.js';
import { ACCEPT } from '../src/viewer/index.js';
import source from './full-viewer.js?raw';

// #region demo
// This project's all-in-one component for all 15 formats: pdf.js, docx-preview,
// SheetJS, JSZip (+ own PPTX renderer), cfb (.doc/.ppt) and UTIF.js (TIFF) behind one API.
import { createDocViewer } from '../src/viewer/index.js';

let viewer;

async function render(buffer, name, out) {
  const box = document.createElement('div');
  box.style.height = '70vh';
  out.append(box);

  viewer?.destroy();
  viewer = createDocViewer(box);       // toolbar, zoom, pages, download
  await viewer.open(new Blob([buffer]), name); // any supported format
}
// #endregion

setup({ sample: 'strategy-deck.pptx', accept: ACCEPT, render, source });
