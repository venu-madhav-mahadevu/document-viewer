import { setup } from './shared.js';
import source from './full-viewer.js?raw';

// #region demo
// This project's all-in-one component: pdf.js + docx-preview + SheetJS + JSZip
// (with its own PPTX renderer) behind a single API.
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

setup({ sample: 'strategy-deck.pptx', render, source });
