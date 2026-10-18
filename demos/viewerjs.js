import { setup } from './shared.js';
import source from './viewerjs.js?raw';

// #region demo
import Viewer from 'viewerjs';
import 'viewerjs/dist/viewer.css';

let viewer;

async function render(buffer, name, out) {
  viewer?.destroy();

  const img = document.createElement('img');
  img.src = URL.createObjectURL(new Blob([buffer], { type: name.endsWith('.svg') ? 'image/svg+xml' : '' }));
  img.alt = name;

  const box = document.createElement('div');
  box.style.height = '60vh';
  box.append(img);
  out.append(box);

  // Inline viewer with zoom, rotate, flip and drag-to-pan.
  viewer = new Viewer(img, {
    inline: true,
    navbar: false,
    title: false,
    toolbar: { zoomIn: 1, zoomOut: 1, oneToOne: 1, reset: 1, rotateLeft: 1, rotateRight: 1, flipHorizontal: 1 },
  });
}
// #endregion

setup({ sample: 'photo.jpg', accept: 'image/*', render, source });
