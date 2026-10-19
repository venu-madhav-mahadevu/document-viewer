import { setup } from './shared.js';
import source from './utif.js?raw';

// #region demo
import UTIF from 'utif';

async function render(buffer, name, out) {
  // 1. Read the page list (cheap: no pixels decoded yet).
  const pages = UTIF.decode(buffer).filter((ifd) => ifd.t256); // t256 = width

  for (const [i, ifd] of pages.entries()) {
    // 2. Decode one page and convert any color mode to RGBA pixels.
    UTIF.decodeImage(buffer, ifd, pages);
    const rgba = UTIF.toRGBA8(ifd);

    // 3. Draw the pixels on a canvas.
    const canvas = document.createElement('canvas');
    canvas.className = 'page';
    canvas.width = ifd.width;
    canvas.height = ifd.height;
    canvas.getContext('2d').putImageData(
      new ImageData(new Uint8ClampedArray(rgba.buffer), ifd.width, ifd.height), 0, 0);
    canvas.title = `Page ${i + 1} of ${pages.length}`;
    out.append(canvas);
  }
  // The full viewer (src/viewer/renderers/tiff.js) adds lazy page loading for
  // large scans and decodes JPEG-compressed TIFFs with the browser's JPEG decoder.
}
// #endregion

// Zoom bar for the demo output (scans often need a closer look at fine print).
function addZoom(out) {
  const steps = [0.25, 0.33, 0.5, 0.67, 0.75, 1, 1.25, 1.5, 2, 3, 4];
  const pages = [...out.querySelectorAll('canvas')];
  if (!pages.length) return;
  const fit = () => Math.min(1, (out.clientWidth - 32) / Math.max(...pages.map((c) => c.width)));
  let zoom = fit();
  const bar = document.createElement('div');
  bar.className = 'zoombar';
  const minus = Object.assign(document.createElement('button'), { type: 'button', textContent: '−', title: 'Zoom out' });
  const label = Object.assign(document.createElement('button'), { type: 'button', title: 'Fit to width' });
  const plus = Object.assign(document.createElement('button'), { type: 'button', textContent: '+', title: 'Zoom in' });
  const apply = (z) => {
    zoom = Math.min(4, Math.max(0.1, z));
    pages.forEach((c) => { c.style.width = `${Math.round(c.width * zoom)}px`; c.style.maxWidth = 'none'; });
    label.textContent = `${Math.round(zoom * 100)}%`;
  };
  minus.onclick = () => apply([...steps].reverse().find((s) => s < zoom - 0.001) ?? zoom);
  plus.onclick = () => apply(steps.find((s) => s > zoom + 0.001) ?? zoom);
  label.onclick = () => apply(fit());
  out.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    (e.deltaY < 0 ? plus : minus).onclick();
  }, { passive: false });
  bar.append(minus, label, plus);
  out.prepend(bar);
  apply(zoom);
}

setup({
  sample: 'scanned-contract.tif',
  accept: '.tif,.tiff',
  render: async (buffer, name, out) => { await render(buffer, name, out); addZoom(out); },
  source,
});
