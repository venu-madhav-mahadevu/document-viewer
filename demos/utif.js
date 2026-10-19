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

setup({ sample: 'scanned-contract.tif', accept: '.tif,.tiff', render, source });
