import UTIF from 'utif';
import { el } from '../sanitize.js';

/**
 * TIFF / multi-page scans (UTIF.js, MIT). Browsers can't display TIFF, so each
 * page is decoded to pixels and drawn on a canvas. Supports the compressions
 * scanners use: CCITT G3/G4 (fax), LZW, PackBits, Deflate and JPEG.
 */
export async function renderTiff({ buffer, mount }) {
  // Skip reduced-resolution previews (NewSubfileType bit 0) that some scanners embed.
  const ifds = UTIF.decode(buffer).filter((ifd) => ifd.t256 && !(ifd.t254 && (ifd.t254[0] & 1)));
  if (!ifds.length) throw new Error('This TIFF file contains no readable images.');

  const root = el('div', 'dv-tiff');
  mount.append(root);

  const pages = [];
  for (const ifd of ifds) {
    UTIF.decodeImage(buffer, ifd, ifds);
    const rgba = UTIF.toRGBA8(ifd);
    const canvas = el('canvas', 'dv-page dv-tiff-page');
    canvas.width = ifd.width;
    canvas.height = ifd.height;
    canvas.getContext('2d').putImageData(
      new ImageData(new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, ifd.width * ifd.height * 4), ifd.width, ifd.height),
      0, 0,
    );
    canvas.setAttribute('role', 'img');
    canvas.setAttribute('aria-label', `Page ${pages.length + 1}`);
    root.append(canvas);
    pages.push(canvas);
  }

  return {
    pages: () => pages,
    pageWidth: (canvas) => canvas.width,
    pageSize: (n) => ({ w: pages[n - 1].width, h: pages[n - 1].height }),
    naturalWidth: () => Math.max(...pages.map((c) => c.width)),
    thumbnail(n, width) {
      const src = pages[n - 1];
      const thumb = el('canvas');
      thumb.width = width * 2;
      thumb.height = Math.round((width * 2 * src.height) / src.width);
      thumb.getContext('2d').drawImage(src, 0, 0, thumb.width, thumb.height);
      return thumb;
    },
    printNodes: () => ({
      nodes: pages.map((canvas, i) => {
        const img = el('img');
        img.src = canvas.toDataURL('image/png');
        img.alt = `Page ${i + 1}`;
        return img;
      }),
      landscape: pages[0].width > pages[0].height,
    }),
    destroy() {},
  };
}
