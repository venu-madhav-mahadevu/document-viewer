import UTIF from 'utif';
import { el } from '../sanitize.js';

/**
 * TIFF / multi-page scans (UTIF.js, MIT). Browsers can't display TIFF, so pages
 * are decoded to pixels and drawn on canvases. Supports the compressions
 * scanners use: CCITT G3/G4 (fax), LZW, PackBits, Deflate and JPEG.
 *
 * Large scans: pages are decoded only when they come near the screen, one at a
 * time with a pause between them (so the page never freezes), and pages far
 * off-screen are released, so memory stays flat however long the scan is.
 */
const KEEP_MARGIN = '1500px 0px'; // decode pages within this distance of the screen

export async function renderTiff({ buffer, mount, scroller }) {
  // Reading the page list is cheap: no pixels are decoded here.
  const ifds = UTIF.decode(buffer).filter((ifd) => ifd.t256 && !(ifd.t254 && (ifd.t254[0] & 1)));
  if (!ifds.length) throw new Error('This TIFF file contains no readable images.');

  const sizes = ifds.map((ifd) => ({ w: ifd.t256[0], h: ifd.t257 ? ifd.t257[0] : ifd.t256[0] }));
  const root = el('div', 'dv-tiff');
  mount.append(root);

  const pages = sizes.map((s, i) => {
    const holder = el('div', 'dv-page dv-tiff-page');
    holder.dataset.page = String(i + 1);
    holder.style.width = `${s.w}px`;
    holder.style.height = `${s.h}px`;
    holder.setAttribute('role', 'img');
    holder.setAttribute('aria-label', `Page ${i + 1}`);
    root.append(holder);
    return holder;
  });

  /** Decode page n into a canvas of the given width (full size when omitted). */
  function decodePage(n, width) {
    const ifd = ifds[n - 1];
    UTIF.decodeImage(buffer, ifd, ifds);
    const rgba = UTIF.toRGBA8(ifd);
    delete ifd.data; // UTIF keeps the decoded bits on the page object; free them now
    const full = el('canvas');
    full.width = ifd.width;
    full.height = ifd.height;
    full.getContext('2d').putImageData(
      new ImageData(new Uint8ClampedArray(rgba.buffer, rgba.byteOffset, ifd.width * ifd.height * 4), ifd.width, ifd.height),
      0, 0,
    );
    if (!width || width >= ifd.width) return full;
    const small = el('canvas');
    small.width = width;
    small.height = Math.round((width * ifd.height) / ifd.width);
    small.getContext('2d').drawImage(full, 0, 0, small.width, small.height);
    full.width = 0; // release the full-size pixels
    return small;
  }

  // One page at a time, yielding between pages so scrolling stays smooth.
  const wanted = new Set();
  let busy = false;
  const pump = () => {
    if (busy) return;
    const n = [...wanted][0];
    if (n == null) return;
    wanted.delete(n);
    busy = true;
    setTimeout(() => {
      try {
        const holder = pages[n - 1];
        if (holder.isConnected && !holder.firstChild) holder.append(decodePage(n));
      } catch (err) {
        console.warn(`TIFF page ${n} could not be decoded`, err);
      } finally {
        busy = false;
        pump();
      }
    }, 0);
  };

  const observer = new IntersectionObserver((entries) => {
    for (const e of entries) {
      const n = Number(e.target.dataset.page);
      if (e.isIntersecting) {
        wanted.add(n);
      } else {
        wanted.delete(n);
        const canvas = e.target.firstChild;
        if (canvas) { canvas.width = 0; canvas.remove(); } // far away: free its memory
      }
    }
    pump();
  }, { root: scroller, rootMargin: KEEP_MARGIN });
  pages.forEach((p) => observer.observe(p));

  const yieldToBrowser = () => new Promise((r) => setTimeout(r, 0));

  return {
    pages: () => pages,
    pageWidth: (holder) => sizes[Number(holder.dataset.page) - 1].w,
    pageSize: (n) => sizes[n - 1],
    naturalWidth: () => Math.max(...sizes.map((s) => s.w)),
    async thumbnail(n, width) {
      await yieldToBrowser();
      return decodePage(n, width * 2);
    },
    async printNodes() {
      const nodes = [];
      for (let n = 1; n <= ifds.length; n++) {
        await yieldToBrowser();
        const canvas = decodePage(n, Math.min(sizes[n - 1].w, 1700)); // ~200 dpi is plenty for paper
        const img = el('img');
        img.src = canvas.toDataURL('image/png');
        img.alt = `Page ${n}`;
        canvas.width = 0;
        nodes.push(img);
      }
      return { nodes, landscape: sizes[0].w > sizes[0].h };
    },
    destroy() {
      observer.disconnect();
      wanted.clear();
      pages.forEach((p) => { const c = p.firstChild; if (c) c.width = 0; });
    },
  };
}
