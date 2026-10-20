import * as pdfjsLib from 'pdfjs-dist';
import { el } from '../sanitize.js';

// The worker is copied to public/pdfjs by scripts/copy-pdfjs-assets.mjs; pdf_viewer.css is in angular.json.
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL('pdfjs/pdf.worker.min.mjs', document.baseURI).href;

// Absolute URLs: pdf.js may resolve these from inside the worker.
const asset = (dir) => new URL(`pdfjs/${dir}/`, document.baseURI).href;
const CSS_UNITS = pdfjsLib.PixelsPerInch.PDF_TO_CSS_UNITS; // 96/72 so 100% = real size

/**
 * Continuous-scroll PDF view. Pages are laid out as sized placeholders and only
 * rendered when they approach the viewport, so 500-page statements stay fast.
 */
export async function renderPdf({ buffer, mount, scroller, askPassword }) {
  const task = pdfjsLib.getDocument({
    data: new Uint8Array(buffer),
    isEvalSupported: false,
    enableXfa: false,
    cMapUrl: asset('cmaps'),
    cMapPacked: true,
    standardFontDataUrl: asset('standard_fonts'),
    wasmUrl: asset('wasm'),
    iccUrl: asset('iccs'),
  });
  task.onPassword = (update, reason) => {
    const retry = reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD;
    askPassword(retry).then((pw) => (pw == null ? task.destroy() : update(pw)));
  };
  const pdf = await task.promise;

  const root = el('div', 'dv-pdf');
  mount.append(root);

  const sizes = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    // Page sizes are cheap to read and let placeholders have the right height.
    const page = await pdf.getPage(i);
    const vp = page.getViewport({ scale: 1 });
    sizes.push({ w: vp.width, h: vp.height });
  }

  let zoom = 1;
  let pages = [];
  const rendered = new Map(); // pageNumber -> { task, textLayer, done }
  const pageText = new Map(); // pageNumber -> lowercase text, for search

  const observer = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) renderPage(Number(e.target.dataset.page));
  }, { root: scroller, rootMargin: '800px 0px' });

  function layout() {
    observer.disconnect();
    for (const r of rendered.values()) { r.task?.cancel(); r.textLayer?.cancel(); }
    rendered.clear();
    root.replaceChildren();
    pages = sizes.map((s, i) => {
      const p = el('div', 'dv-page dv-pdf-page');
      p.dataset.page = String(i + 1);
      p.style.width = `${Math.floor(s.w * zoom * CSS_UNITS)}px`;
      p.style.height = `${Math.floor(s.h * zoom * CSS_UNITS)}px`;
      root.append(p);
      observer.observe(p);
      return p;
    });
  }

  function renderPage(n) {
    if (rendered.has(n)) return rendered.get(n).done;
    const entry = {};
    rendered.set(n, entry);
    entry.done = (async () => {
      const holder = pages[n - 1];
      const page = await pdf.getPage(n);
      const scale = zoom * CSS_UNITS;
      const vp = page.getViewport({ scale });
      // iOS Safari draws nothing on a canvas over ~16.7 million pixels, which a
      // zoomed page on a 3x-density iPhone exceeds. Lower the sharpness instead.
      const MAX_CANVAS_PIXELS = 16_777_216;
      let dpr = window.devicePixelRatio || 1;
      if (vp.width * vp.height * dpr * dpr > MAX_CANVAS_PIXELS) {
        dpr = Math.sqrt(MAX_CANVAS_PIXELS / (vp.width * vp.height));
      }

      const canvas = el('canvas');
      canvas.width = Math.floor(vp.width * dpr);
      canvas.height = Math.floor(vp.height * dpr);
      canvas.style.width = `${Math.floor(vp.width)}px`;
      canvas.style.height = `${Math.floor(vp.height)}px`;

      const textDiv = el('div', 'textLayer');
      holder.style.setProperty('--scale-factor', String(scale));
      holder.style.setProperty('--total-scale-factor', String(scale));
      holder.replaceChildren(canvas, textDiv);

      entry.task = page.render({
        canvas,
        viewport: vp,
        transform: dpr !== 1 ? [dpr, 0, 0, dpr, 0, 0] : undefined,
      });
      entry.textLayer = new pdfjsLib.TextLayer({
        textContentSource: page.streamTextContent(),
        container: textDiv,
        viewport: vp,
      });
      try {
        await Promise.all([entry.task.promise, entry.textLayer.render()]);
      } catch (err) {
        if (!(err instanceof pdfjsLib.RenderingCancelledException)) console.warn(err);
      }
    })();
    return entry.done;
  }

  async function textOf(n) {
    if (!pageText.has(n)) {
      const content = await (await pdf.getPage(n)).getTextContent();
      pageText.set(n, content.items.map((item) => item.str).join('\n').toLowerCase());
    }
    return pageText.get(n);
  }

  /** Render one page off-screen to a canvas of the given CSS width. */
  // intent 'print' renders without requestAnimationFrame pacing, so it also
  // completes in background tabs (where browsers pause animation frames).
  async function snapshot(n, cssWidth, quality = 1, intent = 'display') {
    const page = await pdf.getPage(n);
    const scale = cssWidth / page.getViewport({ scale: 1 }).width;
    const vp = page.getViewport({ scale: scale * quality });
    const canvas = el('canvas');
    canvas.width = Math.floor(vp.width);
    canvas.height = Math.floor(vp.height);
    await page.render({ canvas, viewport: vp, intent }).promise;
    return canvas;
  }

  layout();

  return {
    get zoom() { return zoom; },
    pages: () => pages,
    pageSize: (n) => sizes[n - 1],
    ensureRendered: renderPage,
    setZoom(z) {
      const ratio = scroller.scrollTop / Math.max(1, scroller.scrollHeight);
      zoom = z;
      layout();
      scroller.scrollTop = ratio * scroller.scrollHeight;
    },
    fitWidthZoom(available) {
      const widest = Math.max(...sizes.map((s) => s.w));
      return available / (widest * CSS_UNITS);
    },
    /** Occurrences per page, read from the text content (works for pages not yet drawn). */
    async searchText(term) {
      const needle = term.toLowerCase();
      const hits = [];
      for (let n = 1; n <= pdf.numPages; n++) {
        const text = await textOf(n);
        let count = 0;
        for (let i = text.indexOf(needle); i !== -1; i = text.indexOf(needle, i + needle.length)) count++;
        if (count) hits.push({ page: n, count });
      }
      return hits;
    },
    thumbnail: (n, width) => snapshot(n, width, Math.min(window.devicePixelRatio || 1, 2)),
    async printNodes() {
      const imgs = [];
      for (let n = 1; n <= pdf.numPages; n++) {
        const canvas = await snapshot(n, 816, 2, 'print'); // ~Letter width at 2x for crisp print
        const img = el('img');
        img.src = canvas.toDataURL('image/png');
        img.alt = `Page ${n}`;
        imgs.push(img);
      }
      return { nodes: imgs, landscape: sizes[0].w > sizes[0].h };
    },
    destroy() {
      observer.disconnect();
      for (const r of rendered.values()) { r.task?.cancel(); r.textLayer?.cancel(); }
      task.destroy(); // pdf.js v6: teardown lives on the loading task, not the document
    },
  };
}
