import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import 'pdfjs-dist/web/pdf_viewer.css';
import { el } from '../sanitize.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

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
        await Promise.all([entry.task.promise, entry.textLayer.render(), addLinks(page, vp, holder)]);
      } catch (err) {
        if (!(err instanceof pdfjsLib.RenderingCancelledException)) console.warn(err);
      }
    })();
    return entry.done;
  }

  /** Clickable links over the page: web links open in a new tab, internal ones scroll to their page. */
  async function addLinks(page, vp, holder) {
    const links = (await page.getAnnotations({ intent: 'display' })).filter((a) => a.subtype === 'Link');
    if (!links.length) return;
    const layer = el('div', 'dv-pdf-links');
    for (const link of links) {
      // PDF coordinates -> page pixels with the viewport's matrix (handles zoom and rotated pages)
      const [a0, b0, c0, d0, e0, f0] = vp.transform;
      const at = (x, y) => [a0 * x + c0 * y + e0, b0 * x + d0 * y + f0];
      const [x1, y1] = at(link.rect[0], link.rect[1]);
      const [x2, y2] = at(link.rect[2], link.rect[3]);
      const a = el('a', 'dv-pdf-link');
      Object.assign(a.style, {
        left: `${Math.min(x1, x2)}px`, top: `${Math.min(y1, y2)}px`,
        width: `${Math.abs(x2 - x1)}px`, height: `${Math.abs(y2 - y1)}px`,
      });
      const url = link.url || link.unsafeUrl;
      if (url && /^(https?:|mailto:)/i.test(url)) {
        a.href = url;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        a.title = url;
      } else if (link.dest) {
        a.href = '#';
        a.addEventListener('click', async (e) => {
          e.preventDefault();
          const dest = typeof link.dest === 'string' ? await pdf.getDestination(link.dest) : link.dest;
          if (!dest) return;
          const n = typeof dest[0] === 'number' ? dest[0] + 1 : (await pdf.getPageIndex(dest[0])) + 1;
          pages[n - 1]?.scrollIntoView({ block: 'start' });
        });
        a.title = 'Go to this place in the document';
      } else continue;
      layer.append(a);
    }
    if (holder.isConnected) holder.append(layer);
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
    /** Bookmarks (the PDF's outline) as a flat list: { title, level, page }. */
    async outline() {
      const tree = await pdf.getOutline().catch(() => null);
      const out = [];
      const pageOf = async (dest) => {
        try {
          const d = typeof dest === 'string' ? await pdf.getDestination(dest) : dest;
          if (!d) return null;
          return typeof d[0] === 'number' ? d[0] + 1 : (await pdf.getPageIndex(d[0])) + 1;
        } catch { return null; }
      };
      const walk = async (items, level) => {
        for (const item of items || []) {
          if (out.length >= 1000) return;
          out.push({ title: item.title || '(untitled)', level, page: await pageOf(item.dest) });
          await walk(item.items, level + 1);
        }
      };
      await walk(tree, 0);
      return out;
    },
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
