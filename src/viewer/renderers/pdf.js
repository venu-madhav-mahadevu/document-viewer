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
export async function renderPdf({ buffer, mount, scroller, onState, askPassword }) {
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
  const rendered = new Map(); // pageNumber -> { task, textLayer }

  const observer = new IntersectionObserver((entries) => {
    for (const e of entries) {
      const n = Number(e.target.dataset.page);
      if (e.isIntersecting) renderPage(n);
    }
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

  async function renderPage(n) {
    if (rendered.has(n)) return;
    const holder = pages[n - 1];
    const entry = {};
    rendered.set(n, entry);
    const page = await pdf.getPage(n);
    const scale = zoom * CSS_UNITS;
    const vp = page.getViewport({ scale });
    const dpr = window.devicePixelRatio || 1;

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
  }

  function currentPage() {
    const top = scroller.scrollTop + scroller.clientHeight / 3;
    let n = 1;
    for (const p of pages) { if (p.offsetTop <= top) n = Number(p.dataset.page); else break; }
    return n;
  }

  const emit = () => onState({ page: currentPage(), pages: pdf.numPages });
  scroller.addEventListener('scroll', emit, { passive: true });

  layout();
  emit();

  return {
    get zoom() { return zoom; },
    setZoom(z) {
      const ratio = scroller.scrollTop / Math.max(1, scroller.scrollHeight);
      zoom = z;
      layout();
      scroller.scrollTop = ratio * scroller.scrollHeight;
      emit();
    },
    fitWidthZoom(available) {
      const widest = Math.max(...sizes.map((s) => s.w));
      return available / (widest * CSS_UNITS);
    },
    goToPage(n) {
      const p = pages[Math.min(Math.max(n, 1), pages.length) - 1];
      if (p) scroller.scrollTop = p.offsetTop - 16;
    },
    destroy() {
      scroller.removeEventListener('scroll', emit);
      observer.disconnect();
      for (const r of rendered.values()) { r.task?.cancel(); r.textLayer?.cancel(); }
      task.destroy(); // pdf.js v6: teardown lives on the loading task, not the document
    },
  };
}
