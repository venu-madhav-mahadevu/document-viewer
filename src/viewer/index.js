import './viewer.css';
import { detectKind } from './detect.js';
import { el } from './sanitize.js';

// Renderers are loaded on demand so opening a PDF never downloads the spreadsheet engine.
const RENDERERS = {
  pdf: () => import('./renderers/pdf.js').then((m) => m.renderPdf),
  docx: () => import('./renderers/docx.js').then((m) => m.renderDocx),
  sheet: () => import('./renderers/sheet.js').then((m) => m.renderSheet),
  pptx: () => import('./renderers/pptx.js').then((m) => m.renderPptx),
  image: () => import('./renderers/basic.js').then((m) => m.renderImage),
  text: () => import('./renderers/basic.js').then((m) => m.renderText),
};

const LABELS = { pdf: 'PDF', docx: 'Word', sheet: 'Spreadsheet', pptx: 'Presentation', image: 'Image', text: 'Text' };
const ZOOM_STEPS = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];

/**
 * Create a document viewer inside `container`.
 *
 *   const viewer = createDocViewer(document.getElementById('viewer'));
 *   await viewer.open(file);                          // File from <input> / drag-drop
 *   await viewer.open(blob, 'statement.pdf');         // Blob or ArrayBuffer + name
 *   await viewer.open('/api/docs/123', 'report.xlsx'); // same-origin URL
 *
 * Everything renders in the browser. No document content is sent anywhere.
 *
 * @param {HTMLElement} container
 * @param {{ maxBytes?: number, download?: boolean, fetchInit?: RequestInit }} [options]
 */
export function createDocViewer(container, options = {}) {
  const { maxBytes = 200 * 1024 * 1024, download = true, fetchInit } = options;

  const root = el('div', 'dv-root');
  const toolbar = el('div', 'dv-toolbar');
  toolbar.setAttribute('role', 'toolbar');
  const title = el('div', 'dv-title');
  const badge = el('span', 'dv-badge');
  const nameEl = el('span', 'dv-name');
  title.append(badge, nameEl);

  const pageBox = el('div', 'dv-group dv-pagebox');
  const pageInput = el('input', 'dv-pageinput');
  Object.assign(pageInput, { type: 'number', min: '1', ariaLabel: 'Page number' });
  const pageTotal = el('span', 'dv-pagetotal');
  pageBox.append(pageInput, pageTotal);

  const zoomBox = el('div', 'dv-group');
  const btn = (label, aria, cls = 'dv-btn') => {
    const b = el('button', cls, label);
    b.type = 'button';
    b.setAttribute('aria-label', aria);
    b.title = aria;
    return b;
  };
  const zoomOut = btn('−', 'Zoom out');
  const zoomLabel = btn('100%', 'Fit to width', 'dv-btn dv-zoomlabel');
  const zoomIn = btn('+', 'Zoom in');
  zoomBox.append(zoomOut, zoomLabel, zoomIn);

  const downloadBtn = btn('Download', 'Download original file', 'dv-btn dv-download');
  downloadBtn.hidden = !download;

  toolbar.append(title, pageBox, zoomBox, downloadBtn);
  const scroller = el('div', 'dv-scroller');
  scroller.tabIndex = 0;
  const stage = el('div', 'dv-stage');
  scroller.append(stage);
  const status = el('div', 'dv-status');
  status.setAttribute('role', 'status');
  root.append(toolbar, scroller, status);
  container.replaceChildren(root);

  let current = null; // { controller, kind, blob, name }
  let zoom = 1;
  let openToken = 0;

  function setStatus(message, kind = 'info') {
    status.replaceChildren();
    status.className = `dv-status dv-status-${kind}`;
    status.hidden = !message;
    if (!message) return;
    if (kind === 'loading') status.append(el('span', 'dv-spinner'));
    status.append(el('span', null, message));
  }

  function applyZoom(z) {
    zoom = Math.min(Math.max(z, ZOOM_STEPS[0]), ZOOM_STEPS[ZOOM_STEPS.length - 1]);
    zoomLabel.textContent = `${Math.round(zoom * 100)}%`;
    if (!current) return;
    if (current.controller.setZoom) current.controller.setZoom(zoom);
    // Spreadsheets zoom the grid only, so the sheet tabs stay a usable size.
    else if (current.kind === 'sheet') stage.querySelector('.dv-sheet-grid').style.zoom = String(zoom);
    else stage.style.zoom = String(zoom);
  }

  function fitWidth() {
    if (!current) return;
    const available = scroller.clientWidth - 48;
    const c = current.controller;
    if (c.fitWidthZoom) applyZoom(c.fitWidthZoom(available));
    else if (c.naturalWidth) applyZoom(Math.min(available / c.naturalWidth(), current.kind === 'image' ? 1 : 2));
    else applyZoom(1);
  }

  const step = (dir) => {
    const next = dir > 0 ? ZOOM_STEPS.find((s) => s > zoom + 0.001) : [...ZOOM_STEPS].reverse().find((s) => s < zoom - 0.001);
    applyZoom(next ?? zoom);
  };
  zoomIn.addEventListener('click', () => step(1));
  zoomOut.addEventListener('click', () => step(-1));
  zoomLabel.addEventListener('click', fitWidth);
  pageInput.addEventListener('change', () => current?.controller.goToPage?.(Number(pageInput.value)));
  downloadBtn.addEventListener('click', () => {
    if (!current) return;
    const a = el('a');
    a.href = URL.createObjectURL(current.blob);
    a.download = current.name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  });
  root.addEventListener('keydown', (e) => {
    if (e.target instanceof HTMLInputElement) return;
    if ((e.ctrlKey || e.metaKey) && (e.key === '=' || e.key === '+')) { e.preventDefault(); step(1); }
    else if ((e.ctrlKey || e.metaKey) && e.key === '-') { e.preventDefault(); step(-1); }
    else if ((e.ctrlKey || e.metaKey) && e.key === '0') { e.preventDefault(); fitWidth(); }
  });

  function onState({ page, pages }) {
    pageBox.hidden = !pages;
    if (document.activeElement !== pageInput) pageInput.value = String(page);
    pageInput.max = String(pages);
    pageTotal.textContent = `/ ${pages}`;
  }

  function askPassword(retry) {
    return new Promise((resolve) => {
      setStatus('');
      const form = el('form', 'dv-password');
      form.append(el('p', null, retry ? 'Incorrect password. Try again.' : 'This document is password protected.'));
      const input = el('input');
      Object.assign(input, { type: 'password', placeholder: 'Password', autocomplete: 'off' });
      const ok = btn('Open', 'Open document', 'dv-btn dv-primary');
      ok.type = 'submit';
      const cancel = btn('Cancel', 'Cancel');
      form.append(input, ok, cancel);
      stage.replaceChildren(form);
      input.focus();
      form.addEventListener('submit', (e) => { e.preventDefault(); stage.replaceChildren(); setStatus('Opening…', 'loading'); resolve(input.value); });
      cancel.addEventListener('click', () => { stage.replaceChildren(); resolve(null); });
    });
  }

  function reset() {
    try { current?.controller.destroy?.(); } catch (err) { console.warn('Viewer cleanup failed', err); }
    current = null;
    stage.replaceChildren();
    stage.style.zoom = '';
    stage.className = 'dv-stage';
    pageBox.hidden = true;
    zoom = 1;
    zoomLabel.textContent = '100%';
  }

  async function toBlob(source, name) {
    if (typeof source === 'string') {
      const url = new URL(source, location.href);
      if (url.origin !== location.origin) throw new Error('Only same-origin document URLs can be opened.');
      const res = await fetch(url, { credentials: 'same-origin', ...fetchInit });
      if (!res.ok) throw new Error(`Could not load the document (HTTP ${res.status}).`);
      return { blob: await res.blob(), name: name || decodeURIComponent(url.pathname.split('/').pop() || 'document') };
    }
    if (source instanceof ArrayBuffer) return { blob: new Blob([source]), name: name || 'document' };
    return { blob: source, name: name || source.name || 'document' };
  }

  return {
    element: root,

    /** Open a File, Blob, ArrayBuffer or same-origin URL. */
    async open(source, name) {
      const token = ++openToken;
      reset();
      setStatus('Opening…', 'loading');
      try {
        const { blob, name: fileName } = await toBlob(source, name);
        if (blob.size > maxBytes) throw new Error(`This file is larger than ${Math.round(maxBytes / 1048576)} MB.`);
        const buffer = await blob.arrayBuffer();
        const { kind, legacy } = await detectKind(fileName, buffer);
        if (token !== openToken) return;

        nameEl.textContent = fileName;
        badge.textContent = LABELS[kind] || 'File';
        badge.dataset.kind = kind;
        root.dataset.kind = kind;

        if (kind === 'unsupported') {
          current = { controller: {}, kind, blob, name: fileName };
          throw new Error(legacy
            ? `${legacy} files can't be previewed. Save it as a modern Office file (.docx / .pptx) or PDF, or download it.`
            : 'This file type can\'t be previewed. You can still download it.');
        }

        const render = await RENDERERS[kind]();
        stage.classList.add(`dv-stage-${kind}`);
        const controller = await render({ buffer, name: fileName, mount: stage, scroller, onState, askPassword });
        if (token !== openToken) { controller.destroy?.(); return; }
        current = { controller, kind, blob, name: fileName };
        setStatus('');
        if (kind === 'sheet' || kind === 'text') applyZoom(1);
        else fitWidth();
        scroller.scrollTop = 0;
        scroller.focus({ preventScroll: true });
      } catch (err) {
        if (token !== openToken) return;
        console.error(err);
        const msg = err?.name === 'PasswordException' ? 'The document could not be opened without its password.'
          : err?.name === 'InvalidPDFException' ? 'This PDF is damaged or not a valid PDF.'
            : err?.message || 'This document could not be opened.';
        setStatus(msg, 'error');
      }
    },

    zoomIn: () => step(1),
    zoomOut: () => step(-1),
    fitWidth,
    setZoom: applyZoom,

    destroy() {
      openToken++;
      reset();
      container.replaceChildren();
    },
  };
}
