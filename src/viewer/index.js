import './viewer.css';
import { detectKind } from './detect.js';
import { FORMATS, SUPPORTED_EXTENSIONS } from './formats.js'; // all libraries and formats, declared once
import { el } from './sanitize.js';
import { icon, setIcon } from './icons.js';
import { findRanges, paintHighlights, clearHighlights, revealRange } from './search.js';
import { createThumbnails, cloneThumbnail } from './thumbs.js';
import { printDocument } from './print.js';

/** 1536 → "1.5 KB", 52428800 → "50 MB". */
export function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KB', 'MB', 'GB'];
  let value = bytes / 1024;
  let i = 0;
  while (value >= 1024 && i < units.length - 1) { value /= 1024; i++; }
  return `${value >= 100 ? Math.round(value) : value.toFixed(1).replace(/\.0$/, '')} ${units[i]}`;
}

/** 350 → "0.35 s", 2140 → "2.1 s", 21700 → "22 s". */
export function formatDuration(ms) {
  if (ms < 1000) return `${(ms / 1000).toFixed(2)} s`;
  if (ms < 10000) return `${(ms / 1000).toFixed(1)} s`;
  return `${Math.round(ms / 1000)} s`;
}

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
 * @param {{ maxBytes?: number, download?: boolean, print?: boolean, fetchInit?: RequestInit, onBack?: () => void, badge?: boolean }} [options]
 *   onBack: when given, a Back button appears at the left of the toolbar and calls it.
 *   badge: show the format label (PDF, Word, Spreadsheet...) next to the file name. Off by default.
 */
export function createDocViewer(container, options = {}) {
  const { maxBytes = 200 * 1024 * 1024, download = true, print = true, fetchInit, onBack, badge: showBadge = false } = options;

  const iconButton = (name, label, extraClass = '') => {
    const b = el('button', `dv-icon-btn ${extraClass}`.trim());
    b.type = 'button';
    b.title = label;
    b.setAttribute('aria-label', label);
    b.append(icon(name));
    return b;
  };
  const textButton = (label, aria, cls = 'dv-btn') => {
    const b = el('button', cls, label);
    b.type = 'button';
    b.setAttribute('aria-label', aria);
    b.title = aria;
    return b;
  };

  // ---------- Layout ----------
  const root = el('div', 'dv-root');
  root.tabIndex = -1;

  const toolbar = el('div', 'dv-toolbar');
  toolbar.setAttribute('role', 'toolbar');
  const thumbsBtn = iconButton('thumbnails', 'Show page thumbnails');
  thumbsBtn.setAttribute('aria-pressed', 'false');
  const title = el('div', 'dv-title');
  const badge = el('span', 'dv-badge');
  badge.hidden = !showBadge; // kept in the DOM (the compare and device pages read it)
  const nameEl = el('span', 'dv-name');
  const sizeEl = el('span', 'dv-size');
  title.append(badge, nameEl, sizeEl);
  const searchBtn = iconButton('search', 'Find in document (Ctrl+F)');
  const printBtn = iconButton('print', 'Print (Ctrl+P)');
  const downloadBtn = iconButton('download', 'Download original file');
  const fullscreenBtn = iconButton('fullscreen', 'Full screen');
  printBtn.hidden = !print;
  downloadBtn.hidden = !download;
  fullscreenBtn.hidden = !document.fullscreenEnabled;
  const actions = el('div', 'dv-actions');
  actions.append(searchBtn, printBtn, downloadBtn, fullscreenBtn);
  const backBtn = el('button', 'dv-back');
  backBtn.type = 'button';
  backBtn.append(icon('back'), el('span', null, 'Back'));
  backBtn.title = 'Back';
  backBtn.hidden = typeof onBack !== 'function';
  backBtn.addEventListener('click', () => onBack?.());
  toolbar.append(backBtn, thumbsBtn, title, actions);

  const searchBox = el('form', 'dv-search');
  searchBox.hidden = true;
  searchBox.setAttribute('role', 'search');
  const searchInput = el('input');
  Object.assign(searchInput, { type: 'search', placeholder: 'Find in document', autocomplete: 'off' });
  searchInput.setAttribute('aria-label', 'Find in document');
  const searchCount = el('span', 'dv-search-count');
  searchCount.setAttribute('aria-live', 'polite');
  const prevBtn = iconButton('up', 'Previous match (Shift+Enter)');
  const nextBtn = iconButton('down', 'Next match (Enter)');
  const closeSearchBtn = iconButton('close', 'Close search (Esc)');
  searchBox.append(searchInput, searchCount, prevBtn, nextBtn, closeSearchBtn);

  const body = el('div', 'dv-body');
  const thumbsPanel = el('nav', 'dv-thumbs');
  thumbsPanel.setAttribute('aria-label', 'Page thumbnails');
  thumbsPanel.hidden = true;
  const scroller = el('div', 'dv-scroller');
  scroller.tabIndex = 0;
  const stage = el('div', 'dv-stage');
  scroller.append(stage);
  body.append(thumbsPanel, scroller);

  // Floating page / zoom pill, Google Drive style.
  const pill = el('div', 'dv-pill');
  pill.hidden = true;
  const pageBox = el('div', 'dv-pill-group dv-pagebox');
  const pageLabel = el('span', 'dv-pill-label', 'Page');
  const pageInput = el('input', 'dv-pageinput');
  Object.assign(pageInput, { type: 'number', min: '1' });
  pageInput.setAttribute('aria-label', 'Page number');
  const pageTotal = el('span', 'dv-pagetotal');
  pageBox.append(pageLabel, pageInput, pageTotal);
  const zoomBox = el('div', 'dv-pill-group');
  const zoomOut = iconButton('minus', 'Zoom out (Ctrl+−)');
  const zoomLabel = textButton('100%', 'Fit to width (Ctrl+0)', 'dv-btn dv-zoomlabel');
  const zoomIn = iconButton('plus', 'Zoom in (Ctrl++)');
  zoomBox.append(zoomOut, zoomLabel, zoomIn);
  pill.append(pageBox, el('span', 'dv-pill-sep'), zoomBox);

  const status = el('div', 'dv-status');
  status.setAttribute('role', 'status');
  status.hidden = true;
  root.append(toolbar, searchBox, body, pill, status);
  container.replaceChildren(root);

  // ---------- State ----------
  let current = null; // { controller, kind, blob, name }
  let zoom = 1;
  let openToken = 0;
  let thumbs = null;
  let search = { term: '', hits: [], index: -1, token: 0 };

  function setStatus(message, kind = 'info') {
    status.replaceChildren();
    status.className = `dv-status dv-status-${kind}`;
    status.hidden = !message;
    if (!message) return;
    if (kind === 'loading') status.append(el('span', 'dv-spinner'));
    status.append(el('span', null, message));
  }

  // ---------- Pages ----------
  const pageEls = () => current?.controller.pages?.() || [];

  function currentPage() {
    const els = pageEls();
    if (!els.length) return 0;
    const line = scroller.getBoundingClientRect().top + scroller.clientHeight / 3;
    let n = 1;
    for (let i = 0; i < els.length; i++) {
      if (els[i].getBoundingClientRect().top <= line) n = i + 1;
      else break;
    }
    return n;
  }

  function goToPage(n) {
    const els = pageEls();
    const target = els[Math.min(Math.max(n, 1), els.length) - 1];
    if (!target) return;
    scroller.scrollTop += target.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 12;
    updatePageUi();
  }

  function updatePageUi() {
    const total = pageEls().length;
    pageBox.hidden = !total;
    if (!total) return;
    const n = currentPage();
    if (document.activeElement !== pageInput) pageInput.value = String(n);
    pageInput.max = String(total);
    pageTotal.textContent = `/ ${total}`;
    thumbs?.setActive(n);
  }

  let scrollQueued = false;
  scroller.addEventListener('scroll', () => {
    if (scrollQueued) return;
    scrollQueued = true;
    requestAnimationFrame(() => { scrollQueued = false; updatePageUi(); });
  }, { passive: true });

  function buildThumbnails() {
    thumbs?.destroy();
    thumbs = null;
    const c = current.controller;
    const els = pageEls();
    thumbsBtn.hidden = !els.length;
    if (!els.length) { showThumbs(false); return; }
    thumbs = createThumbnails(thumbsPanel, {
      count: els.length,
      ratio: (n) => {
        if (c.pageSize) { const s = c.pageSize(n); return s.h / s.w; }
        const p = pageEls()[n - 1];
        return p.offsetHeight / Math.max(1, p.offsetWidth);
      },
      render: (n, width) => (c.thumbnail
        ? c.thumbnail(n, width)
        : cloneThumbnail(pageEls()[n - 1], width, c.pageWidth(pageEls()[n - 1]))),
      onPick: (n) => {
        goToPage(n);
        if (root.clientWidth < 700) showThumbs(false); // the panel overlays the page on phones
      },
    });
    // Open by default on wide screens, like Drive's preview.
    showThumbs(root.clientWidth >= 900 && els.length > 1);
  }

  function showThumbs(on) {
    thumbsPanel.hidden = !on;
    thumbsBtn.setAttribute('aria-pressed', String(on));
    root.classList.toggle('dv-has-thumbs', on);
  }

  // ---------- Zoom ----------
  function applyZoom(z) {
    zoom = Math.min(Math.max(z, ZOOM_STEPS[0]), ZOOM_STEPS[ZOOM_STEPS.length - 1]);
    zoomLabel.textContent = `${Math.round(zoom * 100)}%`;
    if (!current) return;
    if (current.controller.setZoom) current.controller.setZoom(zoom);
    // Spreadsheets zoom the grid only, so the sheet tabs stay a usable size.
    else if (current.kind === 'sheet') stage.querySelector('.dv-sheet-grid').style.zoom = String(zoom);
    else stage.style.zoom = String(zoom);
    updatePageUi();
    if (search.hits.length && current.kind === 'pdf') selectHit(search.index); // PDF pages were re-laid out
  }

  // While true, the document re-fits whenever the viewer is resized (window
  // resize, phone rotation, thumbnails toggled). Manual zoom turns it off.
  let autoFit = false;

  function fitWidth() {
    if (!current?.controller) return;
    autoFit = true;
    const available = scroller.clientWidth - 48;
    if (available <= 0) return; // not laid out yet; the resize observer will fit later
    const c = current.controller;
    if (c.fitWidthZoom) applyZoom(c.fitWidthZoom(available));
    else if (c.naturalWidth) applyZoom(Math.min(available / c.naturalWidth(), current.kind === 'image' ? 1 : 2));
    else applyZoom(1);
  }

  const step = (dir) => {
    autoFit = false;
    const next = dir > 0 ? ZOOM_STEPS.find((s) => s > zoom + 0.001) : [...ZOOM_STEPS].reverse().find((s) => s < zoom - 0.001);
    applyZoom(next ?? zoom);
  };

  // Ctrl + mouse wheel and touchpad pinch (which browsers report as Ctrl + wheel)
  // zoom the document instead of the whole page, like Google Docs. One step per
  // 150 ms so a fast wheel doesn't jump straight to 400%.
  let lastWheelZoom = 0;
  scroller.addEventListener('wheel', (e) => {
    if (!e.ctrlKey || !current?.controller) return;
    e.preventDefault();
    const now = performance.now();
    if (now - lastWheelZoom < 150) return;
    lastWheelZoom = now;
    step(e.deltaY < 0 ? 1 : -1);
  }, { passive: false });

  let lastWidth = 0;
  new ResizeObserver(() => {
    const width = scroller.clientWidth;
    if (width === lastWidth) return;
    lastWidth = width;
    if (autoFit) fitWidth();
  }).observe(scroller);

  // ---------- Search ----------
  function openSearch() {
    if (!current?.controller) return;
    searchBox.hidden = false;
    searchInput.focus();
    searchInput.select();
  }

  function closeSearch() {
    searchBox.hidden = true;
    search = { term: '', hits: [], index: -1, token: search.token + 1 };
    clearHighlights();
    searchCount.textContent = '';
    scroller.focus({ preventScroll: true });
  }

  function updateCount() {
    const { term, hits, index } = search;
    searchCount.textContent = !term ? '' : hits.length ? `${index + 1} of ${hits.length}` : 'No results';
    prevBtn.disabled = nextBtn.disabled = hits.length < 2;
  }

  async function runSearch(term) {
    const token = ++search.token;
    clearHighlights();
    search = { term, hits: [], index: -1, token };
    if (!term || !current) return updateCount();
    const c = current.controller;
    let hits;
    if (c.searchText) {
      searchCount.textContent = 'Searching…';
      const perPage = await c.searchText(term);
      hits = perPage.flatMap(({ page, count }) => Array.from({ length: count }, (_, k) => ({ page, k })));
    } else {
      hits = findRanges(stage, term);
    }
    if (token !== search.token) return;
    search.hits = hits;
    if (hits.length) await selectHit(0);
    updateCount();
  }

  async function selectHit(i) {
    const { hits, term } = search;
    if (!hits.length) return;
    search.index = (i + hits.length) % hits.length;
    const hit = hits[search.index];
    if (hit instanceof Range) {
      paintHighlights(hits, hit);
      revealRange(hit);
    } else {
      // PDF: jump to the page, make sure its text layer exists, then highlight there.
      const token = search.token;
      goToPage(hit.page);
      await current.controller.ensureRendered(hit.page);
      if (token !== search.token) return;
      const ranges = findRanges(pageEls()[hit.page - 1], term);
      if (ranges.length) {
        const range = ranges[Math.min(hit.k, ranges.length - 1)];
        paintHighlights(ranges, range);
        revealRange(range);
      }
    }
    updateCount();
  }

  let searchTimer;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => runSearch(searchInput.value.trim()), 200);
  });
  searchBox.addEventListener('submit', (e) => e.preventDefault());
  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      if (searchInput.value.trim() !== search.term) runSearch(searchInput.value.trim());
      else selectHit(search.index + (e.shiftKey ? -1 : 1));
    } else if (e.key === 'Escape') {
      e.preventDefault();
      closeSearch();
    }
  });
  prevBtn.addEventListener('click', () => selectHit(search.index - 1));
  nextBtn.addEventListener('click', () => selectHit(search.index + 1));
  closeSearchBtn.addEventListener('click', closeSearch);
  searchBtn.addEventListener('click', () => (searchBox.hidden ? openSearch() : closeSearch()));

  // ---------- Print / download / full screen ----------
  async function printCurrent() {
    if (!current?.controller || !print) return;
    setStatus('Preparing to print…', 'loading');
    try {
      const c = current.controller;
      const job = c.printNodes ? await c.printNodes() : { nodes: [...stage.children].map((n) => n.cloneNode(true)) };
      setStatus('');
      await printDocument({
        title: current.name,
        nodes: job.nodes,
        paged: !!FORMATS[current.kind]?.paged,
        landscape: job.landscape,
      });
    } catch (err) {
      console.error(err);
      setStatus('Printing failed.', 'error');
    }
  }

  function downloadCurrent() {
    if (!current) return;
    const a = el('a');
    a.href = URL.createObjectURL(current.blob);
    a.download = current.name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function toggleFullscreen() {
    if (document.fullscreenElement === root) document.exitFullscreen();
    else root.requestFullscreen?.().catch(() => {});
  }

  document.addEventListener('fullscreenchange', () => {
    const on = document.fullscreenElement === root;
    setIcon(fullscreenBtn, on ? 'fullscreenExit' : 'fullscreen');
    fullscreenBtn.title = on ? 'Exit full screen' : 'Full screen';
    fullscreenBtn.setAttribute('aria-label', fullscreenBtn.title);
    if (current) setTimeout(fitWidth, 100);
  });

  // ---------- Wiring ----------
  zoomIn.addEventListener('click', () => step(1));
  zoomOut.addEventListener('click', () => step(-1));
  zoomLabel.addEventListener('click', fitWidth);
  pageInput.addEventListener('change', () => goToPage(Number(pageInput.value)));
  thumbsBtn.addEventListener('click', () => showThumbs(thumbsPanel.hidden));
  printBtn.addEventListener('click', printCurrent);
  downloadBtn.addEventListener('click', downloadCurrent);
  fullscreenBtn.addEventListener('click', toggleFullscreen);

  root.addEventListener('keydown', (e) => {
    const mod = e.ctrlKey || e.metaKey;
    const inField = e.target instanceof HTMLInputElement;
    if (mod && e.key.toLowerCase() === 'f') { e.preventDefault(); openSearch(); }
    else if (mod && e.key.toLowerCase() === 'p') { e.preventDefault(); printCurrent(); }
    else if (mod && (e.key === '=' || e.key === '+')) { e.preventDefault(); step(1); }
    else if (mod && e.key === '-') { e.preventDefault(); step(-1); }
    else if (mod && e.key === '0') { e.preventDefault(); fitWidth(); }
    else if (e.key === 'Escape' && !searchBox.hidden) { e.preventDefault(); closeSearch(); }
    else if (!inField && !mod && pageEls().length && (e.key === 'ArrowRight' || e.key === 'ArrowLeft')) {
      e.preventDefault();
      goToPage(currentPage() + (e.key === 'ArrowRight' ? 1 : -1));
    }
  });

  function askPassword(retry) {
    return new Promise((resolve) => {
      setStatus('');
      const form = el('form', 'dv-password');
      form.append(el('p', null, retry ? 'Incorrect password. Try again.' : 'This document is password protected.'));
      const input = el('input');
      Object.assign(input, { type: 'password', placeholder: 'Password', autocomplete: 'off' });
      const ok = textButton('Open', 'Open document', 'dv-btn dv-primary');
      ok.type = 'submit';
      const cancel = textButton('Cancel', 'Cancel');
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
    thumbs?.destroy();
    thumbs = null;
    showThumbs(false);
    thumbsBtn.hidden = true;
    if (!searchBox.hidden) closeSearch();
    stage.replaceChildren();
    stage.style.zoom = '';
    stage.className = 'dv-stage';
    pill.hidden = true;
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
      const started = performance.now();
      root.classList.remove('dv-preloading');
      reset();
      setStatus('Opening…', 'loading');
      try {
        const { blob, name: fileName } = await toBlob(source, name);
        if (token !== openToken) return;
        nameEl.textContent = fileName;
        nameEl.title = fileName;
        sizeEl.textContent = formatBytes(blob.size);
        if (blob.size > maxBytes) {
          throw new Error(`This file is ${formatBytes(blob.size)}; the limit is ${formatBytes(maxBytes)}.`);
        }
        const buffer = await blob.arrayBuffer();
        const { kind } = await detectKind(fileName, buffer);
        if (token !== openToken) return;

        badge.textContent = FORMATS[kind]?.label || 'File';
        badge.dataset.kind = FORMATS[kind]?.badge || kind;
        root.dataset.kind = kind;

        if (kind === 'unsupported') {
          current = { controller: null, kind, blob, name: fileName };
          throw new Error(`This file type isn't supported. Supported formats: ${SUPPORTED_EXTENSIONS.map((e) => `.${e}`).join(', ')}.`);
        }

        const render = await FORMATS[kind].load();
        stage.classList.add(`dv-stage-${kind}`);
        const controller = await render({
          buffer, name: fileName, mount: stage, scroller, askPassword, download,
          onContentChange: () => { if (search.term) runSearch(search.term); },
        });
        if (token !== openToken) { controller.destroy?.(); return; }
        current = { controller, kind, blob, name: fileName };
        setStatus('');
        // Time from open() to the document being on screen, e.g. "49.2 MB · Loaded in 2.1 s".
        const loadMs = Math.round(performance.now() - started);
        sizeEl.textContent = `${formatBytes(blob.size)} · Loaded in ${formatDuration(loadMs)}`;
        sizeEl.title = `Loaded in ${loadMs.toLocaleString()} ms`;
        pill.hidden = false;
        autoFit = false;
        if (kind === 'sheet' || kind === 'text') applyZoom(1);
        else fitWidth();
        scroller.scrollTop = 0;
        buildThumbnails();
        if (pageEls().length) fitWidth(); // thumbnails panel changed the available width
        updatePageUi();
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

    /**
     * Show the loading spinner before open() is called, e.g. while the app is
     * still downloading the file. Call again to update the message ("45%").
     */
    loading(message, name, kind = 'loading') {
      if (!root.classList.contains('dv-preloading')) {
        openToken++;
        reset();
        root.classList.add('dv-preloading');
      }
      if (name) { nameEl.textContent = name; nameEl.title = name; sizeEl.textContent = ''; badge.textContent = ''; }
      setStatus(message, kind);
    },

    zoomIn: () => step(1),
    zoomOut: () => step(-1),
    fitWidth,
    setZoom: (z) => { autoFit = false; applyZoom(z); },
    goToPage,
    find: (term) => { openSearch(); searchInput.value = term; return runSearch(term); },
    print: printCurrent,

    destroy() {
      openToken++;
      reset();
      container.replaceChildren();
    },
  };
}

// Supported formats, e.g. for <input type="file" accept={ACCEPT}>.
export { SUPPORTED_EXTENSIONS, ACCEPT, FORMATS, LIBRARIES } from './formats.js';
