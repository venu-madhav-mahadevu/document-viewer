import '../demo.css';
import hljs from 'highlight.js/lib/core';
import javascript from 'highlight.js/lib/languages/javascript';
import 'highlight.js/styles/github.css';
import { createDocViewer, formatBytes, formatDuration } from '../../src/viewer/index.js';
import source from './convert.js?raw';

// #region demo
import { WorkerBrowserConverter } from '@matbee/libreoffice-converter/browser';

// Engine files are served from our own origin (copied by scripts/copy-lo-assets.mjs).
const lo = (file) => new URL(`./lo/${file}`, location.href).href;

const options = {
  sofficeJs: lo('soffice.js'),
  sofficeWasm: lo('soffice.wasm'),     // ~147 MB
  sofficeData: lo('soffice.data'),     // ~100 MB (fonts, filters)
  sofficeWorkerJs: lo('soffice.worker.js'),
  browserWorkerJs: lo('browser.worker.global.js'),
  onProgress: (p) => showProgress(p.percent, p.message),
};
let converter;

// Downloads (or reads from the browser cache) and starts LibreOffice in a Web Worker.
async function startEngine() {
  converter = new WorkerBrowserConverter(options);
  await converter.initialize();
}

// .doc / .ppt / .xls -> PDF bytes, entirely inside the browser.
async function toPdf(buffer, name) {
  const ext = name.split('.').pop().toLowerCase();
  const result = await converter.convert(new Uint8Array(buffer), { inputFormat: ext, outputFormat: 'pdf' }, name);
  return result.data; // Uint8Array, opened below with pdf.js (createDocViewer)
}

// LibreOffice occasionally stops responding: give up after 20 s, restart the
// engine (a few seconds once cached) and try once more.
async function toPdfSafe(buffer, name) {
  for (let attempt = 1; ; attempt++) {
    let timer;
    const timeout = new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('LibreOffice did not respond')), 20000); });
    try {
      return await Promise.race([toPdf(buffer.slice(0), name), timeout]);
    } catch (err) {
      if (attempt > 1) throw err;
      showProgress(0, 'LibreOffice stopped responding. Restarting the engine…');
      converter.worker?.terminate(); // force-stop the stuck engine (destroy() would wait for it)
      await startEngine();
    } finally {
      clearTimeout(timer);
    }
  }
}
// #endregion

const $ = (id) => document.getElementById(id);
const status = $('status');
const stats = $('stats');
const fastViewer = createDocViewer($('fast'), { download: false });
let pdfViewer;
let ready = false;
let current; // { buffer, name }

function showProgress(percent, message) {
  $('progress').hidden = false;
  $('bar').style.width = `${Math.max(2, Math.min(100, percent || 0))}%`;
  if (message) status.textContent = message;
}

function stat(label, value) {
  let row = stats.querySelector(`[data-k="${label}"]`);
  if (!row) {
    row = document.createElement('div');
    row.dataset.k = label;
    row.append(document.createElement('dt'), document.createElement('dd'));
    row.firstChild.textContent = label;
    stats.append(row);
  }
  row.lastChild.textContent = value;
}

async function show(buffer, name) {
  current = { buffer, name };
  const t0 = performance.now();
  await fastViewer.open(new Blob([buffer]), name);
  stat('Fast preview', formatDuration(performance.now() - t0));
  if (ready) queueConvert();
}

// One conversion at a time; a newer selection replaces any waiting one.
let queue = Promise.resolve();
function queueConvert() {
  const job = current;
  queue = queue.then(() => (job === current ? convert() : undefined));
}

async function convert() {
  const { buffer, name } = current;
  const box = $('pdf');
  pdfViewer?.destroy();
  box.replaceChildren();
  pdfViewer = createDocViewer(box, { download: true });
  pdfViewer.loading(`Converting ${name} with LibreOffice…`, name);
  status.textContent = `Converting ${name}…`;
  const t0 = performance.now();
  try {
    const pdf = await toPdfSafe(buffer, name);
    const ms = performance.now() - t0;
    const pdfName = name.replace(/\.[^.]+$/, '') + '.pdf';
    stat('Convert to PDF', formatDuration(ms));
    stat('PDF size', formatBytes(pdf.byteLength));
    status.textContent = `${name} converted in ${formatDuration(ms)}`;
    $('progress').hidden = true;
    await pdfViewer.open(new File([pdf], pdfName, { type: 'application/pdf' }));
  } catch (err) {
    console.error(err);
    status.textContent = `Conversion failed: ${err.message}`;
    pdfViewer.loading(`Conversion failed: ${err.message}`, name, 'error');
    $('progress').hidden = true;
  }
}

$('start').addEventListener('click', async () => {
  if (!window.crossOriginIsolated) {
    status.textContent = 'This page is not cross-origin isolated yet (needed by LibreOffice). Reload the page once and try again.';
    return;
  }
  $('start').disabled = true;
  $('start').textContent = 'Starting…';
  const t0 = performance.now();
  try {
    await startEngine();
    ready = true;
    $('progress').hidden = true;
    $('start').textContent = 'Engine ready';
    stat('Engine start', formatDuration(performance.now() - t0));
    if (current) queueConvert();
  } catch (err) {
    console.error(err);
    $('start').disabled = false;
    $('start').textContent = 'Start LibreOffice engine';
    status.textContent = `Engine failed to start: ${err.message}`;
  }
});

async function loadSample() {
  const name = $('sample').value;
  const res = await fetch(new URL(`../../samples/${name}`, location.href));
  show(await res.arrayBuffer(), name);
}
$('sample').addEventListener('change', loadSample);
$('file').addEventListener('change', async () => {
  const file = $('file').files[0];
  if (file) show(await file.arrayBuffer(), file.name);
});

if (/Android|iPhone|iPad|iPod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)) {
  $('note').insertAdjacentHTML('beforeend', ' <strong>You are on a phone or tablet: this will probably be too heavy.</strong>');
}

const m = /\/\/ #region demo\r?\n([\s\S]*?)\/\/ #endregion/.exec(source);
$('code').textContent = m[1].trim();
hljs.registerLanguage('javascript', javascript);
hljs.highlightElement($('code'));

loadSample();
