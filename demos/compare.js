import './sw-cleanup.js';
// Runs every library against one sample of each of the 15 supported formats,
// live in the browser, and shows the results as a matrix.
import './demo.css';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorkerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import { renderAsync as docxPreview } from 'docx-preview';
import mammoth from 'mammoth';
import DOMPurify from 'dompurify';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import JSZip from 'jszip';
import UTIF from 'utif';
import { renderDoc, renderPpt } from '../src/viewer/renderers/legacy.js';
import { createDocViewer } from '../src/viewer/index.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

const SAMPLES = [
  ['.pdf', 'quarterly-report.pdf'], ['.doc', 'onboarding-guide.doc'], ['.docx', 'onboarding-guide.docx'],
  ['.docm', 'onboarding-macro.docm'], ['.xls', 'branches-legacy.xls'], ['.xlsx', 'regional-sales.xlsx'],
  ['.xlsm', 'branches-macro.xlsm'], ['.csv', 'accounts.csv'], ['.ppt', 'strategy-deck.ppt'],
  ['.pptx', 'strategy-deck.pptx'], ['.jpg', 'photo.jpg'], ['.jpeg', 'photo-chart.jpeg'],
  ['.tif', 'scanned-contract.tif'], ['.tiff', 'color-chart.tiff'], ['.txt', 'readme.txt'],
];

const table = (rows) => {
  const t = document.createElement('table');
  t.className = 'data';
  for (const r of rows.slice(0, 20)) {
    const tr = t.insertRow();
    for (const v of r.slice(0, 8)) tr.insertCell().textContent = v ?? '';
  }
  return t;
};
const canvasOf = (w, h) => Object.assign(document.createElement('canvas'), { width: w, height: h, className: 'page' });

// Each adapter renders into `out` and returns a short description of what it produced.
const LIBS = [
  { id: 'viewer', name: 'Docs Viewer component', note: 'all six libraries combined', async run(buf, name, out) {
    const host = document.createElement('div');
    host.style.height = '420px';
    out.append(host);
    const v = createDocViewer(host, { download: false });
    await v.open(new Blob([buf]), name);
    const err = host.querySelector('.dv-status-error:not([hidden])');
    if (err) throw new Error(err.textContent);
    const n = host.querySelectorAll('.dv-pdf-page, section.docx, .dv-grid, .dv-pptx-slide, .dv-image img, .dv-text pre, .dv-tiff-page, .dv-legacy-page, .dv-legacy-slide').length;
    return `${host.querySelector('.dv-badge').textContent}, ${n} page/view(s)`;
  } },
  { id: 'pdfjs', name: 'pdf.js', async run(buf, name, out) {
    const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false,
      standardFontDataUrl: new URL('../pdfjs/standard_fonts/', location.href).href }).promise;
    const page = await pdf.getPage(1);
    const vp = page.getViewport({ scale: 0.8 });
    const c = canvasOf(vp.width, vp.height);
    out.append(c);
    await page.render({ canvas: c, viewport: vp }).promise;
    return `${pdf.numPages} page(s)`;
  } },
  { id: 'docx-preview', name: 'docx-preview', async run(buf, name, out) {
    await docxPreview(buf, out, null, { inWrapper: true, breakPages: true });
    return `${out.querySelectorAll('section.docx').length} page(s)`;
  } },
  { id: 'mammoth', name: 'mammoth', async run(buf, name, out) {
    const { value } = await mammoth.convertToHtml({ arrayBuffer: buf });
    const div = document.createElement('div');
    div.className = 'doc-html';
    div.innerHTML = DOMPurify.sanitize(value);
    out.append(div);
    if (!div.textContent.trim()) throw new Error('No text extracted');
    return `${div.textContent.trim().length} characters of HTML text`;
  } },
  { id: 'sheetjs', name: 'SheetJS CE', async run(buf, name, out) {
    const wb = name.endsWith('.csv') || name.endsWith('.txt')
      ? XLSX.read(new TextDecoder().decode(buf), { type: 'string' })
      : XLSX.read(buf, { type: 'array' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    out.append(table(rows));
    return `${wb.SheetNames.length} sheet(s), ${rows.length} row(s)`;
  } },
  { id: 'exceljs', name: 'ExcelJS', async run(buf, name, out) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    const ws = wb.worksheets[0];
    const rows = [];
    ws.eachRow((r) => rows.push(r.values.slice(1).map((v) => (v && typeof v === 'object' ? v.result ?? v.text ?? '' : v))));
    out.append(table(rows));
    return `${wb.worksheets.length} sheet(s), ${rows.length} row(s)`;
  } },
  { id: 'papaparse', name: 'Papa Parse', async run(buf, name, out) {
    const { data, errors } = Papa.parse(new TextDecoder().decode(buf), { skipEmptyLines: true });
    if (errors.length) throw new Error(errors[0].message);
    out.append(table(data));
    return `${data.length} row(s)`;
  } },
  { id: 'jszip', name: 'JSZip (PPTX text)', async run(buf, name, out) {
    const zip = await JSZip.loadAsync(buf);
    const slides = Object.keys(zip.files).filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p));
    if (!slides.length) throw new Error('Not a PowerPoint file (no ppt/slides/)');
    for (const p of slides.sort()) {
      const xml = await zip.file(p).async('string');
      const text = [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join(' / ');
      out.append(Object.assign(document.createElement('p'), { textContent: `${p.split('/').pop()}: ${text}` }));
    }
    return `${slides.length} slide(s)`;
  } },
  { id: 'cfb', name: 'cfb + reader', note: '.doc / .ppt', async run(buf, name, out) {
    const isPpt = name.endsWith('.ppt');
    await (isPpt ? renderPpt : renderDoc)({ buffer: buf, mount: out, download: false });
    return isPpt ? `${out.querySelectorAll('.dv-legacy-slide').length} slide(s)` : 'text and tables';
  } },
  { id: 'utif', name: 'UTIF.js', async run(buf, name, out) {
    const ifds = UTIF.decode(buf).filter((i) => i.t256);
    if (!ifds.length) throw new Error('No TIFF pages');
    UTIF.decodeImage(buf, ifds[0], ifds);
    const rgba = UTIF.toRGBA8(ifds[0]);
    const c = canvasOf(ifds[0].width, ifds[0].height);
    c.getContext('2d').putImageData(new ImageData(new Uint8ClampedArray(rgba.buffer), c.width, c.height), 0, 0);
    out.append(c);
    return `${ifds.length} page(s)`;
  } },
  { id: 'img', name: 'Browser <img>', note: 'also Viewer.js', async run(buf, name, out) {
    const bitmap = await createImageBitmap(new Blob([buf]));
    const c = canvasOf(bitmap.width, bitmap.height);
    c.getContext('2d').drawImage(bitmap, 0, 0);
    out.append(c);
    return `${bitmap.width} x ${bitmap.height} image`;
  } },
];

// The six libraries the Docs Viewer component uses (shown in green), listed first.
const USED = new Set(['pdfjs', 'docx-preview', 'sheetjs', 'jszip', 'cfb', 'utif']);
const ORDER = ['viewer', 'pdfjs', 'docx-preview', 'sheetjs', 'jszip', 'cfb', 'utif', 'mammoth', 'exceljs', 'papaparse', 'img'];
LIBS.sort((a, b) => ORDER.indexOf(a.id) - ORDER.indexOf(b.id));

/** Output that "opened" but is mostly binary junk is not a real success. */
function looksLikeJunk(text) {
  // Raw source of another format shown as text (e.g. a PDF read as CSV rows).
  if (/%PDF-|\bendobj\b|\bendstream\b|JFIF|\[Content_Types\]\.xml/.test(text)) return true;
  if (text.length < 20) return false;
  const bad = (text.match(/[�\u0000-\u0008\u000E-\u001F]/g) || []).length;
  return bad / text.length > 0.02;
}

const withTimeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(`No result after ${ms / 1000} s`)), ms))]);

async function runOne(lib, buf, name) {
  const out = document.createElement('div');
  out.className = 'cmp-out';
  document.getElementById('scratch').append(out);
  const t0 = performance.now();
  try {
    const detail = await withTimeout(lib.run(buf.slice(0), name, out), 20000);
    const ms = Math.round(performance.now() - t0);
    const state = looksLikeJunk(out.textContent) ? 'junk' : 'ok';
    return { state, ms, detail: state === 'junk' ? `Opened, but the output is unreadable (${detail})` : detail, out };
  } catch (err) {
    return { state: 'fail', ms: Math.round(performance.now() - t0), detail: String(err?.message || err).slice(0, 160), out };
  }
}

// ---------- UI ----------
const files = SAMPLES.map(([ext, file]) => ({ ext, file, label: ext }));
const results = new Map(); // `${libId}|${col}` -> result
const grid = document.getElementById('grid');
const statusEl = document.getElementById('status');
const preview = document.getElementById('preview');
const previewTitle = document.getElementById('preview-title');

function cellText(r) {
  if (!r) return '';
  return r.state === 'ok' ? `✅ ${r.ms} ms` : r.state === 'junk' ? '⚠ junk' : '❌';
}

function drawGrid() {
  const t = document.createElement('table');
  t.className = 'cmp';
  const head = t.createTHead().insertRow();
  head.appendChild(document.createElement('th')).textContent = 'Library';
  for (const f of files) {
    const th = document.createElement('th');
    th.textContent = f.label;
    th.title = f.file;
    head.append(th);
  }
  head.appendChild(document.createElement('th')).textContent = 'Works';
  const body = t.createTBody();
  const groupRow = (text, cls) => {
    const tr = body.insertRow();
    tr.className = 'cmp-group ' + cls;
    const td = tr.insertCell();
    td.colSpan = files.length + 2;
    td.textContent = text;
  };
  for (const lib of LIBS) {
    if (lib.id === 'pdfjs') groupRow('The 6 libraries used inside the Docs Viewer component', 'cmp-group-used');
    if (lib.id === 'mammoth') groupRow('Optional alternatives (not used in the viewer)', 'cmp-group-optional');
    const tr = body.insertRow();
    tr.className = lib.id === 'viewer' ? 'cmp-row-viewer' : USED.has(lib.id) ? 'cmp-row-used' : 'cmp-row-optional';
    const th = document.createElement('th');
    th.scope = 'row';
    th.append(lib.name);
    if (USED.has(lib.id)) th.append(Object.assign(document.createElement('span'), { className: 'cmp-badge', textContent: 'Used in viewer' }));
    if (lib.note) th.append(Object.assign(document.createElement('small'), { textContent: lib.note }));
    tr.append(th);
    let ok = 0;
    files.forEach((f, col) => {
      const r = results.get(`${lib.id}|${col}`);
      if (r?.state === 'ok') ok++;
      const td = tr.insertCell();
      td.className = `cmp-${r?.state || 'pending'}`;
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = cellText(r) || '·';
      btn.title = r ? `${lib.name} + ${f.file}: ${r.detail}` : 'Not run yet';
      btn.disabled = !r;
      btn.onclick = () => showPreview(lib, f, r);
      td.append(btn);
    });
    const total = tr.insertCell();
    total.className = 'cmp-total';
    total.textContent = `${ok} / ${files.length}`;
  }
  grid.replaceChildren(t);
}

function showPreview(lib, f, r) {
  previewTitle.textContent = `${lib.name} + ${f.file}: ${r.state === 'ok' ? 'works' : r.state === 'junk' ? 'opens but unreadable' : 'fails'} (${r.ms} ms)`;
  const detail = Object.assign(document.createElement('p'), { className: 'cmp-detail', textContent: r.detail });
  preview.replaceChildren(detail, r.out);
  r.out.hidden = false;
  preview.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

async function load(f) {
  if (f.blob) return f.blob.arrayBuffer();
  return (await fetch(new URL(`../samples/${f.file}`, location.href))).arrayBuffer();
}

async function runAll(cols) {
  const run = document.getElementById('run');
  run.disabled = true;
  const total = LIBS.length * cols.length;
  let done = 0;
  for (const col of cols) {
    const f = files[col];
    const buf = await load(f);
    for (const lib of LIBS) {
      statusEl.textContent = `Testing ${lib.name} on ${f.file}… (${++done} of ${total})`;
      const r = await runOne(lib, buf, f.file);
      r.out.hidden = true;
      results.set(`${lib.id}|${col}`, r);
      drawGrid();
      await new Promise((res) => setTimeout(res, 0));
    }
  }
  statusEl.textContent = `Done: ${total} tests. Click any cell to see what that library produced.`;
  run.disabled = false;
}

document.getElementById('run').onclick = () => runAll(files.map((_, i) => i));
document.getElementById('own').onchange = (e) => {
  const added = [...e.target.files].map((file) => ({ ext: file.name.slice(file.name.lastIndexOf('.')), file: file.name, label: `${file.name.slice(0, 14)}${file.name.length > 14 ? '…' : ''}`, blob: file }));
  const start = files.length;
  files.push(...added);
  drawGrid();
  runAll(added.map((_, i) => start + i));
  e.target.value = '';
};
drawGrid();
