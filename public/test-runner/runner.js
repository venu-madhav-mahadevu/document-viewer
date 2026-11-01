// Runs automatic checks on every test-pack file in this site's own viewer (in an iframe on the
// same site), in the visitor's browser. Works for the Stage / PROD build and the Angular build,
// which both open ?src=samples/... documents.
'use strict';

const $ = (sel) => document.querySelector(sel);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const APP = new URL('../', location.href);
const CONTENT = '.dv-pdf-page, section.docx, .dv-grid, .dv-pptx-slide, .dv-image img, .dv-text pre, .dv-tiff-page, .dv-legacy-page, .dv-legacy-slide';

// What each file must show. `n(sel)` counts elements in the document view (not the thumbnails).
const WORD_X = (pages) => async (t) => {
  t.eq('pages', t.n('section.docx'), pages);
  t.ok('header on every page', t.n('section.docx header') >= pages, `${t.n('section.docx header')} headers`);
  t.ok('floating shapes and pictures', t.n('.dv-docx-float, .dv-docx-float-pic') >= 3, `${t.n('.dv-docx-float, .dv-docx-float-pic')} floating objects`);
  t.ok('pictures shown', t.imgs() >= 2, `${t.imgs()} pictures`);
  t.ok('test text found', t.text().includes('Docs Viewer test pack'), '');
};
const EXCEL = async (t) => {
  t.eq('sheet tabs', t.n('.dv-sheet-tab'), 6);
  t.ok('cells shown', t.n('.dv-grid td') > 50, `${t.n('.dv-grid td')} cells`);
  await t.tab('Colors & fonts');
  const fills = t.all('.dv-grid td').filter((td) => /gradient/.test(td.style.background + td.style.backgroundImage)).length;
  t.ok('pattern and gradient fills', fills >= 2, `${fills} cells`);
  await t.tab('Conditional formatting');
  t.eq('icon set arrows', t.n('.dv-cf-icon'), 3);
  const scale = t.all('.dv-grid td').filter((td) => td.style.backgroundColor && !/255, 255, 255/.test(td.style.backgroundColor)).length;
  t.ok('color scale and highlight rules', scale >= 4, `${scale} colored cells`);
  const bars = t.all('.dv-grid td').filter((td) => /linear-gradient\(to right/.test(td.style.backgroundImage)).length;
  t.eq('data bars', bars, 3);
  await t.tab('Notes & links');
  t.eq('filter buttons', t.n('.dv-filter'), 3);
  t.ok('cell note', t.n('.dv-has-note') >= 1, `${t.n('.dv-has-note')} notes`);
  t.ok('links', t.n('.dv-grid a, .dv-link') >= 2, `${t.n('.dv-grid a, .dv-link')} links`);
  t.ok('picture in the sheet', t.n('.dv-sheet-pic') >= 1, `${t.n('.dv-sheet-pic')} pictures`);
};
const SLIDES = (charts) => async (t) => {
  t.eq('slides', t.n('.dv-pptx-slide, .dv-legacy-slide'), 10);
  t.ok('charts drawn', t.charts.size >= charts, `${t.charts.size} charts`);
  t.ok('shapes drawn', t.n('.dv-pptx-geom, .dv-ppt-geom') >= 5, `${t.n('.dv-pptx-geom, .dv-ppt-geom')} shapes`);
  t.ok('speaker notes', t.n('.dv-notes') >= 1, `${t.n('.dv-notes')} notes`);
  t.ok('pictures shown', t.imgs() >= 2, `${t.imgs()} pictures`);
};
const IMAGE = (w, h) => async (t) => {
  const img = t.q('.dv-image img');
  t.ok('picture decoded', img && img.naturalWidth > 0, img ? `${img.naturalWidth} × ${img.naturalHeight}` : 'no picture');
  if (w) t.eq('full size', `${img?.naturalWidth} × ${img?.naturalHeight}`, `${w} × ${h}`);
};
const TIFF = (pages) => async (t) => {
  t.eq('pages', t.n('.dv-tiff-page'), pages);
  await t.until(() => t.all('.dv-tiff-page canvas').some((c) => c.width > 0), 30000);
  t.ok('page drawn', t.all('.dv-tiff-page canvas').some((c) => c.width > 0), `${t.n('.dv-tiff-page canvas')} drawn`);
};

const TESTS = [
  ['test-word.docx', WORD_X(7)],
  ['test-word.docm', WORD_X(7)],
  ['test-word.doc', async (t) => {
    t.eq('pages', t.n('.dv-legacy-page'), 11);
    t.eq('pictures (inline and floating)', t.n('img.dv-doc-picture'), 3);
    t.eq('shapes, arrows and text boxes', t.n('.dv-doc-shape'), 6);
    t.eq('equation', t.n('math'), 1);
    t.eq('header on every page', t.n('.dv-doc-header'), 11);
    const footers = t.all('.dv-doc-footer').map((f) => f.textContent.trim());
    t.ok('footer page numbers', footers.length === 11 && footers.every((f, i) => f === `Page ${i + 1}`), footers.slice(0, 3).join(', '));
    t.ok('hyperlink', /example\.com/.test(t.q('a.dv-doc-link')?.href || ''), t.q('a.dv-doc-link')?.href || 'no link');
    t.ok('landscape last page', (() => { const p = t.all('.dv-legacy-page').at(-1); return p && p.offsetWidth > p.offsetHeight; })(), '');
  }],
  ['test-excel.xlsx', EXCEL],
  ['test-excel.xlsm', EXCEL],
  ['test-excel.xls', EXCEL],
  ['test-data.csv', async (t) => {
    t.ok('rows shown', t.n('.dv-grid tr') > 10, `${t.n('.dv-grid tr')} rows`);
    t.ok('values shown', /\d/.test(t.text()), '');
  }],
  ['test-slides.pptx', SLIDES(2)],
  ['test-slides.ppt', SLIDES(2)],
  ['test-pdf.pdf', async (t) => {
    t.eq('pages', t.n('.dv-pdf-page'), 9);
    t.eq('bookmarks', frame.contentDocument.querySelectorAll('.dv-outline-item').length, 9); // side panel, outside the pages
    await t.until(() => t.n('.dv-pdf-link') >= 2, 15000);
    t.ok('clickable links', t.n('.dv-pdf-link') >= 2, `${t.n('.dv-pdf-link')} links`);
    await t.until(() => t.n('.textLayer span') > 20, 15000);
    t.ok('selectable text', t.n('.textLayer span') > 20, `${t.n('.textLayer span')} text pieces`);
  }],
  ['test-pdf-protected.pdf', async (t) => {
    t.ok('asks for the password', !!t.q('.dv-password input'), '');
    await t.type('.dv-password input', 'wrong');
    await t.until(() => /Incorrect/.test(t.q('.dv-password')?.textContent || ''), 10000);
    t.ok('wrong password refused', /Incorrect/.test(t.q('.dv-password')?.textContent || ''), '');
    await t.type('.dv-password input', 'test1234');
    await t.until(() => t.n('.dv-pdf-page') >= 9, 20000);
    t.eq('opens with the right password', t.n('.dv-pdf-page'), 9);
  }, '.dv-password input'],
  ['test-image.jpg', IMAGE(1600, 1000)],
  ['test-image.jpeg', IMAGE()],
  ['test-image-progressive.jpeg', IMAGE()],
  ['test-image-rotated-exif.jpg', IMAGE()],
  ['test-image-cmyk.jpg', IMAGE()],
  ['test-image-large.jpg', IMAGE(8000, 5000)],
  ['test-scan-multipage.tif', TIFF(5)],
  ...['test-tiff-uncompressed.tiff', 'test-tiff-lzw.tiff', 'test-tiff-packbits.tiff', 'test-tiff-deflate.tiff', 'test-tiff-jpeg.tif',
    'test-tiff-fax-g4.tif', 'test-tiff-fax-g3.tif', 'test-tiff-grayscale.tiff', 'test-tiff-cmyk.tiff', 'test-tiff-16bit.tif'].map((f) => [f, TIFF(1)]),
];

// ---------- running
const frame = $('#frame');
const rows = $('#results tbody');
let results = [];
let running = false;

function addRow(file, check, pass, seen) {
  results.push({ file, check, result: pass ? 'Pass' : 'Fail', seen });
  const tr = document.createElement('tr');
  if (!pass) tr.className = 'fail';
  const cells = [file, check, null, seen];
  cells.forEach((v, i) => {
    const td = document.createElement('td');
    if (i === 2) { td.className = 'c'; const chip = document.createElement('span'); chip.className = `chip ${pass ? 'Pass' : 'Fail'}`; chip.textContent = pass ? 'Pass' : 'Fail'; td.append(chip); }
    else { td.textContent = v ?? ''; if (i === 3) td.className = 'seen'; }
    tr.append(td);
  });
  rows.append(tr);
  tr.scrollIntoView({ block: 'nearest' });
  const pass2 = results.filter((r) => r.result === 'Pass').length;
  $('#t-pass').textContent = pass2;
  $('#t-fail').textContent = results.length - pass2;
}

function load(file) {
  return new Promise((resolve) => {
    const errors = [];
    const onload = () => {
      const w = frame.contentWindow;
      try {
        w.addEventListener('error', (e) => errors.push(e.message || 'error'));
        w.addEventListener('unhandledrejection', (e) => errors.push(String(e.reason?.message || e.reason)));
        const orig = w.console.error.bind(w.console);
        w.console.error = (...a) => { errors.push(a.map(String).join(' ').slice(0, 200)); orig(...a); };
      } catch { /* other origin: not expected */ }
      resolve(errors);
    };
    frame.addEventListener('load', onload, { once: true });
    // a page that never finishes loading must not stall the run: go on and let 'opens' report it
    setTimeout(() => { frame.removeEventListener('load', onload); resolve(errors); }, 60000);
    frame.src = new URL(`?src=${encodeURIComponent(`samples/test-pack/${file}`)}&t=${Date.now()}`, APP).href;
  });
}

function tester(file) {
  const doc = () => frame.contentDocument;
  const stage = () => doc()?.querySelector('.dv-stage') || doc();
  const t = {
    charts: new Set(),
    win: () => frame.contentWindow,
    q: (sel) => stage()?.querySelector(sel),
    all: (sel) => [...(stage()?.querySelectorAll(sel) || [])],
    n: (sel) => t.all(sel).length,
    imgs: () => t.all('img').filter((i) => i.naturalWidth > 0).length,
    // pages, slides and their links / charts are drawn when they scroll into view: visit them all
    scrollAll: async () => {
      const sc = doc()?.querySelector('.dv-scroller');
      if (!sc) return;
      const note = () => t.all('.dv-chart').forEach((c) => t.charts.add(c)); // far-away slides are freed, so count while passing
      for (let y = 0; y < sc.scrollHeight; y += Math.max(sc.clientHeight * 0.8, 200)) { sc.scrollTop = y; await sleep(250); note(); }
      sc.scrollTop = sc.scrollHeight; await sleep(600); note();
      sc.scrollTop = 0; await sleep(300);
    },
    text: () => stage()?.innerText || '',
    ok: (check, pass, seen) => addRow(file, check, !!pass, seen),
    eq: (check, got, want) => addRow(file, check, got === want, `${got}${got === want ? '' : ` (expected ${want})`}`),
    until: async (fn, ms = 15000) => { const end = Date.now() + ms; while (Date.now() < end) { try { if (fn()) return true; } catch { /* not ready */ } await sleep(150); } return false; },
    tab: async (name) => {
      const tab = [...(doc()?.querySelectorAll('.dv-sheet-tab') || [])].find((b) => b.textContent.trim() === name);
      tab?.click();
      await sleep(900);
    },
    type: async (sel, text) => {
      const input = doc()?.querySelector(sel);
      if (!input) return;
      input.value = text;
      input.dispatchEvent(new Event('input', { bubbles: true }));
      (input.form?.requestSubmit ? input.form.requestSubmit() : input.form?.submit());
      await sleep(400);
    },
  };
  return t;
}

async function runAll() {
  if (running) return;
  running = true;
  $('#run').disabled = true; $('#save').disabled = true; $('#only').disabled = true;
  results = []; rows.replaceChildren(); $('#tiles').hidden = false;
  const only = $('#only').value;
  const list = TESTS.filter(([f]) => !only || f === only);
  const started = Date.now();
  let opened = 0;
  const timer = setInterval(() => { $('#t-time').textContent = `${Math.round((Date.now() - started) / 1000)} s`; }, 500);
  for (const [k, [file, checks, ready]] of list.entries()) {
    $('#status').textContent = `Checking ${file} (${k + 1} of ${list.length})…`;
    $('#progress').style.width = `${(k / list.length) * 100}%`;
    const errors = await load(file);
    const t = tester(file);
    const shown = await t.until(() => {
      const d = frame.contentDocument;
      return d && (d.querySelector(ready || CONTENT) || d.querySelector('.dv-status-error:not([hidden])'));
    }, file.includes('large') ? 120000 : 60000);
    const failed = frame.contentDocument?.querySelector('.dv-status-error:not([hidden])')?.textContent;
    t.ok('opens', shown && !failed, failed || (shown ? '' : 'did not open within the time limit'));
    if (shown && !failed) {
      opened++;
      await sleep(1200); // let pictures, charts and fonts finish
      await t.scrollAll();
      try { await checks(t); } catch (err) { t.ok('checks ran', false, String(err?.message || err)); }
    }
    t.ok('no errors', errors.length === 0, errors.slice(0, 2).join(' | '));
    $('#t-files').textContent = opened;
  }
  clearInterval(timer);
  $('#progress').style.width = '100%';
  $('#t-time').textContent = `${Math.round((Date.now() - started) / 1000)} s`;
  const pass = results.filter((r) => r.result === 'Pass').length;
  $('#status').textContent = `Done: ${pass} of ${results.length} checks pass on ${list.length} file(s) in ${navigator.userAgent.match(/(Firefox|Edg|Chrome|Safari)\/[\d.]+/)?.[0] || 'this browser'}.`;
  frame.src = 'about:blank';
  running = false;
  $('#run').disabled = false; $('#save').disabled = false; $('#only').disabled = false;
}

function save() {
  const data = { site: APP.href, browser: navigator.userAgent, date: new Date().toISOString(), results };
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify(data, null, 1)], { type: 'application/json' }));
  a.download = `docs-viewer-test-run-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

$('#only').append(...TESTS.map(([f]) => new Option(f, f)));
$('#run').addEventListener('click', runAll);
$('#save').addEventListener('click', save);
if (new URLSearchParams(location.search).has('autorun')) runAll();
