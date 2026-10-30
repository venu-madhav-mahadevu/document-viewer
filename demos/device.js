import './sw-cleanup.js';
// Device check: open this page on a phone or tablet to see whether the viewer
// works there - browser features it needs, then all 15 formats opened for real.
import './demo.css';
import { createDocViewer } from '../src/viewer/index.js';

const $ = (id) => document.getElementById(id);

// ---------- Device and browser ----------
function describeDevice() {
  const ua = navigator.userAgent;
  const iOS = /iPhone|iPad|iPod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const android = /Android/.test(ua);
  const iosVersion = (ua.match(/OS (\d+)[_.](\d+)/) || []).slice(1, 3).join('.');
  const androidVersion = (ua.match(/Android (\d+(?:\.\d+)?)/) || [])[1];
  const chrome = (ua.match(/(?:Chrome|CriOS)\/(\d+)/) || [])[1];
  const safari = (ua.match(/Version\/(\d+(?:\.\d+)?).*Safari/) || [])[1];
  const firefox = (ua.match(/(?:Firefox|FxiOS)\/(\d+)/) || [])[1];
  const edge = (ua.match(/(?:Edg|EdgiOS|EdgA)\/(\d+)/) || [])[1];
  const samsung = (ua.match(/SamsungBrowser\/(\d+)/) || [])[1];
  const platform = iOS ? `iOS / iPadOS ${iosVersion || '?'}` : android ? `Android ${androidVersion || '?'}` : 'Desktop';
  const browser = samsung ? `Samsung Internet ${samsung}` : edge ? `Edge ${edge}` : firefox ? `Firefox ${firefox}`
    : /CriOS/.test(ua) ? `Chrome ${chrome} (iOS, Safari engine)` : chrome ? `Chrome ${chrome}`
      : safari ? `Safari ${safari}` : 'Unknown';
  return { iOS, platform, browser, ua };
}

// ---------- Features the viewer uses ----------
async function moduleWorkerWorks() {
  try {
    const url = URL.createObjectURL(new Blob(['self.postMessage(1)'], { type: 'text/javascript' }));
    const w = new Worker(url, { type: 'module' });
    const ok = await Promise.race([new Promise((r) => { w.onmessage = () => r(true); w.onerror = () => r(false); }), new Promise((r) => setTimeout(() => r(false), 3000))]);
    w.terminate();
    URL.revokeObjectURL(url);
    return ok;
  } catch { return false; }
}

const FEATURES = [
  { name: 'Promise.withResolvers', why: 'Required by pdf.js 6 (all PDFs)', required: true, test: () => typeof Promise.withResolvers === 'function' },
  { name: 'Module Web Workers', why: 'Excel files are read in the background (without them, on the page itself)', required: false, test: moduleWorkerWorks },
  { name: 'WebAssembly', why: 'pdf.js image decoders', required: true, test: () => typeof WebAssembly === 'object' },
  { name: 'createImageBitmap', why: 'JPEG-compressed TIFF pages', required: true, test: () => typeof createImageBitmap === 'function' },
  { name: 'structuredClone', why: 'Used by pdf.js', required: true, test: () => typeof structuredClone === 'function' },
  { name: 'IntersectionObserver', why: 'Draws pages as they scroll into view', required: true, test: () => 'IntersectionObserver' in window },
  { name: 'ResizeObserver', why: 'Re-fits pages on rotate / resize', required: true, test: () => 'ResizeObserver' in window },
  { name: 'CSS zoom', why: 'Zoom for Word, PowerPoint, TIFF, images', required: true, test: () => CSS.supports('zoom', '2') },
  { name: 'CSS Custom Highlight API', why: 'Search highlights (search still jumps to matches without it)', required: false, test: () => typeof CSS !== 'undefined' && 'highlights' in CSS },
  { name: 'Fullscreen', why: 'Full-screen button (hidden when missing; iPhone Safari has none)', required: false, test: () => !!document.fullscreenEnabled },
  { name: 'Printing', why: 'Print button', required: false, test: () => typeof window.print === 'function' },
];

const SAMPLES = [
  ['.pdf', 'quarterly-report.pdf'], ['.doc', 'onboarding-guide.doc'], ['.docx', 'onboarding-guide.docx'],
  ['.docm', 'onboarding-macro.docm'], ['.xls', 'branches-legacy.xls'], ['.xlsx', 'regional-sales.xlsx'],
  ['.xlsm', 'branches-macro.xlsm'], ['.csv', 'accounts.csv'], ['.ppt', 'strategy-deck.ppt'],
  ['.pptx', 'strategy-deck.pptx'], ['.jpg', 'photo.jpg'], ['.jpeg', 'photo-chart.jpeg'],
  ['.tif', 'scanned-contract.tif'], ['.tiff', 'color-chart.tiff'], ['.txt', 'readme.txt'],
];

function row(cells, cls) {
  const tr = document.createElement('tr');
  if (cls) tr.className = cls;
  for (const c of cells) tr.insertCell().textContent = c;
  return tr;
}

async function checkFeatures() {
  const body = $('features');
  let missingRequired = 0;
  for (const f of FEATURES) {
    const ok = await f.test();
    if (!ok && f.required) missingRequired++;
    body.append(row([ok ? '✅' : f.required ? '❌' : '⚠', f.name, f.required ? 'Required' : 'Optional', f.why], ok ? 'dc-ok' : f.required ? 'dc-fail' : 'dc-warn'));
  }
  return missingRequired;
}

async function testFormats() {
  const btn = $('test');
  btn.disabled = true;
  const body = $('formats');
  body.replaceChildren();
  const host = $('stage');
  let passed = 0;
  for (const [ext, file] of SAMPLES) {
    $('progress').textContent = `Opening ${file}…`;
    const box = document.createElement('div');
    box.style.height = '300px';
    host.replaceChildren(box);
    const viewer = createDocViewer(box, { download: false });
    const t0 = performance.now();
    let result = '❌';
    let detail = '';
    try {
      const blob = await (await fetch(new URL(`../samples/${file}`, location.href))).blob();
      await viewer.open(blob, file);
      const err = box.querySelector('.dv-status-error:not([hidden])');
      const items = box.querySelectorAll('.dv-pdf-page, section.docx, .dv-grid, .dv-pptx-slide, .dv-image img, .dv-text pre, .dv-tiff-page, .dv-legacy-page, .dv-legacy-slide').length;
      if (err) detail = err.textContent;
      else if (!items) detail = 'Nothing was drawn';
      else { result = '✅'; passed++; detail = `${box.querySelector('.dv-badge').textContent}, ${items} page/view(s)`; }
    } catch (e) {
      detail = String(e?.message || e);
    }
    body.append(row([result, ext, `${Math.round(performance.now() - t0)} ms`, detail], result === '✅' ? 'dc-ok' : 'dc-fail'));
    viewer.destroy();
    await new Promise((r) => setTimeout(r, 0));
  }
  host.replaceChildren();
  $('progress').textContent = `${passed} of ${SAMPLES.length} formats work on this device.`;
  $('verdict-formats').textContent = `${passed} / ${SAMPLES.length} formats work`;
  $('verdict-formats').className = passed === SAMPLES.length ? 'dc-pill dc-pill-ok' : 'dc-pill dc-pill-fail';
  btn.disabled = false;
}

const device = describeDevice();
$('platform').textContent = device.platform;
$('browser').textContent = device.browser;
$('screen').textContent = `${screen.width} × ${screen.height} · ${window.devicePixelRatio}x density${navigator.deviceMemory ? ` · ~${navigator.deviceMemory} GB memory` : ''}`;
$('ua').textContent = device.ua;
if (device.iOS) $('ios-note').hidden = false;

checkFeatures().then((missing) => {
  $('verdict-features').textContent = missing ? `${missing} required feature(s) missing` : 'All required features present';
  $('verdict-features').className = missing ? 'dc-pill dc-pill-fail' : 'dc-pill dc-pill-ok';
});
$('test').addEventListener('click', testFormats);
