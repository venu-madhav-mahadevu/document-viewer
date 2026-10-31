// Tests the built library (dist/) the way an application uses it: the demo page loads
// dist/docs-viewer.js and opens every test file. Run `npm run build` first.
// usage: node test/library.test.cjs [chromium|firefox|webkit]
const http = require('http');
const fs = require('fs');
const path = require('path');
const pw = require('playwright');

const ROOT = path.join(__dirname, '..');
const TYPES = { '.html': 'text/html', '.js': 'text/javascript', '.mjs': 'text/javascript', '.css': 'text/css', '.json': 'application/json',
  '.wasm': 'application/wasm', '.woff2': 'font/woff2', '.svg': 'image/svg+xml', '.pdf': 'application/pdf', '.bcmap': 'application/octet-stream' };
const server = http.createServer((req, res) => {
  const file = path.join(ROOT, decodeURIComponent(new URL(req.url, 'http://x').pathname));
  if (!file.startsWith(ROOT) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); res.end(); return; }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(file)] || 'application/octet-stream' });
  fs.createReadStream(file).pipe(res);
});

const CONTENT = '.dv-pdf-page, section.docx, .dv-grid, .dv-pptx-slide, .dv-image img, .dv-text pre, .dv-tiff-page, .dv-legacy-page, .dv-legacy-slide';
const engine = pw[process.argv[2] || 'chromium'];

(async () => {
  if (!fs.existsSync(path.join(ROOT, 'dist', 'docs-viewer.js'))) { console.error('Run npm run build first.'); process.exit(1); }
  await new Promise((r) => server.listen(0, '127.0.0.1', r));
  const base = `http://127.0.0.1:${server.address().port}/demo/index.html?dist`;
  const browser = await engine.launch();
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
  page.on('requestfailed', (r) => errors.push(`request failed: ${r.url()}`));
  page.on('response', (r) => { if (r.status() >= 400) errors.push(`HTTP ${r.status()}: ${r.url()}`); });

  let pass = 0; let fail = 0;
  const check = (name, ok, detail = '') => { ok ? pass++ : fail++; console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}${detail ? `  ${detail}` : ''}`); };
  const files = fs.readdirSync(path.join(ROOT, 'test', 'files')).sort();
  for (const f of files) {
    errors.length = 0;
    await page.goto(`${base}&file=${encodeURIComponent(f)}`);
    let shown;
    if (/protected/.test(f)) {
      // password PDF: asks first, opens with the right password
      shown = await page.waitForSelector('.dv-password input[type=password]', { timeout: 30000 }).then(() => true).catch(() => false);
      if (shown) { await page.fill('.dv-password input[type=password]', 'test1234'); await page.keyboard.press('Enter'); }
    }
    shown = await page.waitForSelector(CONTENT, { timeout: f.includes('large') ? 120000 : 60000 }).then(() => true).catch(() => false);
    await page.waitForTimeout(800);
    const pages = await page.$$eval(CONTENT, (els) => els.length);
    check(f, shown && pages > 0 && errors.length === 0, `${pages} page(s)${errors.length ? ` · ${errors.slice(0, 2).join(' | ')}` : ''}`);
  }
  // assets served from next to the library: fonts and pdf.js data
  const fonts = await page.evaluate(() => document.getElementById('dv-office-fonts')?.href || '');
  check('fonts load from dist/fonts', /\/dist\/fonts\/office-fonts\.css$/.test(fonts), fonts);
  await browser.close();
  server.close();
  console.log(`\n${pass}/${pass + fail} passed`);
  process.exit(fail ? 1 : 0);
})();
