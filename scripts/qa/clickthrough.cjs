// Opens every sample from the home gallery by clicking it, like a person would, in Chrome, Firefox
// and Safari (iPhone), and reports documents that fail to open or log errors.
// usage: node scripts/qa/clickthrough.cjs <app url>
const pw = require('playwright');

const CONTENT = '.dv-pdf-page, section.docx, .dv-grid, .dv-pptx-slide, .dv-image img, .dv-text pre, .dv-tiff-page, .dv-legacy-page, .dv-legacy-slide';
// NG05604 = service worker registration failed: Playwright's WebKit on Windows has no service workers (real Safari does)
const IGNORE = /Missing Info for XLS Record|Cannot load system font|TT: undefined function|Unknown field name|downloadable font|NG05604/;

(async () => {
  const [base] = process.argv.slice(2);
  let failed = 0;
  for (const [name, engine, opts] of [['Chrome', 'chromium', { viewport: { width: 1366, height: 900 } }],
    ['Firefox', 'firefox', { viewport: { width: 1366, height: 900 } }], ['Safari iPhone', 'webkit', pw.devices['iPhone 13']]]) {
    const b = await pw[engine].launch();
    const p = await (await b.newContext(opts)).newPage();
    const errs = [];
    p.on('pageerror', (e) => errs.push(e.message.slice(0, 160)));
    p.on('console', (m) => { if (m.type() === 'error' && !IGNORE.test(m.text())) errs.push(m.text().slice(0, 160)); });
    await p.goto(base);
    await p.waitForSelector('.sample-group button');
    const names = await p.$$eval('.sample-group button', (bs) => bs.map((x) => x.querySelector('.sample-name')?.textContent || x.textContent).filter((t) => !/large|test-|CHECKLIST/i.test(t)));
    const bad = [];
    let opened = 0;
    for (const label of names) {
      await p.goto(base); await p.waitForSelector('.sample-group button');
      const btn = p.locator('.sample-group button', { hasText: label }).first();
      if (!(await btn.count())) continue;
      await btn.click();
      // a password-protected PDF opens to its password box
      const ok = await p.waitForSelector(`${CONTENT}, .dv-password`, { timeout: 60000 }).then(() => true).catch(() => false);
      if (ok) opened++; else bad.push(label);
      await p.waitForTimeout(400);
    }
    console.log(`${name}: ${opened}/${opened + bad.length} opened by clicking; errors: ${errs.length ? JSON.stringify([...new Set(errs)].slice(0, 4)) : 'none'}${bad.length ? `; did not open: ${bad.join(', ')}` : ''}`);
    failed += bad.length + (errs.length ? 1 : 0);
    await b.close();
  }
  process.exitCode = failed ? 1 : 0;
})();
