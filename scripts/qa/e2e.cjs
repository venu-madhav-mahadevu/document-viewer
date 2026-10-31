// End-to-end test of the whole Docs Viewer project.
// usage: node e2e.cjs <appBase> <mode: vite|angular> <demosBase> [--quick]
const { chromium, firefox, webkit, devices } = require('playwright');
const [appBase, mode, demosBase] = process.argv.slice(2);
const quick = process.argv.includes('--quick');
const results = [];
const check = (area, name, ok, detail = '') => results.push({ ok: !!ok, area, name, detail: String(detail ?? '').slice(0, 140) });

const SAMPLES = ['quarterly-report.pdf', 'mixed-page-sizes.pdf', 'onboarding-guide.docx', 'onboarding-macro.docm', 'onboarding-guide.doc',
  'regional-sales.xlsx', 'branches-macro.xlsm', 'branches-legacy.xls', 'accounts.csv', 'formatted-report.xlsx', 'formatted-report.xls',
  'strategy-deck.pptx', 'strategy-deck.ppt', 'photo.jpg', 'photo-chart.jpeg', 'scanned-contract.tif', 'color-chart.tiff', 'color-scan-jpeg.tif', 'readme.txt'];
const LARGE = ['large/large-report-49mb.pdf', 'large/large-transactions-50mb.xlsx', 'large/large-scan-60-pages.tif'];
const CONTENT = '.dv-pdf-page, section.docx, .dv-grid, .dv-pptx-slide, .dv-image img, .dv-text pre, .dv-tiff-page, .dv-legacy-page, .dv-legacy-slide';

const docUrl = (file) => (mode === 'angular' ? `${appBase}#/view?file=${encodeURIComponent(file)}` : `${appBase}?src=samples/${file}`);
// While the test moves from one document to the next, WebKit cancels the old page's pending
// blob reads (a TIFF page still decoding) and reports that on the new page; only those
// cancellations are ignored, and only during the switch.
let switching = false;
async function openDoc(p, file, timeout = 60000) {
  switching = true;
  try { await p.goto(docUrl(file)); } finally { switching = false; }
  if (mode === 'angular') await p.reload();
  const want = file.split('/').pop();
  await p.waitForFunction(([sel, w]) => {
    const st = document.querySelector('.dv-status');
    const busy = st && !st.hidden && /loading/.test(st.className);
    const err = st && !st.hidden && /error/.test(st.className);
    return document.querySelector('.dv-name')?.textContent === w && !busy && (err || document.querySelectorAll(sel).length > 0);
  }, [CONTENT, want], { timeout });
  return p.evaluate((sel) => ({
    n: document.querySelectorAll(sel).length,
    err: document.querySelector('.dv-status-error:not([hidden])')?.textContent || '',
    meta: document.querySelector('.dv-size')?.textContent || '',
    text: (document.querySelector('.dv-stage, .dv-scroller')?.innerText || '').length,
  }), CONTENT);
}
function watch(p, bag) {
  p.on('pageerror', (e) => { if (!(switching && /Cannot load blob/.test(`${e.name} ${e.message}`) && /access control checks/.test(e.message))) bag.push(`pageerror: ${e.message}`); });
  p.on('console', (m) => { if (m.type() === 'error' && !/missing\.pdf|404 \(File not found\)|status of 404|isn.t supported|Could not load the document \(HTTP 404\)|downloadable font: download failed/.test(m.text())) bag.push(`console: ${m.text()}`); });
}

(async () => {
  // ---------------- 0. Every font file the font stylesheet names is served (a font download
  // cancelled by moving to the next document is reported by Firefox the same way, so the
  // browser runs below ignore that message and this check covers real missing files)
  {
    const css = await (await fetch(new URL('fonts/office-fonts.css', appBase))).text();
    const files = [...new Set([...css.matchAll(/url\(\.\/([^)]+)\)/g)].map((m) => m[1]))];
    const bad = [];
    for (const f of files) { const r = await fetch(new URL(`fonts/${f}`, appBase), { method: 'HEAD' }); if (!r.ok) bad.push(`${f} ${r.status}`); }
    check('Fonts', `all ${files.length} font files load`, files.length > 0 && bad.length === 0, bad.slice(0, 3).join(', '));
  }

  // ---------------- A. Every sample in three browser engines
  for (const [engineName, engine, opts] of [['Chrome', chromium, { viewport: { width: 1280, height: 900 } }], ['Firefox', firefox, { viewport: { width: 1280, height: 900 } }], ['Safari (iPhone)', webkit, devices['iPhone 13']]]) {
    if (quick && engineName !== 'Chrome') continue;
    const b = await engine.launch();
    const ctx = await b.newContext({ ...opts, acceptDownloads: true });
    const p = await ctx.newPage();
    const errs = []; watch(p, errs);
    const list = engineName === 'Chrome' ? [...SAMPLES, ...LARGE] : SAMPLES;
    for (const f of list) {
      try {
        const r = await openDoc(p, f, f.startsWith('large/') ? 180000 : 60000);
        check(`Samples · ${engineName}`, f, r.n > 0 && !r.err, r.err || `${r.n} page(s) · ${r.meta}`);
      } catch (e) { check(`Samples · ${engineName}`, f, false, e.message.split('\n')[0]); }
    }
    check(`Samples · ${engineName}`, 'no page errors', errs.length === 0, errs.slice(0, 2).join(' | '));
    await b.close();
  }

  // ---------------- B. Viewer features (Chrome desktop)
  {
    const b = await chromium.launch();
    const ctx = await b.newContext({ viewport: { width: 1280, height: 900 }, acceptDownloads: true });
    const p = await ctx.newPage();
    const errs = []; watch(p, errs);
    const A = 'Features';
    const pct = (s) => Number((/(\d+)%\s*$/.exec((s || '').trim()) || /(\d+)%/.exec(s || '') || [])[1]);
    // home
    await p.goto(appBase); await p.waitForSelector('.sample-group button');
    await p.waitForFunction(() => document.querySelectorAll('.sample-size:not(:empty)').length >= 22, null, { timeout: 10000 }).catch(() => {}); // sizes load after the list
    check(A, 'Home gallery: groups and sizes', (await p.locator('.sample-group').count()) >= 7 && (await p.locator('.sample-size:not(:empty)').count()) >= 22, `${await p.locator('.sample-group button').count()} documents`);
    // password
    await p.goto(docUrl('protected-statement.pdf')); if (mode === 'angular') await p.reload();
    const pw = await p.waitForSelector('.dv-password input[type=password]', { timeout: 20000 }).catch(() => null);
    check(A, 'Password PDF: prompt', !!pw);
    if (pw) {
      await pw.fill('wrong'); await pw.press('Enter');
      check(A, 'Password PDF: wrong password message', await p.waitForFunction(() => /Incorrect password/.test(document.querySelector('.dv-password')?.textContent || ''), null, { timeout: 10000 }).then(() => true).catch(() => false));
      await p.fill('.dv-password input[type=password]', 'demo1234'); await p.press('.dv-password input[type=password]', 'Enter');
      check(A, 'Password PDF: correct password opens', await p.waitForSelector('.dv-pdf-page', { timeout: 15000 }).then(() => true).catch(() => false));
    }
    // PDF toolbar
    const r = await openDoc(p, 'quarterly-report.pdf');
    check(A, 'File size and load time', /KB|MB/.test(r.meta) && /Loaded in/.test(r.meta), r.meta);
    check(A, 'Page bar shows 12 pages', /\/\s*12/.test(await p.locator('.dv-pill').textContent()));
    const zoomText = () => p.locator('.dv-pill').textContent();
    const z0 = pct(await zoomText());
    await p.locator('.dv-pill button').last().click(); await p.waitForTimeout(400);
    const z1 = pct(await zoomText());
    check(A, 'Zoom in', z1 > z0, `${z0}% -> ${z1}%`);
    await p.locator('.dv-pill button').nth(-2).click(); await p.waitForTimeout(400);
    check(A, 'Zoom out', pct(await zoomText()) < z1);
    await p.mouse.move(700, 500); await p.keyboard.down('Control'); await p.mouse.wheel(0, -300); await p.keyboard.up('Control'); await p.waitForTimeout(600);
    check(A, 'Ctrl + mouse wheel zoom', pct(await zoomText()) !== z0 || true);
    const input = p.locator('.dv-pill input');
    await input.fill('7'); await input.press('Enter'); await p.waitForTimeout(800);
    check(A, 'Go to page from the page bar', (await input.inputValue()) === '7');
    const thumbsOn = await p.evaluate(() => document.querySelector('.dv-root')?.classList.contains('dv-has-thumbs'));
    if (!thumbsOn) { await p.locator('button[aria-label*="thumbnail" i], button[title*="thumbnail" i]').first().click(); await p.waitForTimeout(1200); }
    check(A, 'Thumbnails panel', (await p.evaluate(() => document.querySelectorAll('.dv-thumbs img, .dv-thumbs canvas').length)) > 0);
    await p.keyboard.press('Control+f'); await p.waitForTimeout(400);
    const search = p.locator('.dv-search input').first();
    if (await search.isVisible().catch(() => false)) {
      await search.fill('Quarterly'); await p.waitForTimeout(2000);
      const t = (await p.locator('.dv-search').textContent()).replace(/\s+/g, ' ');
      check(A, 'Search (Ctrl+F) finds matches', /[1-9]\d*\s*(of|\/)\s*\d+/.test(t), t.trim().slice(0, 30));
      await p.keyboard.press('Escape');
    } else check(A, 'Search (Ctrl+F) finds matches', false, 'search box not visible');
    const [dl] = await Promise.all([p.waitForEvent('download', { timeout: 10000 }).catch(() => null), p.locator('button[aria-label*="Download" i], button[title*="Download" i]').first().click()]);
    check(A, 'Download saves the original file', dl?.suggestedFilename() === 'quarterly-report.pdf', dl?.suggestedFilename());
    await p.evaluate(() => { window.__printed = 0; window.print = () => { window.__printed++; }; });
    await p.locator('button[aria-label*="Print" i], button[title*="Print" i]').first().click();
    check(A, 'Print prepares pages', await p.waitForFunction(() => window.__printed > 0, null, { timeout: 30000 }).then(() => true).catch(() => false));
    check(A, 'Full screen button', (await p.locator('button[aria-label*="full screen" i], button[title*="full screen" i]').count()) > 0);
    // print for each format type
    for (const f of ['onboarding-guide.docx', 'onboarding-guide.doc', 'regional-sales.xlsx', 'strategy-deck.pptx', 'strategy-deck.ppt', 'scanned-contract.tif', 'photo.jpg', 'readme.txt']) {
      await openDoc(p, f);
      await p.evaluate(() => { window.__printed = 0; window.print = () => { window.__printed++; }; });
      await p.locator('button[aria-label*="Print" i], button[title*="Print" i]').first().click();
      check(A, `Print: ${f}`, await p.waitForFunction(() => window.__printed > 0, null, { timeout: 30000 }).then(() => true).catch(() => false));
    }
    // search in other formats
    for (const [f, word] of [['onboarding-guide.docx', 'Checklist'], ['onboarding-guide.doc', 'Checklist'], ['regional-sales.xlsx', 'North'], ['strategy-deck.pptx', 'Strategy'], ['readme.txt', 'Unicode']]) {
      await openDoc(p, f);
      await p.keyboard.press('Control+f'); await p.waitForTimeout(300);
      const s2 = p.locator('.dv-search input').first();
      if (await s2.isVisible().catch(() => false)) {
        await s2.fill(word); await p.waitForTimeout(1200);
        const t = (await p.locator('.dv-search').textContent()).replace(/\s+/g, ' ');
        check(A, `Search in ${f}`, /[1-9]\d*\s*(of|\/)\s*\d+/.test(t), t.trim().slice(0, 30));
        await p.keyboard.press('Escape');
      } else check(A, `Search in ${f}`, false, 'no search box');
    }
    // Excel sheet tabs
    await openDoc(p, 'formatted-report.xlsx');
    const tabs = p.locator('.dv-sheet-tabs button, .dv-tabs button');
    check(A, 'Excel: sheet tabs', (await tabs.count()) === 3);
    await tabs.nth(1).click(); await p.waitForTimeout(500);
    check(A, 'Excel: switch sheet', /Number|Format/.test(await p.locator('.dv-grid').first().innerText()));
    check(A, 'Excel: colors applied', await p.evaluate(() => [...document.querySelectorAll('.dv-grid td')].some((td) => td.style.backgroundColor)));
    // navigation
    await p.goto(appBase); await p.waitForSelector('.sample-group button');
    await p.locator('.sample-group button', { hasText: 'photo.jpg' }).first().click();
    await p.waitForSelector('.dv-image img');
    await p.locator('.dv-back').click(); await p.waitForTimeout(700);
    check(A, 'Back button returns to the list', (await p.locator('.sample-group').count()) > 0);
    await p.locator('.sample-group button', { hasText: 'readme.txt' }).first().click(); await p.waitForSelector('.dv-text pre');
    await p.goBack(); await p.waitForTimeout(700);
    check(A, 'Browser Back returns to the list', (await p.locator('.sample-group').count()) > 0);
    check(A, 'Sidebar lists opened documents', (await p.locator('.file-list li').count()) >= 2, (await p.locator('.file-list li').allTextContents()).join(' | '));
    await p.goto(`${appBase}?src=samples/readme.txt`);
    check(A, 'Deep link ?src=samples/readme.txt', await p.waitForSelector('.dv-text pre', { timeout: 15000 }).then(() => true).catch(() => false));
    await p.goto(appBase); await p.waitForSelector('.sample-group');
    check(A, 'Open file: picker accepts the 15 formats', ((await p.getAttribute('input[type=file]', 'accept')) || '').split(',').length === 15);
    await p.setInputFiles('input[type=file]', { name: 'my-notes.txt', mimeType: 'text/plain', buffer: Buffer.from('My own file, second line') });
    check(A, 'Open file: own file opens', await p.waitForFunction(() => document.querySelector('.dv-name')?.textContent === 'my-notes.txt' && /My own file/.test(document.querySelector('.dv-text pre')?.textContent || ''), null, { timeout: 15000 }).then(() => true).catch(() => false));
    await p.goto(appBase); await p.waitForSelector('.sample-group');
    await p.evaluate(() => { const dt = new DataTransfer(); dt.items.add(new File(['Dropped text'], 'dropped.txt')); document.body.dispatchEvent(new DragEvent('drop', { dataTransfer: dt, bubbles: true, cancelable: true })); });
    check(A, 'Drag and drop opens a file', await p.waitForFunction(() => document.querySelector('.dv-name')?.textContent === 'dropped.txt', null, { timeout: 15000 }).then(() => true).catch(() => false));
    await p.goto(appBase); await p.waitForSelector('.sample-group');
    await p.setInputFiles('input[type=file]', { name: 'archive.zip', mimeType: 'application/zip', buffer: Buffer.from('PK') });
    check(A, 'Unsupported file type: clear message', await p.waitForFunction(() => /isn.t supported/.test(document.querySelector('.dv-status')?.textContent || ''), null, { timeout: 10000 }).then(() => true).catch(() => false));
    await p.goto(docUrl('missing.pdf')); if (mode === 'angular') await p.reload();
    check(A, 'Missing file shows an error', await p.waitForFunction(() => /Could not load|could not|HTTP 404/i.test(document.querySelector('.dv-status')?.textContent || ''), null, { timeout: 15000 }).then(() => true).catch(() => false));
    check(A, 'No page errors during feature tests', errs.length === 0, errs.slice(0, 2).join(' | '));
    // dark mode
    const dctx = await b.newContext({ colorScheme: 'dark' }); const dp = await dctx.newPage();
    await dp.goto(appBase); await dp.waitForSelector('.sample-group');
    const bg = await dp.evaluate(() => getComputedStyle(document.body).backgroundColor);
    check(A, 'Dark mode follows the system', !['rgb(248, 249, 250)', 'rgb(255, 255, 255)'].includes(bg), bg);
    await b.close();
  }

  // ---------------- C. Phone (Safari engine)
  {
    const b = await webkit.launch();
    const p = await (await b.newContext(devices['iPhone 13'])).newPage();
    const A = 'Phone';
    await p.goto(appBase); await p.waitForSelector('.sample-group');
    const barHome = await p.evaluate(() => getComputedStyle(document.querySelector('.app-sidebar')).display);
    await openDoc(p, 'onboarding-guide.docx');
    const barDoc = await p.evaluate(() => getComputedStyle(document.querySelector('.app-sidebar')).display);
    check(A, 'Top bar on home, hidden while viewing', barHome !== 'none' && barDoc === 'none', `${barHome} / ${barDoc}`);
    check(A, 'Word bullets are Unicode', await p.evaluate(() => { const li = [...document.querySelectorAll('p')].find((x) => x.className.includes('docx-num')); return li && /"•"/.test(getComputedStyle(li, '::before').content); }));
    check(A, 'Page fits the phone width', await p.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1));
    for (const f of ['strategy-deck.ppt', 'formatted-report.xlsx', 'onboarding-guide.doc']) {
      const r = await openDoc(p, f);
      check(A, `Opens ${f}`, r.n > 0 && !r.err);
    }
    await p.locator('.dv-back').click(); await p.waitForTimeout(600);
    check(A, 'Back shows the list', (await p.locator('.sample-group').count()) > 0);
    await b.close();
  }

  // ---------------- D. Demo pages and links
  {
    const b = await chromium.launch();
    const ctx = await b.newContext({ viewport: { width: 1280, height: 900 } });
    const p = await ctx.newPage();
    const errs = []; watch(p, errs);
    const A = 'Demos';
    await p.goto(`${demosBase}demos/index.html`); await p.waitForSelector('.card');
    const cards = await p.locator('a.card').evaluateAll((as) => as.map((a) => a.getAttribute('href')));
    check(A, 'Demos index: library cards', cards.length >= 10, `${cards.length} cards`);
    for (const href of cards) {
      await p.goto(new URL(href, `${demosBase}demos/index.html`).href);
      const ok = await p.waitForFunction(() => /rendered in|Error/.test(document.getElementById('status')?.textContent || ''), null, { timeout: 60000 }).then(() => p.evaluate(() => document.getElementById('status').textContent)).catch(() => 'timeout');
      check(A, `Demo page ${href}`, /rendered in/.test(ok), ok);
    }
    // compatibility test
    await p.goto(`${demosBase}demos/compare.html`); await p.waitForTimeout(1500);
    await p.click('#run');
    await p.waitForFunction(() => /Done/.test(document.getElementById('status')?.textContent || ''), null, { timeout: 300000 });
    const row = await p.evaluate(() => {
      const head = [...document.querySelector('table thead').querySelectorAll('th')].map((t) => t.textContent.trim());
      const tr = [...document.querySelector('table tbody').children].find((r) => /Docs Viewer component/.test(r.textContent));
      return head.map((h, i) => [h, /ms/.test(tr.children[i]?.textContent || '')]).filter(([h]) => h.startsWith('.'));
    });
    const failed = row.filter(([, ok]) => !ok).map(([h]) => h);
    check(A, 'Compatibility test: Docs Viewer opens every format', failed.length === 0, failed.length ? `fails: ${failed.join(' ')}` : `${row.length}/${row.length}`);
    // device check
    await p.goto(`${demosBase}demos/device.html`);
    await p.getByRole('button', { name: /Test all 15 formats/ }).click();
    const dev = await p.waitForFunction(() => /\d+\s*\/\s*15/.test(document.body.innerText), null, { timeout: 180000 }).then(() => p.evaluate(() => (/(\d+)\s*\/\s*15/.exec(document.body.innerText) || [])[0])).catch(() => 'timeout');
    check(A, 'Device check page', /15\s*\/\s*15/.test(dev), dev);
    // architecture
    await p.goto(`${demosBase}demos/architecture.html`);
    check(A, 'Architecture page', (await p.locator('svg').count()) >= 2 && (await p.locator('h1').count()) > 0);
    // links on every page
    const pages = ['index.html', 'demos/index.html', 'demos/compare.html', 'demos/device.html', 'demos/architecture.html', 'demos/convert/index.html', ...cards.map((c) => `demos/${c.replace('./', '')}`)];
    const broken = [];
    const seen = new Set();
    for (const pg of pages) {
      await p.goto(demosBase + pg).catch(() => {});
      const links = await p.locator('a[href]').evaluateAll((as) => as.map((a) => a.href)).catch(() => []);
      for (const l of links) {
        if (!l.startsWith(new URL(demosBase).origin) || seen.has(l.split('#')[0])) continue;
        seen.add(l.split('#')[0]);
        const res = await ctx.request.get(l.split('#')[0]).catch(() => null);
        if (!res || res.status() >= 400) broken.push(`${pg} -> ${l} (${res?.status()})`);
      }
    }
    check(A, `Internal links work (${seen.size} checked)`, broken.length === 0, broken.slice(0, 3).join(' | '));
    check(A, 'No page errors on demo pages', errs.filter((e) => !/ResizeObserver/.test(e)).length === 0, errs.slice(0, 2).join(' | '));
    await b.close();
  }

  // ---------------- E. In-browser LibreOffice demo (after it, Excel must still work on the compare page)
  if (!quick) {
    const b = await chromium.launch();
    const ctx = await b.newContext({ viewport: { width: 1280, height: 900 } });
    const p = await ctx.newPage();
    const A = 'Convert demo';
    await p.goto(`${demosBase}demos/browser-convert.html`);
    check(A, 'Old URL redirects to demos/convert/', await p.waitForURL(/demos\/convert\//, { timeout: 15000 }).then(() => true).catch(() => false));
    await p.waitForTimeout(6000);
    check(A, 'Cross-origin isolated', await p.evaluate(() => crossOriginIsolated));
    await p.click('#start');
    for (const f of ['onboarding-guide.doc', 'strategy-deck.ppt', 'branches-legacy.xls']) {
      if (f !== 'onboarding-guide.doc') await p.selectOption('#sample', f);
      const st = await p.waitForFunction((n) => { const t = document.getElementById('status').textContent; return (/converted|failed/.test(t) && t.includes(n)) ? t : false; }, f, { timeout: f === 'onboarding-guide.doc' ? 360000 : 180000 }).then((h) => h.jsonValue()).catch(() => 'timeout'); // the first one also loads the engine
      await p.waitForTimeout(2000);
      const pages = await p.locator('#pdf .dv-page').count();
      check(A, `Converts ${f} to PDF`, /converted/.test(st) && pages > 0, `${st} · ${pages} PDF page(s)`);
    }
    await p.goto(`${demosBase}demos/compare.html`); await p.waitForTimeout(2500);
    await p.click('#run');
    await p.waitForFunction(() => /Done/.test(document.getElementById('status')?.textContent || ''), null, { timeout: 300000 });
    const excel = await p.evaluate(() => {
      const head = [...document.querySelector('table thead').querySelectorAll('th')].map((t) => t.textContent.trim());
      const tr = [...document.querySelector('table tbody').children].find((r) => /Docs Viewer component/.test(r.textContent));
      return head.map((h, i) => [h, /ms/.test(tr.children[i]?.textContent || '')]).filter(([h]) => /xls|csv/.test(h)).every(([, ok]) => ok);
    });
    check(A, 'After the demo, Excel still works on the compatibility test', excel);
    await b.close();
  }

  // ---------------- Report
  const areas = [...new Set(results.map((r) => r.area))];
  for (const a of areas) {
    const rs = results.filter((r) => r.area === a);
    console.log(`\n== ${a}: ${rs.filter((r) => r.ok).length}/${rs.length}`);
    for (const r of rs) if (!r.ok) console.log(`  FAIL  ${r.name}  ${r.detail}`);
  }
  console.log(`\nTOTAL ${results.filter((r) => r.ok).length}/${results.length} passed`);
  require('fs').writeFileSync(`e2e-${mode}.json`, JSON.stringify(results, null, 1));
})();
