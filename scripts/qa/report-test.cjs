// Checks the test report page: opened from the app's sidebar link, all tabs, phone width, dark mode.
// usage: node report-test.cjs <app base url> <out prefix>
const pw = require('playwright');
// expected numbers come from the same data file the page reads
const DATA = require('../../public/test-report/data/test-results.json');
const E = (() => {
  const na = DATA.filter((t) => t.result === 'N/A').length;
  const pass = DATA.filter((t) => t.result === 'Pass').length;
  const f = new Map();
  for (const t of DATA) { const k = t.family + '|' + t.feature; f.set(k, (f.get(k) || 0) + (t.result === 'Fail' ? 1 : 0)); }
  return { pass, applies: DATA.length - na, rate: Math.round((pass / (DATA.length - na)) * 100), works: [...f.values()].filter((v) => !v).length, problems: [...f.values()].filter((v) => v).length };
})();
(async () => {
  const [base, out] = process.argv.slice(2);
  const res = [];
  const ok = (name, cond, info = '') => res.push(`${cond ? 'PASS' : 'FAIL'}  ${name} ${info}`);
  for (const [bname, engine] of [['Chrome', 'chromium'], ['Firefox', 'firefox'], ['Safari', 'webkit']]) {
    const b = await pw[engine].launch();
    const ctx = await b.newContext({ viewport: { width: 1366, height: 900 } });
    const p = await ctx.newPage();
    const errs = [];
    p.on('pageerror', (e) => errs.push(e.message));
    p.on('console', (m) => { if (m.type() === 'error') errs.push(m.text().slice(0, 160)); });
    await p.goto(base);
    await p.waitForTimeout(2500); // service worker (Angular) installs on the first visit
    await p.reload();
    await p.waitForTimeout(1500);
    const link = p.getByRole('link', { name: 'Test results report' });
    ok(`${bname}: sidebar link shown`, await link.isVisible());
    await link.click();
    await p.waitForSelector('#works-t tbody tr', { timeout: 20000 }).catch(() => {});
    ok(`${bname}: report opens from the link`, /test-report\/index\.html/.test(p.url()), p.url());
    const take = await p.textContent('#take');
    ok(`${bname}: summary line`, take.includes(`${E.pass} of ${E.applies} tests pass (${E.rate}%)`), JSON.stringify(take.slice(0, 80)));
    ok(`${bname}: What works rows`, (await p.$$eval('#works-t tbody tr', (r) => r.length)) === E.works);
    // family tiles count only tests that can be checked (N/A shown separately)
    const tiles = await p.$$eval('#works-fam .f', (fs) => fs.map((f) => f.innerText.replace(/\s+/g, ' ')));
    ok(`${bname}: family tiles leave out N/A`, tiles.every((t) => { const m = /^\S+ ([\d,]+) of ([\d,]+) tests pass/.exec(t); return m && (E.problems || m[1] === m[2]); }), tiles.join(' | '));
    await p.click('#tab-results');
    ok(`${bname}: Results tiles`, (await p.textContent('#k-rate')) === `${E.rate}%`);
    ok(`${bname}: all 506 tests listed`, (await p.$$eval('#tests tbody tr', (r) => r.length)) === 506);
    await p.fill('#q', 'bullets'); await p.selectOption('#fres', 'Fail');
    ok(`${bname}: search + filter`, /of 506 tests/.test(await p.textContent('#count')), await p.textContent('#count'));
    await p.click('#tab-matrix');
    const mxRows = await p.$$eval('#mx tbody tr', (r) => r.length);
    ok(`${bname}: matrix rows`, mxRows === Math.max(E.problems, 1), String(mxRows)); // no problems: one row saying so
    if (E.problems) {
      await p.click('#mx button.Fail >> nth=0');
      ok(`${bname}: matrix cell detail`, /Viewer showed:/.test(await p.textContent('#mx-detail')));
    } else ok(`${bname}: matrix says there are no problems`, /No open problems/.test(await p.textContent('#mx')));
    await p.click('#tab-plan');
    ok(`${bname}: plan tiles`, (await p.textContent('#p-weeks')) === '2 weeks' && /^\d+%$/.test(await p.textContent('#p-wk1')) && /^\d+%$/.test(await p.textContent('#p-wk2')), `${await p.textContent('#p-weeks')} / ${await p.textContent('#p-wk1')} / ${await p.textContent('#p-wk2')}`);
    ok(`${bname}: 2-week plan table`, (await p.$$eval('#two-weeks tbody tr', (r) => r.length)) === 10);
    ok(`${bname}: no Claude, no 7-10 week estimate`, !/claude|7–10|phases 1/i.test(await p.evaluate(() => document.documentElement.outerHTML)));
    ok(`${bname}: link to a tab (#plan)`, p.url().endsWith('#plan'));
    if (bname === 'Chrome') {
      await p.screenshot({ path: `${out}-plan.png`, fullPage: true });
      await p.click('#tab-works'); await p.screenshot({ path: `${out}-works.png` });
      await p.click('#tab-matrix'); await p.screenshot({ path: `${out}-matrix.png` });
      // back link returns to the app
      await p.click('.back'); await p.waitForTimeout(1500);
      ok('Chrome: back link returns to the viewer', !/test-report/.test(p.url()), p.url());
      // phone width + dark mode
      const ph = await b.newContext({ ...pw.devices['iPhone 13'], colorScheme: 'dark' });
      const pp = await ph.newPage();
      await pp.goto(base.replace(/\/?$/, '/') + 'test-report/index.html#matrix');
      await pp.waitForSelector('#mx tbody tr', { timeout: 20000 });
      const wide = await pp.evaluate(() => document.documentElement.scrollWidth - innerWidth);
      ok('Phone: no sideways page scroll', wide <= 0, String(wide));
      ok('Phone: opens the tab named in the link', await pp.isVisible('#matrix'));
      ok('Dark mode: dark background', (await pp.evaluate(() => getComputedStyle(document.body).backgroundColor)) === 'rgb(31, 31, 31)');
      await pp.screenshot({ path: `${out}-phone-dark.png` });
    }
    ok(`${bname}: no errors`, errs.length === 0, JSON.stringify(errs.slice(0, 3)));
    await b.close();
  }
  console.log(res.join('\n'));
  console.log(`${res.filter((r) => r.startsWith('PASS')).length}/${res.length} passed`);
})();
