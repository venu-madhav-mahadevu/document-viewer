// Angular build only: checks the service worker (installed on the first visit, controls the next
// one) and that the app and an already-opened document still open offline.
// usage: node scripts/qa/offline.cjs <angular app url>
const { chromium } = require('playwright');

(async () => {
  const [base] = process.argv.slice(2);
  const results = [];
  const ok = (name, cond, info = '') => results.push(`${cond ? 'PASS' : 'FAIL'}  ${name} ${info}`);
  const b = await chromium.launch();
  const ctx = await b.newContext();
  const p = await ctx.newPage();
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message));
  await p.goto(base);
  await p.waitForSelector('.sample-group button');
  await p.waitForFunction(() => navigator.serviceWorker?.controller || navigator.serviceWorker?.ready, null, { timeout: 20000 }).catch(() => {});
  await p.waitForTimeout(4000);
  await p.reload();
  await p.waitForSelector('.sample-group button');
  ok('Service worker controls the page on the next visit', await p.evaluate(() => !!navigator.serviceWorker?.controller));
  await p.goto(`${base}#/view?file=quarterly-report.pdf`);
  ok('Online: a document opens', await p.waitForSelector('.dv-pdf-page canvas', { timeout: 30000 }).then(() => true).catch(() => false));
  await p.waitForTimeout(1500);
  await ctx.setOffline(true);
  await p.goto(base).catch(() => {});
  ok('Offline: the app still opens (from the service worker cache)', await p.waitForSelector('.sample-group button', { timeout: 15000 }).then(() => true).catch(() => false));
  await p.goto(`${base}#/view?file=quarterly-report.pdf`).catch(() => {});
  await p.reload().catch(() => {});
  ok('Offline: a document opened before still opens', await p.waitForSelector('.dv-pdf-page canvas', { timeout: 30000 }).then(() => true).catch(() => false));
  ok('No page errors', errs.length === 0, JSON.stringify(errs.slice(0, 3)));
  console.log(results.join('\n'));
  console.log(`${results.filter((r) => r.startsWith('PASS')).length}/${results.length} passed`);
  await b.close();
})();
