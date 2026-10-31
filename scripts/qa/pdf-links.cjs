const { chromium } = require('playwright');
(async () => {
  const base = process.argv[2];
  const b = await chromium.launch(); const ctx = await b.newContext({ viewport: { width: 1300, height: 900 } }); const p = await ctx.newPage();
  const errs = []; p.on('pageerror', (e) => errs.push(e.message));
  await p.goto(`${base}?src=samples/test-pack/test-pdf.pdf`); await p.waitForSelector('.dv-stage .dv-pdf-page canvas');
  await p.fill('.dv-pageinput', '4'); await p.press('.dv-pageinput', 'Enter');
  await p.waitForSelector('.dv-pdf-link', { timeout: 15000 }).catch(() => {});
  console.log('links', await p.$$eval('.dv-pdf-link', (a) => a.map((x) => [x.getAttribute('href'), x.target, x.style.width])));
  const [pop] = await Promise.all([ctx.waitForEvent('page', { timeout: 5000 }).catch(() => null), p.click('.dv-pdf-link[target=_blank]')]);
  console.log('external link opens:', pop && pop.url()); if (pop) await pop.close();
  await p.click('.dv-pdf-link[href="#"]'); await p.waitForTimeout(1500);
  console.log('after internal link, page box shows:', await p.$eval('.dv-pageinput', (e) => e.value));
  console.log('errors', errs); await b.close();
})();
