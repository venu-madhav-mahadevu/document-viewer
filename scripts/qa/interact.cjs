// Interactive checks on the test pack: PDF search, text selection, links, bookmarks, zoom, page
// box, password prompt; image zoom; TIFF thumbnails; slide count, thumbnails and speaker notes.
// usage: node scripts/qa/interact.cjs <app url>   (prints a JSON summary)
const { chromium } = require('playwright');

(async () => {
  const [base] = process.argv.slice(2);
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 900 } });
  const p = await ctx.newPage();
  const out = {};
  const open = async (f) => {
    await p.goto(`${base}?src=samples/test-pack/${f}`);
    await p.waitForSelector('.dv-stage .dv-page, .dv-stage section.docx, .dv-password', { timeout: 90000 });
    await p.waitForTimeout(1500);
  };
  const search = async (term) => {
    await p.keyboard.press('Control+f'); await p.waitForTimeout(300);
    if (!(await p.isVisible('.dv-search input'))) await p.click('button[title^="Find in document"]');
    await p.fill('.dv-search input', term); await p.keyboard.press('Enter'); await p.waitForTimeout(1500);
    const r = await p.$eval('.dv-search-count', (e) => e.textContent);
    await p.keyboard.press('Escape');
    return r;
  };

  await open('test-pdf.pdf');
  out.pdfPages = await p.$eval('.dv-pagebox', (e) => e.textContent);
  out.pdfBookmarks = await p.$$eval('.dv-outline-item', (x) => x.length);
  out.pdfSearchFINDME = await search('FINDME');
  out.pdfSelect = await p.evaluate(() => {
    const t = document.querySelector('.dv-stage .dv-pdf-page .textLayer span'); const r = document.createRange(); r.selectNodeContents(t);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r); return s.toString();
  });
  const before = await p.$eval('.dv-stage .dv-pdf-page canvas', (c) => c.width);
  for (let i = 0; i < 3; i++) { await p.click('.dv-pill button[title^="Zoom in"]').catch(() => {}); await p.waitForTimeout(400); }
  await p.waitForTimeout(1500);
  out.pdfZoomSharper = (await p.$eval('.dv-stage .dv-pdf-page canvas', (c) => c.width)) > before;
  await p.fill('.dv-pageinput', '4'); await p.press('.dv-pageinput', 'Enter');
  await p.waitForSelector('.dv-pdf-link', { timeout: 15000 }).catch(() => {});
  out.pdfLinks = await p.$$eval('.dv-pdf-link', (a) => a.map((x) => x.getAttribute('href')));

  await open('test-pdf-protected.pdf');
  out.pwPrompt = !!(await p.$('.dv-password'));
  await p.fill('.dv-password input', 'wrong'); await p.press('.dv-password input', 'Enter'); await p.waitForTimeout(1500);
  out.pwWrong = await p.$eval('.dv-password', (e) => /Incorrect/.test(e.textContent)).catch(() => false);
  await p.fill('.dv-password input', 'test1234'); await p.press('.dv-password input', 'Enter');
  await p.waitForSelector('.dv-stage .dv-pdf-page', { timeout: 30000 }).catch(() => {});
  out.pwRight = await p.$eval('.dv-pagebox', (e) => e.textContent).catch(() => 'none');

  for (const f of ['test-image.jpg', 'test-image-large.jpg']) {
    const t0 = Date.now(); await open(f);
    out[f] = { openMs: Date.now() - t0, size: await p.$eval('.dv-stage .dv-page img', (i) => [i.naturalWidth, i.naturalHeight]) };
  }
  await open('test-scan-multipage.tif');
  await p.click('button[title="Show page thumbnails"]').catch(() => {}); await p.waitForTimeout(1200);
  out.tif = { pages: await p.$eval('.dv-pagebox', (e) => e.textContent), thumbs: await p.$$eval('.dv-thumb-frame', (t) => t.length) };
  for (const f of ['test-slides.pptx', 'test-slides.ppt']) {
    await open(f);
    out[f] = { pages: await p.$eval('.dv-pagebox', (e) => e.textContent), notes: await p.$$eval('.dv-notes', (n) => n.length), charts: await p.$$eval('.dv-chart', (c) => c.length) };
  }
  await open('test-word.docx');
  out.docxSearch = await search('Square bullets item');
  console.log(JSON.stringify(out, null, 1));
  await b.close();
})();
