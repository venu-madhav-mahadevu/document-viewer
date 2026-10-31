// Screenshots every page / slide / sheet of every test-pack file, plus each sheet's cell data,
// for comparing one build or round with another.
// usage: node scripts/qa/shoot.cjs <app url> <out dir> [chromium|webkit|firefox]   (ONLY=regex limits files)
// Works for the Stage build (?src=samples/…) and the Angular build (redirects ?src= to its view).
const pw = require('playwright');
const fs = require('fs');
const path = require('path');

const FILES = ['test-word.docx', 'test-word.docm', 'test-word.doc', 'test-excel.xlsx', 'test-excel.xlsm', 'test-excel.xls',
  'test-data.csv', 'test-slides.pptx', 'test-slides.ppt', 'test-pdf.pdf', 'test-image.jpg', 'test-image.jpeg',
  'test-image-progressive.jpeg', 'test-image-rotated-exif.jpg', 'test-image-cmyk.jpg', 'test-image-large.jpg',
  'test-scan-multipage.tif', 'test-tiff-uncompressed.tiff', 'test-tiff-lzw.tiff', 'test-tiff-packbits.tiff',
  'test-tiff-deflate.tiff', 'test-tiff-jpeg.tif', 'test-tiff-fax-g4.tif', 'test-tiff-fax-g3.tif',
  'test-tiff-grayscale.tiff', 'test-tiff-cmyk.tiff', 'test-tiff-16bit.tif'];
const PAGES = '.dv-stage .dv-page, .dv-stage section.docx';

(async () => {
  const [base, out, engine = 'chromium'] = process.argv.slice(2);
  fs.mkdirSync(out, { recursive: true });
  const only = process.env.ONLY ? new RegExp(process.env.ONLY) : null;
  const b = await pw[engine].launch();
  const ctx = await b.newContext({ viewport: { width: 1400, height: 2600 } });
  const info = {};
  for (const f of FILES.filter((x) => !only || only.test(x))) {
    const p = await ctx.newPage();
    const errs = [];
    p.on('pageerror', (e) => errs.push(e.message));
    p.on('console', (m) => { if (m.type() === 'error' || m.type() === 'warning') errs.push(m.text().slice(0, 200)); });
    const t0 = Date.now();
    await p.goto(`${base}?src=samples/test-pack/${f}`);
    await p.waitForSelector(`${PAGES}, .dv-grid`, { timeout: 90000 }).catch(() => errs.push('TIMEOUT'));
    await p.waitForTimeout(2500);
    await p.addStyleTag({ content: '.dv-pill{display:none!important}' }); // the floating page / zoom bar
    const rec = { ms: Date.now() - t0, errs, shots: [] };
    const stem = f.replace(/\./g, '_');
    const tabs = await p.$$('.dv-sheet-tab');
    if (tabs.length) {
      rec.tabs = await p.$$eval('.dv-sheet-tab', (x) => x.map((t) => t.textContent));
      rec.sheets = {};
      for (let i = 0; i < tabs.length; i++) {
        await tabs[i].click(); await p.waitForTimeout(400);
        await p.$eval('.dv-sheet-grid', (g) => { g.style.maxHeight = 'none'; g.style.height = 'auto'; g.style.overflow = 'visible'; });
        const shot = path.join(out, `${stem}-sheet${i + 1}.png`);
        await p.locator('.dv-grid').first().screenshot({ path: shot });
        rec.shots.push(shot);
        rec.sheets[rec.tabs[i]] = await p.$$eval('.dv-grid tbody tr', (rows) => rows.map((tr) => [...tr.children].map((c) => {
          const s = getComputedStyle(c);
          return { t: c.textContent, bg: s.backgroundColor, color: s.color, fw: s.fontWeight, fs: s.fontStyle, td: s.textDecorationLine,
            size: s.fontSize, font: s.fontFamily.split(',')[0], ta: s.textAlign, va: s.verticalAlign, html: c.innerHTML.length < 300 ? c.innerHTML : '' };
        })));
      }
    } else {
      const n = await p.locator(PAGES).count();
      rec.pages = n;
      for (let i = 0; i < Math.min(n, 12); i++) {
        const loc = p.locator(PAGES).nth(i);
        await loc.scrollIntoViewIfNeeded(); await p.waitForTimeout(700);
        const shot = path.join(out, `${stem}-p${i + 1}.png`);
        await loc.screenshot({ path: shot });
        rec.shots.push(shot);
        rec[`size${i + 1}`] = await loc.evaluate((e) => [e.offsetWidth, e.offsetHeight]);
      }
      rec.pill = await p.$eval('.dv-pagebox', (e) => e.textContent).catch(() => null);
    }
    info[f] = rec;
    console.log(f, `${rec.ms}ms`, rec.pages ?? rec.tabs?.length, errs.length ? JSON.stringify(errs.slice(0, 3)) : '');
    await p.close();
  }
  fs.writeFileSync(path.join(out, 'info.json'), JSON.stringify(info, null, 1));
  await b.close();
})();
