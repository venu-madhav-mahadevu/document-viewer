import { setup } from './shared.js';
import source from './pptx-jszip.js?raw';

// #region demo
import JSZip from 'jszip';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

async function render(buffer, name, out) {
  // A .pptx is a zip of XML files: ppt/slides/slide1.xml, slide2.xml, ...
  const zip = await JSZip.loadAsync(buffer);
  const slideNo = (p) => Number(/slide(\d+)\.xml$/.exec(p)[1]);
  const slides = Object.keys(zip.files)
    .filter((p) => /^ppt\/slides\/slide\d+\.xml$/.test(p))
    .sort((a, b) => slideNo(a) - slideNo(b));

  for (const path of slides) {
    const xml = new DOMParser().parseFromString(await zip.file(path).async('string'), 'application/xml');

    const card = document.createElement('section');
    card.className = 'slide';
    card.appendChild(document.createElement('h3')).textContent = `Slide ${slideNo(path)}`;

    // Text: every <a:p> paragraph is made of <a:t> runs.
    for (const p of xml.getElementsByTagNameNS(A, 'p')) {
      const text = [...p.getElementsByTagNameNS(A, 't')].map((t) => t.textContent).join('');
      if (text) card.appendChild(document.createElement('p')).textContent = text;
    }

    // Images: the slide's .rels file points at files in ppt/media/.
    const rels = await zip.file(path.replace('slides/', 'slides/_rels/') + '.rels')?.async('string');
    for (const [, target] of (rels || '').matchAll(/Target="\.\.\/media\/([^"]+)"/g)) {
      const blob = await zip.file(`ppt/media/${target}`).async('blob');
      const img = card.appendChild(document.createElement('img'));
      img.src = URL.createObjectURL(blob);
    }
    out.append(card);
  }
}
// #endregion

setup({ sample: 'strategy-deck.pptx', accept: '.pptx', render, source });
