import CFB from 'cfb';
import { el } from '../sanitize.js';
import { clonePagesForPrint, PAGE_WIDTH, PAGE_WIDTH_LANDSCAPE } from '../print.js';
import { readWordDocument, readPageSetup, paragraphStyle, charStyle } from './doc-format.js';
import { readPresentation } from './ppt-format.js';
import { bulletCode } from '../bullets.js';
import { shrinkToFit } from '../fit.js';

/*
 * Word 97-2003 (.doc) and PowerPoint 97-2003 (.ppt) preview.
 *
 * These are binary "compound files" (CFB, read with SheetJS's open-source
 * `cfb`). No open-source browser library reproduces their layout, so this
 * shows the content - text, paragraphs, tables, slide titles/bullets and
 * pictures - laid out as pages. For .doc, character and paragraph formatting,
 * styles and lists are read too (doc-format.js).
 */

const cp1252 = new TextDecoder('windows-1252');
const utf16 = new TextDecoder('utf-16le');

function openCompound(buffer) {
  try {
    return CFB.read(new Uint8Array(buffer), { type: 'array' });
  } catch {
    throw new Error('This file is damaged or is not an Office 97-2003 file.');
  }
}

function stream(cfb, name) {
  const entry = CFB.find(cfb, name);
  return entry?.content ? new Uint8Array(entry.content) : null;
}

const view = (bytes) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);


// ===================================================================== .doc

/** Main-document text of a Word 97-2003 file, via the FIB and piece table. */
function wordText(cfb) {
  const wd = stream(cfb, 'WordDocument');
  if (!wd) throw new Error('This is not a Word document.');
  const dv = view(wd);
  if (dv.getUint16(0, true) !== 0xa5ec) throw new Error('This is not a Word document.');
  if (dv.getUint16(2, true) < 0x00c1) {
    throw new Error('Word 6.0/95 documents can\'t be previewed. Please save it as .docx or PDF.');
  }
  const flags = dv.getUint16(0x0a, true);
  if (flags & 0x0100) throw new Error('This Word document is password protected and can\'t be previewed.');

  const table = stream(cfb, flags & 0x0200 ? '1Table' : '0Table');
  if (!table) throw new Error('This Word document is damaged.');
  const tv = view(table);
  const ccpText = dv.getInt32(76, true); // characters in the main document
  const fcClx = dv.getUint32(0x1a2, true);
  const lcbClx = dv.getUint32(0x1a6, true);

  // Clx: skip property runs (0x01), then the piece table (0x02).
  let p = fcClx;
  while (p < fcClx + lcbClx && table[p] === 0x01) p += 3 + tv.getUint16(p + 1, true);
  if (table[p] !== 0x02) throw new Error('This Word document is damaged.');
  const lcb = tv.getUint32(p + 1, true);
  const plc = p + 5;
  const pieces = (lcb - 4) / 12;

  let text = '';
  for (let i = 0; i < pieces && text.length < ccpText; i++) {
    const cpStart = tv.getInt32(plc + i * 4, true);
    const cpEnd = tv.getInt32(plc + (i + 1) * 4, true);
    const pcd = plc + (pieces + 1) * 4 + i * 8;
    const raw = tv.getUint32(pcd + 2, true);
    const length = cpEnd - cpStart;
    if (raw & 0x40000000) {
      const offset = (raw & 0x3fffffff) / 2; // 8-bit (Windows-1252) text
      text += cp1252.decode(wd.subarray(offset, offset + length));
    } else {
      text += utf16.decode(wd.subarray(raw, raw + length * 2)); // UTF-16 text
    }
  }
  return text.slice(0, ccpText);
}

/** Drop field codes ({ PAGE }, { HYPERLINK "…" }) but keep their displayed result. */
function stripFields(text) {
  let out = '';
  const stack = []; // per open field: true while inside its code part
  for (const ch of text) {
    if (ch === '\x13') { stack.push(true); continue; }
    if (ch === '\x14') { if (stack.length) stack[stack.length - 1] = false; continue; }
    if (ch === '\x15') { stack.pop(); continue; }
    if (stack.some((inCode) => inCode)) continue;
    out += ch;
  }
  return out;
}

const CLEAN = (s) => s
  .replace(/[\x01\x02\x05\x08]/g, '') // pictures, footnote/comment refs, drawn objects
  .replace(/\x1e/g, '-')             // non-breaking hyphen
  .replace(/\x1f/g, '')              // optional hyphen
  .replace(/\xa0/g, ' ');

/**
 * Split Word text into paragraphs, tables and page breaks.
 * Paragraphs end with \r, table cells with \x07, and a row ends with an extra \x07.
 */
function wordBlocks(text) {
  const units = text.match(/[^\r\x07]*[\r\x07]|[^\r\x07]+$/g) || [];
  const blocks = [];
  let table = null;
  let row = null;
  let cellParas = [];
  let lastWasCellEnd = false;

  const pushParagraph = (s) => {
    for (const [i, part] of s.split('\x0c').entries()) {
      if (i > 0) blocks.push({ type: 'break' });
      if (part !== '' || i === 0) blocks.push({ type: 'p', text: part });
    }
  };

  for (const unit of units) {
    const end = unit.slice(-1);
    const body = end === '\r' || end === '\x07' ? unit.slice(0, -1) : unit;
    if (end === '\x07') {
      if (lastWasCellEnd && body === '' && cellParas.length === 0) {
        if (!table) { table = { type: 'table', rows: [] }; blocks.push(table); }
        if (row) table.rows.push(row);
        row = null;
        lastWasCellEnd = false;
        continue;
      }
      if (!table) { table = { type: 'table', rows: [] }; blocks.push(table); }
      row ??= [];
      row.push([...cellParas, body].join('\n'));
      cellParas = [];
      lastWasCellEnd = true;
      continue;
    }
    lastWasCellEnd = false;
    if (row) { cellParas.push(body); continue; } // extra paragraph inside a cell
    if (table) table = null;                     // a normal paragraph ends the table
    pushParagraph(body);
  }
  if (row && table) table.rows.push(row);
  return blocks;
}

function appendWithBreaks(parent, text) {
  text.split('\x0b').forEach((line, i) => {
    if (i) parent.append(el('br'));
    parent.append(document.createTextNode(line));
  });
}

/** Like wordBlocks, for formatted paragraphs from readWordDocument. */
function formattedBlocks(paras) {
  const blocks = [];
  let table = null;
  let row = null;
  let cellParas = [];
  let lastWasCellEnd = false;
  const textOf = (para) => para.segments.map((s) => s.text).join('');

  // A page break (\x0c) inside a paragraph splits it.
  const pushParagraph = (para) => {
    let part = { ...para, segments: [] };
    const parts = [part];
    for (const seg of para.segments) {
      seg.text.split('\x0c').forEach((text, i) => {
        if (i > 0) { part = { ...para, list: undefined, segments: [] }; parts.push(part); }
        if (text) part.segments.push({ ...seg, text });
      });
    }
    if (para.pap.pageBreakBefore) blocks.push({ type: 'break' });
    parts.forEach((pp, i) => {
      if (i > 0) blocks.push({ type: 'break' });
      if (pp.segments.length || i === 0) blocks.push({ type: 'p', para: pp });
    });
  };

  for (const para of paras) {
    if (para.end === '\x07') {
      if (lastWasCellEnd && textOf(para) === '' && cellParas.length === 0) {
        if (!table) { table = { type: 'table', rows: [] }; blocks.push(table); }
        if (row) table.rows.push(row);
        row = null;
        lastWasCellEnd = false;
        continue;
      }
      if (!table) { table = { type: 'table', rows: [] }; blocks.push(table); }
      row ??= [];
      row.push([...cellParas, para]);
      cellParas = [];
      lastWasCellEnd = true;
      continue;
    }
    lastWasCellEnd = false;
    if (row) { cellParas.push(para); continue; }
    if (table) table = null;
    pushParagraph(para);
  }
  if (row && table) table.rows.push(row);
  return blocks;
}

/** One formatted paragraph: paragraph styles, list marker and styled text runs. */
function renderParagraph(para) {
  const p = el('p');
  Object.assign(p.style, paragraphStyle(para.pap, para.list));
  if (para.styleChp?.size) p.style.fontSize = `${para.styleChp.size}pt`;
  if (para.pap.sti >= 1 && para.pap.sti <= 9) {
    p.setAttribute('role', 'heading');
    p.setAttribute('aria-level', String(para.pap.sti));
  }
  if (para.list) {
    const marker = el('span', 'dv-doc-marker', para.list.marker);
    const first = para.pap.first ?? para.list.pap.first;
    if (first < 0) marker.style.minWidth = `${(-first / 15).toFixed(1)}px`;
    Object.assign(marker.style, charStyle({ ...(para.segments[0]?.chp || para.styleChp), ...para.list.chp, fontName: '' }));
    p.append(marker);
  }
  if (!para.segments.length) p.append(el('br'));
  for (const seg of para.segments) {
    const span = el('span');
    Object.assign(span.style, charStyle(seg.chp));
    appendWithBreaks(span, seg.text);
    p.append(span);
  }
  return p;
}

export async function renderDoc({ buffer, mount }) {
  const cfb = openCompound(buffer);
  const plain = wordText(cfb); // also checks the version, encryption and structure

  // Formatted reading first; plain text if this file's formatting can't be read.
  let blocks;
  let formatted = true;
  let setup = null; // page size and margins from the file
  try {
    const wd = stream(cfb, 'WordDocument');
    const table = stream(cfb, view(wd).getUint16(0x0a, true) & 0x0200 ? '1Table' : '0Table');
    blocks = formattedBlocks(readWordDocument(wd, table));
    try { setup = readPageSetup(wd, table); } catch { setup = null; }
  } catch (err) {
    console.warn('Word formatting could not be read; showing plain text.', err);
    formatted = false;
    blocks = wordBlocks(CLEAN(stripFields(plain)));
  }

  const root = el('div', 'dv-legacy');
  mount.append(root);

  // Lay the content out on Letter-sized pages, like a PDF: a new page starts at
  // each page break in the file and whenever a page is full.
  const tw = (v) => Math.round(v / 15); // twips -> CSS px
  const PAGE_WIDTH_PX = setup ? tw(setup.width) : 816;
  const PAGE_HEIGHT = setup ? tw(setup.height) : 1056; // padding included
  const pages = [];
  let page;
  const newPage = () => {
    page = el('div', 'dv-page dv-legacy-page');
    if (setup) {
      Object.assign(page.style, { width: `${PAGE_WIDTH_PX}px`, minHeight: `${PAGE_HEIGHT}px`,
        padding: `${tw(setup.top)}px ${tw(setup.right)}px ${tw(setup.bottom)}px ${tw(setup.left)}px` });
    }
    pages.push(page);
    root.append(page);
  };
  const place = (node) => {
    page.append(node);
    // Full: move this block to a fresh page (a block taller than a page stays where it is).
    if (page.scrollHeight > PAGE_HEIGHT && page.childElementCount > 1) {
      node.remove();
      newPage();
      page.append(node);
    }
  };
  newPage();

  let prev = null; // previous formatted paragraph
  for (const block of blocks) {
    if (block.type !== 'p') prev = null;
    if (block.type === 'break') {
      if (page.childElementCount) newPage();
    } else if (block.type === 'p' && formatted) {
      const node = renderParagraph(block.para);
      // "Don't add space between paragraphs of the same style" (lists, mostly).
      if (prev && block.para.pap.contextual && prev.para.pap.istd === block.para.pap.istd) {
        node.style.marginTop = '0px';
        prev.node.style.marginBottom = '0px';
      }
      prev = { para: block.para, node };
      place(node);
    } else if (block.type === 'p') {
      const p = el('p');
      if (block.text.trim() === '') p.append(el('br'));
      else appendWithBreaks(p, block.text);
      place(p);
    } else if (block.rows.length) {
      const t = el('table', formatted ? 'dv-legacy-table dv-doc-table' : 'dv-legacy-table');
      for (const cells of block.rows) {
        const tr = t.insertRow();
        for (const cell of cells) {
          const td = tr.insertCell();
          if (formatted) { for (const para of cell) td.append(renderParagraph(para)); continue; }
          cell.split('\n').forEach((line, i) => {
            if (i) td.append(el('br'));
            appendWithBreaks(td, line);
          });
        }
      }
      place(t);
    }
  }

  // "Page 2 of 5" at the bottom of each page, as in a printed PDF.
  pages.forEach((p, i) => p.append(el('span', 'dv-legacy-footer', `Page ${i + 1} of ${pages.length}`)));

  const widthOf = () => PAGE_WIDTH_PX;
  return {
    pages: () => pages,
    pageWidth: widthOf,
    naturalWidth: widthOf,
    printNodes: () => ({ nodes: clonePagesForPrint(pages, PAGE_WIDTH, widthOf) }),
    destroy() {},
  };
}

// ===================================================================== .ppt

const REC = {
  DOCUMENT_ATOM: 0x03e9,
  SLIDE: 0x03ee,
  NOTES: 0x03f0,
  SLIDE_PERSIST: 0x03f3,
  MAIN_MASTER: 0x03f8,
  TEXT_HEADER: 0x0f9f,
  TEXT_CHARS: 0x0fa0,
  TEXT_BYTES: 0x0fa8,
  SLIDE_LIST: 0x0ff0,
};
const TITLE_TYPES = new Set([0, 6]);

function pptContent(cfb) {
  const doc = stream(cfb, 'PowerPoint Document');
  if (!doc) throw new Error('This is not a PowerPoint presentation.');
  const dv = view(doc);
  const slides = [];       // from the slide list, in presentation order
  const boxes = [];        // free text boxes, one list per Slide record
  let size = { w: 10, h: 7.5 };
  let current = null;
  let textType = 4;

  const readText = (type, start, length) => (type === REC.TEXT_CHARS
    ? utf16.decode(doc.subarray(start, start + length))
    : cp1252.decode(doc.subarray(start, start + length)));

  function walk(start, end, ctx) {
    let p = start;
    while (p + 8 <= end) {
      const verInst = dv.getUint16(p, true);
      const type = dv.getUint16(p + 2, true);
      const length = dv.getUint32(p + 4, true);
      const body = p + 8;
      const next = body + length;
      if (next > end) break;
      if ((verInst & 0xf) === 0xf) {
        if (type === REC.SLIDE_LIST) walk(body, next, { slideList: (verInst >> 4) === 0 });
        else if (type === REC.SLIDE) { const box = []; boxes.push(box); walk(body, next, { box }); }
        else if (type !== REC.NOTES && type !== REC.MAIN_MASTER) walk(body, next, ctx);
      } else if (type === REC.DOCUMENT_ATOM && length >= 8) {
        size = { w: dv.getInt32(body, true), h: dv.getInt32(body + 4, true) };
      } else if (ctx.slideList) {
        if (type === REC.SLIDE_PERSIST) { current = { title: [], body: [] }; slides.push(current); }
        else if (type === REC.TEXT_HEADER) textType = dv.getUint32(body, true);
        else if ((type === REC.TEXT_CHARS || type === REC.TEXT_BYTES) && current) {
          (TITLE_TYPES.has(textType) ? current.title : current.body).push(readText(type, body, length));
        }
      } else if (ctx.box && type === REC.TEXT_HEADER) {
        textType = dv.getUint32(body, true);
      } else if (ctx.box && (type === REC.TEXT_CHARS || type === REC.TEXT_BYTES)) {
        ctx.box.push({ title: TITLE_TYPES.has(textType), text: readText(type, body, length) });
      }
      p = next;
    }
  }
  walk(0, doc.length, {});

  if (!slides.length) boxes.forEach(() => slides.push({ title: [], body: [] }));
  // Text stored in the slide's own shapes (free text boxes, and placeholders in
  // files written by some apps). Titles keep their title marker.
  if (boxes.length === slides.length) {
    slides.forEach((s, i) => boxes[i].forEach((b) => (b.title ? s.title : s.body).push(b.text)));
  }
  return { slides, size };
}

/** JPEG and PNG pictures stored in the "Pictures" stream, with their byte offsets. */
function pptPictures(cfb) {
  const pics = stream(cfb, 'Pictures');
  if (!pics) return [];
  const dv = view(pics);
  const out = [];
  let p = 0;
  while (p + 8 <= pics.length) {
    const inst = dv.getUint16(p, true) >> 4;
    const type = dv.getUint16(p + 2, true);
    const length = dv.getUint32(p + 4, true);
    const body = p + 8;
    const next = body + length;
    if (next > pics.length) break;
    let mime = null;
    let uids = 1;
    if (type === 0xf01d || type === 0xf02a) { mime = 'image/jpeg'; uids = inst === 0x46b || inst === 0x6e3 ? 2 : 1; }
    else if (type === 0xf01e) { mime = 'image/png'; uids = inst === 0x6e1 ? 2 : 1; }
    if (mime) out.push({ offset: p, blob: new Blob([pics.subarray(body + 16 * uids + 1, next)], { type: mime }) });
    p = next;
  }
  return out;
}

const lines = (texts) => texts.flatMap((t) => t.split(/[\r\n]/)).map((s) => s.replace(/\x0b/g, ' ').trim()).filter(Boolean);

const bulletText = (code, fontName = '') => (code === undefined ? '•' : bulletCode(code, fontName));

/** One slide drawn like PowerPoint: positioned shapes, pictures and formatted text. */
function renderDrawnSlide(slide, size, W, H, urlFor) {
  const scale = W / size.w;          // px per master unit
  const ptPx = scale * 8;            // px per point (576 master units per inch, 72 pt per inch)
  const card = el('section', 'dv-page dv-legacy-slide dv-ppt-slide');
  card.style.width = `${W}px`;
  card.style.height = `${H}px`;
  card.style.background = slide.background || '#ffffff';
  const px = (v) => `${(v * scale).toFixed(2)}px`;

  for (const sh of slide.shapes) {
    const box = el('div', 'dv-ppt-shape');
    Object.assign(box.style, { left: px(sh.x), top: px(sh.y), width: px(Math.max(sh.w, 1)), height: px(Math.max(sh.h, 1)) });
    const tf = [];
    if (sh.rot) tf.push(`rotate(${sh.rot}deg)`);
    if (sh.flipH || sh.flipV) tf.push(`scale(${sh.flipH ? -1 : 1}, ${sh.flipV ? -1 : 1})`);
    if (tf.length) box.style.transform = tf.join(' ');

    if (sh.kind === 'line') {
      const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
      svg.setAttribute('class', 'dv-ppt-line');
      svg.setAttribute('viewBox', `0 0 ${Math.max(sh.w, 1)} ${Math.max(sh.h, 1)}`);
      svg.setAttribute('preserveAspectRatio', 'none');
      const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
      line.setAttribute('x1', '0'); line.setAttribute('y1', '0');
      line.setAttribute('x2', String(sh.w)); line.setAttribute('y2', String(sh.h));
      line.setAttribute('stroke', sh.line || '#000000');
      line.setAttribute('stroke-width', String(Math.max(sh.lineWidth * scale, 1)));
      line.setAttribute('vector-effect', 'non-scaling-stroke');
      svg.append(line);
      box.append(svg);
      card.append(box);
      continue;
    }

    if (sh.fill) {
      box.style.backgroundColor = sh.fill;
      if (sh.fillOpacity < 1) box.style.backgroundColor = `color-mix(in srgb, ${sh.fill} ${Math.round(sh.fillOpacity * 100)}%, transparent)`;
    }
    if (sh.line) box.style.border = `${Math.max(sh.lineWidth * scale, 1).toFixed(2)}px solid ${sh.line}`;
    if (sh.kind === 'ellipse') box.style.borderRadius = '50%';
    if (sh.kind === 'roundRect') box.style.borderRadius = px(Math.min(sh.w, sh.h) * ((sh.adjust ?? 3600) / 21600));
    if (sh.picture) {
      const img = el('img', 'dv-ppt-picture');
      img.src = urlFor(sh.picture);
      img.alt = '';
      box.append(img);
    }

    if (sh.text) {
      const t = sh.text;
      const inner = el('div', 'dv-ppt-text');
      // Placeholders, titles and body text shrink to their box; other text boxes only when
      // their text would run off the bottom of the slide.
      inner.dataset.fit = t.placeholder ? 'box' : 'slide';
      inner.style.padding = t.insets.map(px).join(' ');
      inner.style.justifyContent = ['flex-start', 'center', 'flex-end'][t.anchor % 3];
      if (!t.wrap) inner.style.whiteSpace = 'pre';
      if (sh.flipH || sh.flipV) inner.style.transform = `scale(${sh.flipH ? -1 : 1}, ${sh.flipV ? -1 : 1})`; // keep text readable
      for (const para of t.paragraphs) {
        const p = el('div', 'dv-ppt-p');
        const pf = para.pf;
        const firstSize = (para.runs[0]?.size ?? para.emptyCf.size ?? 18) * ptPx;
        p.style.fontSize = `${firstSize.toFixed(2)}px`;
        if (pf.align) p.style.textAlign = ['left', 'center', 'right', 'justify'][pf.align] || 'left';
        if (pf.leftMargin !== undefined) p.style.paddingLeft = px(pf.leftMargin);
        if (pf.leftMargin !== undefined || pf.indent !== undefined) p.style.textIndent = px((pf.indent ?? 0) - (pf.leftMargin ?? 0));
        if (pf.lineSpacing !== undefined) p.style.lineHeight = pf.lineSpacing >= 0 ? String((pf.lineSpacing / 100) * 1.2) : px(-pf.lineSpacing);
        const space = (v) => (v >= 0 ? `${((v / 100) * firstSize * 1.2).toFixed(2)}px` : px(-v));
        if (pf.spaceBefore) p.style.marginTop = space(pf.spaceBefore);
        if (pf.spaceAfter) p.style.marginBottom = space(pf.spaceAfter);

        if (pf.hasBullet && !para.empty) {
          const b = el('span', 'dv-ppt-bullet', bulletText(pf.bulletChar, para.bulletFontName));
          const gap = (pf.leftMargin ?? 0) - (pf.indent ?? 0);
          if (gap > 0) b.style.minWidth = px(gap);
          const first = para.runs[0] || para.emptyCf;
          b.style.color = (pf.bulletHasColor && pf.bulletColor) || first.color || '';
          if (pf.bulletHasSize && pf.bulletSize > 0) b.style.fontSize = `${pf.bulletSize}%`;
          p.append(b);
        }
        if (para.empty) p.append(el('br'));
        for (const run of para.runs) {
          const span = el('span');
          const st = span.style;
          if (run.bold) st.fontWeight = '700';
          if (run.italic) st.fontStyle = 'italic';
          if (run.underline) st.textDecoration = 'underline';
          if (run.color) st.color = run.color;
          if (run.size) st.fontSize = `${(run.size * ptPx).toFixed(2)}px`;
          if (run.fontName) st.fontFamily = `"${run.fontName.replace(/"/g, '')}", sans-serif`;
          if (run.position > 0) { st.verticalAlign = 'super'; st.fontSize = '0.7em'; }
          if (run.position < 0) { st.verticalAlign = 'sub'; st.fontSize = '0.7em'; }
          appendWithBreaks(span, run.text);
          p.append(span);
        }
        inner.append(p);
      }
      box.append(inner);
    }
    card.append(box);
  }
  return card;
}

export async function renderPpt({ buffer, mount }) {
  const cfb = openCompound(buffer);
  const pictureEntries = pptPictures(cfb);
  const urls = new Map(); // blob -> object URL
  const urlFor = (blob) => { if (!urls.has(blob)) urls.set(blob, URL.createObjectURL(blob)); return urls.get(blob); };
  const result = (pages, W) => ({
    pages: () => pages,
    pageWidth: () => W,
    naturalWidth: () => W,
    printNodes: () => ({ nodes: clonePagesForPrint(pages, PAGE_WIDTH_LANDSCAPE, () => W), landscape: true }),
    destroy() { urls.forEach((u) => URL.revokeObjectURL(u)); },
  });

  // Slides drawn like PowerPoint first; the simple title + bullets view if that fails.
  try {
    const drawn = readPresentation(stream(cfb, 'PowerPoint Document'), new Map(pictureEntries.map((e) => [e.offset, e.blob])));
    const W = 960;
    const H = Math.round((W * drawn.size.h) / drawn.size.w) || 540;
    const root = el('div', 'dv-legacy dv-legacy-ppt');
    mount.append(root);
    const pages = drawn.slides.map((slide, i) => {
      const card = renderDrawnSlide(slide, drawn.size, W, H, urlFor);
      card.setAttribute('aria-label', `Slide ${i + 1}`);
      root.append(card);
      shrinkToFit(card);
      return card;
    });
    return result(pages, W);
  } catch (err) {
    console.warn('Slides could not be drawn; showing slide text.', err);
    mount.replaceChildren();
  }

  const { slides, size } = pptContent(cfb);
  const pictures = pictureEntries.map((e) => e.blob);
  if (!slides.length && !pictures.length) throw new Error('No slides could be read from this presentation.');

  const W = 960;
  const H = Math.round((W * size.h) / size.w) || 540;
  const root = el('div', 'dv-legacy dv-legacy-ppt');
  mount.append(root);

  const pages = slides.map((slide, i) => {
    const card = el('section', 'dv-page dv-legacy-slide');
    card.style.width = `${W}px`;
    card.style.minHeight = `${H}px`;
    card.setAttribute('aria-label', `Slide ${i + 1}`);
    const titleLines = lines(slide.title);
    if (titleLines.length) card.append(el('h2', null, titleLines.join(' ')));
    const bodyLines = lines(slide.body);
    if (bodyLines.length) {
      const ul = el('ul');
      bodyLines.forEach((line) => ul.append(el('li', null, line)));
      card.append(ul);
    }
    if (!titleLines.length && !bodyLines.length) card.append(el('p', 'dv-legacy-empty', 'No text on this slide'));
    card.append(el('span', 'dv-legacy-num', String(i + 1)));
    root.append(card);
    return card;
  });

  if (pictures.length) {
    const gallery = el('section', 'dv-legacy-pictures');
    gallery.append(el('h3', null, `Pictures in this presentation (${pictures.length})`));
    const grid = el('div', 'dv-legacy-grid');
    for (const blob of pictures) {
      const img = el('img');
      img.src = urlFor(blob);
      img.alt = '';
      grid.append(img);
    }
    gallery.append(grid);
    root.append(gallery);
  }

  return result(pages, W);
}
