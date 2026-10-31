// @ts-nocheck -- plain JavaScript kept as .ts so Angular bundles its Web Worker.
import * as XLSX from 'xlsx';
import { el } from '../sanitize.js';
import { parseWorkbook } from './sheet-parse.js';

const MAX_ROWS = 5000;

/** Cell background from its fill; text turns white on dark fills unless the font color is known. */
function applyFill(td, style, hasFontColor) {
  const rgb = style?.patternType && style.patternType !== 'none' ? style.fgColor?.rgb : null;
  if (!rgb || !/^[0-9a-f]{6}$/i.test(rgb)) return;
  td.style.backgroundColor = `#${rgb}`;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16));
  if (!hasFontColor && 0.299 * r + 0.587 * g + 0.114 * b < 140) td.style.color = '#ffffff';
}

const FORMAT_COLORS = { red: '#ff0000', blue: '#0000ff', green: '#00a000', yellow: '#c0a000', magenta: '#ff00ff', cyan: '#00a0a0', white: '#ffffff', black: '#000000' };

/** Color from the number format, e.g. "0.0%;[Red]-0.0%" shows negatives in red. */
function formatColor(cell) {
  if (!cell?.z || typeof cell.v !== 'number' || !/\[[a-z]/i.test(cell.z)) return null;
  const sections = cell.z.match(/(?:"[^"]*"|\\.|[^;])+/g) || []; // split on ; outside quotes
  const section = cell.v < 0 && sections[1] ? sections[1] : cell.v === 0 && sections[2] ? sections[2] : sections[0];
  const name = /\[(red|blue|green|yellow|magenta|cyan|white|black)\]/i.exec(section || '')?.[1];
  return name ? FORMAT_COLORS[name.toLowerCase()] : null;
}

/** Fonts, alignment, wrapping and borders (from xlsx-styles.js). */
function applyFormat(td, f) {
  const st = td.style;
  if (f.b) st.fontWeight = '700';
  if (f.i) st.fontStyle = 'italic';
  const deco = [f.u && 'underline', f.strike && 'line-through'].filter(Boolean);
  if (deco.length) st.textDecoration = deco.join(' ');
  if (f.u === 'double') st.textDecorationStyle = 'double';
  if (f.color) st.color = f.color;
  if (f.size) st.fontSize = `${f.size}em`;
  if (f.font) st.fontFamily = `"${f.font.replace(/"/g, '')}", "Segoe UI", Arial, sans-serif`;
  if (f.align) st.textAlign = { centerContinuous: 'center', distributed: 'justify', fill: 'left' }[f.align] || f.align;
  if (f.valign) st.verticalAlign = { center: 'middle', distributed: 'middle', justify: 'middle' }[f.valign] || f.valign;
  if (f.wrap) { st.whiteSpace = 'pre-wrap'; st.overflowWrap = 'anywhere'; }
  if (f.indent) st.paddingLeft = `${f.indent * 9 + 4}px`;
  if (f.gradient) st.backgroundImage = f.gradient;
  if (f.pattern) st.background = patternCss(f.pattern);
  for (const [side, value] of Object.entries(f.border || {})) st[`border${side[0].toUpperCase()}${side.slice(1)}`] = value;
}
const unxml = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

/**
 * Text with mixed formats in one cell (Excel "rich text": red word, bold word, x², H₂O).
 * Built from the cell's runs with textContent, so the text can never become HTML.
 */
function richText(xml) {
  const frag = document.createDocumentFragment();
  for (const [, run] of xml.matchAll(/<r>([\s\S]*?)<\/r>/g)) {
    const text = unxml([...run.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => m[1]).join(''));
    const pr = /<rPr>([\s\S]*?)<\/rPr>/.exec(run)?.[1] || '';
    const has = (tag) => new RegExp(`<${tag}(?:\\s+val="(?:1|true)")?\\s*/>`).test(pr);
    const valign = /<vertAlign val="(superscript|subscript)"/.exec(pr)?.[1];
    const node = el(valign === 'superscript' ? 'sup' : valign === 'subscript' ? 'sub' : 'span', null, text);
    if (has('b')) node.style.fontWeight = '700';
    if (has('i')) node.style.fontStyle = 'italic';
    const u = /<u(?:\s+val="(\w+)")?\s*\/>/.exec(pr);
    const deco = [u && 'underline', has('strike') && 'line-through'].filter(Boolean);
    if (deco.length) node.style.textDecoration = deco.join(' ');
    if (u?.[1] === 'double') node.style.textDecorationStyle = 'double';
    const rgb = /<color\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/.exec(pr)?.[1];
    if (rgb) node.style.color = `#${rgb.slice(-6)}`;
    frag.append(node);
  }
  return frag;
}

/** Sheet named by an in-workbook link ("#'Formulas'!A1", or a .xls "file.xlsx#Formulas!A1"). */
function linkedSheet(link, names) {
  const raw = unxml(link.location || (link.Target || '').split('#')[1] || '').replace(/^#/, '');
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!/.exec(raw);
  const sheet = m ? (m[1] ? m[1].replace(/''/g, "'") : m[2]) : raw;
  return names.includes(sheet) ? sheet : null;
}

/** A cell's content: rich text, a link, rotated text, or the accounting layout ($ at the left edge). */
function fillCell(td, cell, text, format, names, goTo) {
  let content = cell?.t === 's' && cell.r && /<r>/.test(cell.r) ? richText(cell.r) : null;
  // Accounting formats ("_($* #,##0.00_)") put the currency symbol at the left edge of the cell.
  const z = cell?.z ? cell.z.replace(/"[^"]*"/g, '') : '';
  const acct = typeof cell?.v === 'number' && /\*[ -]/.test(z) ? /^(\s*\(?\s*)([$€£¥₹])\s*(.*?)\s*$/.exec(text) : null;
  if (acct) {
    content = document.createDocumentFragment();
    content.append(el('span', 'dv-acct-sym', acct[2]), el('span', null, `${acct[1].trim()}${acct[3]}`));
    td.classList.add('dv-acct');
  }
  if (cell?.l) {
    const sheet = linkedSheet(cell.l, names);
    const url = (cell.l.Target || '').trim();
    if (sheet || /^(https?:|mailto:)/i.test(url)) {
      const a = el('a', 'dv-link');
      if (content) a.append(content); else a.textContent = text;
      if (sheet) {
        a.href = '#';
        a.title = `Go to sheet ${sheet}`;
        a.addEventListener('click', (e) => { e.preventDefault(); goTo(sheet); });
      } else {
        a.href = url;
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        a.title = url;
      }
      content = a;
    }
  }
  if (format?.rotate) {
    // Excel: 1-90 = degrees up (counter-clockwise), 91-180 = degrees down, 255 = stacked letters.
    const box = el('span', 'dv-rot');
    if (content) box.append(content); else box.textContent = text;
    if (format.rotate === 255) box.classList.add('dv-rot-stack');
    else box.style.transform = `rotate(${format.rotate <= 90 ? -format.rotate : format.rotate - 90}deg)`;
    // centered in the (taller) row so the turned text stays inside its cell
    if (!format.valign) td.style.verticalAlign = 'middle';
    if (!format.align) td.style.textAlign = 'center';
    td.style.overflow = 'visible';
    content = box;
  }
  if (content) { td.textContent = ''; td.append(content); }
  // Cell note / comment: red corner mark, text on hover.
  const note = cell?.c?.map((c) => c.t).filter(Boolean).join('\n');
  if (note) { td.classList.add('dv-has-note'); td.title = note; }
}

/** Excel pattern fills (stripes, grids, dots) as CSS backgrounds. */
function patternCss({ pattern, fg, bg }) {
  const dark = /^dark/.test(pattern);
  const w = dark ? 3 : 1.5; const period = dark ? 6 : 6;
  const stripes = (angle) => `repeating-linear-gradient(${angle}deg, ${fg} 0 ${w}px, ${bg} ${w}px ${period}px)`;
  if (/Up$/.test(pattern)) return stripes(-45);
  if (/Down$/.test(pattern)) return stripes(45);
  if (/Horizontal$/.test(pattern)) return stripes(0);
  if (/Vertical$/.test(pattern)) return stripes(90);
  if (/Grid$|Trellis$/.test(pattern)) return `${stripes(0)}, ${stripes(90).replace(bg, 'transparent')}`;
  const dots = { gray0625: 1, gray125: 1.2, lightGray: 1.5, mediumGray: 2, darkGray: 2.6 }[pattern] || 1.5;
  return `radial-gradient(${fg} ${dots}px, ${bg} ${dots}px) 0 0 / 5px 5px`;
}

const ICON_SETS = {
  Arrows: { glyphs: ['↓', '↘', '→', '↗', '↑'], colors: ['#e3463b', '#f08a24', '#f5b400', '#8cc63f', '#3aa457'] },
  ArrowsGray: { glyphs: ['↓', '↘', '→', '↗', '↑'], colors: ['#8a8a8a'] },
  TrafficLights: { glyphs: ['●'], colors: ['#e3463b', '#f5b400', '#3aa457', '#3aa457'] },
  Signs: { glyphs: ['◆', '▲', '●'], colors: ['#e3463b', '#f5b400', '#3aa457'] },
  Symbols: { glyphs: ['✖', '!', '✔'], colors: ['#e3463b', '#f5b400', '#3aa457'] },
  Flags: { glyphs: ['⚑'], colors: ['#e3463b', '#f5b400', '#3aa457'] },
  Rating: { glyphs: ['▁', '▃', '▅', '▆', '█'], colors: ['#2f6eba'] },
  Quarters: { glyphs: ['○', '◔', '◑', '◕', '●'], colors: ['#555555'] },
  RedToBlack: { glyphs: ['●'], colors: ['#222222', '#7a7a7a', '#d8a0a0', '#e3463b'] },
};
/** Icon for position i (0 = lowest) of an icon set such as "3Arrows" or "5Quarters". */
function iconFor(set, i) {
  const n = Number(set[0]) || 3;
  const family = ICON_SETS[set.slice(1).replace(/\d+$/, '')] || ICON_SETS.TrafficLights;
  // spread n icons over the 5-icon arrow / rating scales; others are listed low to high
  const k = family.glyphs.length === 5 && n < 5 ? [[0, 2, 4], [0, 1, 3, 4]][n - 3][i] : i;
  return { glyph: family.glyphs[Math.min(k, family.glyphs.length - 1)] ?? family.glyphs[0], color: family.colors[Math.min(family.colors.length === 1 ? 0 : (family.colors.length === 5 ? k : i), family.colors.length - 1)] };
}

const hexRgb = (c) => { const h = c.replace('#', ''); return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16)); };
const mix = (a, b, t) => `rgb(${hexRgb(a).map((v, i) => Math.round(v + (hexRgb(b)[i] - v) * t)).join(',')})`;
const shiftRef = (ref, dr, dc) => ref.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, ac, col, ar, row) =>
  `${ac}${ac ? col : XLSX.utils.encode_col(XLSX.utils.decode_col(col) + dc)}${ar}${ar ? row : Number(row) + dr}`);

/**
 * Conditional formatting (Excel "Home > Conditional Formatting"): works out each cell's look from
 * the rules and the cell values. Returns "row:col" -> { fill, color, bold, italic, bar, icon }.
 */
function conditionalLooks(ws, rules) {
  const looks = new Map();
  const valueAt = (r, c) => ws[XLSX.utils.encode_cell({ r, c })]?.v;
  for (const rule of rules) {
    const cells = [];
    let origin = null;
    for (const ref of rule.ranges) {
      const rg = XLSX.utils.decode_range(ref.replace(/\$/g, ''));
      origin ||= rg.s;
      for (let r = rg.s.r; r <= Math.min(rg.e.r, rg.s.r + 5000); r++) for (let c = rg.s.c; c <= rg.e.c; c++) cells.push({ r, c, v: valueAt(r, c) });
    }
    const nums = cells.map((x) => x.v).filter((v) => typeof v === 'number').sort((a, b) => a - b);
    const min = nums[0]; const max = nums.at(-1);
    const percentile = (p) => { if (!nums.length) return 0; const i = (nums.length - 1) * p; const lo = Math.floor(i); return nums[lo] + (nums[Math.min(lo + 1, nums.length - 1)] - nums[lo]) * (i - lo); };
    const point = (cfvo, fallback) => {
      const v = Number(cfvo?.val);
      switch (cfvo?.type) {
        case 'min': return min; case 'max': return max;
        case 'num': return v; case 'percent': return min + ((max - min) * v) / 100;
        case 'percentile': return percentile(v / 100);
        default: return fallback;
      }
    };
    const set = (x, look) => {
      const key = `${x.r}:${x.c}`;
      if (looks.get(key)?.stop) return;
      looks.set(key, { ...look, ...looks.get(key), stop: rule.stopIfTrue });
    };
    const matchesText = (v) => String(v ?? '').toLowerCase();
    const numOf = (f) => (/^".*"$/.test(f) ? f.slice(1, -1) : Number(f));
    for (const x of cells) {
      const v = x.v;
      if (rule.type === 'colorScale' && typeof v === 'number' && rule.colors.length >= 2 && max !== undefined) {
        const pts = rule.cfvo.map((c, i) => point(c, i ? max : min));
        let i = 0;
        while (i < pts.length - 2 && v > pts[i + 1]) i++;
        const span = pts[i + 1] - pts[i];
        const t = span ? Math.min(Math.max((v - pts[i]) / span, 0), 1) : 0;
        set(x, { fill: mix(rule.colors[i], rule.colors[i + 1], t) });
      } else if (rule.type === 'dataBar' && typeof v === 'number' && max !== undefined) {
        const lo = point(rule.cfvo[0], min); const hi = point(rule.cfvo[1], max);
        const t = hi > lo ? Math.min(Math.max((v - lo) / (hi - lo), 0), 1) : 1;
        set(x, { bar: { pct: 10 + t * 80, color: rule.colors[0] || '#638ec6' } }); // Excel draws 10%-90% of the cell
      } else if (rule.type === 'iconSet' && typeof v === 'number' && max !== undefined) {
        const n = rule.cfvo.length || 3;
        let i = 0;
        for (let k = 1; k < n; k++) if (v >= point(rule.cfvo[k], max)) i = k;
        set(x, { icon: iconFor(rule.iconSet || '3TrafficLights1', rule.reverse ? n - 1 - i : i), hideValue: !rule.showValue });
      } else if (rule.dxf) {
        let hit = false;
        const [f1, f2] = rule.formulas.map(numOf);
        switch (rule.type) {
          case 'cellIs': {
            const a = typeof f1 === 'number' && typeof v !== 'number' ? NaN : v;
            hit = { greaterThan: a > f1, lessThan: a < f1, greaterThanOrEqual: a >= f1, lessThanOrEqual: a <= f1, equal: a === f1 || String(v) === String(f1),
              notEqual: a !== f1, between: a >= Math.min(f1, f2) && a <= Math.max(f1, f2), notBetween: !(a >= Math.min(f1, f2) && a <= Math.max(f1, f2)) }[rule.operator] ?? false;
            break;
          }
          case 'duplicateValues': case 'uniqueValues': {
            const n = cells.filter((y) => y.v != null && String(y.v).toLowerCase() === String(v).toLowerCase()).length;
            hit = v != null && v !== '' && (rule.type === 'duplicateValues' ? n > 1 : n === 1);
            break;
          }
          case 'top10': {
            if (typeof v !== 'number') break;
            const k = rule.percent ? Math.max(1, Math.floor((nums.length * rule.rank) / 100)) : rule.rank;
            hit = rule.bottom ? v <= nums[Math.min(k, nums.length) - 1] : v >= nums[Math.max(nums.length - k, 0)];
            break;
          }
          case 'aboveAverage': {
            if (typeof v !== 'number') break;
            const avg = nums.reduce((s, n) => s + n, 0) / nums.length;
            hit = rule.aboveAverage ? v > avg : v < avg;
            break;
          }
          case 'containsText': hit = matchesText(v).includes((rule.text || '').toLowerCase()); break;
          case 'notContainsText': hit = !matchesText(v).includes((rule.text || '').toLowerCase()); break;
          case 'beginsWith': hit = matchesText(v).startsWith((rule.text || '').toLowerCase()); break;
          case 'endsWith': hit = matchesText(v).endsWith((rule.text || '').toLowerCase()); break;
          case 'containsBlanks': hit = v == null || String(v).trim() === ''; break;
          case 'notContainsBlanks': hit = !(v == null || String(v).trim() === ''); break;
          case 'expression': {
            // The common "COUNTIF(range, cell) > n" duplicate test; refs move with the cell like Excel's.
            const f = shiftRef(rule.formulas[0] || '', x.r - origin.r, x.c - origin.c);
            const m = /^COUNTIF\(([$A-Z0-9:]+),\s*([$A-Z]+\$?\d+)\)\s*(>=|<=|>|<|=)\s*(\d+)$/i.exec(f.replace(/\s+/g, ''));
            if (m) {
              const rg = XLSX.utils.decode_range(m[1].replace(/\$/g, ''));
              const target = ws[m[2].replace(/\$/g, '')]?.v;
              let n = 0;
              for (let r = rg.s.r; r <= rg.e.r; r++) for (let c = rg.s.c; c <= rg.e.c; c++) if (target != null && valueAt(r, c) === target) n++;
              const k = Number(m[4]);
              hit = { '>': n > k, '<': n < k, '>=': n >= k, '<=': n <= k, '=': n === k }[m[3]];
            }
            break;
          }
          default: break;
        }
        if (hit) set(x, { fill: rule.dxf.fill, color: rule.dxf.color, bold: rule.dxf.bold, italic: rule.dxf.italic });
      }
    }
  }
  return looks;
}

/** Pictures on the sheet, placed over the grid at their anchor cell (positions measured after layout). */
function placePictures(grid, table, pics) {
  grid.querySelector('.dv-sheet-pics')?.remove();
  if (!pics?.length) return;
  const layer = el('div', 'dv-sheet-pics');
  const headCells = new Map([...table.tHead.rows[0].cells].map((th) => [Number(th.dataset.c), th]));
  const rowEls = new Map([...table.tBodies[0].rows].map((tr) => [Number(tr.dataset.r), tr]));
  // offsets: EMU (.xlsx) or a fraction of the cell (.xls)
  const colX = (c, off, frac) => { const th = headCells.get(c); return th ? th.offsetLeft + (frac != null ? frac * th.offsetWidth : off / 9525) : null; };
  const rowY = (r, off, frac) => { const tr = rowEls.get(r); return tr ? tr.offsetTop + (frac != null ? frac * tr.offsetHeight : off / 9525) : null; };
  for (const p of pics) {
    const x = colX(p.from.col, p.from.colOff, p.from.colFrac); const y = rowY(p.from.row, p.from.rowOff, p.from.rowFrac);
    if (x == null || y == null) continue;
    let w = p.cx / 9525; let h = p.cy / 9525;
    if (p.to) { const x2 = colX(p.to.col, p.to.colOff, p.to.colFrac); const y2 = rowY(p.to.row, p.to.rowOff, p.to.rowFrac); if (x2 != null) w = x2 - x; if (y2 != null) h = y2 - y; }
    const img = el('img', 'dv-sheet-pic');
    img.src = p.src;
    img.alt = 'Picture';
    Object.assign(img.style, { left: `${table.offsetLeft + x}px`, top: `${table.offsetTop + y}px`, width: `${w}px`, height: `${h}px` });
    layer.append(img);
  }
  grid.append(layer);
}

/** Keep frozen rows (Freeze panes) on top while scrolling: each row sits below the ones above it. */
function freezeRows(table) {
  let top = table.tHead?.offsetHeight || 0;
  for (const tr of table.querySelectorAll('tbody tr.dv-frozen')) {
    for (const cell of tr.children) { cell.style.position = 'sticky'; cell.style.top = `${top}px`; cell.style.zIndex = cell.tagName === 'TH' ? '3' : '2'; } // z 2: above pictures (z 1), like Excel's frozen panes
    top += tr.offsetHeight;
  }
}
const MAX_COLS = 256;
const DEFAULT_COL_WIDTH = 100; // px, matches .dv-grid col in viewer.css
const ROWHEAD_WIDTH = 46; // px, .dv-grid col.dv-rowhead-col

/**
 * Parse in a Web Worker so a 50 MB workbook never freezes the page, and read
 * only the rows we display (`sheetRows`); `!fullref` still reports the full size.
 */
function parseInWorker(buffer, isCsv) {
  return new Promise((resolve, reject) => {
    let worker;
    try {
      worker = new Worker(new URL('./sheet.worker', import.meta.url), { type: 'module' });
    } catch (err) {
      reject(Object.assign(err, { workerFailed: true }));
      return;
    }
    worker.onmessage = ({ data }) => {
      worker.terminate();
      if (data.ok) resolve(data.wb);
      else reject(new Error(data.message)); // the file itself could not be read
    };
    worker.onerror = (e) => {
      worker.terminate();
      // The worker could not start or load (not a problem with the file).
      reject(Object.assign(new Error(e.message || 'Spreadsheet worker failed.'), { workerFailed: true }));
    };
    worker.postMessage({ buffer, isCsv, maxRows: MAX_ROWS });
  });
}

/** Worker first (keeps big files from freezing the page); the page itself if workers are blocked. */
async function parse(buffer, isCsv) {
  try {
    return await parseInWorker(buffer, isCsv);
  } catch (err) {
    if (!err.workerFailed) throw err;
    console.warn('Spreadsheet worker unavailable; reading on the main thread.', err);
    return parseWorkbook(buffer, isCsv, MAX_ROWS);
  }
}

/**
 * Spreadsheet view with Google-Sheets-style column letters, row numbers and
 * sheet tabs. Cells are written with textContent only, so nothing in a cell
 * can ever be interpreted as HTML.
 */
export async function renderSheet({ buffer, name, mount, onContentChange, download }) {
  const isCsv = /\.csv$/i.test(name);
  const wb = await parse(buffer, isCsv);

  const root = el('div', 'dv-sheet');
  const grid = el('div', 'dv-sheet-grid');
  const tabs = el('div', 'dv-sheet-tabs');
  tabs.setAttribute('role', 'tablist');
  root.append(grid, tabs);
  mount.append(root);

  const visible = wb.SheetNames.filter((_, i) => !(wb.Workbook?.Sheets?.[i]?.Hidden));
  const buttons = visible.map((sheetName) => {
    const b = el('button', 'dv-sheet-tab', sheetName);
    b.setAttribute('role', 'tab');
    b.addEventListener('click', () => show(sheetName));
    tabs.append(b);
    return b;
  });

  function show(sheetName) {
    buttons.forEach((b) => b.setAttribute('aria-selected', String(b.textContent === sheetName)));
    grid.replaceChildren(buildTable(wb.Sheets[sheetName], download, wb.CellStyles?.[sheetName], isCsv, visible, show));
    const table = grid.querySelector('table.dv-grid');
    if (table) { freezeRows(table); placePictures(grid, table, wb.CellStyles?.[sheetName]?.['!pics']); }
    grid.scrollTo(0, 0);
    onContentChange?.();
  }

  show(visible[0] ?? wb.SheetNames[0]);
  return {
    // Prints the sheet that is currently open, like a spreadsheet app does.
    printNodes: () => ({ nodes: [...grid.children].map((n) => n.cloneNode(true)) }),
    destroy() {},
  };
}

/** Column width in pixels: Excel stores it in characters of the default font (~7 px each). */
function columnWidth(info) {
  const px = info?.width ? Math.round(info.width * 7 + 5) : info?.wpx;
  return px ? Math.min(Math.max(px, 20), 800) : DEFAULT_COL_WIDTH;
}

function buildTable(ws, download, formats, isCsv, names = [], goTo = () => {}) {
  const frag = document.createDocumentFragment();
  if (!ws || !ws['!ref']) {
    frag.append(el('p', 'dv-note', 'This sheet is empty.'));
    return frag;
  }
  // `!fullref` = the sheet's real size when only the first MAX_ROWS were read.
  const full = XLSX.utils.decode_range(ws['!fullref'] || ws['!ref']);
  const range = {
    s: full.s,
    e: { r: Math.min(full.e.r, full.s.r + MAX_ROWS - 1), c: Math.min(full.e.c, full.s.c + MAX_COLS - 1) },
  };
  // Rows and columns hidden in Excel stay hidden.
  const cols = [];
  for (let c = range.s.c; c <= range.e.c; c++) if (!ws['!cols']?.[c]?.hidden) cols.push(c);
  const visibleCols = new Set(cols);

  const merges = new Map();
  const covered = new Set();
  for (const m of ws['!merges'] || []) {
    let cs = 0;
    for (let c = m.s.c; c <= m.e.c; c++) if (visibleCols.has(c)) cs++;
    let rs = 0;
    for (let r = m.s.r; r <= m.e.r; r++) if (!ws['!rows']?.[r]?.hidden) rs++;
    merges.set(`${m.s.r}:${m.s.c}`, { rs: Math.max(rs, 1), cs: Math.max(cs, 1) });
    for (let r = m.s.r; r <= m.e.r; r++)
      for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) covered.add(`${r}:${c}`);
  }

  // Freeze panes: the first rows / columns stay in view while scrolling.
  const freeze = formats?.['!freeze'] || {};
  const leftOf = new Map(); // frozen column -> its left edge (px)

  const table = el('table', 'dv-grid');
  const colgroup = el('colgroup');
  colgroup.append(el('col', 'dv-rowhead-col'));
  const head = el('tr');
  head.append(el('th', 'dv-corner'));
  let tableWidth = ROWHEAD_WIDTH;
  // CSV files have no column widths: size each column to its longest value (within reason).
  const fitWidth = (c) => {
    let longest = 0;
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 200); r++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell) for (const line of XLSX.utils.format_cell(cell).split('\n')) longest = Math.max(longest, line.length);
    }
    return Math.min(Math.max(longest * 7 + 16, 60), 420);
  };
  for (const c of cols) {
    const col = el('col');
    const width = isCsv && !ws['!cols']?.[c] ? fitWidth(c) : columnWidth(ws['!cols']?.[c]);
    col.style.width = `${width}px`;
    const th = el('th', null, XLSX.utils.encode_col(c));
    th.dataset.c = String(c);
    if (c < (freeze.cols || 0)) { leftOf.set(c, tableWidth); th.classList.add('dv-frozen-col'); th.style.left = `${tableWidth}px`; }
    tableWidth += width;
    colgroup.append(col);
    head.append(th);
  }
  // A fixed table width makes the browser keep the column widths (long text no longer widens a column).
  table.style.width = `${tableWidth}px`;
  const thead = el('thead');
  thead.append(head);
  table.append(colgroup, thead);

  const isEmpty = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({ r, c })];
    return (!cell || cell.v == null || cell.v === '') && !covered.has(`${r}:${c}`) && !merges.has(`${r}:${c}`);
  };

  const looks = formats?.['!cf'] ? conditionalLooks(ws, formats['!cf']) : null;
  const filter = ws['!autofilter']?.ref ? XLSX.utils.decode_range(ws['!autofilter'].ref) : null;

  const tbody = el('tbody');
  for (let r = range.s.r; r <= range.e.r; r++) {
    if (ws['!rows']?.[r]?.hidden) continue;
    const tr = el('tr');
    tr.dataset.r = String(r);
    if (r < (freeze.rows || 0)) tr.classList.add('dv-frozen', ...(r === freeze.rows - 1 ? ['dv-frozen-last'] : []));
    const rowHeight = ws['!rows']?.[r]?.hpx;
    if (rowHeight && rowHeight > 24) tr.style.height = `${Math.round(rowHeight)}px`; // taller rows from Excel
    tr.append(el('th', null, String(r + 1)));
    cols.forEach((c, i) => {
      const key = `${r}:${c}`;
      if (covered.has(key)) return;
      let cell = ws[XLSX.utils.encode_cell({ r, c })];
      // .xls rich text comes from xls-format.js in the same run form as .xlsx
      const rich = formats?.['!rich']?.[XLSX.utils.encode_cell({ r, c })];
      if (rich && cell && !cell.r) cell = { ...cell, r: rich };
      const text = cell ? XLSX.utils.format_cell(cell) : '';
      const td = el('td', null, text);
      if (cell && cell.t === 'n') td.className = 'dv-num';
      const format = formats?.[XLSX.utils.encode_cell({ r, c })];
      if (format) applyFormat(td, format);
      if (cell?.s) applyFill(td, cell.s, !!format?.color);
      const numberColor = formatColor(cell);
      if (numberColor) td.style.color = numberColor;
      const look = looks?.get(key);
      if (look) {
        if (look.fill) td.style.backgroundColor = look.fill;
        if (look.color) td.style.color = look.color;
        if (look.bold) td.style.fontWeight = '700';
        if (look.italic) td.style.fontStyle = 'italic';
        if (look.bar) td.style.backgroundImage = `linear-gradient(to right, ${look.bar.color}, ${look.bar.color}66 ${look.bar.pct}%, transparent ${look.bar.pct}%)`;
      }
      // filter buttons on the header row of a filtered range (shown, not working: the viewer is read-only)
      if (filter && r === filter.s.r && c >= filter.s.c && c <= filter.e.c) td.classList.add('dv-filter');
      const m = merges.get(key);
      if (m) { td.rowSpan = m.rs; td.colSpan = m.cs; }
      if (isCsv && text.includes('\n')) td.style.whiteSpace = 'pre-wrap'; // Excel shows line breaks in a CSV cell
      // Like Excel, left-aligned text runs on into empty cells to its right instead of being cut off.
      else if (text && cell?.t === 's' && !m && !format?.wrap && !format?.align && cols[i + 1] != null && isEmpty(r, cols[i + 1])) td.classList.add('dv-spill');
      if (text.length > 30) td.title = text; // full text on hover when it is cut off
      fillCell(td, cell, text, format, names, goTo);
      if (look?.icon) {
        if (look.hideValue) td.textContent = '';
        const icon = el('span', 'dv-cf-icon', look.icon.glyph);
        icon.style.color = look.icon.color;
        td.prepend(icon);
      }
      if (leftOf.has(c)) { td.classList.add('dv-frozen-col'); td.style.left = `${leftOf.get(c)}px`; }
      tr.append(td);
    });
    tbody.append(tr);
  }
  table.append(tbody);
  frag.append(table);

  if (full.e.r > range.e.r || full.e.c > range.e.c) {
    frag.append(el('p', 'dv-note',
      `Showing the first ${(range.e.r - range.s.r + 1).toLocaleString()} of ${(full.e.r - full.s.r + 1).toLocaleString()} rows` +
      ` and ${range.e.c - range.s.c + 1} of ${full.e.c - full.s.c + 1} columns.${download ? ' Download the file to see everything.' : ''}`));
  }
  return frag;
}
