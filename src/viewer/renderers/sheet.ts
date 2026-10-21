// @ts-nocheck -- plain JavaScript kept as .ts so Angular bundles its Web Worker.
import * as XLSX from 'xlsx';
import { el } from '../sanitize.js';

const MAX_ROWS = 5000;

/** Cell background from its fill; text turns white on dark fills unless the font color is known. */
function applyFill(td, style, hasFontColor) {
  const rgb = style?.patternType && style.patternType !== 'none' ? style.fgColor?.rgb : null;
  if (!rgb || !/^[0-9a-f]{6}$/i.test(rgb)) return;
  td.style.backgroundColor = `#${rgb}`;
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16));
  if (!hasFontColor && 0.299 * r + 0.587 * g + 0.114 * b < 140) td.style.color = '#ffffff';
}

/** Fonts, alignment, wrapping and borders (from xlsx-styles.js). */
function applyFormat(td, f) {
  const st = td.style;
  if (f.b) st.fontWeight = '700';
  if (f.i) st.fontStyle = 'italic';
  const deco = [f.u && 'underline', f.strike && 'line-through'].filter(Boolean);
  if (deco.length) st.textDecoration = deco.join(' ');
  if (f.color) st.color = f.color;
  if (f.size) st.fontSize = `${f.size}em`;
  if (f.font) st.fontFamily = `"${f.font.replace(/"/g, '')}", "Segoe UI", Arial, sans-serif`;
  if (f.align) st.textAlign = { centerContinuous: 'center', distributed: 'justify', fill: 'left' }[f.align] || f.align;
  if (f.valign) st.verticalAlign = { center: 'middle', distributed: 'middle', justify: 'middle' }[f.valign] || f.valign;
  if (f.wrap) { st.whiteSpace = 'pre-wrap'; st.overflowWrap = 'anywhere'; }
  if (f.indent) st.paddingLeft = `${f.indent * 9 + 4}px`;
  for (const [side, value] of Object.entries(f.border || {})) st[`border${side[0].toUpperCase()}${side.slice(1)}`] = value;
}
const MAX_COLS = 256;

/**
 * Parse in a Web Worker so a 50 MB workbook never freezes the page, and read
 * only the rows we display (`sheetRows`); `!fullref` still reports the full size.
 */
function parseInWorker(buffer, isCsv) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('./sheet.worker', import.meta.url), { type: 'module' });
    worker.onmessage = ({ data }) => {
      worker.terminate();
      if (data.ok) resolve(data.wb);
      else reject(new Error(data.message));
    };
    worker.onerror = (e) => {
      worker.terminate();
      reject(new Error(e.message || 'This spreadsheet could not be read.'));
    };
    worker.postMessage({ buffer, isCsv, maxRows: MAX_ROWS });
  });
}

/**
 * Spreadsheet view with Google-Sheets-style column letters, row numbers and
 * sheet tabs. Cells are written with textContent only, so nothing in a cell
 * can ever be interpreted as HTML.
 */
export async function renderSheet({ buffer, name, mount, onContentChange, download }) {
  const wb = await parseInWorker(buffer, /\.csv$/i.test(name));

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
    grid.replaceChildren(buildTable(wb.Sheets[sheetName], download, wb.CellStyles?.[sheetName]));
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

function buildTable(ws, download, formats) {
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

  const merges = new Map();
  const covered = new Set();
  for (const m of ws['!merges'] || []) {
    merges.set(`${m.s.r}:${m.s.c}`, { rs: m.e.r - m.s.r + 1, cs: m.e.c - m.s.c + 1 });
    for (let r = m.s.r; r <= m.e.r; r++)
      for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) covered.add(`${r}:${c}`);
  }

  const table = el('table', 'dv-grid');
  const colgroup = el('colgroup');
  colgroup.append(el('col', 'dv-rowhead-col'));
  const head = el('tr');
  head.append(el('th', 'dv-corner'));
  for (let c = range.s.c; c <= range.e.c; c++) {
    const col = el('col');
    const wpx = ws['!cols']?.[c]?.wpx;
    if (wpx) col.style.width = `${Math.min(Math.max(wpx, 40), 600)}px`;
    colgroup.append(col);
    head.append(el('th', null, XLSX.utils.encode_col(c)));
  }
  const thead = el('thead');
  thead.append(head);
  table.append(colgroup, thead);

  const tbody = el('tbody');
  for (let r = range.s.r; r <= range.e.r; r++) {
    const tr = el('tr');
    tr.append(el('th', null, String(r + 1)));
    for (let c = range.s.c; c <= range.e.c; c++) {
      const key = `${r}:${c}`;
      if (covered.has(key)) continue;
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      const td = el('td', null, cell ? XLSX.utils.format_cell(cell) : '');
      if (cell && cell.t === 'n') td.className = 'dv-num';
      const format = formats?.[XLSX.utils.encode_cell({ r, c })];
      if (format) applyFormat(td, format);
      if (cell?.s) applyFill(td, cell.s, !!format?.color);
      const m = merges.get(key);
      if (m) { td.rowSpan = m.rs; td.colSpan = m.cs; }
      tr.append(td);
    }
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
