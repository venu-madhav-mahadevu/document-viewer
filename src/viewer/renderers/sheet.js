import * as XLSX from 'xlsx';
import { el } from '../sanitize.js';

const MAX_ROWS = 5000;
const MAX_COLS = 256;

/**
 * Spreadsheet view with Google-Sheets-style column letters, row numbers and
 * sheet tabs. Cells are written with textContent only, so nothing in a cell
 * can ever be interpreted as HTML.
 */
export async function renderSheet({ buffer, name, mount }) {
  const isText = /\.(csv|tsv)$/i.test(name);
  const wb = isText
    ? XLSX.read(new TextDecoder().decode(buffer), { type: 'string', raw: true })
    : XLSX.read(buffer, { type: 'array', cellDates: true, cellFormula: false, cellHTML: false });

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
    grid.replaceChildren(buildTable(wb.Sheets[sheetName]));
    grid.scrollTo(0, 0);
  }

  show(visible[0] ?? wb.SheetNames[0]);
  return { destroy() {} };
}

function buildTable(ws) {
  const frag = document.createDocumentFragment();
  if (!ws || !ws['!ref']) {
    frag.append(el('p', 'dv-note', 'This sheet is empty.'));
    return frag;
  }
  const full = XLSX.utils.decode_range(ws['!ref']);
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
      `Showing the first ${range.e.r - range.s.r + 1} rows and ${range.e.c - range.s.c + 1} columns. ` +
      'Download the file to see everything.'));
  }
  return frag;
}
