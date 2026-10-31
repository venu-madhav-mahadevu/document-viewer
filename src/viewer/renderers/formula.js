/*
 * A small spreadsheet formula evaluator for showing results a file did not save: numbers, text,
 * TRUE/FALSE, cell and range references (also 'Other sheet'!A1), + - * / ^ & and comparisons,
 * and common functions. Anything else throws, and the cell keeps its saved value.
 */
import * as XLSX from 'xlsx';

const FN = {
  SUM: (a) => nums(a).reduce((s, n) => s + n, 0),
  AVERAGE: (a) => { const n = nums(a); if (!n.length) throw new Error('#DIV/0!'); return n.reduce((s, x) => s + x, 0) / n.length; },
  MIN: (a) => Math.min(...nums(a)), MAX: (a) => Math.max(...nums(a)),
  COUNT: (a) => nums(a).length, COUNTA: (a) => a.flat(Infinity).filter((v) => v != null && v !== '').length,
  IF: (a) => (truthy(a[0]) ? a[1] : a[2] ?? false),
  AND: (a) => a.flat(Infinity).every(truthy), OR: (a) => a.flat(Infinity).some(truthy), NOT: (a) => !truthy(a[0]),
  CONCATENATE: (a) => a.flat(Infinity).map(text).join(''), CONCAT: (a) => a.flat(Infinity).map(text).join(''),
  ROUND: (a) => { const f = 10 ** (a[1] || 0); return Math.round(num(a[0]) * f) / f; },
  ABS: (a) => Math.abs(num(a[0])), LEN: (a) => text(a[0]).length,
  UPPER: (a) => text(a[0]).toUpperCase(), LOWER: (a) => text(a[0]).toLowerCase(), TRIM: (a) => text(a[0]).trim().replace(/\s+/g, ' '),
  LEFT: (a) => text(a[0]).slice(0, a[1] ?? 1), RIGHT: (a) => text(a[0]).slice(-(a[1] ?? 1)),
};
const nums = (a) => a.flat(Infinity).filter((v) => typeof v === 'number');
const num = (v) => (typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : v === '' || v == null ? 0 : Number.isNaN(Number(v)) ? (() => { throw new Error('#VALUE!'); })() : Number(v));
const text = (v) => (v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'number' ? String(+v.toPrecision(15)) : String(v));
const truthy = (v) => (typeof v === 'string' ? v.toUpperCase() === 'TRUE' : !!v);

function tokenize(f) {
  const out = [];
  const re = /\s*(?:("(?:[^"]|"")*")|('(?:[^']|'')+'|[A-Za-z_][\w.]*)!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(])|(\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)|([A-Za-z_][\w.]*)\s*\(|(TRUE|FALSE)\b|(<>|<=|>=|[-+*/^&=<>(),:]))/y;
  let m;
  while (re.lastIndex < f.length && (m = re.exec(f))) {
    if (m[1]) out.push({ t: 'str', v: m[1].slice(1, -1).replace(/""/g, '"') });
    else if (m[3]) out.push({ t: 'ref', sheet: m[2].replace(/^'|'$/g, '').replace(/''/g, "'"), v: m[3] });
    else if (m[4]) out.push({ t: 'ref', v: m[4] });
    else if (m[5]) out.push({ t: 'num', v: Number(m[5]) });
    else if (m[6]) out.push({ t: 'fn', v: m[6].toUpperCase() });
    else if (m[7]) out.push({ t: 'bool', v: m[7] === 'TRUE' });
    else out.push({ t: 'op', v: m[8] });
  }
  if (re.lastIndex < f.trim().length && out.length === 0) throw new Error('unsupported');
  return out;
}

export function evaluate(formula, sheetName, wb) {
  const toks = tokenize(formula.replace(/^=/, ''));
  let i = 0;
  const peek = () => toks[i];
  const take = (v) => { if (toks[i]?.v === v) { i++; return true; } return false; };
  const cellValue = (sheet, a) => {
    const ws = wb.Sheets[sheet];
    if (!ws) throw new Error('#REF!');
    const c = ws[a.replace(/\$/g, '')];
    return c ? c.v : null;
  };
  const refValue = (tok) => {
    const sheet = tok.sheet || sheetName;
    if (!tok.v.includes(':')) return cellValue(sheet, tok.v);
    const rg = XLSX.utils.decode_range(tok.v.replace(/\$/g, ''));
    const rows = [];
    for (let r = rg.s.r; r <= rg.e.r; r++) {
      const row = [];
      for (let c = rg.s.c; c <= rg.e.c; c++) row.push(cellValue(sheet, XLSX.utils.encode_cell({ r, c })));
      rows.push(row);
    }
    return rows;
  };
  // precedence: comparison < & < + - < * / < ^ < unary
  const primary = () => {
    const tk = toks[i++];
    if (!tk) throw new Error('end');
    if (tk.t === 'num' || tk.t === 'str' || tk.t === 'bool') return tk.v;
    if (tk.t === 'ref') return refValue(tk);
    if (tk.t === 'fn') {
      const args = [];
      if (!take(')')) {
        do args.push(compare()); while (take(','));
        if (!take(')')) throw new Error('paren');
      }
      const fn = FN[tk.v];
      if (!fn) throw new Error(`unsupported ${tk.v}`);
      return fn(args);
    }
    if (tk.v === '(') { const v = compare(); if (!take(')')) throw new Error('paren'); return v; }
    if (tk.v === '-') return -num(unary());
    if (tk.v === '+') return num(unary());
    throw new Error('syntax');
  };
  const unary = () => primary();
  const power = () => { let v = unary(); while (take('^')) v = num(v) ** num(unary()); return v; };
  const product = () => { let v = power(); for (;;) { if (take('*')) v = num(v) * num(power()); else if (take('/')) { const d = num(power()); if (!d) throw new Error('#DIV/0!'); v = num(v) / d; } else return v; } };
  const sum = () => { let v = product(); for (;;) { if (take('+')) v = num(v) + num(product()); else if (take('-')) v = num(v) - num(product()); else return v; } };
  const concat = () => { let v = sum(); while (take('&')) v = text(v) + text(sum()); return v; };
  const compare = () => {
    let v = concat();
    for (;;) {
      const op = peek()?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v) ? toks[i++].v : null;
      if (!op) return v;
      const w = concat();
      const [a, b] = typeof v === 'string' || typeof w === 'string' ? [text(v).toLowerCase(), text(w).toLowerCase()] : [num(v), num(w)];
      v = { '=': a === b, '<>': a !== b, '<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b }[op];
    }
  };
  const v = compare();
  if (i !== toks.length) throw new Error('trailing');
  return Array.isArray(v) ? v[0]?.[0] : v;
}
