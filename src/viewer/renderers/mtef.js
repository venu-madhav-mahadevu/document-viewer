/**
 * Equation Editor 3.0 / MathType equations (the "Equation Native" stream of an embedded
 * equation, MTEF version 3) drawn as MathML, which browsers lay out themselves.
 */

const MATHML = 'http://www.w3.org/1998/Math/MathML';
const u16 = (b, p) => b[p] | (b[p + 1] << 8);

// MTEF v3 template selectors used here (others are drawn as their slots side by side)
const T = { ANGLE: 0, PAREN: 1, BRACE: 2, BRACK: 3, BAR: 4, DBAR: 5, FLOOR: 6, CEILING: 7, ROOT: 13, FRACT: 14, SCRIPT: 15, UBAR: 16, OBAR: 17, SUM: 29, ISUM: 30, PROD: 31, IPROD: 32, LIM: 39, SLFRACT: 41 };
const FENCES = { [T.ANGLE]: ['⟨', '⟩'], [T.PAREN]: ['(', ')'], [T.BRACE]: ['{', '}'], [T.BRACK]: ['[', ']'], [T.BAR]: ['|', '|'], [T.DBAR]: ['‖', '‖'], [T.FLOOR]: ['⌊', '⌋'], [T.CEILING]: ['⌈', '⌉'] };
// Greek letters in the Symbol font (typefaces 4 and 5)
const GREEK = 'ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ';
const greek = (c) => {
  if (c >= 0x41 && c <= 0x5a) return GREEK[c - 0x41];
  if (c >= 0x61 && c <= 0x7a) return GREEK[c - 0x61].toLowerCase().replace('ϑ', 'θ').replace('ς', 'ς');
  return String.fromCharCode(c);
};
const SYMBOL = { 0xb1: '±', 0xb4: '×', 0xb8: '÷', 0xb9: '≠', 0xa3: '≤', 0xb3: '≥', 0xbb: '≈', 0xa5: '∞', 0xd6: '√', 0xe5: '∑', 0xf2: '∫', 0xb6: '∂', 0xae: '→', 0xac: '←', 0xd7: '·', 0x2d: '−' };

/**
 * Reads an Equation Native stream into a tree: lines (arrays of items), where an item is
 * { char, face } or { tmpl, variation, slots: [line…] } or { pile: [line…] }. Null if the data is
 * not MTEF version 3.
 */
export function readMtef(bytes) {
  if (!bytes || bytes.length < 34) return null;
  const hdr = u16(bytes, 0); // EQNOLEFILEHDR size (28)
  let p = hdr;
  if (bytes[p] !== 3) return null; // MTEF version
  p += 5; // version, platform, product, product version, subversion
  const nudge = (opts) => {
    if (!(opts & 0x08)) return;
    const dx = bytes[p]; const dy = bytes[p + 1];
    p += 2;
    if (dx === 128 && dy === 128) p += 4;
  };

  // one object list, up to its END record
  const list = () => {
    const items = [];
    while (p < bytes.length) {
      const tag = bytes[p++];
      const type = tag & 0x0f;
      const opts = tag >> 4;
      if (type === 0) break; // END
      if (type === 1) { // LINE: a row of items (an empty slot when "null")
        nudge(opts);
        if (opts & 0x02) ruler();
        items.push({ line: opts & 0x01 ? [] : list() });
      } else if (type === 2) { // CHAR
        nudge(opts);
        const face = bytes[p++] - 128;
        const code = u16(bytes, p); p += 2;
        const item = { char: code, face };
        if (opts & 0x01) item.embell = list().map((e) => e.embell).filter(Boolean);
        items.push(item);
      } else if (type === 3) { // TMPL
        nudge(opts);
        const tmpl = bytes[p++]; const variation = bytes[p++]; p++; // template options
        items.push({ tmpl, variation, slots: list().map((s) => s.line ?? s.pile ?? []) });
      } else if (type === 4) { // PILE: lines stacked
        nudge(opts);
        p += 2; // alignments
        if (opts & 0x02) ruler();
        items.push({ pile: list().map((s) => s.line ?? []) });
      } else if (type === 5) { // MATRIX
        nudge(opts);
        p += 3;
        const rows = bytes[p++]; const cols = bytes[p++];
        p += Math.ceil(((rows + 1) * 2) / 8) + Math.ceil(((cols + 1) * 2) / 8);
        const cells = list().map((s) => s.line ?? []);
        items.push({ matrix: Array.from({ length: rows }, (_, r) => cells.slice(r * cols, r * cols + cols)) });
      } else if (type === 6) { // EMBELL (dot, bar, prime… over a character)
        nudge(opts);
        items.push({ embell: bytes[p++] });
      } else if (type === 7) ruler();
      else if (type === 8) { p += 2; while (p < bytes.length && bytes[p++] !== 0); } // FONT
      else if (type === 9) { // SIZE
        const s = bytes[p++];
        if (s === 101) p += 2; else if (s === 100) p += 3; else p += 1;
      } else if (type >= 10 && type <= 14) continue; // size shortcuts
      else break;
    }
    return items;
  };
  function ruler() { const n = bytes[p++]; p += n * 3; }

  const top = list();
  const flatten = (items) => items.flatMap((it) => (it.line ? flatten(it.line) : [it]));
  return flatten(top);
}

/** The equation as a <math> element. */
export function mtefToMathML(items, doc = document) {
  const m = (name, ...kids) => {
    const e = doc.createElementNS(MATHML, name);
    for (const k of kids) e.append(k);
    return e;
  };
  const row = (line) => {
    const out = [];
    for (const it of line || []) {
      if (it.line) { out.push(row(it.line)); continue; }
      if (it.char !== undefined) {
        const c = it.char;
        const face = it.face;
        let text = face === 4 || face === 5 ? greek(c) : face === 6 && SYMBOL[c] ? SYMBOL[c] : String.fromCharCode(c);
        if (face === 6 && !SYMBOL[c] && c > 0x7f) text = String.fromCharCode(c);
        let node;
        if (/^[0-9.]$/.test(text)) node = m('mn', text);
        else if (face === 1 && /^[A-Za-z]$/.test(text)) node = m('mtext', text);
        else if (/^[A-Za-zα-ωΑ-Ω]$/.test(text)) node = m('mi', text);
        else if (face === 1 && text === ' ') node = m('mspace');
        else node = m('mo', text);
        if (node.localName === 'mspace') node.setAttribute('width', '0.3em');
        if (it.embell?.length) {
          const accent = { 2: '˙', 3: '¨', 4: '⃛', 5: '′', 6: '″', 7: '‴', 8: '^', 9: '~', 10: '¯', 11: '→' }[it.embell[0]];
          if (accent) { node = m('mover', node, m('mo', accent)); node.setAttribute('accent', 'true'); }
        }
        out.push(node);
        continue;
      }
      if (it.pile) { out.push(m('mtable', ...it.pile.map((l) => m('mtr', m('mtd', row(l)))))); continue; }
      if (it.matrix) { out.push(m('mtable', ...it.matrix.map((r) => m('mtr', ...r.map((cell) => m('mtd', row(cell))))))); continue; }
      if (it.tmpl === undefined) continue;
      const s = it.slots;
      switch (it.tmpl) {
        case T.FRACT: case T.SLFRACT: out.push(m('mfrac', row(s[0]), row(s[1]))); break;
        case T.ROOT: out.push(s[1]?.length ? m('mroot', row(s[0]), row(s[1])) : m('msqrt', row(s[0]))); break;
        case T.SCRIPT: {
          // subscript / superscript of the item before it
          const base = out.pop() || m('mrow');
          const sub = s[0]?.length ? row(s[0]) : null;
          const sup = s[1]?.length ? row(s[1]) : null;
          out.push(sub && sup ? m('msubsup', base, sub, sup) : sup ? m('msup', base, sup) : sub ? m('msub', base, sub) : base);
          break;
        }
        case T.UBAR: out.push(m('munder', row(s[0]), m('mo', '_'))); break;
        case T.OBAR: { const o = m('mover', row(s[0]), m('mo', '¯')); o.setAttribute('accent', 'true'); out.push(o); break; }
        case T.SUM: case T.ISUM: case T.PROD: case T.IPROD: {
          const op = m('mo', it.tmpl <= T.ISUM ? '∑' : '∏');
          const lim = s[1]?.length || s[2]?.length ? m('munderover', op, row(s[1]), row(s[2])) : op;
          out.push(lim, row(s[0]));
          break;
        }
        case T.LIM: out.push(m('munder', m('mi', 'lim'), row(s[1])), row(s[0])); break;
        default: {
          const f = FENCES[it.tmpl];
          if (f) out.push(m('mrow', m('mo', f[0]), row(s[0]), m('mo', f[1])));
          else out.push(...s.map(row));
        }
      }
    }
    return m('mrow', ...out);
  };
  const math = m('math', row(items));
  math.setAttribute('display', 'inline');
  return math;
}
