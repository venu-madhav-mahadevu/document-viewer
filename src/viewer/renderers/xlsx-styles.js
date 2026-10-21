/*
 * Cell formatting for .xlsx / .xlsm that SheetJS CE does not hand over: which style each
 * cell uses, fonts (bold, italic, underline, strike, color, size, name), alignment,
 * wrapping and borders. Runs in the spreadsheet Web Worker (plain text parsing, no DOM)
 * and only for the rows that are displayed.
 */

const decoder = new TextDecoder();
const attr = (tag, name) => new RegExp(`\\b${name}="([^"]*)"`).exec(tag)?.[1];
const unxml = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');

function fileText(files, path) {
  const f = files[path] || files[`/${path}`];
  return f?.content ? decoder.decode(f.content) : '';
}

/** Theme colors (dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink) in Excel's index order. */
function themeColors(files) {
  const xml = fileText(files, 'xl/theme/theme1.xml');
  const scheme = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(xml)?.[0] || '';
  const get = (name) => {
    const block = new RegExp(`<a:${name}>([\\s\\S]*?)</a:${name}>`).exec(scheme)?.[1] || '';
    return attr(block, 'lastClr') || attr(block, 'val') || null;
  };
  const [dk1, lt1, dk2, lt2] = ['dk1', 'lt1', 'dk2', 'lt2'].map(get);
  return [lt1, dk1, lt2, dk2, ...['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(get)];
}

function colorOf(tag, theme) {
  if (!tag) return null;
  const rgb = attr(tag, 'rgb');
  if (rgb) return `#${rgb.slice(-6)}`;
  const t = attr(tag, 'theme');
  if (t != null && theme[Number(t)]) return `#${theme[Number(t)]}`;
  return null;
}

const BORDER_WIDTH = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const BORDER_STYLE = { dotted: 'dotted', hair: 'dotted', dashed: 'dashed', mediumDashed: 'dashed', dashDot: 'dashed', dashDotDot: 'dashed', mediumDashDot: 'dashed', mediumDashDotDot: 'dashed', slantDashDot: 'dashed', double: 'double' };

function readBorders(styles, theme) {
  const block = /<borders\b[\s\S]*?<\/borders>/.exec(styles)?.[0] || '';
  return [...block.matchAll(/<border\b[^>]*?(?:\/>|>([\s\S]*?)<\/border>)/g)].map((m) => {
    const out = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const s = new RegExp(`<${side}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${side}>)`).exec(m[1] || '');
      const style = s && attr(s[1], 'style');
      if (!style || style === 'none') continue;
      const color = colorOf(/<color\b[^>]*>/.exec(s[2] || '')?.[0], theme) || '#000000';
      out[side] = `${BORDER_WIDTH[style] || 1}px ${BORDER_STYLE[style] || 'solid'} ${color}`;
    }
    return out;
  });
}

/**
 * @param wb SheetJS workbook read with { cellStyles: true, bookFiles: true }
 * @returns {Record<string, Record<string, object>>} sheet name -> cell address -> style
 */
export function readCellStyles(wb, maxRows) {
  const files = wb.files;
  if (!files) return null;
  const styles = fileText(files, 'xl/styles.xml');
  const book = fileText(files, 'xl/workbook.xml');
  const rels = fileText(files, 'xl/_rels/workbook.xml.rels');
  if (!styles || !book) return null;

  const theme = themeColors(files);
  const fonts = wb.Styles?.Fonts || [];
  const xfs = wb.Styles?.CellXf || [];
  const borders = readBorders(styles, theme);
  const defaultSize = fonts[0]?.sz || 11;
  const defaultFont = fonts[0]?.name || '';

  const cache = new Map();
  const styleFor = (index) => {
    if (cache.has(index)) return cache.get(index);
    const xf = xfs[index];
    let out = null;
    if (xf) {
      const s = {};
      const font = fonts[xf.fontId] || {};
      if (font.bold) s.b = true;
      if (font.italic) s.i = true;
      if (font.underline) s.u = true;
      if (font.strike) s.strike = true;
      const rgb = font.color?.rgb;
      if (rgb && !/^0{6}$/.test(rgb) && !(font.color.theme === 1)) s.color = `#${rgb.slice(-6)}`;
      if (font.sz && font.sz !== defaultSize) s.size = font.sz / defaultSize; // relative to the sheet's normal size
      if (font.name && font.name !== defaultFont) s.font = font.name;
      const al = xf.alignment;
      if (al?.horizontal && al.horizontal !== 'general') s.align = al.horizontal;
      if (al?.vertical) s.valign = al.vertical;
      if (al?.wrapText) s.wrap = true;
      if (al?.indent) s.indent = Number(al.indent);
      const border = borders[xf.borderId];
      if (border && Object.keys(border).length) s.border = border;
      if (Object.keys(s).length) out = s;
    }
    cache.set(index, out);
    return out;
  };

  const target = {};
  for (const [tag] of rels.matchAll(/<Relationship\b[^>]*>/g)) target[attr(tag, 'Id')] = attr(tag, 'Target');
  const result = {};
  for (const [tag] of book.matchAll(/<sheet\b[^>]*>/g)) {
    const name = unxml(attr(tag, 'name') || '');
    let path = target[attr(tag, 'r:id')];
    if (!path) continue;
    path = path.startsWith('/') ? path.slice(1) : `xl/${path}`.replace(/xl\/\.\.\//, '');
    const xml = fileText(files, path);
    if (!xml) continue;
    const map = {};
    for (const [cell] of xml.matchAll(/<c\b[^>]*>/g)) {
      const ref = attr(cell, 'r');
      const s = attr(cell, 's');
      if (!ref || !s || s === '0') continue;
      if (Number(/\d+/.exec(ref)?.[0]) > maxRows) break; // rows are stored in order
      const style = styleFor(Number(s));
      if (style) map[ref] = style;
    }
    result[name] = map;
  }
  return result;
}
