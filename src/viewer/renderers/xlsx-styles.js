/*
 * Cell formatting for .xlsx / .xlsm that SheetJS CE does not hand over: which style each
 * cell uses, fonts (bold, italic, underline, strike, color, size, name), alignment,
 * rotation, wrapping, borders, gradient / pattern fills, frozen panes, conditional
 * formatting rules and pictures. Runs in the spreadsheet Web Worker (plain text parsing, no DOM)
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

const rgbOf = (tag, theme) => colorOf(tag, theme);

/** Fills that SheetJS does not describe fully: gradients and patterns (stripes, dots…). */
function readFills(styles, theme) {
  const block = /<fills\b[\s\S]*?<\/fills>/.exec(styles)?.[0] || '';
  return [...block.matchAll(/<fill\b[^>]*>([\s\S]*?)<\/fill>/g)].map(([, f]) => {
    const grad = /<gradientFill\b([^>]*)>([\s\S]*?)<\/gradientFill>/.exec(f);
    if (grad) {
      const stops = [...grad[2].matchAll(/<stop\b[^>]*position="([\d.]+)"[^>]*>([\s\S]*?)<\/stop>/g)]
        .map((m) => `${rgbOf(/<color\b[^>]*>/.exec(m[2])?.[0], theme) || '#ffffff'} ${Number(m[1]) * 100}%`);
      const path = attr(grad[1], 'type') === 'path';
      return { gradient: path ? `radial-gradient(${stops.join(', ')})` : `linear-gradient(${Number(attr(grad[1], 'degree') || 0) + 90}deg, ${stops.join(', ')})` };
    }
    const pat = /<patternFill\b([^>]*)>([\s\S]*?)<\/patternFill>/.exec(f);
    const type = pat && attr(pat[1], 'patternType');
    if (!type || type === 'solid' || type === 'none') return null;
    return {
      pattern: type,
      fg: rgbOf(/<fgColor\b[^>]*>/.exec(pat[2])?.[0], theme) || '#000000',
      bg: rgbOf(/<bgColor\b[^>]*>/.exec(pat[2])?.[0], theme) || '#ffffff',
    };
  });
}

/** Differential formats used by conditional formatting rules: fill, font color, bold. */
function readDxfs(styles, theme) {
  const block = /<dxfs\b[\s\S]*?<\/dxfs>/.exec(styles)?.[0] || '';
  return [...block.matchAll(/<dxf>([\s\S]*?)<\/dxf>/g)].map(([, d]) => {
    const out = {};
    const font = /<font>([\s\S]*?)<\/font>/.exec(d)?.[1] || '';
    const color = rgbOf(/<color\b[^>]*>/.exec(font)?.[0], theme);
    if (color) out.color = color;
    if (/<b\s*\/>/.test(font)) out.bold = true;
    if (/<i\s*\/>/.test(font)) out.italic = true;
    const fill = /<fill>([\s\S]*?)<\/fill>/.exec(d)?.[1] || '';
    // in a dxf the fill color is the background color (Excel writes bgColor, some tools fgColor)
    const bg = rgbOf(/<bgColor\b[^>]*>/.exec(fill)?.[0], theme) || rgbOf(/<fgColor\b[^>]*>/.exec(fill)?.[0], theme);
    if (bg) out.fill = bg;
    return out;
  });
}

/** Conditional formatting rules of a sheet: ranges, rule type and what they show. */
function readConditional(xml, dxfs, theme) {
  const rules = [];
  for (const [, sqref, body] of xml.matchAll(/<conditionalFormatting\b[^>]*sqref="([^"]*)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g)) {
    for (const m of body.matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const a = m[1]; const inner = m[2] || '';
      const cfvo = [...inner.matchAll(/<cfvo\b[^>]*>/g)].map(([t]) => ({ type: attr(t, 'type'), val: attr(t, 'val') }));
      rules.push({
        ranges: sqref.split(/\s+/), type: attr(a, 'type'), priority: Number(attr(a, 'priority') || 0),
        operator: attr(a, 'operator'), text: attr(a, 'text') && unxml(attr(a, 'text')),
        rank: Number(attr(a, 'rank') || 10), percent: attr(a, 'percent') === '1', bottom: attr(a, 'bottom') === '1',
        aboveAverage: attr(a, 'aboveAverage') !== '0', stopIfTrue: attr(a, 'stopIfTrue') === '1',
        formulas: [...inner.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((f) => unxml(f[1])),
        dxf: attr(a, 'dxfId') != null ? dxfs[Number(attr(a, 'dxfId'))] : null,
        cfvo, colors: [...inner.matchAll(/<color\b[^>]*>/g)].map(([t]) => rgbOf(t, theme)),
        iconSet: /<iconSet\b[^>]*iconSet="([^"]*)"/.exec(inner)?.[1] || (/<iconSet\b/.test(inner) ? '3TrafficLights1' : null),
        reverse: /<iconSet\b[^>]*reverse="1"/.test(inner), showValue: !/showValue="0"/.test(inner),
      });
    }
  }
  return rules.sort((x, y) => x.priority - y.priority);
}

const base64 = (bytes) => {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
};
const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp' };
const relTargets = (xml, base) => {
  const out = {};
  for (const [tag] of xml.matchAll(/<Relationship\b[^>]*>/g)) {
    const t = attr(tag, 'Target') || '';
    const parts = (t.startsWith('/') ? t.slice(1) : `${base}/${t}`).split('/');
    const path = [];
    for (const p of parts) { if (p === '..') path.pop(); else if (p !== '.') path.push(p); }
    out[attr(tag, 'Id')] = { path: path.join('/'), type: attr(tag, 'Type') || '' };
  }
  return out;
};

/** Pictures placed on a sheet (Insert > Picture): anchor cell, offsets and size, as data URLs. */
function readPictures(files, sheetPath, xml) {
  const drawingId = /<drawing\b[^>]*r:id="([^"]*)"/.exec(xml)?.[1];
  if (!drawingId) return [];
  const dir = sheetPath.slice(0, sheetPath.lastIndexOf('/'));
  const sheetRels = relTargets(fileText(files, `${dir}/_rels/${sheetPath.slice(dir.length + 1)}.rels`), dir);
  const drawingPath = sheetRels[drawingId]?.path;
  const drawing = drawingPath && fileText(files, drawingPath);
  if (!drawing) return [];
  const ddir = drawingPath.slice(0, drawingPath.lastIndexOf('/'));
  const drels = relTargets(fileText(files, `${ddir}/_rels/${drawingPath.slice(ddir.length + 1)}.rels`), ddir);
  const pics = [];
  const n = (s, tag) => Number(new RegExp(`<xdr:${tag}>(-?\\d+)</xdr:${tag}>`).exec(s)?.[1] || 0);
  for (const [, kind, anchor] of drawing.matchAll(/<xdr:(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g)) {
    const embed = /<a:blip\b[^>]*r:embed="([^"]*)"/.exec(anchor)?.[1];
    const target = embed && drels[embed]?.path;
    const file = target && (files[target] || files[`/${target}`]);
    const mime = IMAGE_TYPES[(target || '').split('.').pop().toLowerCase()];
    if (!file?.content || !mime || file.content.length > 8 * 1024 * 1024) continue;
    const from = /<xdr:from>([\s\S]*?)<\/xdr:from>/.exec(anchor)?.[1] || '';
    const to = /<xdr:to>([\s\S]*?)<\/xdr:to>/.exec(anchor)?.[1];
    const ext = /<xdr:ext\b[^>]*>/.exec(anchor)?.[0];
    pics.push({
      kind,
      from: { col: n(from, 'col'), colOff: n(from, 'colOff'), row: n(from, 'row'), rowOff: n(from, 'rowOff') },
      to: to ? { col: n(to, 'col'), colOff: n(to, 'colOff'), row: n(to, 'row'), rowOff: n(to, 'rowOff') } : null,
      cx: ext ? Number(attr(ext, 'cx')) : 0, cy: ext ? Number(attr(ext, 'cy')) : 0,
      src: `data:${mime};base64,${base64(file.content)}`,
    });
  }
  return pics;
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
  const fills = readFills(styles, theme);
  const dxfs = readDxfs(styles, theme);
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
      if (font.underline) s.u = font.underline === 2 || font.underline === 'double' ? 'double' : true;
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
      if (al?.textRotation && Number(al.textRotation)) s.rotate = Number(al.textRotation);
      const fill = fills[xf.fillId];
      if (fill?.gradient) s.gradient = fill.gradient;
      if (fill?.pattern) s.pattern = fill;
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
    // Frozen rows / columns ("Freeze panes"), under a key no cell address can have.
    const pane = /<pane\b[^>]*state="frozen(?:Split)?"[^>]*>/.exec(xml)?.[0];
    if (pane) map['!freeze'] = { rows: Number(attr(pane, 'ySplit') || 0), cols: Number(attr(pane, 'xSplit') || 0) };
    const cf = readConditional(xml, dxfs, theme);
    if (cf.length) map['!cf'] = cf;
    try { const pics = readPictures(files, path, xml); if (pics.length) map['!pics'] = pics; } catch { /* pictures are optional */ }
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
