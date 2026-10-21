import { renderAsync } from 'docx-preview';
import JSZip from 'jszip';
import { el, sanitizeLinks } from '../sanitize.js';
import { clonePagesForPrint, PAGE_WIDTH } from '../print.js';
import { bulletGlyph } from '../bullets.js';

// Bullets in Word are characters from the Windows "Symbol" / "Wingdings" fonts
// (private-use codes such as U+F0B7) followed by a tab. Phones have neither font,
// so they show a different glyph and spacing. Map them to normal Unicode bullets
// and turn the tab into a fixed hanging indent, so lists look the same everywhere.
const unescapeCss = (s) => s.replace(/\\([0-9a-fA-F]{1,6})\s?/g, (_, h) => String.fromCodePoint(parseInt(h, 16)));

function fixListMarkers(styleRoot) {
  for (const style of styleRoot.querySelectorAll('style')) {
    let css = style.textContent;
    if (!css.includes('docx-num-')) continue;
    const hanging = {}; // list class -> hanging indent, from "p.docx-num-1-0 { text-indent: -18pt }"
    for (const m of css.matchAll(/p\.(docx-num-[\w-]+)\s*\{[^}]*?text-indent:\s*-([\d.]+)(pt|px)/g)) hanging[m[1]] = `${m[2]}${m[3]}`;
    css = css.replace(/(p\.(docx-num-[\w-]+):before\s*\{)([^}]*)\}/g, (all, open, cls, body) => {
      const font = /font-family:\s*([^;]+)/.exec(body)?.[1] || '';
      const symbolFont = /wingdings|webdings|symbol/i.test(font);
      let mapped = false;
      let next = body.replace(/content:\s*([^;]+)/, (_, value) => `content: ${value.replace(/"((?:[^"\\]|\\.)*)"/g, (q, str) => {
        const raw = unescapeCss(str).replace(/\t/g, ''); // the tab becomes the fixed indent below
        const text = raw ? bulletGlyph(raw, symbolFont ? font : '') : '';
        if (text !== raw || symbolFont) mapped = true;
        return JSON.stringify(text);
      })}`);
      if (mapped) next = next.replace(/font-family:\s*[^;]+;?/, 'font-family: inherit;');
      const width = hanging[cls];
      next += `; display: inline-block; text-indent: 0; ${width ? `min-width: ${width};` : 'padding-right: 0.5em;'}`;
      return `${open}${next}}`;
    });
    style.textContent = css;
  }
}

/** Paragraph style ids with Word's "Don't add space between paragraphs of the same style". */
async function contextualStyles(buffer) {
  const xml = await (await JSZip.loadAsync(buffer)).file('word/styles.xml')?.async('string');
  if (!xml) return new Set();
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const styles = new Map();
  for (const st of doc.getElementsByTagNameNS(W, 'style')) {
    const id = st.getAttributeNS(W, 'styleId');
    const cs = st.getElementsByTagNameNS(W, 'contextualSpacing')[0];
    const val = cs?.getAttributeNS(W, 'val');
    styles.set(id, {
      own: cs ? !(val === '0' || val === 'false') : undefined,
      basedOn: st.getElementsByTagNameNS(W, 'basedOn')[0]?.getAttributeNS(W, 'val'),
    });
  }
  const on = (id, depth = 0) => {
    const st = styles.get(id);
    if (!st || depth > 20) return false;
    return st.own ?? (st.basedOn ? on(st.basedOn, depth + 1) : false);
  };
  return new Set([...styles.keys()].filter((id) => on(id)));
}

const styleClass = (id) => `docx_${id.replace(/[ .]+/g, '-').replace(/[&]+/g, 'and').toLowerCase()}`;

// Layout width in CSS px, unaffected by the viewer's zoom.
const widthOf = (node) => parseFloat(getComputedStyle(node).width) || node.offsetWidth || 1;

export async function renderDocx({ buffer, mount }) {
  const styles = el('div');
  const body = el('div', 'dv-docx');
  mount.append(styles, body);

  const contextual = contextualStyles(buffer).catch(() => new Set());
  const doc = await renderAsync(buffer, body, styles, {
    className: 'docx',
    inWrapper: true,
    breakPages: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
    renderEndnotes: true,
    renderComments: false,
    renderChanges: false,
    useBase64URL: false,
    experimental: false,
  });
  sanitizeLinks(body);
  fixListMarkers(styles);
  // docx-preview always hyphenates; Word only does when the document turns it on.
  if (doc?.settingsPart?.settings?.autoHyphenation) body.classList.add('dv-hyphenate');
  // Contextual spacing: no space between consecutive paragraphs of the same such style.
  for (const id of await contextual) {
    for (const p of body.querySelectorAll(`p.${CSS.escape(styleClass(id))}`)) {
      const prev = p.previousElementSibling;
      if (prev?.localName === 'p' && prev.classList.contains(styleClass(id))) {
        prev.style.marginBottom = '0px';
        p.style.marginTop = '0px';
      }
    }
  }
  // Word's thin (0.5 pt) table borders shrink below one screen pixel when the page is
  // zoomed out to fit a phone, and Safari then drops them: keep them at least 1 screen px.
  for (const cell of body.querySelectorAll('td, th')) {
    const cs = getComputedStyle(cell);
    for (const side of ['Top', 'Right', 'Bottom', 'Left']) {
      const width = parseFloat(cs[`border${side}Width`]);
      if (cs[`border${side}Style`] !== 'none' && width > 0 && width <= 1.5) {
        cell.style[`border${side}Width`] = `max(${width}px, calc(1px / var(--dv-zoom, 1)))`;
      }
    }
  }

  const pages = () => [...body.querySelectorAll('section.docx')];

  return {
    pages,
    pageWidth: widthOf,
    naturalWidth: () => Math.max(...pages().map(widthOf), 1),
    printNodes: () => ({ nodes: clonePagesForPrint(pages(), PAGE_WIDTH, widthOf, 'docx-wrapper') }),
    destroy() {},
  };
}
