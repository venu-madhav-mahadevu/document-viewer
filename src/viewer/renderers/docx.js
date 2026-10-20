import { renderAsync } from 'docx-preview';
import { el, sanitizeLinks } from '../sanitize.js';
import { clonePagesForPrint, PAGE_WIDTH } from '../print.js';

// Bullets in Word are characters from the Windows "Symbol" / "Wingdings" fonts
// (private-use codes such as U+F0B7) followed by a tab. Phones have neither font,
// so they show a different glyph and spacing. Map them to normal Unicode bullets
// and turn the tab into a fixed hanging indent, so lists look the same everywhere.
const SYMBOL_BULLETS = { 0xb7: '•', 0xa8: '♦', 0x6f: 'o', 0xde: '⇒', 0xae: '→' };
const WINGDINGS_BULLETS = { 0xa7: '▪', 0x6c: '●', 0x6e: '■', 0x71: '❑', 0x75: '◆', 0x76: '❖', 0xd8: '➢', 0xfc: '✓', 0xa8: '◻', 0x9f: '•', 0x77: '⬥' };
const unescapeCss = (s) => s.replace(/\\([0-9a-fA-F]{1,6})\s?/g, (_, h) => String.fromCodePoint(parseInt(h, 16)));

function fixListMarkers(styleRoot) {
  for (const style of styleRoot.querySelectorAll('style')) {
    let css = style.textContent;
    if (!css.includes('docx-num-')) continue;
    const hanging = {}; // list class -> hanging indent, from "p.docx-num-1-0 { text-indent: -18pt }"
    for (const m of css.matchAll(/p\.(docx-num-[\w-]+)\s*\{[^}]*?text-indent:\s*-([\d.]+)(pt|px)/g)) hanging[m[1]] = `${m[2]}${m[3]}`;
    css = css.replace(/(p\.(docx-num-[\w-]+):before\s*\{)([^}]*)\}/g, (all, open, cls, body) => {
      const font = /font-family:\s*([^;]+)/.exec(body)?.[1] || '';
      const map = /wingdings/i.test(font) ? WINGDINGS_BULLETS : /symbol/i.test(font) ? SYMBOL_BULLETS : null;
      let mapped = false;
      let next = body.replace(/content:\s*([^;]+)/, (_, value) => `content: ${value.replace(/"((?:[^"\\]|\\.)*)"/g, (q, str) => {
        const text = [...unescapeCss(str)].map((ch) => {
          let code = ch.codePointAt(0);
          if (code === 9) return '';          // the tab becomes the fixed indent below
          if (code >= 0xf000 && code <= 0xf0ff) code -= 0xf000;
          else if (!map) return ch;
          mapped = true;
          return map?.[code] ?? (code < 0x20 || code > 0x7e ? '•' : String.fromCharCode(code));
        }).join('');
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

// Layout width in CSS px, unaffected by the viewer's zoom.
const widthOf = (node) => parseFloat(getComputedStyle(node).width) || node.offsetWidth || 1;

export async function renderDocx({ buffer, mount }) {
  const styles = el('div');
  const body = el('div', 'dv-docx');
  mount.append(styles, body);

  await renderAsync(buffer, body, styles, {
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
