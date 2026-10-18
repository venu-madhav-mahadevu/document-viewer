import { renderAsync } from 'docx-preview';
import { el, sanitizeLinks } from '../sanitize.js';
import { clonePagesForPrint, PAGE_WIDTH } from '../print.js';

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

  const pages = () => [...body.querySelectorAll('section.docx')];

  return {
    pages,
    pageWidth: widthOf,
    naturalWidth: () => Math.max(...pages().map(widthOf), 1),
    printNodes: () => ({ nodes: clonePagesForPrint(pages(), PAGE_WIDTH, widthOf, 'docx-wrapper') }),
    destroy() {},
  };
}
