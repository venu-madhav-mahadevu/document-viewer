import { el } from './sanitize.js';

// Printable widths in CSS px at 96 dpi (Letter portrait / landscape, no margin).
export const PAGE_WIDTH = 816;
export const PAGE_WIDTH_LANDSCAPE = 1056;

/** Clone HTML pages (Word sections, slides) scaled to fill a printed page. */
export function clonePagesForPrint(pageEls, targetWidth, naturalWidth, wrapperClass) {
  return pageEls.map((page) => {
    const clone = page.cloneNode(true);
    clone.style.zoom = String(targetWidth / naturalWidth(page));
    clone.style.margin = '0';
    clone.style.boxShadow = 'none';
    if (!wrapperClass) return clone;
    const wrap = el('div', wrapperClass);
    wrap.append(clone);
    return wrap;
  });
}

/**
 * Print only the document (not the surrounding app) by moving a printable
 * copy to the top of <body> and hiding everything else via print CSS.
 *
 * @param {{ title: string, nodes: Node[], paged: boolean, landscape?: boolean }} job
 */
export async function printDocument({ title, nodes, paged, landscape = false }) {
  const box = el('div', paged ? 'dv-print' : 'dv-print dv-print-flow');
  for (const node of nodes) {
    const page = el('div', paged ? 'dv-print-page' : 'dv-print-block');
    page.append(node);
    box.append(page);
  }
  const pageRule = el('style');
  pageRule.textContent = `@page { size: ${landscape ? 'landscape' : 'auto'}; margin: ${paged ? '0' : '0.5in'}; }`;
  document.head.append(pageRule);
  document.body.append(box);

  // Wait for images (PDF page snapshots) so they are not blank in the print.
  await Promise.all([...box.querySelectorAll('img')].map((img) =>
    img.complete ? null : new Promise((r) => { img.onload = img.onerror = r; })));

  const previousTitle = document.title;
  document.title = title; // becomes the default file name for "Save as PDF"

  const cleanup = () => {
    box.remove();
    pageRule.remove();
    document.title = previousTitle;
    window.removeEventListener('afterprint', cleanup);
  };
  window.addEventListener('afterprint', cleanup);
  window.print();
  // Some browsers don't fire afterprint reliably; print() blocks in the rest.
  setTimeout(() => { if (box.isConnected) cleanup(); }, 1000);
}
