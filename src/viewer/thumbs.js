import { el } from './sanitize.js';

const THUMB_WIDTH = 120;

/**
 * Page thumbnail strip. Thumbnails are drawn lazily as they scroll into the
 * panel, so a 500-page PDF doesn't render 500 previews up front.
 *
 * @param {HTMLElement} panel
 * @param {{ count: number, ratio: (n: number) => number, render: (n: number, width: number) => Promise<Node>|Node, onPick: (n: number) => void }} opts
 */
export function createThumbnails(panel, { count, ratio, render, onPick }) {
  panel.replaceChildren();
  const items = [];

  const observer = new IntersectionObserver(async (entries) => {
    for (const e of entries) {
      if (!e.isIntersecting || e.target.dataset.drawn) continue;
      e.target.dataset.drawn = '1';
      observer.unobserve(e.target);
      try {
        const node = await render(Number(e.target.dataset.page), THUMB_WIDTH);
        e.target.querySelector('.dv-thumb-frame').replaceChildren(node);
      } catch (err) {
        console.warn('thumbnail failed', err);
      }
    }
  }, { root: panel, rootMargin: '300px 0px' });

  for (let n = 1; n <= count; n++) {
    const button = el('button', 'dv-thumb');
    button.type = 'button';
    button.dataset.page = String(n);
    button.setAttribute('aria-label', `Go to page ${n}`);
    const frame = el('div', 'dv-thumb-frame');
    frame.style.width = `${THUMB_WIDTH}px`;
    frame.style.height = `${Math.round(THUMB_WIDTH * ratio(n))}px`;
    button.append(frame, el('span', 'dv-thumb-num', String(n)));
    button.addEventListener('click', () => onPick(n));
    panel.append(button);
    observer.observe(button);
    items.push(button);
  }

  let active = 0;
  return {
    setActive(n) {
      if (n === active) return;
      items[active - 1]?.removeAttribute('aria-current');
      active = n;
      const item = items[n - 1];
      if (!item) return;
      item.setAttribute('aria-current', 'page');
      // Keep the active thumbnail visible without yanking the main view.
      const p = panel.getBoundingClientRect(), r = item.getBoundingClientRect();
      if (r.top < p.top || r.bottom > p.bottom) panel.scrollTop += r.top - p.top - p.height / 3;
    },
    destroy() {
      observer.disconnect();
      panel.replaceChildren();
    },
  };
}

/** Thumbnail for an HTML page (Word section, slide): a scaled, inert clone. */
export function cloneThumbnail(pageEl, width, naturalWidth) {
  const clone = pageEl.cloneNode(true);
  clone.removeAttribute('id');
  clone.querySelectorAll('[id]').forEach((n) => n.removeAttribute('id'));
  clone.style.zoom = String(width / naturalWidth);
  clone.style.margin = '0';
  clone.style.boxShadow = 'none';
  const wrap = el('div', 'dv-thumb-clone');
  wrap.setAttribute('aria-hidden', 'true');
  wrap.inert = true;
  wrap.append(clone);
  return wrap;
}
