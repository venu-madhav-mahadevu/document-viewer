/*
 * PowerPoint's "Shrink text on overflow": when a text box's content is taller than the
 * box, PowerPoint scales the text down until it fits. Boxes marked with `data-fit` get
 * the same treatment here (font size and line spacing shrink together).
 *   data-fit="box" (or "1"): shrink whenever the text is taller than its box
 *   data-fit="slide":        shrink only when the text would run off the bottom of the slide
 * Call after the slide is in the page, so sizes can be measured.
 *
 * Uses a CSS transform (same result in Chrome, Safari and Firefox): the text is laid out
 * 1/scale times wider, then scaled down, so lines break where a smaller font would break.
 */
export function shrinkToFit(root) {
  for (const box of root.querySelectorAll('[data-fit]')) {
    if (!box.clientHeight || box.scrollHeight <= box.clientHeight + 1) continue;
    if (box.dataset.fit === 'slide') {
      const slide = box.closest('.dv-page');
      const shape = box.offsetParent;
      if (!slide || !shape || shape.offsetTop + box.scrollHeight <= slide.clientHeight) continue;
    }
    const cs = getComputedStyle(box);
    const avail = box.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
    const inner = document.createElement('div');
    inner.className = 'dv-fit';
    inner.append(...box.childNodes);
    box.append(inner);

    const fits = (scale) => {
      inner.style.width = `${100 / scale}%`;
      return inner.offsetHeight * scale <= avail + 1;
    };
    let lo = 0.25;
    let hi = 1;
    for (let i = 0; i < 10; i++) {
      const mid = (lo + hi) / 2;
      if (fits(mid)) lo = mid;
      else hi = mid;
    }
    fits(lo);
    const height = inner.offsetHeight;
    inner.style.transform = `scale(${lo})`;
    inner.style.transformOrigin = 'top left';
    inner.style.marginBottom = `${-height * (1 - lo)}px`; // the box sees the scaled height (keeps centering)
  }
}
