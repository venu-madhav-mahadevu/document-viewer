// In-document search using the CSS Custom Highlight API: matches are painted
// without modifying the document's DOM, so renderers are never disturbed.

const ALL = 'dv-search';
const CURRENT = 'dv-search-current';
const supported = typeof CSS !== 'undefined' && 'highlights' in CSS;

/** Case-insensitive matches inside individual text nodes under `root`. */
export function findRanges(root, term) {
  const needle = term.toLowerCase();
  const ranges = [];
  if (!needle || !root) return ranges;
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
    acceptNode: (node) => (node.parentElement?.closest('style, script') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT),
  });
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.data.toLowerCase();
    for (let i = text.indexOf(needle); i !== -1; i = text.indexOf(needle, i + needle.length)) {
      const range = document.createRange();
      range.setStart(node, i);
      range.setEnd(node, i + needle.length);
      ranges.push(range);
    }
  }
  return ranges;
}

export function paintHighlights(ranges, current) {
  if (!supported) return;
  CSS.highlights.set(ALL, new Highlight(...ranges));
  if (current) CSS.highlights.set(CURRENT, new Highlight(current));
  else CSS.highlights.delete(CURRENT);
}

export function clearHighlights() {
  if (!supported) return;
  CSS.highlights.delete(ALL);
  CSS.highlights.delete(CURRENT);
}

/** Scroll a match into the middle of whatever scroll container holds it. */
export function revealRange(range) {
  const target = range.startContainer.parentElement;
  target?.scrollIntoView({ block: 'center', inline: 'nearest' });
}
