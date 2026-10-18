const SAFE_PROTOCOLS = new Set(['http:', 'https:', 'mailto:']);

/**
 * Documents can carry hyperlinks with arbitrary targets (including `javascript:`).
 * Neutralise anything that isn't a plain web/mail link or an in-document anchor,
 * and make external links open in a new tab without referrer/opener access.
 */
export function sanitizeLinks(root) {
  for (const a of root.querySelectorAll('a[href]')) {
    const href = a.getAttribute('href').trim();
    if (href.startsWith('#')) continue;
    let ok = false;
    try { ok = SAFE_PROTOCOLS.has(new URL(href, 'https://invalid.local').protocol); } catch { /* bad URL */ }
    if (!ok) { a.removeAttribute('href'); continue; }
    a.target = '_blank';
    a.rel = 'noopener noreferrer';
  }
  // Belt and braces: no document format we render should produce these.
  root.querySelectorAll('script, iframe, object, embed, form').forEach((el) => el.remove());
  root.querySelectorAll('*').forEach((el) => {
    for (const attr of [...el.attributes]) if (/^on/i.test(attr.name)) el.removeAttribute(attr.name);
  });
}

export function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text != null) node.textContent = text;
  return node;
}
