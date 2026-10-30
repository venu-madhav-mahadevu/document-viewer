/*
 * Office bullets are characters from the Windows "Symbol" and "Wingdings" fonts,
 * often stored as private-use codes (U+F0xx). Phones and most browsers don't have
 * those fonts, so the raw character shows instead ("§", "Ø", "ü", or an empty box).
 * bulletGlyph() turns them into normal Unicode bullets that every device can draw.
 * Used by the .doc, .docx/.docm, .ppt and .pptx renderers.
 */

const WINGDINGS = {
  0x6c: '●', 0x6e: '■', 0x6f: '□', 0x71: '❑', 0x72: '❒', 0x73: '◆', 0x74: '⧫', 0x75: '◆', 0x76: '❖',
  0x77: '⬥', 0x9f: '•', 0xa7: '▪', 0xa8: '◻', 0xd8: '➢', 0xe0: '➔', 0xe8: '➔', 0xf0: '⇨', 0xfb: '✗', 0xfc: '✓', 0xfe: '☑',
};
const SYMBOL = { 0xb7: '•', 0xa8: '♦', 0xae: '→', 0xde: '⇒', 0x6f: 'o', 0x2d: '–' };

// Characters that are almost always a Wingdings bullet when they appear alone as a bullet,
// even if the file doesn't say which font (e.g. "§" is Wingdings' small square).
const LIKELY_WINGDINGS = new Set([0xa7, 0xd8, 0xfc, 0x71, 0x76, 0x6e, 0x6c, 0xa8, 0x9f]);

/**
 * @param {string} text   the bullet text (usually one character)
 * @param {string} [font] the bullet's font name, when known
 */
export function bulletGlyph(text, font = '') {
  if (!text) return '•';
  const wingdings = /wingdings|webdings/i.test(font);
  const symbol = /symbol/i.test(font);
  const single = [...text].length === 1;
  return [...text].map((ch) => {
    let code = ch.codePointAt(0);
    const privateUse = code >= 0xf000 && code <= 0xf0ff;
    if (privateUse) code -= 0xf000;
    if (wingdings) return WINGDINGS[code] ?? '•';
    if (symbol) return SYMBOL[code] ?? (code >= 0x20 && code < 0x7f ? String.fromCharCode(code) : '•');
    if (privateUse) return WINGDINGS[code] ?? SYMBOL[code] ?? '•';
    if (single && !font && LIKELY_WINGDINGS.has(code) && code > 0x7f) return WINGDINGS[code];
    if (single && code === 0xa7) return '▪'; // "§" is never meant as a bullet
    return ch;
  }).join('');
}

export function bulletCode(code, font = '') {
  return bulletGlyph(String.fromCodePoint(code), font);
}
