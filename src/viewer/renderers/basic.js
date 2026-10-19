import { el } from '../sanitize.js';

const IMAGE_MIME = { jpg: 'image/jpeg', jpeg: 'image/jpeg' };

/** JPEG images are shown via <img>, decoded natively by the browser. */
export async function renderImage({ buffer, name, mount }) {
  const ext = (name.split('.').pop() || '').toLowerCase();
  const url = URL.createObjectURL(new Blob([buffer], { type: IMAGE_MIME[ext] || '' }));
  const frame = el('div', 'dv-page dv-image');
  const img = el('img');
  img.alt = name;
  // load/error rather than img.decode(): decode() never settles in a hidden tab.
  await new Promise((resolve, reject) => {
    img.onload = resolve;
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('This image could not be decoded.')); };
    img.src = url;
  });
  frame.append(img);
  mount.append(frame);
  return {
    naturalWidth: () => img.naturalWidth,
    destroy() { URL.revokeObjectURL(url); },
  };
}

const MAX_TEXT = 5 * 1024 * 1024;

export async function renderText({ buffer, mount }) {
  const truncated = buffer.byteLength > MAX_TEXT;
  let text = new TextDecoder().decode(truncated ? buffer.slice(0, MAX_TEXT) : buffer);
  if (/^\s*[[{]/.test(text) && !truncated) {
    try { text = JSON.stringify(JSON.parse(text), null, 2); } catch { /* not JSON */ }
  }
  const page = el('div', 'dv-page dv-text');
  page.append(el('pre', null, text));
  if (truncated) page.append(el('p', 'dv-note', 'Showing the first 5 MB. Download the file to see everything.'));
  mount.append(page);
  return { destroy() {} };
}
