import { renderAsync } from 'docx-preview';
import { el, sanitizeLinks } from '../sanitize.js';

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

  return {
    naturalWidth: () => Math.max(...[...body.querySelectorAll('section.docx')].map((s) => s.offsetWidth), 1),
    destroy() {},
  };
}
