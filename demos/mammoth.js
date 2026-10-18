import { setup } from './shared.js';
import source from './mammoth.js?raw';

// #region demo
import mammoth from 'mammoth';
import DOMPurify from 'dompurify';

async function render(buffer, name, out) {
  // mammoth produces clean, semantic HTML (headings, lists, tables),
  // ignoring exact page layout - good for search, reading, accessibility.
  const { value: html, messages } = await mammoth.convertToHtml({ arrayBuffer: buffer });

  const doc = document.createElement('article');
  doc.className = 'doc-html';
  doc.innerHTML = DOMPurify.sanitize(html); // always sanitize before innerHTML
  out.append(doc);

  if (messages.length) console.info('mammoth notes:', messages);
}
// #endregion

setup({ sample: 'onboarding-guide.docx', accept: '.docx', render, source });
