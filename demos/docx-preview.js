import { setup } from './shared.js';
import source from './docx-preview.js?raw';

// #region demo
import { renderAsync } from 'docx-preview';

async function render(buffer, name, out) {
  // Converts the .docx XML into paged HTML inside `out`.
  await renderAsync(buffer, out, null, {
    inWrapper: true,        // grey background with page "sheets"
    breakPages: true,       // honour page breaks
    renderHeaders: true,
    renderFooters: true,
  });
}
// #endregion

setup({ sample: 'onboarding-guide.docx', accept: '.docx,.dotx,.docm', render, source });
