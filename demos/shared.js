import './sw-cleanup.js';
// Shared plumbing for the demo pages: file picker, sample loader and a
// "show the code" panel. The interesting code lives in each demo file between
// `// #region demo` and `// #endregion`.
import './demo.css';
import hljs from 'highlight.js/lib/core';
import javascript from 'highlight.js/lib/languages/javascript';
import 'highlight.js/styles/github.css';

hljs.registerLanguage('javascript', javascript);

/**
 * @param {{ sample: string, accept?: string, render: (buffer: ArrayBuffer, name: string, out: HTMLElement) => Promise<void>, source: string }} demo
 */
export function setup({ sample, accept = '', render, source }) {
  const out = document.getElementById('output');
  const status = document.getElementById('status');
  const input = document.getElementById('file');
  if (accept) input.accept = accept;

  async function run(buffer, name) {
    out.replaceChildren();
    status.textContent = `Rendering ${name}…`;
    const t0 = performance.now();
    try {
      await render(buffer, name, out);
      status.textContent = `${name} rendered in ${Math.round(performance.now() - t0)} ms`;
    } catch (err) {
      console.error(err);
      status.textContent = `Error: ${err.message}`;
    }
  }

  input.addEventListener('change', async () => {
    const file = input.files[0];
    if (file) run(await file.arrayBuffer(), file.name);
  });
  document.getElementById('load-sample').addEventListener('click', loadSample);

  async function loadSample() {
    const res = await fetch(new URL(`../samples/${sample}`, location.href));
    run(await res.arrayBuffer(), sample);
  }

  const m = /\/\/ #region demo\r?\n([\s\S]*?)\/\/ #endregion/.exec(source);
  const code = document.getElementById('code');
  code.textContent = (m ? m[1] : source).trim();
  hljs.highlightElement(code);

  loadSample();
}
