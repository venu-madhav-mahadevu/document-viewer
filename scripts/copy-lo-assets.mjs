// Copies the LibreOffice WebAssembly engine (@matbee/libreoffice-converter) into
// public/demos/lo/ (inside the coi-serviceworker scope) so the browser-convert demo serves it from our own origin.
// About 250 MB - kept out of git (.gitignore); copied again only when a file changed.
import { copyFileSync, existsSync, mkdirSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const pkg = join(root, 'node_modules', '@matbee', 'libreoffice-converter');
const dest = join(root, 'public', 'demos', 'lo');

const files = [
  ['wasm/soffice.js', 'soffice.js'],
  ['wasm/soffice.wasm', 'soffice.wasm'],
  ['wasm/soffice.data', 'soffice.data'],
  ['wasm/soffice.worker.js', 'soffice.worker.js'],
  ['dist/browser.worker.global.js', 'browser.worker.global.js'],
];

if (!existsSync(pkg)) {
  console.log('LibreOffice engine not installed; skipping (browser-convert demo will not work)');
} else {
  mkdirSync(dest, { recursive: true });
  for (const [from, to] of files) {
    const src = join(pkg, from);
    const out = join(dest, to);
    if (existsSync(out) && statSync(out).size === statSync(src).size) continue;
    copyFileSync(src, out);
  }
  // The demo page registers this to get cross-origin isolation on hosts that
  // cannot send COOP/COEP headers (GitHub Pages).
  mkdirSync(join(root, 'public', 'demos'), { recursive: true });
  copyFileSync(join(root, 'node_modules', 'coi-serviceworker', 'coi-serviceworker.min.js'), join(root, 'public', 'demos', 'coi-serviceworker.js'));
  console.log('LibreOffice engine copied to public/demos/lo');
}
