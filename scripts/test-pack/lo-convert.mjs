// LibreOffice (WebAssembly, in Node) conversions used by the test pack: docx -> doc, xlsx -> xls, pptx -> ppt.
// usage: node lo-convert.mjs out/test-word-for-doc.docx=doc ...
import { readFileSync, writeFileSync } from 'node:fs';
import { basename } from 'node:path';
import { fileURLToPath } from 'node:url';
const WASM = fileURLToPath(new URL('../../node_modules/@matbee/libreoffice-converter/wasm', import.meta.url));
const { convertDocument } = await import('@matbee/libreoffice-converter');

const timeout = (ms) => new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), ms));
for (const [file, format] of process.argv.slice(2).map((a) => a.split('='))) {
  const t0 = Date.now();
  try {
    const out = await Promise.race([convertDocument(readFileSync(file), { outputFormat: format, inputFormat: file.split('.').pop() }, { wasmPath: WASM }), timeout(120000)]);
    const data = out.data ?? out;
    const dest = format === 'pdf' ? `${file}.pdf` : file.replace(/\.[^.]+$/, `.${format}`);
    writeFileSync(dest, data);
    console.log('ok', basename(dest), data.length, 'bytes', Date.now() - t0, 'ms');
  } catch (err) {
    console.log('FAILED', file, format, err.message);
  }
}
process.exit(0);
