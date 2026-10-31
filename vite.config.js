import { resolve } from 'node:path';
import { defineConfig } from 'vite';

// Library build: dist/docs-viewer.js (ES module) + dist/docs-viewer.css, the format renderers
// as chunks loaded on demand, and static/ (fonts/, pdfjs/) copied next to them.
export default defineConfig({
  base: './',
  publicDir: 'static',
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 2000,
    lib: { entry: resolve('src/viewer/index.js'), formats: ['es'], fileName: () => 'docs-viewer.js', cssFileName: 'docs-viewer' },
  },
  worker: { format: 'es' },
  server: { port: Number(process.env.PORT) || 5185 },
});
