import { readdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { defineConfig } from 'vite';

// The app plus every demos/*.html page.
const demoPages = Object.fromEntries(
  readdirSync('demos').filter((f) => f.endsWith('.html')).map((f) => [`demos/${f.replace('.html', '')}`, resolve('demos', f)]),
);
demoPages['demos/convert/index'] = resolve('demos', 'convert', 'index.html'); // in-browser LibreOffice demo

const isolation = { 'Cross-Origin-Opener-Policy': 'same-origin', 'Cross-Origin-Embedder-Policy': 'require-corp' };

export default defineConfig({
  // Relative base so the built app can be hosted under any path on an internal server.
  base: './',
  build: {
    target: 'es2022',
    chunkSizeWarningLimit: 2000,
    rollupOptions: { input: { main: resolve('index.html'), ...demoPages } },
  },
  // Cross-origin isolation (needed by the LibreOffice WebAssembly engine in
  // demos/convert/). Production hosts send the same two headers;
  // GitHub Pages cannot, so that page also registers coi-serviceworker.
  server: { port: Number(process.env.PORT) || 5180, strictPort: !!process.env.PORT, headers: isolation },
  preview: { headers: isolation },
});
