import { defineConfig } from 'vite';

export default defineConfig({
  // Relative base so the built app can be hosted under any path on an internal server.
  base: './',
  build: { target: 'es2022', chunkSizeWarningLimit: 2000 },
  server: { port: Number(process.env.PORT) || 5180, strictPort: !!process.env.PORT },
});
