// Minimal document-conversion server. No npm dependencies on the server side
// (Node 20+); pdfjs-dist is installed only so the web page can serve it locally.
//
//   POST /api/convert   body = file bytes, header X-File-Name = original name
//                       -> returns application/pdf
//   GET  /              the web demo page
//   GET  /pdfjs/*       pdf.js for the web page
//   GET  /samples/*     demo files (from ../public/samples)
//
// Office files are converted by Gotenberg (LibreOffice) running on BofA's own
// infrastructure. PDFs are passed straight through. Nothing goes to a third party.

import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname, join, basename } from 'node:path';
import { fileURLToPath } from 'node:url';

const PORT = Number(process.env.PORT) || 8080;
const GOTENBERG_URL = process.env.GOTENBERG_URL || 'http://localhost:3000';
const MAX_BYTES = 50 * 1024 * 1024;
const ROOT = fileURLToPath(new URL('.', import.meta.url));
const SAMPLES = process.env.SAMPLES_DIR || join(ROOT, '..', 'public', 'samples');
const PDFJS = join(ROOT, 'node_modules', 'pdfjs-dist');

// Formats LibreOffice converts well. Anything else is rejected up front.
const CONVERTIBLE = new Set([
  '.doc', '.docx', '.docm', '.dotx', '.odt', '.rtf', '.txt',
  '.xls', '.xlsx', '.xlsm', '.xlsb', '.ods', '.csv',
  '.ppt', '.pptx', '.pptm', '.ppsx', '.potx', '.odp',
  '.png', '.jpg', '.jpeg', '.gif', '.bmp',
]);

const SECURITY_HEADERS = {
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  'Cache-Control': 'no-store', // never cache customer documents
};

function send(res, status, body, headers = {}) {
  res.writeHead(status, { ...SECURITY_HEADERS, ...headers });
  res.end(body);
}

const sendError = (res, status, message) =>
  send(res, status, JSON.stringify({ error: message }), { 'Content-Type': 'application/json' });

async function readBody(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > MAX_BYTES) throw Object.assign(new Error('File is larger than 50 MB.'), { status: 413 });
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

async function convert(req, res) {
  const name = basename(decodeURIComponent(req.headers['x-file-name'] || ''));
  const ext = extname(name).toLowerCase();
  if (!name) return sendError(res, 400, 'Missing X-File-Name header.');

  const file = await readBody(req);
  if (!file.length) return sendError(res, 400, 'Empty file.');

  const pdfHeaders = {
    'Content-Type': 'application/pdf',
    'Content-Disposition': `inline; filename="${name.replace(/[^\w.-]/g, '_')}.pdf"`,
  };

  // Already a PDF: no conversion needed.
  if (ext === '.pdf') {
    if (file.subarray(0, 4).toString() !== '%PDF') return sendError(res, 400, 'Not a valid PDF.');
    return send(res, 200, file, pdfHeaders);
  }
  if (!CONVERTIBLE.has(ext)) return sendError(res, 415, `Unsupported file type: ${ext || 'none'}`);

  // Gotenberg LibreOffice route: multipart form with the file in "files".
  const form = new FormData();
  form.append('files', new Blob([file]), name);

  const started = Date.now();
  let upstream;
  try {
    upstream = await fetch(`${GOTENBERG_URL}/forms/libreoffice/convert`, {
      method: 'POST',
      body: form,
      signal: AbortSignal.timeout(90_000),
    });
  } catch (err) {
    console.error('Gotenberg unreachable:', err.message);
    return sendError(res, 503, 'Conversion service is not running. Start it with: docker compose up -d');
  }
  if (!upstream.ok) {
    console.error('Gotenberg error', upstream.status, await upstream.text());
    return sendError(res, 422, 'This file could not be converted.');
  }

  const pdf = Buffer.from(await upstream.arrayBuffer());
  console.log(`converted ${ext} (${file.length} B) -> PDF (${pdf.length} B) in ${Date.now() - started} ms`);
  send(res, 200, pdf, { ...pdfHeaders, 'X-Conversion-Ms': String(Date.now() - started) });
}

const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css' };

async function serveFile(res, path, contentType) {
  try {
    send(res, 200, await readFile(path), { 'Content-Type': contentType });
  } catch {
    sendError(res, 404, 'Not found');
  }
}

const server = createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost');
  try {
    if (req.method === 'POST' && url.pathname === '/api/convert') return await convert(req, res);

    if (req.method === 'GET' && url.pathname === '/api/health') {
      const ok = await fetch(`${GOTENBERG_URL}/health`, { signal: AbortSignal.timeout(3000) }).then((r) => r.ok, () => false);
      return send(res, ok ? 200 : 503, JSON.stringify({ server: 'ok', gotenberg: ok ? 'ok' : 'down' }), { 'Content-Type': 'application/json' });
    }

    // pdf.js served from our own origin (no CDN).
    if (req.method === 'GET' && url.pathname.startsWith('/pdfjs/')) {
      const rel = url.pathname.slice('/pdfjs/'.length);
      if (rel === 'pdf.min.mjs' || rel === 'pdf.worker.min.mjs') {
        return await serveFile(res, join(PDFJS, 'build', rel), 'text/javascript');
      }
      if (/^standard_fonts\/[\w.-]+$/.test(rel)) {
        return await serveFile(res, join(PDFJS, rel), 'application/octet-stream');
      }
      return sendError(res, 404, 'Not found');
    }

    if (req.method === 'GET' && url.pathname.startsWith('/samples/')) {
      const file = basename(decodeURIComponent(url.pathname)); // basename blocks ../ traversal
      return await serveFile(res, join(SAMPLES, file), 'application/octet-stream');
    }

    if (req.method === 'GET') {
      const file = url.pathname === '/' ? 'index.html' : basename(url.pathname);
      const type = MIME[extname(file)];
      if (!type) return sendError(res, 404, 'Not found');
      return await serveFile(res, join(ROOT, 'public', file), type);
    }

    sendError(res, 405, 'Method not allowed');
  } catch (err) {
    console.error(err);
    sendError(res, err.status || 500, err.status ? err.message : 'Server error');
  }
});

server.listen(PORT, () => {
  console.log(`Demo:      http://localhost:${PORT}`);
  console.log(`Gotenberg: ${GOTENBERG_URL}`);
});
