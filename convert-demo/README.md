# Convert-to-PDF demo (recommended approach)

Every document is converted to PDF on **our own server**, then shown with a standard PDF viewer on
each platform. No Google, Microsoft or other third-party service is involved.

```
Word / Excel / PowerPoint / .doc / .ppt / .rtf / images
                │  POST /api/convert   (server.js)
                ▼
     Gotenberg (LibreOffice inside)   ← Docker container, internal only
                │  PDF
                ▼
   Web: pdf.js     iOS: PDFKit     Android: PdfRenderer
```

## Run it

Requires Docker Desktop and Node.js 20+.

```bash
cd convert-demo
npm install              # pdf.js for the web page
docker compose up -d     # starts Gotenberg + the demo server
```

Open **http://localhost:8080** and click a sample, or open your own file.

Without Docker for the app (Gotenberg still in Docker):

```bash
docker run -d --name gotenberg -p 127.0.0.1:3000:3000 gotenberg/gotenberg:8
npm start                # http://localhost:8080
```

The first `docker compose up` downloads the Gotenberg image, which is large (around 1.5 GB)
because it contains LibreOffice.

## Files

| File | What it is |
|---|---|
| `server.js` | About 150 lines of Node.js with no server dependencies. `POST /api/convert` returns a PDF |
| `public/index.html` | Web page: uploads the file and draws the returned PDF with pdf.js |
| `docker-compose.yml` | Gotenberg + the demo server. Gotenberg is not exposed outside the compose network |
| `mobile/ios/DocumentViewer.swift` | SwiftUI + PDFKit client for the same API |
| `mobile/android/DocumentViewerActivity.kt` | Kotlin + PdfRenderer client for the same API |

## The API (same for web, iOS and Android)

```
POST /api/convert
X-File-Name: report.docx        (URL-encoded original file name)
Body: raw file bytes

200 application/pdf             converted PDF (PDFs are passed through unchanged)
4xx/5xx application/json        { "error": "..." }
```

```js
// Web
const res = await fetch('/api/convert', { method: 'POST', headers: { 'X-File-Name': encodeURIComponent(file.name) }, body: file });
const pdfBytes = new Uint8Array(await res.arrayBuffer()); // → pdf.js
```

## Libraries

| Layer | Library | License |
|---|---|---|
| Conversion | Gotenberg (`gotenberg/gotenberg:8`) | MIT |
| | LibreOffice (inside Gotenberg) | MPL-2.0 |
| Web viewer | pdf.js (`pdfjs-dist`) | Apache-2.0 |
| iOS viewer | PDFKit | Built into iOS |
| Android viewer | PdfRenderer (or Jetpack `androidx.pdf`) | Built into Android |
| Server | Node.js built-ins only (`http`, `fetch`, `FormData`) | MIT |

## Why pdf.js instead of the browser's built-in viewer

An `<iframe src="file.pdf">` is simpler, but Android Chrome downloads PDFs instead of displaying
them, and iOS Safari often shows only the first page in a frame. pdf.js renders the same way in
every browser.

## Security notes

- Gotenberg is reachable only by the demo server (no published port), and the app binds to
  `127.0.0.1`.
- Responses send `Cache-Control: no-store`, `X-Content-Type-Options: nosniff` and
  `Referrer-Policy: no-referrer`.
- Uploads are limited to 50 MB and an extension allowlist. File names are reduced to a basename, so
  `../` path tricks don't work.
- Mobile clients keep files in the app sandbox and delete converted PDFs after rendering. They
  never hand documents to other apps.
- For production, add authentication on `/api/convert`, pin the Gotenberg image by digest, and run
  it with resource limits.
