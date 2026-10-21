import { Injectable, signal } from '@angular/core';

export interface SampleGroup { title: string; files: [path: string, note: string][] }

/** Demo files in public/samples, grouped for the home-screen gallery. */
export const SAMPLE_GROUPS: SampleGroup[] = [
  { title: 'PDF', files: [
    ['quarterly-report.pdf', '12 pages, selectable text'],
    ['mixed-page-sizes.pdf', 'Letter + A4, portrait + landscape'],
    ['protected-statement.pdf', 'Password: demo1234'],
  ] },
  { title: 'Word', files: [
    ['onboarding-guide.docx', '.docx'],
    ['onboarding-macro.docm', '.docm macro-enabled (macros never run)'],
    ['onboarding-guide.doc', '.doc Word 97-2003'],
  ] },
  { title: 'Excel', files: [
    ['regional-sales.xlsx', '.xlsx, merged cells, 2 sheets'],
    ['branches-macro.xlsm', '.xlsm macro-enabled'],
    ['branches-legacy.xls', '.xls Excel 97-2003'],
    ['accounts.csv', '.csv with quoted commas'],
    ['formatted-report.xlsx', '.xlsx formats and colors: currency, %, dates, fills, fonts, borders (3 sheets)'],
    ['formatted-report.xls', '.xls Excel 97-2003 copy of the formats and colors file'],
  ] },
  { title: 'PowerPoint', files: [
    ['strategy-deck.pptx', '.pptx'],
    ['strategy-deck.ppt', '.ppt PowerPoint 97-2003'],
  ] },
  { title: 'Images', files: [
    ['photo.jpg', '.jpg'],
    ['photo-chart.jpeg', '.jpeg'],
    ['scanned-contract.tif', '.tif 3-page scan (fax G4 compression)'],
    ['color-chart.tiff', '.tiff color (LZW)'],
    ['color-scan-jpeg.tif', '.tif color scan (JPEG compression)'],
  ] },
  { title: 'Text', files: [
    ['readme.txt', '.txt with Unicode'],
  ] },
  { title: 'Large files', files: [
    ['large/large-report-49mb.pdf', '49 MB, 100 pages'],
    ['large/large-transactions-50mb.xlsx', '50 MB, 370,000 rows'],
    ['large/large-scan-60-pages.tif', '36 MB, 60-page scan'],
  ] },
];

const sampleUrl = (path: string) => new URL(`samples/${path}`, document.baseURI).href;

/** Downloads documents (with progress) and keeps the ones opened in this session. */
@Injectable({ providedIn: 'root' })
export class DocumentsService {
  /** Documents opened in this session, shown in the sidebar. */
  readonly opened = signal<{ path: string; file: File }[]>([]);
  /** Exact sample sizes, recorded at build time (scripts/sample-sizes.mjs). */
  readonly sizes = signal<Record<string, number>>({});

  constructor() {
    fetch(new URL('samples/sizes.json', document.baseURI))
      .then((r) => (r.ok ? r.json() : {}))
      .then((sizes) => this.sizes.set(sizes))
      .catch(() => {});
  }

  /** The sample at `path` as a File; reports progress while it downloads. */
  async load(path: string, onProgress: (message: string) => void): Promise<File> {
    const cached = this.opened().find((d) => d.path === path);
    if (cached) return cached.file;

    const name = path.split('/').pop() || path;
    const res = await fetch(sampleUrl(path));
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const total = Number(res.headers.get('content-length')) || this.sizes()[path] || 0;
    const reader = res.body!.getReader();
    const chunks: BlobPart[] = [];
    let received = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      received += value.length;
      onProgress(total
        ? `Loading ${name}… ${Math.min(100, Math.round((received / total) * 100))}%`
        : `Loading ${name}… ${Math.round(received / 1024)} KB`);
    }
    const file = new File(chunks, name);
    this.opened.update((list) => [...list, { path, file }]);
    return file;
  }
}
