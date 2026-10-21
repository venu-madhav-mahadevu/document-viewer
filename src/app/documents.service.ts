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
    ['regional-sales.xlsx', '.xlsx colorful: shading, merged cells, totals, 2 sheets'],
    ['branches-macro.xlsm', '.xlsm macro-enabled, color-coded ratios'],
    ['branches-legacy.xls', '.xls Excel 97-2003 with cell colors'],
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
  { title: 'Test pack (manual testing)', files: [
    ['test-pack/TEST-CHECKLIST.xlsx', 'START HERE: all 506 tests, expected results, Pass/Fail column'],
    ['test-pack/test-word.docx', 'Word: fonts, styles, bullets, tables, pictures, shapes, layout (7 pages)'],
    ['test-pack/test-word.docm', 'Word macro-enabled: same tests'],
    ['test-pack/test-word.doc', 'Word 97-2003: same tests'],
    ['test-pack/test-excel.xlsx', 'Excel: formats, colors, fonts, borders, formulas, rules (6 sheets)'],
    ['test-pack/test-excel.xlsm', 'Excel macro-enabled: same tests'],
    ['test-pack/test-excel.xls', 'Excel 97-2003: same tests'],
    ['test-pack/test-data.csv', 'CSV: quotes, commas, accents, Asian text, numbers, dates'],
    ['test-pack/test-slides.pptx', 'PowerPoint: text, bullets, shapes, pictures, table, charts (10 slides)'],
    ['test-pack/test-slides.ppt', 'PowerPoint 97-2003: same tests'],
    ['test-pack/test-pdf.pdf', 'PDF: fonts, Asian text, drawings, links, forms, page sizes (9 pages)'],
    ['test-pack/test-pdf-protected.pdf', 'PDF with password test1234'],
    ['test-pack/test-image.jpg', 'JPG test chart: colors, gray levels, sharpness'],
    ['test-pack/test-image.jpeg', 'JPEG extension, zoom'],
    ['test-pack/test-image-progressive.jpeg', 'Progressive JPEG'],
    ['test-pack/test-image-rotated-exif.jpg', 'Phone photo rotation (EXIF)'],
    ['test-pack/test-image-cmyk.jpg', 'CMYK JPEG (print / scanner)'],
    ['test-pack/test-image-large.jpg', '8000 x 5000 photo'],
    ['test-pack/test-scan-multipage.tif', '5-page scan, page bar, thumbnails'],
    ['test-pack/test-tiff-16bit.tif', 'TIFF, 16-bit grayscale'],
    ['test-pack/test-tiff-cmyk.tiff', 'TIFF, CMYK color'],
    ['test-pack/test-tiff-deflate.tiff', 'TIFF, Deflate compression'],
    ['test-pack/test-tiff-fax-g3.tif', 'TIFF, black & white fax G3'],
    ['test-pack/test-tiff-fax-g4.tif', 'TIFF, black & white fax G4'],
    ['test-pack/test-tiff-grayscale.tiff', 'TIFF, grayscale'],
    ['test-pack/test-tiff-jpeg.tif', 'TIFF, JPEG compression'],
    ['test-pack/test-tiff-lzw.tiff', 'TIFF, LZW compression'],
    ['test-pack/test-tiff-packbits.tiff', 'TIFF, PackBits compression'],
    ['test-pack/test-tiff-uncompressed.tiff', 'TIFF, no compression'],
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

  /**
   * Files the user opens from their computer (button or drag and drop). They stay in memory
   * in this tab; nothing is uploaded. Returns the key to show the first one (?file=<key>).
   */
  openLocal(files: File[]): string | null {
    let first: string | null = null;
    for (const file of files) {
      const path = `local:${file.name}:${file.size}:${file.lastModified}`;
      if (!this.opened().some((d) => d.path === path)) this.opened.update((list) => [...list, { path, file }]);
      first ??= path;
    }
    return first;
  }

  /** The sample at `path` as a File; reports progress while it downloads. `signal` cancels it. */
  async load(path: string, onProgress: (message: string) => void, signal?: AbortSignal): Promise<File> {
    const cached = this.opened().find((d) => d.path === path);
    if (cached) return cached.file;
    if (path.startsWith('local:')) throw new Error('this file was opened from your computer; please open it again');

    const name = path.split('/').pop() || path;
    const res = await fetch(sampleUrl(path), { signal });
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
