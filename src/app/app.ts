import { ChangeDetectionStrategy, Component, computed, inject, signal } from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { NavigationEnd, Router, RouterLink, RouterOutlet } from '@angular/router';
import { filter, map } from 'rxjs';
import { formatBytes } from '../viewer/format.js';
import { ACCEPT } from '../viewer/formats.js';
import { DocumentsService } from './documents.service';

/** Demo and architecture pages (not yet moved to Angular; served by the staging build). */
const DEMOS_URL = 'https://venu-madhav-mahadevu.github.io/document-viewer/staging/demos/';

/** App shell: sidebar (brand, opened documents, links) and the routed page. */
@Component({
  selector: 'app-root',
  imports: [RouterOutlet, RouterLink],
  changeDetection: ChangeDetectionStrategy.OnPush,
  template: `
    <div class="app" [class.viewing]="viewing()">
      <aside class="app-sidebar" aria-label="Open documents">
        <button type="button" class="brand" title="Home and samples" routerLink="/">Docs Viewer</button>
        <label class="open-btn" title="Open a document from your computer (it stays in your browser)">
          <input type="file" multiple hidden [accept]="accept" (change)="openPicked($event)" />
          Open file
        </label>
        <ul class="file-list">
          @for (doc of docs.opened(); track doc.path) {
            <li>
              <button type="button" [title]="doc.file.name" [attr.aria-current]="doc.path === activeFile()"
                      [routerLink]="['/view']" [queryParams]="{ file: doc.path }">
                <span class="file-name">{{ doc.file.name }}</span>
                <span class="file-size">{{ size(doc.file.size) }}</span>
              </button>
            </li>
          }
        </ul>
        <a class="demos-link" [href]="demosUrl + 'index.html'">Library demos with code</a>
        <a class="demos-link" [href]="demosUrl + 'architecture.html'">Architecture</a>
        <p class="privacy">Documents open in your browser; nothing is uploaded, including files you open yourself.</p>
      </aside>
      <main class="main">
        <router-outlet />
      </main>
    </div>
    @if (dragging()) {
      <div class="drop-overlay">Drop to open (stays in your browser)</div>
    }
  `,
})
export class App {
  protected readonly docs = inject(DocumentsService);
  private readonly router = inject(Router);
  protected readonly demosUrl = DEMOS_URL;
  protected readonly accept = ACCEPT;
  /** A file is being dragged over the page. */
  protected readonly dragging = signal(false);
  private dragDepth = 0;

  private readonly url = toSignal(
    this.router.events.pipe(filter((e) => e instanceof NavigationEnd), map(() => this.router.url)),
    { initialValue: this.router.url },
  );
  /** A document is open: on phones the sidebar hides so the document gets the whole screen. */
  protected readonly viewing = computed(() => this.url().startsWith('/view'));
  protected readonly activeFile = computed(() => this.router.parseUrl(this.url()).queryParams['file'] ?? null);

  constructor() {
    // Drag and drop your own files anywhere on the page (read in the browser only).
    const hasFiles = (e: DragEvent) => [...(e.dataTransfer?.types || [])].includes('Files');
    window.addEventListener('dragenter', (e) => { if (hasFiles(e)) { this.dragDepth++; this.dragging.set(true); } });
    window.addEventListener('dragleave', () => { if (--this.dragDepth <= 0) { this.dragDepth = 0; this.dragging.set(false); } });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault();
      this.dragDepth = 0;
      this.dragging.set(false);
      this.open([...(e.dataTransfer?.files || [])]);
    });

    // Old deep links: /?src=samples/report.pdf -> #/view?file=report.pdf
    const src = new URLSearchParams(location.search).get('src');
    if (src) {
      history.replaceState(null, '', location.pathname + location.hash);
      const file = src.replace(/^\.?\/?samples\//, '');
      void this.router.navigate(['/view'], { queryParams: file !== src ? { file } : { src } });
    }
  }

  /** Files chosen with the Open file button. */
  protected openPicked(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.open([...(input.files || [])]);
    input.value = '';
  }

  private open(files: File[]): void {
    const key = this.docs.openLocal(files);
    if (key) void this.router.navigate(['/view'], { queryParams: { file: key } });
  }

  protected size(bytes: number): string {
    return formatBytes(bytes);
  }
}
