import { ChangeDetectionStrategy, Component, computed, inject } from '@angular/core';
import { toSignal } from '@angular/core/rxjs-interop';
import { NavigationEnd, Router, RouterLink, RouterOutlet } from '@angular/router';
import { filter, map } from 'rxjs';
import { formatBytes } from '../viewer/index.js';
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
        <p class="privacy">Documents open in your browser; nothing is uploaded. Use Download to save a copy.</p>
      </aside>
      <main class="main">
        <router-outlet />
      </main>
    </div>
  `,
})
export class App {
  protected readonly docs = inject(DocumentsService);
  private readonly router = inject(Router);
  protected readonly demosUrl = DEMOS_URL;

  private readonly url = toSignal(
    this.router.events.pipe(filter((e) => e instanceof NavigationEnd), map(() => this.router.url)),
    { initialValue: this.router.url },
  );
  /** A document is open: on phones the sidebar hides so the document gets the whole screen. */
  protected readonly viewing = computed(() => this.url().startsWith('/view'));
  protected readonly activeFile = computed(() => this.router.parseUrl(this.url()).queryParams['file'] ?? null);

  constructor() {
    // View only: files cannot be opened from the user's computer. Block drops so the
    // browser doesn't open a dropped file on its own either.
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => e.preventDefault());

    // Old deep links: /?src=samples/report.pdf -> #/view?file=report.pdf
    const src = new URLSearchParams(location.search).get('src');
    if (src) {
      history.replaceState(null, '', location.pathname + location.hash);
      const file = src.replace(/^\.?\/?samples\//, '');
      void this.router.navigate(['/view'], { queryParams: file !== src ? { file } : { src } });
    }
  }

  protected size(bytes: number): string {
    return formatBytes(bytes);
  }
}
