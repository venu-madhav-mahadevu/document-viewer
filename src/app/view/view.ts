import { ChangeDetectionStrategy, Component, effect, inject, input, signal, untracked, viewChild } from '@angular/core';
import { Location } from '@angular/common';
import { Router } from '@angular/router';
import { DocViewerComponent } from '../doc-viewer/doc-viewer';
import { DocumentsService } from '../documents.service';

/**
 * View page: /view?file=<sample path>  (or ?src=<same-origin URL>).
 * Downloads the document with a progress spinner, then shows it in <doc-viewer>.
 */
@Component({
  selector: 'app-view',
  imports: [DocViewerComponent],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { class: 'viewer' },
  template: `<doc-viewer [src]="doc()" [download]="true" [showBack]="true" (back)="goBack()" />`,
})
export class ViewPage {
  /** Route query parameters (withComponentInputBinding). */
  readonly file = input<string | undefined>(undefined);
  readonly src = input<string | undefined>(undefined);

  private readonly docs = inject(DocumentsService);
  private readonly router = inject(Router);
  private readonly location = inject(Location);
  private readonly viewer = viewChild(DocViewerComponent);
  protected readonly doc = signal<File | string | null>(null);

  constructor() {
    effect(() => {
      const viewer = this.viewer();
      if (!viewer) return;
      const file = this.file();
      const src = this.src();
      if (src) { this.doc.set(src); return; }
      if (!file) return;
      const name = file.split('/').pop() || file;
      untracked(() => {
        // Spinner straight away: large samples take seconds to download.
        viewer.loading(`Loading ${name}…`, name);
        this.docs
          .load(file, (message) => viewer.loading(message, name))
          .then((f) => { if (this.file() === file) this.doc.set(f); })
          .catch((err: Error) => viewer.loading(`Could not load ${name} (${err.message}).`, name, 'error'));
      });
    });
  }

  /** Back: return to the list (same as the browser Back button). */
  protected goBack(): void {
    const id = (history.state as { navigationId?: number } | null)?.navigationId ?? 1;
    if (id > 1) this.location.back();
    else void this.router.navigate(['/']);
  }
}
