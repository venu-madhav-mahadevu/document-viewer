import { ChangeDetectionStrategy, Component, computed, effect, inject, input, resource, viewChild } from '@angular/core';
import { Location } from '@angular/common';
import { Router } from '@angular/router';
import { DocViewerComponent } from '../doc-viewer/doc-viewer';
import { DocumentsService } from '../documents.service';

/**
 * View page: /view?file=<sample path>  (or ?src=<same-origin URL>).
 * The document is loaded with resource(): loading and error states come built in, and
 * switching to another document cancels the download still in progress (abortSignal).
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

  /** The sample file, downloaded with progress; re-runs (and cancels the old download) when ?file= changes. */
  private readonly sample = resource({
    params: () => (this.src() ? undefined : this.file()),
    loader: ({ params: path, abortSignal }) =>
      this.docs.load(path, (message) => this.viewer()?.loading(message, this.nameOf(path)), abortSignal),
  });

  /** What <doc-viewer> shows: a same-origin URL, or the downloaded sample. */
  protected readonly doc = computed<File | string | null>(() =>
    this.src() || (this.sample.hasValue() ? this.sample.value() : null) || null);

  constructor() {
    // Spinner while downloading (large samples take seconds), message on failure.
    effect(() => {
      const viewer = this.viewer();
      const path = this.file();
      if (!viewer || !path || this.src()) return;
      const name = this.nameOf(path);
      const status = this.sample.status();
      if (status === 'loading' || status === 'reloading') viewer.loading(`Loading ${name}…`, name);
      else if (status === 'error') {
        viewer.loading(`Could not load ${name} (${(this.sample.error() as Error | undefined)?.message}).`, name, 'error');
      }
    });
  }

  private nameOf(path: string): string {
    return path.split('/').pop() || path;
  }

  /** Back: return to the list (same as the browser Back button). */
  protected goBack(): void {
    const id = (history.state as { navigationId?: number } | null)?.navigationId ?? 1;
    if (id > 1) this.location.back();
    else void this.router.navigate(['/']);
  }
}
