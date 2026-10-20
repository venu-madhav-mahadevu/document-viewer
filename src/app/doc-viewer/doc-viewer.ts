import {
  AfterViewInit, ChangeDetectionStrategy, Component, ElementRef, OnDestroy, effect, input, output, viewChild,
} from '@angular/core';
import { createDocViewer, type DocViewer as Viewer } from '../../viewer/index.js';

/**
 * <doc-viewer>: the document viewer as an Angular component.
 * Opens all 15 formats with pdf.js, docx-preview, SheetJS CE, JSZip, cfb and UTIF.js.
 *
 *   <doc-viewer [src]="file" [download]="true" [showBack]="true" (back)="goHome()" />
 *
 * `src` can be a File, Blob, ArrayBuffer or a same-origin URL.
 */
@Component({
  selector: 'doc-viewer',
  template: '<div #host class="doc-viewer-host"></div>',
  styles: [':host { display: block; height: 100%; min-height: 0; } .doc-viewer-host { height: 100%; }'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class DocViewerComponent implements AfterViewInit, OnDestroy {
  /** File, Blob, ArrayBuffer or same-origin URL to show. */
  readonly src = input<File | Blob | ArrayBuffer | string | null>(null);
  /** File name, when `src` is a Blob or ArrayBuffer. */
  readonly name = input<string | undefined>(undefined);
  readonly download = input(true);
  readonly print = input(true);
  readonly showBack = input(false);
  readonly maxBytes = input<number | undefined>(undefined);
  /** Back button in the toolbar was clicked (only with [showBack]="true"). */
  readonly back = output<void>();

  private readonly host = viewChild.required<ElementRef<HTMLElement>>('host');
  private viewer?: Viewer;
  private shown: unknown = null;
  private pendingLoading?: [string, string | undefined, 'loading' | 'error'];

  constructor() {
    // Open again whenever [src] changes.
    effect(() => {
      const src = this.src();
      const name = this.name();
      if (this.viewer && src && src !== this.shown) {
        this.shown = src;
        void this.viewer.open(src, name);
      }
    });
  }

  ngAfterViewInit(): void {
    this.viewer = createDocViewer(this.host().nativeElement, {
      download: this.download(),
      print: this.print(),
      maxBytes: this.maxBytes(),
      onBack: this.showBack() ? () => this.back.emit() : undefined,
    });
    if (this.pendingLoading) this.viewer.loading(...this.pendingLoading);
    const src = this.src();
    if (src && src !== this.shown) {
      this.shown = src;
      void this.viewer.open(src, this.name());
    }
  }

  /** Spinner with a message (e.g. download progress), or an error message. */
  loading(message: string, name?: string, kind: 'loading' | 'error' = 'loading'): void {
    if (this.viewer) this.viewer.loading(message, name, kind);
    else this.pendingLoading = [message, name, kind]; // shown as soon as the viewer exists
  }

  zoomIn(): void { this.viewer?.zoomIn(); }
  zoomOut(): void { this.viewer?.zoomOut(); }
  fitWidth(): void { this.viewer?.fitWidth(); }
  goToPage(page: number): void { this.viewer?.goToPage(page); }
  printDocument(): void { this.viewer?.print(); }

  ngOnDestroy(): void {
    this.viewer?.destroy();
  }
}
