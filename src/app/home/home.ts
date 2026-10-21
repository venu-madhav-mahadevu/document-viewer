import { ChangeDetectionStrategy, Component, inject } from '@angular/core';
import { NgTemplateOutlet } from '@angular/common';
import { RouterLink } from '@angular/router';
import { formatBytes } from '../../viewer/format.js';
import { DocumentsService, SAMPLE_GROUPS } from '../documents.service';

/** Home: the document gallery. */
@Component({
  selector: 'app-home',
  imports: [RouterLink, NgTemplateOutlet],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { class: 'empty' },
  template: `
    <header class="intro">
      <h1>Documents</h1>
      <p>Select a document to view, or test your own: <button type="button" class="link-btn" (click)="pickFile()">open a file</button> or drop it anywhere on this page. Files never leave your browser.</p>
      <p class="intro-formats">PDF · Word (.doc .docx .docm) · Excel (.xls .xlsx .xlsm .csv) · PowerPoint (.ppt .pptx) · Images (.jpg .jpeg .tif .tiff) · Text (.txt)</p>
    </header>
    <section class="samples" aria-label="Documents">
      <div id="samples">
        @for (group of groups; track group.title; let i = $index) {
          <!-- The first row of groups renders at once; the rest when scrolled into view or when idle. -->
          @if (i < 3) {
            <ng-container *ngTemplateOutlet="groupTpl; context: { $implicit: group }" />
          } @else {
            @defer (on viewport; on idle) {
              <ng-container *ngTemplateOutlet="groupTpl; context: { $implicit: group }" />
            } @placeholder {
              <div class="sample-group sample-group-placeholder" aria-hidden="true"></div>
            }
          }
        }
        <ng-template #groupTpl let-group>
          <div class="sample-group">
            <h3>{{ group.title }}</h3>
            <ul>
              @for (entry of group.files; track entry[0]) {
                <li>
                  <button type="button" [routerLink]="['/view']" [queryParams]="{ file: entry[0] }">
                    <span class="sample-top">
                      <span class="sample-name">{{ fileName(entry[0]) }}</span>
                      <span class="sample-size">{{ size(entry[0]) }}</span>
                    </span>
                    <span class="sample-note">{{ entry[1] }}</span>
                  </button>
                </li>
              }
            </ul>
          </div>
        </ng-template>
      </div>
    </section>
  `,
})
export class HomePage {
  /** Opens the sidebar's file picker (one picker for the whole app). */
  protected pickFile(): void {
    document.querySelector<HTMLInputElement>('.app-sidebar input[type=file]')?.click();
  }

  private readonly docs = inject(DocumentsService);
  protected readonly groups = SAMPLE_GROUPS;

  protected fileName(path: string): string {
    return path.split('/').pop() || path;
  }

  protected size(path: string): string {
    const bytes = this.docs.sizes()[path];
    return bytes ? formatBytes(bytes) : '';
  }
}
