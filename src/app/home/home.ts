import { ChangeDetectionStrategy, Component, inject } from '@angular/core';
import { RouterLink } from '@angular/router';
import { formatBytes } from '../../viewer/index.js';
import { DocumentsService, SAMPLE_GROUPS } from '../documents.service';

/** Home: the document gallery. */
@Component({
  selector: 'app-home',
  imports: [RouterLink],
  changeDetection: ChangeDetectionStrategy.OnPush,
  host: { class: 'empty' },
  template: `
    <header class="intro">
      <h1>Documents</h1>
      <p>Select a document to view. Use the Download button to save the original file.</p>
      <p class="intro-formats">PDF · Word (.doc .docx .docm) · Excel (.xls .xlsx .xlsm .csv) · PowerPoint (.ppt .pptx) · Images (.jpg .jpeg .tif .tiff) · Text (.txt)</p>
    </header>
    <section class="samples" aria-label="Documents">
      <div id="samples">
        @for (group of groups; track group.title) {
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
        }
      </div>
    </section>
  `,
})
export class HomePage {
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
