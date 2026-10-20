// Types for the framework-free viewer (index.js), used by the Angular <doc-viewer> component.

export interface DocViewerOptions {
  /** Largest file the viewer opens. Default 200 MB. */
  maxBytes?: number;
  /** Show the Download button. Default true. */
  download?: boolean;
  /** Show the Print button. Default true. */
  print?: boolean;
  /** Show the format label (PDF, Word, ...) next to the file name. Default false. */
  badge?: boolean;
  /** Options for fetch() when `open()` gets a URL (headers, credentials ...). */
  fetchInit?: RequestInit;
  /** When given, a Back button appears at the left of the toolbar and calls it. */
  onBack?: () => void;
}

export interface DocViewer {
  element: HTMLElement;
  /** Opens a File, Blob, ArrayBuffer or same-origin URL. */
  open(source: File | Blob | ArrayBuffer | string, name?: string): Promise<void>;
  /** Shows the spinner (kind 'loading') or an error message (kind 'error'). */
  loading(message: string, name?: string, kind?: 'loading' | 'error'): void;
  zoomIn(): void;
  zoomOut(): void;
  fitWidth(): void;
  setZoom(zoom: number): void;
  goToPage(page: number): void;
  find(): void;
  print(): void;
  destroy(): void;
}

export function createDocViewer(container: HTMLElement, options?: DocViewerOptions): DocViewer;
export function formatBytes(bytes: number): string;
export function formatDuration(ms: number): string;

export interface FormatInfo {
  extensions: string[];
  label: string;
  library: string;
  paged: boolean;
  badge?: string;
}
export interface LibraryInfo { name: string; npm: string; license: string; url: string }

export const SUPPORTED_EXTENSIONS: string[];
export const ACCEPT: string;
export const FORMATS: Record<string, FormatInfo>;
export const LIBRARIES: Record<string, LibraryInfo>;
