// The only formats the viewer accepts.
export const SUPPORTED_EXTENSIONS = [
  'pdf',
  'doc', 'docx', 'docm',
  'xls', 'xlsx', 'xlsm', 'csv',
  'ppt', 'pptx',
  'jpg', 'jpeg', 'tif', 'tiff',
  'txt',
];

const EXT = {
  pdf: 'pdf',
  docx: 'docx', docm: 'docx',
  doc: 'doc',
  xlsx: 'sheet', xlsm: 'sheet', xls: 'sheet', csv: 'sheet',
  pptx: 'pptx',
  ppt: 'ppt',
  jpg: 'image', jpeg: 'image',
  tif: 'tiff', tiff: 'tiff',
  txt: 'text',
};

/** Value for <input type="file" accept="…">. */
export const ACCEPT = SUPPORTED_EXTENSIONS.map((e) => `.${e}`).join(',');

export function extensionOf(name = '') {
  const m = /\.([a-z0-9]+)$/i.exec(name);
  return m ? m[1].toLowerCase() : '';
}

const startsWith = (bytes, sig) => sig.every((b, i) => bytes[i] === b);

/**
 * Work out which renderer handles a file. Uses the extension first, then sniffs
 * magic bytes so files served without a name (e.g. from an API) still open.
 * Anything outside SUPPORTED_EXTENSIONS is reported as unsupported.
 * @returns {Promise<{kind: string}>}
 */
export async function detectKind(name, buffer) {
  const ext = extensionOf(name);
  if (EXT[ext]) return { kind: EXT[ext] };
  if (ext) return { kind: 'unsupported' }; // has an extension we don't support

  // No extension: identify the content.
  const b = new Uint8Array(buffer.slice(0, 8));
  if (startsWith(b, [0x25, 0x50, 0x44, 0x46])) return { kind: 'pdf' };
  if (startsWith(b, [0xff, 0xd8, 0xff])) return { kind: 'image' };
  if (startsWith(b, [0x49, 0x49, 0x2a, 0x00]) || startsWith(b, [0x4d, 0x4d, 0x00, 0x2a])) return { kind: 'tiff' };
  if (startsWith(b, [0x50, 0x4b, 0x03, 0x04])) {
    try {
      const { default: JSZip } = await import('jszip');
      const zip = await JSZip.loadAsync(buffer);
      if (zip.file('word/document.xml')) return { kind: 'docx' };
      if (zip.file('xl/workbook.xml')) return { kind: 'sheet' };
      if (zip.file('ppt/presentation.xml')) return { kind: 'pptx' };
    } catch { /* not a readable zip */ }
  }
  if (startsWith(b, [0xd0, 0xcf, 0x11, 0xe0])) {
    // Office 97-2003 compound file: tell Word / Excel / PowerPoint apart by their streams.
    try {
      const { default: CFB } = await import('cfb');
      const cfb = CFB.read(new Uint8Array(buffer), { type: 'array' });
      if (CFB.find(cfb, 'WordDocument')) return { kind: 'doc' };
      if (CFB.find(cfb, 'PowerPoint Document')) return { kind: 'ppt' };
      if (CFB.find(cfb, 'Workbook') || CFB.find(cfb, 'Book')) return { kind: 'sheet' };
    } catch { /* not a readable compound file */ }
  }
  if (looksLikeText(new Uint8Array(buffer.slice(0, 4096)))) return { kind: 'text' };
  return { kind: 'unsupported' };
}

function looksLikeText(bytes) {
  for (const c of bytes) if (c === 0) return false;
  return true;
}
