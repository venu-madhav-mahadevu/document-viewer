const EXT = {
  pdf: 'pdf',
  docx: 'docx', docm: 'docx', dotx: 'docx',
  xlsx: 'sheet', xlsm: 'sheet', xlsb: 'sheet', xls: 'sheet', ods: 'sheet', csv: 'sheet', tsv: 'sheet',
  pptx: 'pptx', pptm: 'pptx', ppsx: 'pptx', potx: 'pptx',
  png: 'image', jpg: 'image', jpeg: 'image', gif: 'image', webp: 'image', bmp: 'image', svg: 'image', avif: 'image', ico: 'image',
  txt: 'text', log: 'text', md: 'text', json: 'text', xml: 'text', yaml: 'text', yml: 'text',
  ini: 'text', cfg: 'text', html: 'text', htm: 'text', css: 'text', js: 'text', ts: 'text', sql: 'text',
};

// Legacy binary Office formats have no reliable open-source browser renderer.
const LEGACY = { doc: 'Word 97-2003 (.doc)', ppt: 'PowerPoint 97-2003 (.ppt)', rtf: 'Rich Text (.rtf)' };

export function extensionOf(name = '') {
  const m = /\.([a-z0-9]+)$/i.exec(name);
  return m ? m[1].toLowerCase() : '';
}

const startsWith = (bytes, sig) => sig.every((b, i) => bytes[i] === b);

/**
 * Work out which renderer handles a file. Uses the extension first, then sniffs
 * magic bytes so files served without a name (e.g. from an API) still open.
 * @returns {Promise<{kind: string, legacy?: string}>}
 */
export async function detectKind(name, buffer) {
  const ext = extensionOf(name);
  if (LEGACY[ext]) return { kind: 'unsupported', legacy: LEGACY[ext] };
  if (EXT[ext]) return { kind: EXT[ext] };

  const b = new Uint8Array(buffer.slice(0, 8));
  if (startsWith(b, [0x25, 0x50, 0x44, 0x46])) return { kind: 'pdf' };
  if (startsWith(b, [0x89, 0x50, 0x4e, 0x47]) || startsWith(b, [0xff, 0xd8, 0xff]) ||
      startsWith(b, [0x47, 0x49, 0x46, 0x38])) return { kind: 'image' };
  if (startsWith(b, [0x50, 0x4b, 0x03, 0x04])) {
    try {
      const { default: JSZip } = await import('jszip');
      const zip = await JSZip.loadAsync(buffer);
      if (zip.file('word/document.xml')) return { kind: 'docx' };
      if (zip.file('xl/workbook.xml') || zip.file('xl/workbook.bin')) return { kind: 'sheet' };
      if (zip.file('ppt/presentation.xml')) return { kind: 'pptx' };
    } catch { /* not a readable zip */ }
  }
  if (startsWith(b, [0xd0, 0xcf, 0x11, 0xe0])) return { kind: 'sheet' }; // OLE2: SheetJS reads .xls
  if (looksLikeText(new Uint8Array(buffer.slice(0, 4096)))) return { kind: 'text' };
  return { kind: 'unsupported' };
}

function looksLikeText(bytes) {
  for (const c of bytes) if (c === 0) return false;
  return true;
}
