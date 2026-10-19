import { setup } from './shared.js';
import source from './papaparse.js?raw';

// #region demo
import Papa from 'papaparse';

async function render(buffer, name, out) {
  const text = new TextDecoder().decode(buffer);

  // Handles quotes, embedded commas/newlines, tabs (auto-detects the delimiter).
  const { data, meta, errors } = Papa.parse(text, { header: true, skipEmptyLines: true });

  const table = document.createElement('table');
  table.className = 'data';
  const head = table.createTHead().insertRow();
  for (const field of meta.fields) head.appendChild(document.createElement('th')).textContent = field;

  const body = table.createTBody();
  for (const row of data) {
    const tr = body.insertRow();
    for (const field of meta.fields) tr.insertCell().textContent = row[field]; // text only - safe
  }
  out.append(table);

  if (errors.length) console.warn('CSV issues:', errors);
}
// #endregion

setup({ sample: 'accounts.csv', accept: '.csv', render, source });
