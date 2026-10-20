import './app.css';
import { createDocViewer, ACCEPT, formatBytes } from './viewer/index.js';
import { SAMPLE_GROUPS } from './samples.js';

const fileInput = document.getElementById('file-input');
fileInput.accept = ACCEPT; // only the supported formats
const list = document.getElementById('file-list');
const empty = document.getElementById('empty');
const viewerEl = document.getElementById('viewer');
const overlay = document.getElementById('drop-overlay');

// Public site: view only. Back returns to the file list (same as the browser Back button).
const viewer = createDocViewer(viewerEl, { download: false, onBack: () => goBack() });
const files = [];
let activeIndex = -1;

function renderList() {
  list.replaceChildren(...files.map((file, i) => {
    const li = document.createElement('li');
    const b = document.createElement('button');
    b.type = 'button';
    const label = document.createElement('span');
    label.className = 'file-name';
    label.textContent = file.name;
    const size = document.createElement('span');
    size.className = 'file-size';
    size.textContent = formatBytes(file.size);
    b.append(label, size);
    b.title = file.name;
    b.setAttribute('aria-current', String(i === activeIndex));
    b.addEventListener('click', () => show(i));
    li.append(b);
    return li;
  }));
}

function showViewer() {
  empty.hidden = true;
  viewerEl.hidden = false;
  // One history entry while a document is open, so browser Back returns to the list.
  if (history.state?.view !== 'doc') history.pushState({ view: 'doc' }, '');
}

function show(i) {
  activeIndex = i;
  showViewer();
  renderList();
  viewer.open(files[i]);
}

function goHome() {
  activeIndex = -1;
  renderList();
  viewerEl.hidden = true;
  empty.hidden = false;
}

function goBack() {
  if (history.state?.view === 'doc') history.back(); // popstate below shows the list
  else goHome();
}

window.addEventListener('popstate', () => { if (history.state?.view !== 'doc') goHome(); });

function add(newFiles) {
  // A file that is already in the list (same name and size) is shown again, not duplicated.
  let first = -1;
  for (const file of newFiles) {
    let i = files.findIndex((f) => f.name === file.name && f.size === file.size);
    if (i === -1) { files.push(file); i = files.length - 1; }
    if (first === -1) first = i;
  }
  if (first !== -1) show(first);
}

fileInput.addEventListener('change', () => { add([...fileInput.files]); fileInput.value = ''; });
document.getElementById('empty-open').addEventListener('click', () => fileInput.click());

let dragDepth = 0;
window.addEventListener('dragenter', (e) => { if (e.dataTransfer?.types.includes('Files')) { dragDepth++; overlay.hidden = false; } });
window.addEventListener('dragleave', () => { if (--dragDepth <= 0) { dragDepth = 0; overlay.hidden = true; } });
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  overlay.hidden = true;
  add([...(e.dataTransfer?.files || [])]);
});

const samplesEl = document.getElementById('samples');
// Exact sample sizes, recorded at build time (scripts/sample-sizes.mjs).
const sampleSizes = fetch(new URL('samples/sizes.json', document.baseURI)).then((r) => (r.ok ? r.json() : {})).catch(() => ({}));
for (const group of SAMPLE_GROUPS) {
  const box = document.createElement('div');
  box.className = 'sample-group';
  const h = document.createElement('h3');
  h.textContent = group.title;
  const ul = document.createElement('ul');
  for (const [file, note] of group.files) {
    const li = document.createElement('li');
    const b = document.createElement('button');
    b.type = 'button';
    const name = document.createElement('span');
    name.className = 'sample-name';
    name.textContent = file.split('/').pop();
    const desc = document.createElement('span');
    desc.className = 'sample-note';
    desc.textContent = note;
    const size = document.createElement('span');
    size.className = 'sample-size';
    const top = document.createElement('span');
    top.className = 'sample-top';
    top.append(name, size);
    b.append(top, desc);
    sampleSizes.then((sizes) => { if (sizes[file]) size.textContent = formatBytes(sizes[file]); });
    b.addEventListener('click', async () => {
      const fileName = file.split('/').pop();
      // Spinner straight away: large samples take seconds to download.
      showViewer();
      viewer.loading(`Loading ${fileName}…`, fileName);
      try {
        const res = await fetch(new URL(`samples/${file}`, document.baseURI));
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const total = Number(res.headers.get('content-length')) || 0;
        const reader = res.body.getReader();
        const chunks = [];
        let received = 0;
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          chunks.push(value);
          received += value.length;
          viewer.loading(total
            ? `Loading ${fileName}… ${Math.min(100, Math.round((received / total) * 100))}%`
            : `Loading ${fileName}… ${formatBytes(received)}`);
        }
        add([new File(chunks, fileName)]);
      } catch (err) {
        viewer.loading(`Could not load ${fileName} (${err.message}).`, fileName, 'error');
      }
    });
    li.append(b);
    ul.append(li);
  }
  box.append(h, ul);
  samplesEl.append(box);
}

// Deep link for same-origin documents: /?src=/files/report.pdf
const src = new URLSearchParams(location.search).get('src');
if (src) {
  showViewer();
  viewer.open(src);
}

document.getElementById('home').addEventListener('click', goBack);
