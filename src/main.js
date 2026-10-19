import './app.css';
import { createDocViewer, ACCEPT, formatBytes } from './viewer/index.js';
import { SAMPLE_GROUPS } from './samples.js';

const fileInput = document.getElementById('file-input');
fileInput.accept = ACCEPT; // only the supported formats
const list = document.getElementById('file-list');
const empty = document.getElementById('empty');
const viewerEl = document.getElementById('viewer');
const overlay = document.getElementById('drop-overlay');

const viewer = createDocViewer(viewerEl, { download: false }); // public site: view only
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

function show(i) {
  activeIndex = i;
  empty.hidden = true;
  viewerEl.hidden = false;
  renderList();
  viewer.open(files[i]);
}

function add(newFiles) {
  const start = files.length;
  files.push(...newFiles);
  if (newFiles.length) show(start);
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
    b.append(name, desc);
    b.addEventListener('click', async () => {
      const fileName = file.split('/').pop();
      // Spinner straight away: large samples take seconds to download.
      empty.hidden = true;
      viewerEl.hidden = false;
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
  empty.hidden = true;
  viewerEl.hidden = false;
  viewer.open(src);
}

document.getElementById('home').addEventListener('click', () => {
  activeIndex = -1;
  renderList();
  viewerEl.hidden = true;
  empty.hidden = false;
});
