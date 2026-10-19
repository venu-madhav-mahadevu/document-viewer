import './app.css';
import { createDocViewer, ACCEPT } from './viewer/index.js';
import { SAMPLE_GROUPS } from './samples.js';

const fileInput = document.getElementById('file-input');
fileInput.accept = ACCEPT; // only the supported formats
const list = document.getElementById('file-list');
const empty = document.getElementById('empty');
const viewerEl = document.getElementById('viewer');
const overlay = document.getElementById('drop-overlay');

const viewer = createDocViewer(viewerEl);
const files = [];
let activeIndex = -1;

function renderList() {
  list.replaceChildren(...files.map((file, i) => {
    const li = document.createElement('li');
    const b = document.createElement('button');
    b.type = 'button';
    b.textContent = file.name;
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
    name.textContent = file;
    const desc = document.createElement('span');
    desc.className = 'sample-note';
    desc.textContent = note;
    b.append(name, desc);
    b.addEventListener('click', async () => {
      const res = await fetch(new URL(`samples/${file}`, document.baseURI));
      add([new File([await res.blob()], file)]);
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
