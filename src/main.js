import './app.css';
import { createDocViewer } from './viewer/index.js';

const fileInput = document.getElementById('file-input');
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

// Deep link for same-origin documents: /?src=/files/report.pdf
const src = new URLSearchParams(location.search).get('src');
if (src) {
  empty.hidden = true;
  viewerEl.hidden = false;
  viewer.open(src);
}
