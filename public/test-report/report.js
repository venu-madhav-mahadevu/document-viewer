// Test report: reads the exported checklist and plan (data/*.json) and draws the four tabs.
// Data is written with textContent only, so nothing in it can run as HTML.
const FORMATS = ['DOCX', 'DOCM', 'DOC', 'XLSX', 'XLSM', 'XLS', 'CSV', 'PPTX', 'PPT', 'PDF', 'Image', 'TIFF'];
const MX_FORMATS = ['DOCX', 'DOCM', 'DOC', 'XLSX', 'XLSM', 'XLS', 'PPTX', 'PPT', 'PDF'];
const FAMILIES = ['Word', 'Excel', 'PowerPoint', 'PDF', 'CSV', 'Images'];
const num = new Intl.NumberFormat('en');
const $ = (sel) => document.querySelector(sel);

/** Element helper: el('td', 'cls', 'text'). */
function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}
const rateCell = (pct) => {
  const td = el('td');
  const box = el('div', 'rate');
  const bar = el('div', 'bar');
  const fill = el('i');
  fill.style.width = `${pct}%`;
  bar.append(fill);
  box.append(el('span', 'p', `${pct}%`), bar);
  td.append(box);
  return td;
};
const famOrder = (a, b) => FAMILIES.indexOf(a.family) - FAMILIES.indexOf(b.family);

// ---------- tabs (each one has a link: #works, #results, #matrix, #plan) ----------
const tabs = [...document.querySelectorAll('.tab')];
function show(id, push) {
  for (const t of tabs) {
    const on = t.getAttribute('aria-controls') === id;
    t.setAttribute('aria-selected', String(on));
    document.getElementById(t.getAttribute('aria-controls')).hidden = !on;
  }
  if (push) history.replaceState(null, '', `#${id}`);
}
tabs.forEach((t) => t.addEventListener('click', () => show(t.getAttribute('aria-controls'), true)));
const fromHash = () => { const id = location.hash.slice(1); if (tabs.some((t) => t.getAttribute('aria-controls') === id)) show(id); };
window.addEventListener('hashchange', fromHash);
fromHash();

// ---------- data ----------
async function load(name) {
  const res = await fetch(`data/${name}.json`, { cache: 'no-cache' });
  if (!res.ok) throw new Error(`${name}.json: HTTP ${res.status}`);
  return res.json();
}

Promise.all([load('test-results'), load('fix-plan'), load('two-weeks')])
  .then(([tests, plan, weeks]) => draw(tests, plan, weeks))
  .catch((err) => {
    $('#take').textContent = `The test results could not be loaded (${err.message}).`;
    $('#take').classList.add('error');
  });

function draw(tests, plan, weeks) {
  const pass = tests.filter((t) => t.result === 'Pass').length;
  const fail = tests.filter((t) => t.result === 'Fail').length;
  const rate = Math.round((pass / tests.length) * 100);

  // features: one row per family + feature
  const features = new Map();
  for (const t of tests) {
    const key = `${t.family}|${t.feature}`;
    const f = features.get(key) || { family: t.family, feature: t.feature, tests: 0, fails: 0, formats: [], byFormat: {} };
    f.tests++;
    if (t.result === 'Fail') f.fails++;
    if (!f.formats.includes(t.format)) f.formats.push(t.format);
    f.byFormat[t.format] = t;
    features.set(key, f);
  }
  const works = [...features.values()].filter((f) => !f.fails).sort((a, b) => famOrder(a, b) || a.feature.localeCompare(b.feature));
  const problems = [...features.values()].filter((f) => f.fails).sort((a, b) => famOrder(a, b) || b.fails - a.fails || a.feature.localeCompare(b.feature));

  $('#take').textContent = `${num.format(pass)} of ${num.format(tests.length)} tests pass (${rate}%) on Stage and Angular. `
    + `${num.format(works.length)} features match the original in every format; ${num.format(problems.length)} features still have problems. `
    + 'Images, TIFF scans and CSV pass every test.';

  drawWorks(tests, works);
  drawResults(tests, pass, fail, rate);
  drawMatrix(tests, problems);
  drawPlan(plan, weeks, rate, pass, tests.length);
}

// ---------- What works ----------
function familyTiles(box, tests, field) {
  box.replaceChildren(...FAMILIES.map((family) => {
    const mine = tests.filter((t) => t.family === family);
    const n = mine.filter((t) => t.result === field).length;
    const d = el('div', 'f');
    d.append(el('div', 'k', family), el('div', 'v', num.format(n)),
      el('div', 'k', field === 'Pass' ? `of ${num.format(mine.length)} tests pass` : `${n === 1 ? 'fail' : 'fails'} of ${num.format(mine.length)} tests`));
    return d;
  }));
}

function drawWorks(tests, works) {
  familyTiles($('#works-fam'), tests, 'Pass');
  const q = $('#wq'), fam = $('#wfam');
  fam.append(...FAMILIES.map((f) => new Option(f, f)));
  const render = () => {
    const text = q.value.trim().toLowerCase();
    const rows = works.filter((w) => (!fam.value || w.family === fam.value) && (!text || `${w.family} ${w.feature}`.toLowerCase().includes(text)));
    $('#works-count').textContent = `${num.format(works.length)} features in total; showing ${num.format(rows.length)}.`;
    let last = null;
    $('#works-t tbody').replaceChildren(...rows.map((w) => {
      const tr = el('tr', w.family !== last ? 'first' : '');
      last = w.family;
      const chips = el('div', 'chips');
      chips.append(...w.formats.sort((a, b) => FORMATS.indexOf(a) - FORMATS.indexOf(b)).map((f) => el('span', 'chip Pass', f === 'Image' ? '.jpg' : `.${f.toLowerCase()}`)));
      const c3 = el('td');
      c3.append(chips);
      tr.append(el('td', 'famc', w.family), el('td', 'feat', w.feature), c3, el('td', 'n', num.format(w.tests)));
      return tr;
    }));
  };
  q.addEventListener('input', render);
  fam.addEventListener('change', render);
  render();
}

// ---------- Results ----------
function drawResults(tests, pass, fail, rate) {
  $('#k-tests').textContent = num.format(tests.length);
  $('#k-pass').textContent = num.format(pass);
  $('#k-fail').textContent = num.format(fail);
  $('#k-rate').textContent = `${rate}%`;

  const files = new Map();
  for (const t of tests) {
    const r = files.get(t.file) || { file: t.file, format: t.format, tests: 0, pass: 0, fail: 0 };
    r.tests++;
    if (t.result === 'Pass') r.pass++; else if (t.result === 'Fail') r.fail++;
    files.set(t.file, r);
  }
  const rows = [...files.values()].map((r) => ({ ...r, rate: Math.round((r.pass / r.tests) * 100) })).sort((a, b) => a.rate - b.rate || b.tests - a.tests);
  $('#by-file tbody').replaceChildren(...rows.map((r) => {
    const tr = el('tr');
    tr.append(el('td', null, r.file), el('td', null, r.format), el('td', 'n', num.format(r.tests)), el('td', 'n', num.format(r.pass)), el('td', 'n', num.format(r.fail)), rateCell(r.rate));
    return tr;
  }));

  // all tests: search, filters, sort
  const state = { q: '', res: '', fmt: '', sort: 'id', dir: 1 };
  const ffmt = $('#ffmt');
  ffmt.append(...FORMATS.map((f) => new Option(f === 'Image' ? '.jpg / .jpeg' : `.${f.toLowerCase()}`, f)));
  const idNum = (id) => Number(/(\d+)$/.exec(id)?.[1] ?? 0);
  const byFormat = (a, b) => FORMATS.indexOf(a.format) - FORMATS.indexOf(b.format);
  const cmp = {
    id: (a, b) => byFormat(a, b) || idNum(a.id) - idNum(b.id),
    format: byFormat,
    feature: (a, b) => a.feature.localeCompare(b.feature),
    result: (a, b) => a.result.localeCompare(b.result),
  };
  const render = () => {
    const list = tests.filter((t) => (!state.res || t.result === state.res) && (!state.fmt || t.format === state.fmt)
      && (!state.q || [t.id, t.feature, t.expected, t.seen, t.where, t.file].join(' ').toLowerCase().includes(state.q)))
      .sort((a, b) => state.dir * cmp[state.sort](a, b) || byFormat(a, b) || idNum(a.id) - idNum(b.id));
    document.querySelectorAll('#tests th.sortable').forEach((th) => th.setAttribute('aria-sort', th.dataset.k === state.sort ? (state.dir > 0 ? 'ascending' : 'descending') : 'none'));
    $('#count').textContent = `${num.format(list.length)} of ${num.format(tests.length)} tests`;
    $('#tests tbody').replaceChildren(...list.map((t) => {
      const tr = el('tr');
      const res = el('td');
      res.append(el('span', `chip ${t.result}`, t.result));
      tr.append(el('td', null, t.id), el('td', null, t.format), el('td', null, t.where), el('td', 'feat', t.feature), el('td', 'muted', t.expected), res, el('td', 'muted', t.seen));
      return tr;
    }));
  };
  $('#q').addEventListener('input', (e) => { state.q = e.target.value.trim().toLowerCase(); render(); });
  $('#fres').addEventListener('change', (e) => { state.res = e.target.value; render(); });
  ffmt.addEventListener('change', (e) => { state.fmt = e.target.value; render(); });
  document.querySelectorAll('#tests th.sortable').forEach((th) => {
    const sort = () => { state.dir = state.sort === th.dataset.k ? -state.dir : 1; state.sort = th.dataset.k; render(); };
    th.addEventListener('click', sort);
    th.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sort(); } });
  });
  render();
}

// ---------- Problems matrix ----------
function drawMatrix(tests, problems) {
  familyTiles($('#fam'), tests, 'Fail');
  const detail = $('#mx-detail');
  let last = null;
  $('#mx tbody').replaceChildren(...problems.map((p) => {
    const tr = el('tr', p.family !== last ? 'first' : '');
    last = p.family;
    tr.append(el('td', 'famc', p.family), el('td', 'feat', p.feature));
    for (const f of MX_FORMATS) {
      const td = el('td', 'cell');
      const t = p.byFormat[f];
      if (t) {
        const b = el('button', t.result, t.result);
        b.type = 'button';
        b.setAttribute('aria-pressed', 'false');
        b.setAttribute('aria-label', `${p.feature}, .${f.toLowerCase()}: ${t.result}`);
        b.addEventListener('click', () => {
          document.querySelectorAll('#mx button[aria-pressed="true"]').forEach((x) => x.setAttribute('aria-pressed', 'false'));
          b.setAttribute('aria-pressed', 'true');
          const head = el('b', null, `${t.id} · ${p.feature} in .${f.toLowerCase()}: ${t.result}`);
          detail.replaceChildren(head, el('div', null, `Expected: ${t.expected}`), el('div', null, `Viewer showed: ${t.seen}`));
        });
        td.append(b);
      }
      tr.append(td);
    }
    tr.append(el('td', 'n', num.format(p.fails)));
    return tr;
  }));
}

// ---------- Fix plan ----------
function drawPlan(plan, weeks, rate, pass, total) {
  // pass rate after each step = today's passes plus the tests fixed so far, so the plan always adds up
  const cumulative = (rows, key) => {
    let fixed = 0;
    [...rows].sort((a, b) => a[key] - b[key]).forEach((r) => { fixed += r.tests_fixed; r.pass_after = Math.round(((pass + fixed) / total) * 100); });
  };
  cumulative(weeks, 'day');
  cumulative(plan, 'phase');
  const days = weeks.length;
  const wkCount = Math.max(...weeks.map((w) => w.week));
  const lastOf = (n) => weeks.filter((w) => w.week === n).sort((a, b) => b.day - a.day)[0];
  const fixed = weeks.reduce((t, w) => t + w.tests_fixed, 0);
  const devMin = plan.reduce((t, p) => t + p.days_min, 0);
  const devMax = plan.reduce((t, p) => t + p.days_max, 0);
  $('#p-weeks').textContent = `${wkCount} weeks`;
  $('#p-days').textContent = `${days} working days, review at the end of each week`;
  $('#p-wk1').textContent = `${lastOf(1).pass_after}%`;
  $('#p-today').textContent = `from ${rate}% today`;
  $('#p-wk2').textContent = `${lastOf(wkCount).pass_after}%`;
  $('#p-fixed').textContent = `about ${num.format(fixed)} failing tests fixed`;
  $('#p-dev').textContent = `${Math.ceil(devMin / 5)}–${Math.ceil(devMax / 5)} weeks`;
  $('#p-devdays').textContent = `${devMin}–${devMax} working days for the same work`;

  let wk = null;
  $('#two-weeks tbody').replaceChildren(...[...weeks].sort((a, b) => a.day - b.day).map((w) => {
    const tr = el('tr', w.week !== wk ? 'first' : '');
    wk = w.week;
    tr.append(el('td', 'n', String(w.day)), el('td', 'n', String(w.week)), el('td', 'feat', w.title), el('td', 'muted', w.work), el('td', 'n', num.format(w.tests_fixed)), rateCell(w.pass_after));
    return tr;
  }));

  // developer timeline: phases one after another; dark = shortest, light = longest estimate
  const phases = [...plan].sort((a, b) => a.phase - b.phase);
  const spanWeeks = Math.ceil(devMax / 5);
  const span = spanWeeks * 5;
  const pct = (d) => `${(d / span) * 100}%`;
  let start = 0;
  const rows = phases.map((p) => {
    const row = el('div', 'g-row');
    const track = el('div', 'g-track');
    const grid = el('div', 'g-grid');
    for (let i = 0; i < spanWeeks; i++) grid.append(el('span'));
    const max = el('div', 'g-max');
    max.style.left = pct(start);
    max.style.width = pct(p.days_max);
    const min = el('div', 'g-min');
    min.style.left = pct(start);
    min.style.width = pct(p.days_min);
    const lbl = el('div', 'g-lbl', `${p.days_min}–${p.days_max} days · ${p.pass_after}% after`);
    const end = start + p.days_max;
    if (end / span < 0.6) lbl.style.left = `calc(${pct(end)} + 6px)`;
    else lbl.style.right = `calc(${(1 - start / span) * 100}% + 6px)`;
    track.append(grid, max, min, lbl);
    row.append(el('div', 'g-name', `${p.phase}. ${p.name}`), track);
    start = end;
    return row;
  });
  const axis = el('div', 'g-axis');
  const ticks = el('div', 'ticks');
  const step = spanWeeks > 8 ? 2 : 1;
  for (let w = 0; w <= spanWeeks; w += step) {
    const t = el('span', null, w === 0 ? 'Start' : `Wk ${w}`);
    t.style.left = `${(w / spanWeeks) * 100}%`;
    if (w === 0) t.style.transform = 'none';
    if (w === spanWeeks) t.style.transform = 'translateX(-100%)';
    ticks.append(t);
  }
  axis.append(el('div'), ticks);
  $('#gantt').replaceChildren(...rows, axis);

  $('#phases tbody').replaceChildren(...phases.map((p) => {
    const tr = el('tr');
    tr.append(el('td', 'n', String(p.phase)), el('td', 'feat', p.name), el('td', 'muted', p.covers), el('td', 'n', num.format(p.tests_fixed)),
      el('td', 'n', String(p.days_min)), el('td', 'n', String(p.days_max)), rateCell(p.pass_after));
    return tr;
  }));
}
