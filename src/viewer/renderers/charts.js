/*
 * Office charts (DrawingML chart parts, c:chartSpace) drawn as SVG from the values the file
 * stores with the chart (numCache / strCache). Column, bar, line, area, pie and doughnut, with
 * title, value axis and gridlines, category labels, legend and data labels.
 */
import { svgEl } from './shapes.js';

const kids = (node, name) => (node ? [...node.children].filter((c) => !name || c.localName === name) : []);
const child = (node, ...path) => {
  let cur = node;
  for (const name of path) cur = cur && [...cur.children].find((c) => c.localName === name);
  return cur || null;
};
const val = (node) => node?.getAttribute('val');
const all = (node, name) => (node ? [...node.getElementsByTagName('*')].filter((n) => n.localName === name) : []);

const TYPES = {
  barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area',
  pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'line', radarChart: 'line',
};

/** Cached values of a c:cat / c:val / c:tx reference, in point order. */
function points(ref) {
  const cache = all(ref, 'strCache')[0] || all(ref, 'numCache')[0];
  if (cache) {
    const count = Number(val(child(cache, 'ptCount')) || 0);
    const out = new Array(count).fill(null);
    for (const pt of kids(cache, 'pt')) out[Number(pt.getAttribute('idx'))] = child(pt, 'v')?.textContent ?? null;
    return out;
  }
  const lit = child(ref, 'v');
  return lit ? [lit.textContent] : [];
}

const textOf = (node) => all(node, 't').map((t) => t.textContent).join('');

/**
 * Reads a chart part.
 * @param {Document} doc chart XML
 * @param {(series: number) => string} accent default color for series / slice n
 * @param {(node: Element) => string|null} colorOf color from a fill element (solidFill …)
 */
export function readChart(doc, accent, colorOf) {
  const chart = child(doc.documentElement, 'chart');
  const plot = child(chart, 'plotArea');
  if (!plot) return null;
  const groups = kids(plot).filter((n) => TYPES[n.localName]).map((g) => {
    const type = TYPES[g.localName];
    const series = kids(g, 'ser').map((s, i) => {
      const idx = Number(val(child(s, 'idx')) ?? i);
      const fill = child(s, 'spPr', 'solidFill') || child(s, 'spPr', 'ln', 'solidFill');
      return {
        name: textOf(child(s, 'tx')) || points(child(s, 'tx'))[0] || `Series ${idx + 1}`,
        cats: points(child(s, 'cat') || child(s, 'xVal')),
        values: points(child(s, 'val') || child(s, 'yVal')).map((v) => (v == null || v === '' ? null : Number(v))),
        color: (fill && colorOf(fill)) || accent(idx),
        // per-point colors (pie slices, "vary colors by point")
        pointColors: Object.fromEntries(kids(s, 'dPt').map((d) => [Number(val(child(d, 'idx'))), colorOf(child(d, 'spPr', 'solidFill'))]).filter(([, c]) => c)),
        labels: child(s, 'dLbls') || child(g, 'dLbls'),
      };
    });
    return {
      type, series,
      horizontal: val(child(g, 'barDir')) === 'bar',
      grouping: val(child(g, 'grouping')) || 'clustered',
      varyColors: val(child(g, 'varyColors')) === '1',
      hole: type === 'doughnut' ? Number(val(child(g, 'holeSize')) || 50) / 100 : 0,
      labels: child(g, 'dLbls'),
    };
  });
  if (!groups.length) return null;
  const titleNode = child(chart, 'title');
  let title = titleNode ? textOf(titleNode) : '';
  // Office shows a one-series chart's series name as its title unless the title was deleted.
  const single = groups.length === 1 && groups[0].series.length === 1;
  if (!title && (titleNode || (single && val(child(chart, 'autoTitleDeleted')) !== '1' && groups[0].type !== 'bar'))) title = groups[0].series[0]?.name || '';
  const legend = child(chart, 'legend');
  const size = Number(all(doc.documentElement, 'defRPr').map((d) => d.getAttribute('sz')).find(Boolean) || 1000) / 100;
  return {
    accent, title, groups, legend: legend ? (val(child(legend, 'legendPos')) || 'r') : null, fontPt: size,
    valueAxisDeleted: val(child(plot, 'valAx', 'delete')) === '1',
    gridlines: !!child(plot, 'valAx', 'majorGridlines'),
    numFormat: child(plot, 'valAx', 'numFmt')?.getAttribute('formatCode') || 'General',
  };
}

const niceStep = (span, ticks = 5) => {
  const raw = span / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw || 1));
  return [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => span / s <= ticks) || mag * 10;
};
const fmt = (v, pct) => (pct ? `${Math.round(v * 100)}%` : Number.isInteger(v) ? v.toLocaleString() : (+v.toFixed(2)).toLocaleString());

/** Draws a chart read by readChart into an SVG of w x h px. */
export function drawChart(model, w, h, ptToPx = 96 / 72) {
  const svg = svgEl('svg', { class: 'dv-chart', viewBox: `0 0 ${w} ${h}`, width: w, height: h, role: 'img' });
  const font = Math.max(9, Math.min(model.fontPt * ptToPx, h / 12));
  const text = (x, y, s, attrs = {}) => {
    const t = svgEl('text', { x, y, 'font-size': font, fill: '#595959', 'font-family': 'Calibri, Carlito, "Segoe UI", Arial, sans-serif', ...attrs });
    t.textContent = s;
    svg.append(t);
    return t;
  };
  svg.setAttribute('aria-label', model.title || 'Chart');
  let top = 8; let bottom = h - 8; let left = 8; let right = w - 8;
  if (model.title) { text(w / 2, top + font * 1.1, model.title, { 'text-anchor': 'middle', 'font-size': font * 1.2, fill: '#404040' }); top += font * 1.8; }

  const pie = model.groups.find((g) => g.type === 'pie' || g.type === 'doughnut');
  // legend entries: slices for pie charts, series otherwise
  const entries = pie
    ? pie.series[0].cats.map((c, i) => ({ name: c ?? String(i + 1), color: pie.series[0].pointColors[i] || model.accent(i) }))
    : model.groups.flatMap((g) => g.series.map((s) => ({ name: s.name, color: s.color })));
  if (model.legend && entries.length) {
    const sw = font * 0.75;
    if (model.legend === 'b' || model.legend === 't') {
      const widths = entries.map((e) => sw + 6 + String(e.name).length * font * 0.55 + 16);
      const total = widths.reduce((a, b) => a + b, 0);
      let x = Math.max(left, (w - total) / 2);
      const y = model.legend === 'b' ? bottom - font * 0.4 : top + font;
      entries.forEach((e, i) => {
        svg.append(svgEl('rect', { x, y: y - sw, width: sw, height: sw, fill: e.color }));
        text(x + sw + 6, y, e.name);
        x += widths[i];
      });
      if (model.legend === 'b') bottom -= font * 1.8; else top += font * 1.8;
    } else {
      const longest = Math.max(...entries.map((e) => String(e.name).length));
      const lw = sw + 10 + longest * font * 0.55;
      const x = model.legend === 'l' ? left : right - lw;
      let y = (top + bottom) / 2 - (entries.length * font * 1.4) / 2 + font;
      for (const e of entries) {
        svg.append(svgEl('rect', { x, y: y - sw, width: sw, height: sw, fill: e.color }));
        text(x + sw + 6, y, e.name);
        y += font * 1.4;
      }
      if (model.legend === 'l') left += lw + 8; else right -= lw + 8;
    }
  }

  if (pie) {
    const s = pie.series[0];
    const values = s.values.map((v) => Math.max(0, v || 0));
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const r = Math.max(10, Math.min(right - left, bottom - top) / 2 - 4);
    const cx = (left + right) / 2; const cy = (top + bottom) / 2;
    const showPct = pie.labels && child(pie.labels, 'showPercent') && val(child(pie.labels, 'showPercent')) === '1';
    const showVal = pie.labels && val(child(pie.labels, 'showVal')) === '1';
    const showCat = pie.labels && val(child(pie.labels, 'showCatName')) === '1';
    let a = -Math.PI / 2; // Office starts at 12 o'clock and goes clockwise
    values.forEach((v, i) => {
      const da = (v / total) * Math.PI * 2;
      const color = entries[i]?.color;
      const large = da > Math.PI ? 1 : 0;
      const p = (ang, rr) => [cx + rr * Math.cos(ang), cy + rr * Math.sin(ang)];
      const [x1, y1] = p(a, r); const [x2, y2] = p(a + da, r);
      let d;
      if (pie.hole) {
        const ri = r * pie.hole;
        const [x3, y3] = p(a + da, ri); const [x4, y4] = p(a, ri);
        d = `M${x1} ${y1}A${r} ${r} 0 ${large} 1 ${x2} ${y2}L${x3} ${y3}A${ri} ${ri} 0 ${large} 0 ${x4} ${y4}Z`;
      } else d = values.filter((x) => x > 0).length === 1 && v > 0 ? `M${cx - r} ${cy}a${r} ${r} 0 1 0 ${2 * r} 0a${r} ${r} 0 1 0 ${-2 * r} 0` : `M${cx} ${cy}L${x1} ${y1}A${r} ${r} 0 ${large} 1 ${x2} ${y2}Z`;
      svg.append(svgEl('path', { d, fill: color, stroke: '#ffffff', 'stroke-width': 1.5 }));
      if ((showPct || showVal || showCat) && v > 0) {
        const [lx, ly] = p(a + da / 2, pie.hole ? r * (1 + pie.hole) / 2 : r * 0.62);
        const parts = [showCat && s.cats[i], showVal && fmt(v), showPct && `${Math.round((v / total) * 100)}%`].filter(Boolean);
        text(lx, ly + font * 0.35, parts.join('; '), { 'text-anchor': 'middle', fill: '#ffffff', 'font-weight': '600' });
      }
      a += da;
    });
    return svg;
  }

  // axis charts: categories from the first series
  const cats = model.groups[0].series[0]?.cats || [];
  const n = Math.max(cats.length, ...model.groups.flatMap((g) => g.series.map((s) => s.values.length)));
  const stacked = (g) => g.grouping === 'stacked' || g.grouping === 'percentStacked';
  let lo = 0; let hi = 0;
  for (const g of model.groups) {
    for (let i = 0; i < n; i++) {
      if (stacked(g)) {
        let pos = 0; let neg = 0;
        for (const s of g.series) { const v = s.values[i] || 0; if (v >= 0) pos += v; else neg += v; }
        if (g.grouping === 'percentStacked') { hi = Math.max(hi, 1); lo = Math.min(lo, neg < 0 ? -1 : 0); } else { hi = Math.max(hi, pos); lo = Math.min(lo, neg); }
      } else for (const s of g.series) { const v = s.values[i]; if (v != null) { hi = Math.max(hi, v); lo = Math.min(lo, v); } }
    }
  }
  const pct = model.groups.some((g) => g.grouping === 'percentStacked');
  const step = niceStep(hi - lo || 1);
  hi = Math.ceil(hi / step) * step; lo = Math.floor(lo / step) * step;
  if (hi === lo) hi = lo + step;
  const horizontal = model.groups[0].horizontal && model.groups[0].type === 'bar';
  // room for axis labels
  const labelW = model.valueAxisDeleted ? 0 : Math.max(...[lo, hi].map((v) => fmt(v, pct).length)) * font * 0.6 + 8;
  const catW = horizontal ? Math.max(...cats.map((c) => String(c ?? '').length)) * font * 0.55 + 8 : 0;
  const plot = horizontal
    ? { x: left + catW, y: top, w: right - left - catW, h: bottom - top - (model.valueAxisDeleted ? 0 : font * 1.6) }
    : { x: left + labelW, y: top, w: right - left - labelW, h: bottom - top - font * 1.8 };
  const scale = (v) => (v - lo) / (hi - lo);
  const grid = svgEl('g');
  svg.append(grid);
  for (let v = lo; v <= hi + step / 2; v += step) {
    if (horizontal) {
      const x = plot.x + scale(v) * plot.w;
      if (model.gridlines) grid.append(svgEl('line', { x1: x, x2: x, y1: plot.y, y2: plot.y + plot.h, stroke: '#d9d9d9' }));
      if (!model.valueAxisDeleted) text(x, plot.y + plot.h + font * 1.3, fmt(v, pct), { 'text-anchor': 'middle' });
    } else {
      const y = plot.y + plot.h - scale(v) * plot.h;
      if (model.gridlines) grid.append(svgEl('line', { x1: plot.x, x2: plot.x + plot.w, y1: y, y2: y, stroke: '#d9d9d9' }));
      if (!model.valueAxisDeleted) text(plot.x - 6, y + font * 0.35, fmt(v, pct), { 'text-anchor': 'end' });
    }
  }
  const band = (horizontal ? plot.h : plot.w) / Math.max(n, 1);
  const zero = scale(Math.max(lo, Math.min(0, hi)));
  // category axis line and labels
  if (horizontal) {
    const x0 = plot.x + zero * plot.w;
    grid.append(svgEl('line', { x1: x0, x2: x0, y1: plot.y, y2: plot.y + plot.h, stroke: '#bfbfbf' }));
    cats.forEach((c, i) => text(plot.x - 6, plot.y + plot.h - (i + 0.5) * band + font * 0.35, String(c ?? ''), { 'text-anchor': 'end' }));
  } else {
    const y0 = plot.y + plot.h - zero * plot.h;
    grid.append(svgEl('line', { x1: plot.x, x2: plot.x + plot.w, y1: y0, y2: y0, stroke: '#bfbfbf' }));
    cats.forEach((c, i) => text(plot.x + (i + 0.5) * band, plot.y + plot.h + font * 1.3, String(c ?? ''), { 'text-anchor': 'middle' }));
  }

  for (const g of model.groups) {
    if (g.type === 'bar') {
      const k = stacked(g) ? 1 : g.series.length;
      const gap = band * 0.25; // Office's default gap width is 150% of a bar; this is close at small counts
      const bw = (band - gap * 2) / k;
      for (let i = 0; i < n; i++) {
        let pos = 0; let neg = 0;
        const sum = g.series.reduce((t, s) => t + Math.abs(s.values[i] || 0), 0) || 1;
        g.series.forEach((s, j) => {
          let v = s.values[i];
          if (v == null) return;
          if (g.grouping === 'percentStacked') v /= sum;
          let a0 = 0; let a1 = v;
          if (stacked(g)) { if (v >= 0) { a0 = pos; a1 = pos + v; pos += v; } else { a0 = neg; a1 = neg + v; neg += v; } }
          const s0 = scale(Math.min(a0, a1)); const s1 = scale(Math.max(a0, a1));
          const off = gap + (stacked(g) ? 0 : j * bw);
          const color = (g.varyColors && g.series.length === 1 && (s.pointColors[i] || model.accent?.(i))) || s.color;
          const rect = horizontal
            ? { x: plot.x + s0 * plot.w, y: plot.y + plot.h - (i + 1) * band + off, width: Math.max((s1 - s0) * plot.w, 0.5), height: bw }
            : { x: plot.x + i * band + off, y: plot.y + plot.h - s1 * plot.h, width: bw, height: Math.max((s1 - s0) * plot.h, 0.5) };
          svg.append(svgEl('rect', { ...rect, fill: color }));
          if (s.labels && val(child(s.labels, 'showVal')) === '1') {
            text(horizontal ? rect.x + rect.width + 4 : rect.x + rect.width / 2, horizontal ? rect.y + rect.height / 2 + font * 0.35 : rect.y - 4, fmt(s.values[i]), { 'text-anchor': horizontal ? 'start' : 'middle' });
          }
        });
      }
    } else {
      // line / area: points at band centers
      const base = new Array(n).fill(0);
      for (const s of g.series) {
        const pts = [];
        for (let i = 0; i < n; i++) {
          const v = s.values[i];
          if (v == null) continue;
          const y0 = stacked(g) ? base[i] : 0;
          const y = stacked(g) ? y0 + v : v;
          if (stacked(g)) base[i] = y;
          pts.push([plot.x + (i + 0.5) * band, plot.y + plot.h - scale(y) * plot.h, plot.y + plot.h - scale(y0) * plot.h]);
        }
        if (!pts.length) continue;
        if (g.type === 'area') {
          const d = `M${pts.map((p) => `${p[0]} ${p[1]}`).join('L')}L${pts.map((p) => `${p[0]} ${p[2]}`).reverse().join('L')}Z`;
          svg.append(svgEl('path', { d, fill: s.color, 'fill-opacity': 0.85 }));
        } else {
          svg.append(svgEl('polyline', { points: pts.map((p) => `${p[0]},${p[1]}`).join(' '), fill: 'none', stroke: s.color, 'stroke-width': 2.25, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
          for (const p of pts) svg.append(svgEl('circle', { cx: p[0], cy: p[1], r: 3, fill: s.color }));
        }
      }
    }
  }
  return svg;
}

const ODF_TYPES = { bar: 'bar', line: 'line', area: 'area', circle: 'pie', ring: 'doughnut', scatter: 'line', radar: 'line' };

/**
 * A chart in OpenDocument form (content.xml of a chart object, as LibreOffice embeds it):
 * chart type, title, legend, series colors and the data table, read into the same model as readChart.
 */
export function readOdfChart(xml, accent) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const byLocal = (name, root = doc) => [...root.getElementsByTagName('*')].filter((n) => n.localName === name);
  const a = (node, name) => node && [...node.attributes].find((x) => x.localName === name)?.value;
  const chart = byLocal('chart').find((n) => n.namespaceURI?.includes('chart'));
  if (!chart) return null;
  const styles = new Map(byLocal('style').map((st) => [a(st, 'name'), st]));
  const styleProp = (name, prop) => {
    const st = styles.get(name);
    if (!st) return null;
    for (const p of [...st.children]) { const v = a(p, prop); if (v) return v; }
    return null;
  };
  const type = ODF_TYPES[(a(chart, 'class') || 'chart:bar').split(':').pop()] || 'bar';
  const plot = byLocal('plot-area')[0];
  const plotStyle = a(plot, 'style-name');
  const horizontal = styleProp(plotStyle, 'vertical') === 'true';
  const grouping = styleProp(plotStyle, 'percentage') === 'true' ? 'percentStacked' : styleProp(plotStyle, 'stacked') === 'true' ? 'stacked' : 'clustered';
  // data table: header row = series names, then category + values per row
  const table = byLocal('table')[0];
  const rows = table ? byLocal('table-row', table) : [];
  const cellsOf = (row) => byLocal('table-cell', row).map((c) => ({ text: byLocal('p', c).map((p) => p.textContent).join(' '), value: a(c, 'value') }));
  const header = rows.length ? cellsOf(rows[0]) : [];
  const body = rows.slice(1).map(cellsOf);
  const cats = body.map((r) => r[0]?.text ?? '');
  const series = byLocal('series').map((se, i) => {
    const col = i + 1;
    const pointColors = {};
    let k = 0;
    for (const dp of byLocal('data-point', se)) {
      const n = Number(a(dp, 'repeated') || 1);
      const c = styleProp(a(dp, 'style-name'), 'fill-color');
      for (let j = 0; j < n; j++, k++) if (c) pointColors[k] = c;
    }
    return {
      name: header[col]?.text || `Series ${col}`,
      cats,
      values: body.map((r) => (r[col]?.value != null ? Number(r[col].value) : r[col]?.text ? Number(r[col].text) : null)),
      color: styleProp(a(se, 'style-name'), 'fill-color') || styleProp(a(se, 'style-name'), 'stroke-color') || accent(i),
      pointColors,
      labels: null,
    };
  });
  if (!series.length) return null;
  const legend = byLocal('legend')[0];
  const pos = { bottom: 'b', top: 't', start: 'l', end: 'r', left: 'l', right: 'r' }[a(legend, 'legend-position')] || 'r';
  const title = byLocal('title')[0];
  const size = Number(/([\d.]+)pt/.exec(styleProp(a(legend, 'style-name'), 'font-size') || '')?.[1] || 18);
  const valueAxis = byLocal('axis').find((ax) => a(ax, 'dimension') === 'y');
  return {
    accent,
    title: title ? byLocal('p', title).map((p) => p.textContent).join(' ') : '',
    groups: [{ type, series, horizontal, grouping, varyColors: type === 'pie' || type === 'doughnut', hole: type === 'doughnut' ? 0.5 : 0, labels: null }],
    legend: legend ? pos : null,
    fontPt: size,
    valueAxisDeleted: valueAxis ? styleProp(a(valueAxis, 'style-name'), 'visible') === 'false' : false,
    gridlines: !!(valueAxis && byLocal('grid', valueAxis).length),
  };
}
