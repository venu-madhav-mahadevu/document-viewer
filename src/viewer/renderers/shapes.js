/*
 * Outlines of Office's common preset shapes (arrows, stars, callouts, chevron, polygons…)
 * as SVG paths, shared by the .pptx and .ppt readers.
 */
export const SVG_NS = 'http://www.w3.org/2000/svg';
export const svgEl = (tag, attrs = {}) => {
  const e = document.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  return e;
};

/**
 * Outlines of PowerPoint's common preset shapes as SVG paths in a w x h box.
 * `adj(name, default)` reads the shape's adjustment handles (in 1/100000).
 */
export const PRESET_SHAPES = {
  triangle: (w, h, adj) => { const x = w * adj('adj', 50000) / 100000; return `M${x} 0L${w} ${h}L0 ${h}Z`; },
  rtTriangle: (w, h) => `M0 0L${w} ${h}L0 ${h}Z`,
  diamond: (w, h) => `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z`,
  parallelogram: (w, h, adj) => { const a = Math.min(w, h) * adj('adj', 25000) / 100000; return `M${a} 0L${w} 0L${w - a} ${h}L0 ${h}Z`; },
  trapezoid: (w, h, adj) => { const a = Math.min(w, h) * adj('adj', 25000) / 100000; return `M${a} 0L${w - a} 0L${w} ${h}L0 ${h}Z`; },
  pentagon: (w, h) => polygon(w, h, 5),
  hexagon: (w, h, adj) => { const a = Math.min(w, h) * adj('adj', 25000) / 100000; return `M${a} 0L${w - a} 0L${w} ${h / 2}L${w - a} ${h}L${a} ${h}L0 ${h / 2}Z`; },
  heptagon: (w, h) => polygon(w, h, 7),
  octagon: (w, h, adj) => {
    const a = Math.min(w, h) * adj('adj', 29289) / 100000;
    return `M${a} 0L${w - a} 0L${w} ${a}L${w} ${h - a}L${w - a} ${h}L${a} ${h}L0 ${h - a}L0 ${a}Z`;
  },
  plus: (w, h, adj) => {
    const a = Math.min(w, h) * adj('adj', 25000) / 100000;
    return `M${a} 0L${w - a} 0L${w - a} ${a}L${w} ${a}L${w} ${h - a}L${w - a} ${h - a}L${w - a} ${h}L${a} ${h}L${a} ${h - a}L0 ${h - a}L0 ${a}L${a} ${a}Z`;
  },
  rightArrow: (w, h, adj) => arrow(w, h, adj, 'r'),
  leftArrow: (w, h, adj) => arrow(w, h, adj, 'l'),
  upArrow: (w, h, adj) => arrow(w, h, adj, 'u'),
  downArrow: (w, h, adj) => arrow(w, h, adj, 'd'),
  leftRightArrow: (w, h, adj) => {
    const t = h * adj('adj1', 50000) / 100000; const head = Math.min(w, h) * adj('adj2', 50000) / 100000;
    const y1 = (h - t) / 2; const y2 = (h + t) / 2;
    return `M0 ${h / 2}L${head} 0L${head} ${y1}L${w - head} ${y1}L${w - head} 0L${w} ${h / 2}L${w - head} ${h}L${w - head} ${y2}L${head} ${y2}L${head} ${h}Z`;
  },
  upDownArrow: (w, h, adj) => {
    const t = w * adj('adj1', 50000) / 100000; const head = Math.min(w, h) * adj('adj2', 50000) / 100000;
    const x1 = (w - t) / 2; const x2 = (w + t) / 2;
    return `M${w / 2} 0L${w} ${head}L${x2} ${head}L${x2} ${h - head}L${w} ${h - head}L${w / 2} ${h}L0 ${h - head}L${x1} ${h - head}L${x1} ${head}L0 ${head}Z`;
  },
  chevron: (w, h, adj) => { const a = Math.min(w, h) * adj('adj', 50000) / 100000; return `M0 0L${w - a} 0L${w} ${h / 2}L${w - a} ${h}L0 ${h}L${a} ${h / 2}Z`; },
  homePlate: (w, h, adj) => { const a = Math.min(w, h) * adj('adj', 50000) / 100000; return `M0 0L${w - a} 0L${w} ${h / 2}L${w - a} ${h}L0 ${h}Z`; },
  star4: (w, h, adj) => star(w, h, 4, adj('adj', 12500) / 50000),
  star5: (w, h, adj) => star(w, h, 5, adj('adj', 19098) / 50000),
  star6: (w, h, adj) => star(w, h, 6, adj('adj', 28868) / 50000),
  star8: (w, h, adj) => star(w, h, 8, adj('adj', 38250) / 50000),
  wedgeRectCallout: (w, h, adj) => callout(w, h, adj, 0),
  wedgeRoundRectCallout: (w, h, adj) => callout(w, h, adj, Math.min(w, h) * 0.1667),
  wedgeEllipseCallout: (w, h, adj) => callout(w, h, adj, -1),
  flowChartDecision: (w, h) => `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z`,
  flowChartTerminator: (w, h) => { const r = Math.min(w / 2, h / 2); return `M${r} 0L${w - r} 0A${r} ${h / 2} 0 0 1 ${w - r} ${h}L${r} ${h}A${r} ${h / 2} 0 0 1 ${r} 0Z`; },
  flowChartDocument: (w, h) => `M0 0L${w} 0L${w} ${h * 0.83}C${w * 0.75} ${h * 0.7} ${w * 0.5} ${h} ${w * 0.25} ${h}C${w * 0.12} ${h} 0 ${h * 0.95} 0 ${h * 0.9}Z`,
  heart: (w, h) => `M${w / 2} ${h * 0.25}C${w / 2} ${-h * 0.08} 0 ${-h * 0.08} 0 ${h * 0.3}C0 ${h * 0.6} ${w * 0.3} ${h * 0.78} ${w / 2} ${h}C${w * 0.7} ${h * 0.78} ${w} ${h * 0.6} ${w} ${h * 0.3}C${w} ${-h * 0.08} ${w / 2} ${-h * 0.08} ${w / 2} ${h * 0.25}Z`,
};
PRESET_SHAPES.flowChartPreparation = PRESET_SHAPES.hexagon;
PRESET_SHAPES.flowChartInputOutput = (w, h) => PRESET_SHAPES.parallelogram(w, h, (n, d) => 20000);

function polygon(w, h, n) {
  const pts = [];
  for (let i = 0; i < n; i++) {
    const a = -Math.PI / 2 + (i * 2 * Math.PI) / n;
    pts.push(`${w / 2 + (w / 2) * Math.cos(a)} ${h / 2 + (h / 2) * Math.sin(a)}`);
  }
  return `M${pts.join('L')}Z`;
}
function star(w, h, n, inner) {
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const r = i % 2 ? inner : 1;
    const a = -Math.PI / 2 + (i * Math.PI) / n;
    pts.push(`${w / 2 + (w / 2) * r * Math.cos(a)} ${h / 2 + (h / 2) * r * Math.sin(a)}`);
  }
  return `M${pts.join('L')}Z`;
}
function arrow(w, h, adj, dir) {
  // adj1 = shaft thickness (share of the arrow's width), adj2 = head length (share of the shorter side)
  const horizontal = dir === 'r' || dir === 'l';
  const [L, T] = horizontal ? [w, h] : [h, w];
  const t = T * adj('adj1', 50000) / 100000;
  const head = Math.min(w, h) * adj('adj2', 50000) / 100000;
  const s1 = (T - t) / 2; const s2 = (T + t) / 2;
  // drawn pointing right in (along, across) coordinates, then mapped to the direction
  const pts = [[0, s1], [L - head, s1], [L - head, 0], [L, T / 2], [L - head, T], [L - head, s2], [0, s2]];
  const map = { r: ([a, c]) => [a, c], l: ([a, c]) => [w - a, c], d: ([a, c]) => [c, a], u: ([a, c]) => [c, h - a] }[dir];
  return `M${pts.map(map).map((p) => p.join(' ')).join('L')}Z`;
}
function callout(w, h, adj, radius) {
  // the tail's tip, relative to the shape's center (in 1/100000 of the size)
  const tx = w / 2 + w * adj('adj1', -20833) / 100000;
  const ty = h / 2 + h * adj('adj2', 62500) / 100000;
  if (radius < 0) { // ellipse with a tail
    const a = Math.atan2(ty - h / 2, tx - w / 2);
    const p = (d) => [w / 2 + (w / 2) * Math.cos(a + d), h / 2 + (h / 2) * Math.sin(a + d)];
    const [p1, p2] = [p(-0.25), p(0.25)];
    return `M${p2[0]} ${p2[1]}A${w / 2} ${h / 2} 0 1 1 ${p1[0]} ${p1[1]}L${tx} ${ty}Z`;
  }
  const r = radius;
  // the tail leaves the side nearest its tip
  const side = ty > h ? 'b' : ty < 0 ? 't' : tx < 0 ? 'l' : 'r';
  const along = side === 'b' || side === 't' ? w : h;
  const at = (side === 'b' || side === 't' ? Math.min(Math.max(tx, w * 0.1), w * 0.9) : Math.min(Math.max(ty, h * 0.1), h * 0.9));
  const b1 = Math.max(r, at - along * 0.1); const b2 = Math.min(along - r, at + along * 0.1);
  const seg = {
    t: `L${b1} 0L${tx} ${ty}L${b2} 0`, r: `L${w} ${b1}L${tx} ${ty}L${w} ${b2}`,
    b: `L${b2} ${h}L${tx} ${ty}L${b1} ${h}`, l: `L0 ${b2}L${tx} ${ty}L0 ${b1}`,
  };
  const arc = (x, y) => (r ? `A${r} ${r} 0 0 1 ${x} ${y}` : `L${x} ${y}`);
  return `M${r} 0${side === 't' ? seg.t : ''}L${w - r} 0${arc(w, r)}${side === 'r' ? seg.r : ''}L${w} ${h - r}${arc(w - r, h)}`
    + `${side === 'b' ? seg.b : ''}L${r} ${h}${arc(0, h - r)}${side === 'l' ? seg.l : ''}L0 ${r}${arc(r, 0)}Z`;
}


/** PowerPoint 97-2003 shape type numbers ([MS-ODRAW] MSOSPT) and their preset names. */
export const MSOSPT_PRESET = {
  4: 'diamond', 5: 'triangle', 6: 'rtTriangle', 7: 'parallelogram', 8: 'trapezoid', 9: 'hexagon', 10: 'octagon',
  11: 'plus', 12: 'star5', 13: 'rightArrow', 15: 'homePlate', 55: 'chevron', 56: 'pentagon', 58: 'star8',
  61: 'wedgeRectCallout', 62: 'wedgeRoundRectCallout', 63: 'wedgeEllipseCallout', 66: 'leftArrow', 67: 'downArrow',
  68: 'upArrow', 69: 'leftRightArrow', 70: 'upDownArrow', 74: 'heart', 110: 'flowChartDecision', 111: 'flowChartInputOutput',
  114: 'flowChartDocument', 116: 'flowChartTerminator', 117: 'flowChartPreparation', 187: 'star4',
};
