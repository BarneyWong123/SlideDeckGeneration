"""Build the interactive customer-visit deck as one self-contained HTML file.

Every chart is real SVG drawn from the SAP export, and every mark carries the
invoice lines behind it, so hovering (or tab-focusing) a bar names the products
and the units that make it up.

    python3 scripts/build_interactive_deck.py <data.xlsx> -o customer_visit_deck.html

Design tokens are the Biomed Global brand system; the two chart series colours
(--s-prior / --s-current) are validated for colour-vision deficiency separation.
"""

import argparse
import base64
import json
import pathlib

import extract_deck_data as X

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent

VISIT_DATE = "8 September 2026"

SOURCE_NOTE = (
    "Source: SAP AR invoice export, 5 May 2025 – 4 Sep 2026. Financial year runs "
    "May–Apr, so FY27 YTD = May–Aug 2026; the part-month of September is excluded "
    "from every comparison. Counts are invoiced quantities, including lines invoiced at "
    "RM 0."
)


def b64(path):
    return base64.b64encode(pathlib.Path(path).read_bytes()).decode()


def n(v):
    return f"{v:,.0f}"


def pct(new, old):
    return (new / old - 1) * 100 if old else 0.0


def signed(v, digits=0):
    s = "+" if v > 0 else "−"
    return f"{s}{abs(v):,.{digits}f}%"


# --- page ------------------------------------------------------------------------
CSS = """
:root {
  --brand-dark: #003366;
  --brand-light: #78c8ef;
  --ink: #0f1b2b;
  --ink-body: #2b3138;
  --ink-muted: #5d666e;
  --rule: #e3e8ec;
  --rule-strong: #cfd6dc;
  --surface: #ffffff;
  --surface-sunken: #f4f8fb;
  /* Two chart series, validated for CVD separation against the slide surface. */
  --s-prior: #4fb0de;
  --s-current: #2e69a0;
  --good: #1f6b44;
  --bad: #8c2417;
  /* Chrome around the slide follows the viewer's theme; the slide itself does not,
     because a projected deck is a light artefact wherever it is opened. */
  --page: #eef2f5;
  --page-ink: #2b3138;
  --page-rule: #d6dee4;
  --page-chip: #ffffff;
}
@media (prefers-color-scheme: dark) {
  :root:not([data-theme="light"]) {
    --page: #0f1b2b;
    --page-ink: #c7d2dc;
    --page-rule: #223448;
    --page-chip: #16283c;
  }
}
:root[data-theme="dark"] {
  --page: #0f1b2b;
  --page-ink: #c7d2dc;
  --page-rule: #223448;
  --page-chip: #16283c;
}

* { box-sizing: border-box; }
body {
  margin: 0;
  background: var(--page);
  color: var(--page-ink);
  font-family: 'DM Sans', system-ui, -apple-system, 'Segoe UI', sans-serif;
  font-variant-numeric: tabular-nums;
}

/* --- stage: a 1920x1080 slide scaled to whatever room the window gives it ----- */
#app { min-height: 100vh; display: flex; flex-direction: column; }
#stage-wrap { flex: 1; display: grid; place-items: center; padding: 20px 20px 0; }
#stage {
  width: 1920px; height: 1080px;
  transform-origin: top left;
  background: var(--surface);
  color: var(--ink);
  box-shadow: 0 2px 4px rgba(15,27,43,.10), 0 18px 48px rgba(15,27,43,.16);
  position: relative; overflow: hidden;
}
#stage-box { position: relative; }
.slide { position: absolute; inset: 0; display: none; grid-template-rows: 152px 1fr 96px; }
.slide.on { display: grid; }

.head {
  background: var(--brand-dark); color: #fff; padding: 0 76px;
  display: flex; align-items: center; justify-content: space-between;
}
.head h1 { margin: 0; font-size: 46px; font-weight: 700; letter-spacing: -.01em; text-wrap: balance; }
.head .eyebrow {
  margin-top: 8px; font-size: 25px; font-weight: 500; color: var(--brand-light);
  letter-spacing: .06em; text-transform: uppercase;
}
.head img { height: 62px; width: auto; }

.body { padding: 40px 76px 0; display: flex; flex-direction: column; gap: 26px; min-height: 0; }
.foot {
  margin: 0 76px; border-top: 1px solid var(--rule);
  display: flex; align-items: center; justify-content: space-between;
  color: var(--ink-muted); font-size: 22px; gap: 40px;
}
.foot span:first-child { max-width: 1400px; }

/* --- pieces ------------------------------------------------------------------ */
.label {
  font-size: 24px; font-weight: 500; color: var(--ink-muted);
  letter-spacing: .04em; text-transform: uppercase;
}
.kpis { display: grid; grid-template-columns: repeat(3, 1fr); gap: 26px; }
.kpis.four { grid-template-columns: repeat(4, 1fr); }
.kpi { border: 1px solid var(--rule); border-radius: 14px; padding: 26px 30px;
       display: flex; flex-direction: column; gap: 8px; }
.kpi.tint { background: var(--surface-sunken); border-color: transparent; }
.kpi b { font-size: 60px; font-weight: 700; color: var(--brand-dark); line-height: 1; }
.kpi small { font-size: 24px; color: var(--ink-muted); font-weight: 400; }
.kpi small.bad { color: var(--bad); font-weight: 500; }
.kpi small.good { color: var(--good); font-weight: 500; }

.read { background: var(--surface-sunken); border-radius: 14px; padding: 30px 34px;
        display: flex; flex-direction: column; gap: 18px; }
.read p { margin: 0; font-size: 29px; line-height: 1.4; font-weight: 500; text-wrap: pretty; }
.read dl { margin: 0; border-top: 1px solid var(--rule-strong); padding-top: 18px;
           display: flex; flex-direction: column; gap: 14px; }
.read div { display: flex; justify-content: space-between; font-size: 23px; }
.read dt { color: var(--ink-muted); }
.read dd { margin: 0; font-weight: 700; }

.note { border: 1px solid var(--rule); border-radius: 14px; padding: 24px 28px;
        display: flex; flex-direction: column; gap: 10px; }
.note.dark { background: var(--brand-dark); border-color: transparent; color: #fff; }
.note.dark .label { color: var(--brand-light); }
.note p { margin: 0; font-size: 24px; line-height: 1.4; text-wrap: pretty; }

.asks { display: grid; grid-template-columns: 1fr 1fr; gap: 44px; }
.asks h2 { margin: 0 0 6px; font-size: 34px; color: var(--brand-dark); }
.ask { display: grid; grid-template-columns: 44px 1fr; gap: 18px; align-items: start;
       padding: 16px 0; border-top: 1px solid var(--rule); font-size: 25px; line-height: 1.4; }
.ask i { font-style: normal; font-weight: 700; color: var(--brand-light);
         font-size: 26px; line-height: 1.3; }
.leave { background: var(--surface-sunken); border-radius: 14px; padding: 22px 26px;
         display: flex; flex-direction: column; gap: 8px; margin-top: 12px; }
.leave p { margin: 0; font-size: 25px; font-weight: 500; line-height: 1.35; }

/* --- charts ------------------------------------------------------------------ */
.chart { display: flex; flex-direction: column; gap: 12px; min-height: 0; flex: 1; }
.chart-top { display: flex; align-items: baseline; justify-content: space-between; gap: 30px;
             flex: none; }
.plot { flex: 1; min-height: 0; }
.chart svg { display: block; width: 100%; height: 100%; overflow: visible; }
.chart > .legend, .chart > div:last-child { flex: none; }
.chart svg text { font-family: inherit; }
.legend { display: flex; gap: 30px; font-size: 23px; color: var(--ink-muted); font-weight: 500;
          align-items: center; flex-wrap: wrap; }
.legend span { display: inline-flex; align-items: center; gap: 10px; }
.legend i { width: 22px; height: 12px; border-radius: 3px; display: inline-block; }
.hint { font-size: 21px; color: var(--ink-muted); display: inline-flex; align-items: center; gap: 8px; }
.hint svg { width: 20px; height: 20px; }
.markwrap { outline: none; cursor: default; }
.markwrap .mark { transition: opacity .12s ease; }
.markwrap.hot .mark { opacity: .8; }
.markwrap .ring { fill: none; stroke: var(--brand-dark); stroke-width: 3;
                  opacity: 0; pointer-events: none; }
.markwrap.hot .ring, .markwrap:focus-visible .ring { opacity: 1; }
.markwrap:focus-visible .ring { stroke-width: 4; }

/* --- tooltip ----------------------------------------------------------------- */
#tip {
  position: fixed; z-index: 40; pointer-events: none; opacity: 0;
  transition: opacity .1s ease; max-width: 460px;
  background: var(--surface); color: var(--ink);
  border: 1px solid var(--rule-strong); border-radius: 12px;
  box-shadow: 0 10px 34px rgba(15,27,43,.22);
  padding: 14px 16px; font-size: 14px; line-height: 1.35;
}
#tip.on { opacity: 1; }
#tip .t-head { display: flex; align-items: baseline; gap: 10px; justify-content: space-between;
               padding-bottom: 8px; border-bottom: 1px solid var(--rule); }
#tip .t-name { font-weight: 700; font-size: 15px; }
#tip .t-val { font-weight: 700; font-size: 17px; color: var(--brand-dark); white-space: nowrap; }
#tip .t-sub { font-size: 12px; color: var(--ink-muted); letter-spacing: .04em;
              text-transform: uppercase; padding-top: 8px; }
#tip ul { list-style: none; margin: 8px 0 0; padding: 0;
          display: flex; flex-direction: column; gap: 5px; }
#tip li { display: grid; grid-template-columns: 3px 1fr auto; gap: 9px; align-items: baseline; }
#tip li em { background: currentColor; border-radius: 2px; height: 11px;
             align-self: center; font-style: normal; }
#tip li b { font-weight: 700; color: var(--ink); white-space: nowrap; }
#tip li span { color: var(--ink-body); }

/* --- data table (the tooltip's non-hover equivalent) ------------------------- */
.table-toggle {
  font: inherit; font-size: 21px; color: var(--brand-dark); background: none;
  border: 1px solid var(--rule-strong); border-radius: 999px; padding: 5px 16px;
  cursor: pointer;
}
.table-toggle:hover { background: var(--surface-sunken); }
.table-wrap { overflow: auto; max-height: 340px; }
table { border-collapse: collapse; width: 100%; font-size: 22px; }
th, td { text-align: right; padding: 8px 14px; border-bottom: 1px solid var(--rule); }
th:first-child, td:first-child { text-align: left; }
thead th { color: var(--ink-muted); font-weight: 500; font-size: 20px;
           text-transform: uppercase; letter-spacing: .04em; }

/* --- deck chrome ------------------------------------------------------------- */
#bar {
  display: flex; align-items: center; gap: 14px; padding: 14px 20px 18px;
  justify-content: center; flex-wrap: wrap;
}
#bar button {
  font: inherit; font-size: 14px; color: var(--page-ink); background: var(--page-chip);
  border: 1px solid var(--page-rule); border-radius: 8px; padding: 7px 13px; cursor: pointer;
}
#bar button:hover { border-color: var(--brand-light); }
#bar button[aria-current="true"] { background: var(--brand-dark); border-color: var(--brand-dark);
                                   color: #fff; }
#bar .count { font-size: 13px; color: var(--page-ink); opacity: .7; padding: 0 6px; }
#notes { max-width: 900px; margin: 0 auto 24px; padding: 0 20px; font-size: 14px;
         line-height: 1.5; color: var(--page-ink); opacity: .78; text-align: center; }
@media (prefers-reduced-motion: reduce) { * { transition: none !important; } }
"""

JS = r"""
const D = window.__DECK__;
const FMT = new Intl.NumberFormat('en-US');
const SVGNS = 'http://www.w3.org/2000/svg';

function el(tag, attrs, kids) {
  const n = document.createElementNS(SVGNS, tag);
  for (const k in (attrs || {})) n.setAttribute(k, attrs[k]);
  (kids || []).forEach(c => n.appendChild(c));
  return n;
}
function txt(s) { return document.createTextNode(String(s)); }
function h(tag, cls, text) {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text !== undefined) n.textContent = text;   // labels are data: never innerHTML
  return n;
}

/* ---- tooltip ---------------------------------------------------------------- */
const tip = document.getElementById('tip');
function showTip(ev, spec) {
  tip.replaceChildren();
  const head = h('div', 't-head');
  head.appendChild(h('div', 't-name', spec.name));
  head.appendChild(h('div', 't-val', FMT.format(Math.round(spec.value)) + ' ' + spec.unit));
  tip.appendChild(head);
  if (spec.sub) tip.appendChild(h('div', 't-sub', spec.sub));
  if (spec.items && spec.items.length) {
    const ul = document.createElement('ul');
    spec.items.forEach(it => {
      const li = document.createElement('li');
      li.style.color = spec.color;
      li.appendChild(h('em'));
      li.appendChild(h('span', null, it.name));
      li.appendChild(h('b', null, FMT.format(Math.round(it.units))));
      ul.appendChild(li);
    });
    tip.appendChild(ul);
  }
  tip.classList.add('on');
  moveTip(ev);
}
function moveTip(ev) {
  const r = tip.getBoundingClientRect();
  let x = (ev.clientX || 0) + 16, y = (ev.clientY || 0) + 16;
  if (x + r.width > innerWidth - 8) x = (ev.clientX || 0) - r.width - 16;
  if (y + r.height > innerHeight - 8) y = (ev.clientY || 0) - r.height - 16;
  tip.style.left = Math.max(8, x) + 'px';
  tip.style.top = Math.max(8, y) + 'px';
}
function hideTip() { tip.classList.remove('on'); }

function wire(g, spec) {
  g.setAttribute('tabindex', '0');
  g.setAttribute('role', 'img');
  g.setAttribute('aria-label',
    spec.name + ': ' + FMT.format(Math.round(spec.value)) + ' ' + spec.unit);
  const on = ev => { g.classList.add('hot'); showTip(ev, spec); };
  g.addEventListener('pointerenter', on);
  g.addEventListener('pointermove', moveTip);
  g.addEventListener('pointerleave', () => { g.classList.remove('hot'); hideTip(); });
  g.addEventListener('focus', () => {
    g.classList.add('hot');
    const b = g.getBoundingClientRect();
    showTip({ clientX: b.left + b.width / 2, clientY: b.top }, spec);
  });
  g.addEventListener('blur', () => { g.classList.remove('hot'); hideTip(); });
}

/* ---- shared chart chrome ---------------------------------------------------- */
const GAP = 2;          // surface gap between adjacent fills
const R = 4;            // rounded data-end
const AXIS = '#cfd6dc', GRID = '#e3e8ec', INK = '#5d666e', DARK = '#0f1b2b';

function niceMax(v) {
  if (v <= 0) return 1;
  const p = Math.pow(10, Math.floor(Math.log10(v)));
  for (const s of [1, 1.25, 1.5, 2, 2.5, 3, 4, 5, 7.5, 10]) {
    if (s * p >= v) return s * p;
  }
  return 10 * p;
}
/* A bar rounded only at its data-end, anchored square to the baseline.
   'up' rounds the top, 'right' the right edge, 'flat' nothing (a stack's inner segment). */
function barPath(x, y, w, hgt, dir) {
  if (hgt <= 0.5 || w <= 0.5) return '';
  if (dir === 'flat') return `M${x},${y}h${w}v${hgt}h${-w}Z`;
  if (dir === 'up') {
    const r = Math.min(R, w / 2, hgt);
    return `M${x},${y + hgt}V${y + r}a${r},${r} 0 0 1 ${r},-${r}h${w - 2 * r}a${r},${r} 0 0 1 ${r},${r}V${y + hgt}Z`;
  }
  const r = Math.min(R, hgt / 2, w);
  return `M${x},${y}h${w - r}a${r},${r} 0 0 1 ${r},${r}v${hgt - 2 * r}a${r},${r} 0 0 1 ${-r},${r}h${-(w - r)}Z`;
}

function legendFor(series) {
  const box = h('div', 'legend');
  /* A lone series is already named by the chart title; only >= 2 need swatches. */
  if (series.length > 1) series.forEach(s => {
    const sp = document.createElement('span');
    const i = h('i'); i.style.background = s.color;
    sp.appendChild(i); sp.appendChild(txt(s.name));
    box.appendChild(sp);
  });
  const hint = h('span', 'hint');
  const ic = el('svg', { viewBox: '0 0 24 24', fill: 'none' });
  ic.appendChild(el('path', {
    d: 'M6 3l12 9-5 1.2 3 5.4-2.6 1.5-2.9-5.3L6 18z',
    fill: 'none', stroke: 'currentColor', 'stroke-width': '1.8', 'stroke-linejoin': 'round'
  }));
  hint.appendChild(ic);
  hint.appendChild(txt('Hover or tab a bar for the items behind it'));
  box.appendChild(hint);
  return box;
}

/* Every chart ships a table so no value is reachable only by hovering. */
function tableFor(spec) {
  const wrap = h('div');
  const btn = h('button', 'table-toggle', 'Show the numbers');
  const box = h('div', 'table-wrap');
  box.hidden = true;
  const t = document.createElement('table');
  const thead = document.createElement('thead');
  const hr = document.createElement('tr');
  hr.appendChild(h('th', null, spec.categoryLabel || 'Item'));
  spec.series.forEach(s => hr.appendChild(h('th', null, s.name)));
  thead.appendChild(hr); t.appendChild(thead);
  const tb = document.createElement('tbody');
  spec.data.forEach(d => {
    const tr = document.createElement('tr');
    tr.appendChild(h('td', null, d.label));
    spec.series.forEach(s => tr.appendChild(
      h('td', null, FMT.format(Math.round(d.values[s.key] || 0)))));
    tb.appendChild(tr);
  });
  t.appendChild(tb); box.appendChild(t);
  btn.addEventListener('click', () => {
    box.hidden = !box.hidden;
    btn.textContent = box.hidden ? 'Show the numbers' : 'Hide the numbers';
    btn.setAttribute('aria-expanded', String(!box.hidden));
  });
  btn.setAttribute('aria-expanded', 'false');
  wrap.appendChild(btn); wrap.appendChild(box);
  return wrap;
}

/* ---- vertical columns (grouped or stacked) ---------------------------------- */
function columns(spec) {
  const W = spec.width, padB = 40, padT = 34;
  const H = Math.max(140, spec.height - padB - padT);
  const svg = el('svg', { viewBox: `0 0 ${W} ${H + padB + padT}`, role: 'presentation' });
  const plotH = H;
  const tops = spec.data.map(d => spec.stacked
    ? spec.series.reduce((a, s) => a + (d.values[s.key] || 0), 0)
    : Math.max(...spec.series.map(s => d.values[s.key] || 0)));
  const max = niceMax(Math.max(...tops, 1));
  const y = v => padT + plotH - (v / max) * plotH;

  [0, 0.5, 1].forEach(f => svg.appendChild(el('line', {
    x1: 0, x2: W, y1: y(max * f), y2: y(max * f),
    stroke: f === 0 ? AXIS : GRID, 'stroke-width': f === 0 ? 2 : 1
  })));

  const slot = W / spec.data.length;
  const bandW = slot * 0.62;
  spec.data.forEach((d, i) => {
    const cx = i * slot + slot / 2;
    if (spec.stacked) {
      let cur = padT + plotH;
      const total = spec.series.reduce((a, s) => a + (d.values[s.key] || 0), 0);
      spec.series.forEach((s, si) => {
        const v = d.values[s.key] || 0;
        if (v <= 0) return;
        const hgt = (v / max) * plotH - (si ? GAP : 0);
        const top = cur - hgt;
        const g = el('g', { class: 'markwrap' });
        g.appendChild(el('path', {
          class: 'mark', d: barPath(cx - bandW / 2, top, bandW, hgt,
            si === spec.series.length - 1 ? 'up' : 'flat'),
          fill: s.color
        }));
        g.appendChild(el('rect', {
          class: 'ring', x: cx - bandW / 2 - 2, y: top - 2,
          width: bandW + 4, height: hgt + 4, rx: 6
        }));
        wire(g, {
          name: d.label + ' · ' + s.name, value: v, unit: spec.unit,
          sub: d.sub || spec.sub, color: s.color, items: (d.items || {})[s.key] || []
        });
        svg.appendChild(g);
        cur = top - GAP;
      });
      const lab = el('text', {
        x: cx, y: y(total) - 12, 'text-anchor': 'middle',
        fill: DARK, 'font-size': 26, 'font-weight': 700
      });
      lab.appendChild(txt(FMT.format(Math.round(total))));
      svg.appendChild(lab);
    } else {
      const each = (bandW - GAP * (spec.series.length - 1)) / spec.series.length;
      spec.series.forEach((s, si) => {
        const v = d.values[s.key] || 0;
        const hgt = (v / max) * plotH;
        const x = cx - bandW / 2 + si * (each + GAP);
        const g = el('g', { class: 'markwrap' });
        if (hgt > 0.5) {
          g.appendChild(el('path', {
            class: 'mark', d: barPath(x, padT + plotH - hgt, each, hgt, 'up'), fill: s.color
          }));
        }
        g.appendChild(el('rect', {
          class: 'ring', x: x - 2, y: padT + plotH - hgt - 2,
          width: each + 4, height: Math.max(hgt, 4) + 4, rx: 6
        }));
        wire(g, {
          name: d.label + ' · ' + s.name, value: v, unit: spec.unit,
          sub: d.sub || spec.sub, color: s.color, items: (d.items || {})[s.key] || []
        });
        svg.appendChild(g);
        const lab = el('text', {
          x: x + each / 2, y: padT + plotH - hgt - 10, 'text-anchor': 'middle',
          fill: DARK, 'font-size': 22, 'font-weight': 700
        });
        lab.appendChild(txt(FMT.format(Math.round(v))));
        svg.appendChild(lab);
      });
    }
    const cat = el('text', {
      x: cx, y: padT + plotH + 30, 'text-anchor': 'middle', fill: INK, 'font-size': 24
    });
    cat.appendChild(txt(d.label));
    svg.appendChild(cat);
  });
  return svg;
}

/* ---- horizontal bars -------------------------------------------------------- */
function bars(spec) {
  const W = spec.width;
  const gutter = spec.gutter || 380;
  const valW = 130;
  /* Rows share the measured height, within limits that keep a bar bar-shaped. */
  const rowH = Math.max(34, Math.min(120, spec.height / spec.data.length));
  const H = rowH * spec.data.length;
  const svg = el('svg', { viewBox: `0 0 ${W} ${H + 8}`, role: 'presentation' });
  const plotW = W - gutter - valW;
  const max = niceMax(Math.max(1, ...spec.data.flatMap(
    d => spec.series.map(s => d.values[s.key] || 0))));

  spec.data.forEach((d, i) => {
    const top = i * rowH;
    if (i) svg.appendChild(el('line', {
      x1: 0, x2: W, y1: top, y2: top, stroke: GRID, 'stroke-width': 1
    }));
    const name = el('text', {
      x: 0, y: top + rowH / 2, 'dominant-baseline': 'middle',
      fill: DARK, 'font-size': Math.min(25, Math.max(18, rowH * 0.42)),
      'font-weight': d.emphasis ? 700 : 500
    });
    name.appendChild(txt(d.label));
    svg.appendChild(name);

    const pad = Math.min(22, rowH * 0.34);
    const each = (rowH - pad - GAP * (spec.series.length - 1)) / spec.series.length;
    spec.series.forEach((s, si) => {
      const v = d.values[s.key] || 0;
      const w = (v / max) * plotW;
      const y0 = top + pad / 2 + si * (each + GAP);
      const g = el('g', { class: 'markwrap' });
      if (w > 0.5) {
        g.appendChild(el('path', {
          class: 'mark', d: barPath(gutter, y0, w, each, 'right'), fill: s.color
        }));
      }
      g.appendChild(el('rect', {
        class: 'ring', x: gutter - 2, y: y0 - 2,
        width: Math.max(w, 4) + 4, height: each + 4, rx: 5
      }));
      wire(g, {
        name: d.label + ' · ' + s.name, value: v, unit: spec.unit,
        sub: d.sub || spec.sub, color: s.color, items: (d.items || {})[s.key] || []
      });
      svg.appendChild(g);
      const lab = el('text', {
        x: gutter + Math.max(w, 0) + 12, y: y0 + each / 2, 'dominant-baseline': 'middle',
        fill: DARK, 'font-size': Math.min(24, Math.max(17, each * 0.85)), 'font-weight': 700
      });
      lab.appendChild(txt(spec.decimals
        ? v.toFixed(spec.decimals) + (spec.unit === '%' ? '%' : '')
        : FMT.format(Math.round(v))));
      svg.appendChild(lab);
    });
    if (d.note) {
      const nt = el('text', {
        x: W, y: top + rowH / 2, 'dominant-baseline': 'middle', 'text-anchor': 'end',
        fill: d.noteColor || INK, 'font-size': 23, 'font-weight': 500
      });
      nt.appendChild(txt(d.note));
      svg.appendChild(nt);
    }
  });
  svg.appendChild(el('line', {
    x1: gutter, x2: gutter, y1: 0, y2: H, stroke: AXIS, 'stroke-width': 2
  }));
  return svg;
}

/* ---- mount ------------------------------------------------------------------ */
function mount(spec) {
  const host = document.getElementById(spec.id);
  if (!host) return;
  /* Chrome first, so the plot can be measured against what is actually left over.
     The stage is CSS-transformed, but offset/client sizes stay in slide units. */
  const plot = h('div', 'plot');
  host.appendChild(plot);
  host.appendChild(legendFor(spec.series));
  host.appendChild(tableFor(spec));
  spec.width = plot.clientWidth || 1200;
  spec.height = plot.clientHeight || 300;
  plot.appendChild(spec.horizontal ? bars(spec) : columns(spec));
}

/* Slides are display:none until shown, so measure each one while it is on screen. */
const mounted = new Set();
function mountSlide(section) {
  D.charts.forEach(c => {
    if (mounted.has(c.id) || !section.querySelector('#' + c.id)) return;
    mounted.add(c.id);
    mount(Object.assign({}, c));
  });
}

/* ---- deck navigation --------------------------------------------------------- */
const slides = [...document.querySelectorAll('.slide')];
const chips = [...document.querySelectorAll('#bar button[data-go]')];
const counter = document.getElementById('count');
const notes = document.getElementById('notes');
let at = 0;
function go(i) {
  at = Math.max(0, Math.min(slides.length - 1, i));
  slides.forEach((s, n) => s.classList.toggle('on', n === at));
  mountSlide(slides[at]);
  chips.forEach((c, n) => c.setAttribute('aria-current', String(n === at)));
  counter.textContent = (at + 1) + ' / ' + slides.length;
  notes.textContent = slides[at].dataset.notes || '';
  hideTip();
  location.hash = 'slide-' + (at + 1);
}
chips.forEach((c, n) => c.addEventListener('click', () => go(n)));
document.getElementById('prev').addEventListener('click', () => go(at - 1));
document.getElementById('next').addEventListener('click', () => go(at + 1));
addEventListener('keydown', e => {
  if (e.target.matches('input, textarea')) return;
  if (e.key === 'ArrowRight' || e.key === 'PageDown') { go(at + 1); e.preventDefault(); }
  if (e.key === 'ArrowLeft' || e.key === 'PageUp') { go(at - 1); e.preventDefault(); }
  if (e.key === 'Home') { go(0); }
  if (e.key === 'End') { go(slides.length - 1); }
});

/* ---- fit the 1920x1080 stage into the window -------------------------------- */
const wrap = document.getElementById('stage-wrap');
const box = document.getElementById('stage-box');
const stage = document.getElementById('stage');
function fit() {
  const w = wrap.clientWidth - 8, hgt = wrap.clientHeight - 8;
  const k = Math.min(w / 1920, hgt / 1080);
  stage.style.transform = 'scale(' + k + ')';
  box.style.width = (1920 * k) + 'px';
  box.style.height = (1080 * k) + 'px';
}
addEventListener('resize', fit);
fit();
const start = parseInt((location.hash.match(/slide-(\d+)/) || [])[1], 10);
go(Number.isFinite(start) ? start - 1 : 0);
"""


# --- slide markup ----------------------------------------------------------------
def head(title, eyebrow, logo):
    return f"""  <div class="head">
    <div>
      <h1>{title}</h1>
      <div class="eyebrow">{eyebrow}</div>
    </div>
    <img src="data:image/png;base64,{logo}" alt="Biomed Global">
  </div>"""


def foot(note):
    return f"""  <div class="foot"><span>{note}</span>
    <span style="font-weight:500;white-space:nowrap">Customer visit &middot; {VISIT_DATE}</span></div>"""


def kpi(label, value, sub, cls="", subcls=""):
    return (f'<div class="kpi {cls}"><div class="label">{label}</div>'
            f'<b>{value}</b><small class="{subcls}">{sub}</small></div>')


def build_html(data):
    hus, pil = data["hus"], data["pil"]
    logo = b64(ROOT / "assets" / "logo-white.png")

    charts = []
    slides = []

    # ---------- 01 HUS units by line ----------
    prior = {"key": "fy26", "name": "FY26 full year", "color": "var(--s-prior)"}
    cur = {"key": "fy27", "name": "FY27 YTD (May–Aug)", "color": "var(--s-current)"}
    charts.append({
        "id": "c-hus-line", "horizontal": True, "unit": "units",
        "categoryLabel": "Product line", "gutter": 430,
        "series": [prior, cur],
        "data": [{
            "label": r["label"],
            "values": {"fy26": r["fy26"], "fy27": r["fy27"]},
            "items": {"fy26": r["items26"], "fy27": r["items27"]},
        } for r in hus["by_line"] if r["fy26"] + r["fy27"] >= 10],
    })
    hus_delta = pct(hus["lfl27"], hus["lfl26"])
    slides.append(f"""<section class="slide" data-notes="Everything on this slide is units purchased, not ringgit. Cells, diluent and QC are at or ahead of last year; gel cards are the whole gap. Hover any bar to name the products inside it.">
{head("Hospital Umum Sarawak", "Units purchased &middot; Bio-Rad &amp; Werfen Immucor transfusion", logo)}
  <div class="body">
    <div class="kpis">
      {kpi("FY26 full year", f"{n(hus['fy26_total'])} units", "Cards, cells, reagents and QC")}
      {kpi("FY27 &middot; May&ndash;Aug", f"{n(hus['lfl27'])} units",
         f"{signed(hus_delta, 0)} vs {n(hus['lfl26'])} in the same four months",
         subcls="bad" if hus_delta < 0 else "good")}
      {kpi("Since the last gel card order", "2 months", "Last cards shipped June 2026", cls="tint")}
    </div>
    <div style="display:grid;grid-template-columns:1.45fr 1fr;gap:44px;min-height:0;flex:1;padding-bottom:24px">
      <div class="chart" id="c-hus-line">
        <div class="chart-top">
          <div class="label">Units by product line</div>
          <div class="hint">FY26 is twelve months, FY27 four</div>
        </div>
      </div>
      <div class="read">
        <div class="label">The one-line read</div>
        <p>Bench consumption is holding &mdash; cells, diluent and QC all sit at or above
           last year's pace. The gap is almost entirely in <strong>gel cards</strong>.</p>
        <dl>
          <div><dt>Card units, FY26</dt><dd>{n(hus['cards26'])}</dd></div>
          <div><dt>Card units, FY27 YTD</dt><dd style="color:var(--bad)">{n(hus['cards27'])}</dd></div>
          <div><dt>Cells, reagent &amp; QC, FY27 YTD</dt><dd style="color:var(--good)">{n(hus['bench27'])}</dd></div>
        </dl>
      </div>
    </div>
  </div>
{foot("Figures are invoiced quantities for all Bio-Rad and Werfen Immucor transfusion lines.")}
</section>""")

    # ---------- 02 HUS tender ----------
    charts.append({
        "id": "c-hus-sites", "horizontal": True, "unit": "units",
        "categoryLabel": "Site", "gutter": 320,
        "series": [
            {"key": "fy26", "name": "FY26 May–Aug", "color": "var(--s-prior)"},
            {"key": "fy27", "name": "FY27 May–Aug", "color": "var(--s-current)"},
        ],
        "data": [{
            "label": s["label"],
            "values": {"fy26": s["fy26"], "fy27": s["fy27"]},
            "items": {"fy26": s["items26"], "fy27": s["items27"]},
            "note": signed(s["change"]) if s["change"] is not None else "new",
            "noteColor": "#8c2417" if (s["change"] or 0) < 0 else "#1f6b44",
        } for s in hus["sites"]],
    })
    charts.append({
        "id": "c-tender", "horizontal": True, "unit": "%",
        "categoryLabel": "State", "gutter": 300, "decimals": 2,
        "series": [{"key": "v", "name": "I-LOAN PDN achievement",
                    "color": "var(--s-current)"}],
        "data": [{
            "label": t["label"], "values": {"v": t["value"]},
            "items": {"v": []},
            "sub": f"Rank {i} of {len(hus['tender'])} \u00b7 I-LOAN PDN contract",
            "emphasis": t["label"] == "Sarawak",
        } for i, t in enumerate(hus["tender"], 1)] + [{
            "label": "National total", "values": {"v": hus["tender_total"]},
            "items": {"v": []}, "sub": "All states combined", "emphasis": True,
        }],
    })
    sar = next(s for s in hus["sites"] if s["label"] == "Hospital Umum Sarawak")
    state26 = sum(s["fy26"] for s in hus["sites"])
    state27 = sum(s["fy27"] for s in hus["sites"])
    slides.append(f"""<section class="slide" data-notes="Interim tender achievement. Sarawak as a state is up on units; HUS is the largest site but the only one moving backwards. Hover a site bar to see which lines moved.">
{head("Interim tender achievement", "Transfusion units, May&ndash;Aug FY26 vs FY27 &middot; Sarawak against other states", logo)}
  <div class="body" style="display:grid;grid-template-columns:1fr 1.1fr;gap:48px;padding-bottom:24px">
    <div class="chart" id="c-tender">
      <div class="chart-top">
        <div class="label">I-LOAN PDN achievement by state</div>
      </div>
    </div>
    <div style="display:flex;flex-direction:column;gap:24px;min-height:0">
      <div class="kpis" style="grid-template-columns:1fr 1fr;flex:none">
        {kpi("Sarawak state", f"{signed(pct(state27, state26))}",
             f"{n(state26)} &rarr; {n(state27)} units",
             subcls="good" if state27 >= state26 else "bad")}
        {kpi("Hospital Umum Sarawak", f"{signed(sar['change'])}",
             f"{n(sar['fy26'])} &rarr; {n(sar['fy27'])} units", cls="tint",
             subcls="bad" if sar["change"] < 0 else "good")}
      </div>
      <div class="chart" id="c-hus-sites">
        <div class="chart-top">
          <div class="label">Sarawak sites &middot; units in the same four months</div>
        </div>
      </div>
      <div class="note" style="flex:none">
        <p>Sarawak is ahead on the tender overall &mdash; Sibu and Miri have taken the
           growth while Hospital Umum Sarawak, the largest site, is the only one
           moving backwards.</p>
      </div>
    </div>
  </div>
{foot("Site figures are invoiced units, May&ndash;Aug in each year. Achievement percentages are from the I-LOAN PDN contract report.")}
</section>""")

    # ---------- 03 HUS six months ----------
    cards_by_month = {p["label"]: p for p in hus["cards_monthly"]}
    rows = []
    for p in hus["monthly"]:
        c = cards_by_month.get(p["label"], {"value": 0, "items": []})
        other = p["value"] - c["value"]
        card_names = {i["name"] for i in c["items"]}
        rows.append({
            "label": p["label"],
            "values": {"cards": c["value"], "other": other},
            "items": {
                "cards": c["items"],
                "other": [i for i in p["items"] if i["name"] not in card_names],
            },
            "sub": p["fy"] + (" · prior year" if p["fy"] == "FY26" else ""),
        })
    charts.append({
        "id": "c-hus-months", "unit": "units", "stacked": True, "height": 300,
        "categoryLabel": "Month",
        "series": [
            {"key": "other", "name": "Cells, reagent, QC and accessories",
             "color": "var(--s-current)"},
            {"key": "cards", "name": "Bio-Rad ID gel cards", "color": "var(--s-prior)"},
        ],
        "data": rows,
    })
    slides.append(f"""<section class="slide" data-notes="Six-month view. April was the bulk card intake; since then units have run around 100-150 a month with no cards after June. Hover a segment to see exactly which products were invoiced that month.">
{head("Hospital Umum Sarawak", "Last six months &middot; units purchased per month", logo)}
  <div class="body">
    <div class="chart" id="c-hus-months">
      <div class="chart-top">
        <div class="label">Total transfusion units, March &ndash; August 2026</div>
        <div class="hint">Gel cards split out from everything else</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:26px">
      <div class="note">
        <div class="label">Cards, month by month</div>
        <p>{n(cards_by_month['Apr 26']['value'])} in April,
           {n(cards_by_month['May 26']['value'])} in May,
           {n(cards_by_month['Jun 26']['value'])} in June, then
           <strong>nothing in July or August</strong>.</p>
      </div>
      <div class="note">
        <div class="label">Everything else, month by month</div>
        <p>Cells, reagent and QC were invoiced in every one of the six months &mdash;
           the bench is as busy as last year.</p>
      </div>
      <div class="note dark">
        <div class="label">Ask</div>
        <p>How much card stock is on the shelf, and when does the next requisition go in?</p>
      </div>
    </div>
  </div>
{foot("March and April fall in FY26; May onwards in FY27.")}
</section>""")

    # ---------- 04 PIL tests ----------
    charts.append({
        "id": "c-pil-assay", "horizontal": True, "unit": "tests",
        "categoryLabel": "Assay", "gutter": 340,
        "series": [
            {"key": "fy26", "name": "FY26 full year", "color": "var(--s-prior)"},
            {"key": "fy27", "name": "FY27 YTD (May–Aug)", "color": "var(--s-current)"},
        ],
        "data": [{
            "label": a["label"],
            "values": {"fy26": a["fy26"], "fy27": a["fy27"]},
            "items": {"fy26": a["items26"], "fy27": a["items27"]},
            "sub": f"{a['size']} tests per kit",
        } for a in pil["by_assay"]],
    })
    pil_delta = pct(pil["lfl27_tests"], pil["lfl26_tests"])
    hp = next(a for a in pil["by_assay"] if a["label"].startswith("H. pylori"))
    slides.append(f"""<section class="slide" data-notes="Frame this in tests, not ringgit. 14,650 tests last year; 2,550 in four months. H. pylori is the collapse. Hover a bar to see the kit count behind the test number.">
{head("Premier Integrated Labs &mdash; Kuching",
      "Timberland Medical Centre &middot; SNIBE Maglumi tests purchased", logo)}
  <div class="body">
    <div class="kpis">
      {kpi("FY26 full year", n(pil['fy26_tests']),
           f"tests purchased &middot; {n(pil['fy26_kits'])} reagent kits")}
      {kpi("FY27 &middot; May&ndash;Aug", n(pil['fy27_tests']),
           f"{signed(pil_delta)} vs {n(pil['lfl26_tests'])} in the same four months",
           subcls="bad" if pil_delta < 0 else "good")}
      {kpi("H. pylori kits", f"{n(hp['kits26'])} &rarr; {n(hp['kits27'])}",
           "FY26 full year &rarr; FY27 YTD", cls="tint", subcls="bad")}
    </div>
    <div style="display:grid;grid-template-columns:1.35fr 1fr;gap:44px;min-height:0;flex:1;padding-bottom:24px">
      <div class="chart" id="c-pil-assay">
        <div class="chart-top">
          <div class="label">Tests by assay</div>
          <div class="hint">Tests = kits invoiced &times; tests per kit</div>
        </div>
      </div>
      <div class="read">
        <div class="label">What they run on the Maglumi</div>
        <p>Three CLIA assays only. <strong>H. pylori IgG</strong> carried the platform last
           year and has all but stopped; EBV is the one line still ordering to pace.</p>
        <dl>
          <div><dt>H. pylori tests, FY26 &rarr; FY27 YTD</dt>
               <dd style="color:var(--bad)">{n(hp['fy26'])} &rarr; {n(hp['fy27'])}</dd></div>
          <div><dt>Consumable lines invoiced at RM 0</dt><dd>{pil['free_lines']}</dd></div>
        </dl>
        <p style="font-size:24px;font-weight:400;color:var(--ink-muted)">Reaction cups, starter
           kits, wash and system liquid and light check all invoice at RM 0 &mdash; consistent
           with a reagent-rental placement where only reagents are billed.</p>
      </div>
    </div>
  </div>
{foot("Tests per kit are read from the SAP product name. No SNIBE orders in January 2026.")}
</section>""")

    # ---------- 05 PIL six months ----------
    charts.append({
        "id": "c-pil-months", "unit": "tests", "height": 300,
        "categoryLabel": "Month",
        "series": [{"key": "v", "name": "Maglumi tests purchased",
                    "color": "var(--s-current)"}],
        "data": [{
            "label": p["label"], "values": {"v": p["value"]},
            "items": {"v": p["items"]},
            "sub": p["fy"] + (" · prior year" if p["fy"] == "FY26" else ""),
        } for p in pil["monthly"]],
    })
    vals = [p["value"] for p in pil["monthly"]]
    slides.append(f"""<section class="slide" data-notes="Six-month test volume. Ordering is small-batch and uneven. Hover a column to see which assays made up that month.">
{head("Premier Integrated Labs &mdash; Kuching",
      "Last six months &middot; tests purchased per month", logo)}
  <div class="body">
    <div class="chart" id="c-pil-months">
      <div class="chart-top">
        <div class="label">Maglumi tests purchased, March &ndash; August 2026</div>
        <div class="hint">Average {n(sum(vals) / len(vals))} a month &middot;
          range {n(min(vals))}&ndash;{n(max(vals))}</div>
      </div>
    </div>
    <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:26px">
      <div class="note">
        <div class="label">Ordering pattern</div>
        <p>One or two kits at a time, with no fixed cycle &mdash; a month can swing by a factor
           of three on a single kit.</p>
      </div>
      <div class="note">
        <div class="label">Against contract</div>
        <p style="color:var(--bad);font-weight:700">Contract target &mdash; to be confirmed</p>
        <p>The PIL SNIBE contract was not found in the projects folder or Drive. Send it and
           this becomes an achievement percentage.</p>
      </div>
      <div class="note dark">
        <div class="label">Ask</div>
        <p>Is H. pylori still on the menu, or has the workload moved to another platform?</p>
      </div>
    </div>
  </div>
{foot("Ship-to = Timberland Medical Centre Kuching (all SAP spellings), billed to Premier Integrated Labs Sdn Bhd.")}
</section>""")

    # ---------- 06 asks ----------
    def ask(i, body):
        return f'<div class="ask"><i>{i:02d}</i><div>{body}</div></div>'

    slides.append(f"""<section class="slide" data-notes="Close with three questions per site. Flag that the SNIBE contract for Premier Integrated Labs was not found.">
{head("Conversation starters", f"Both visits &middot; {VISIT_DATE}", logo)}
  <div class="body">
    <div class="asks">
      <div>
        <h2>Hospital Umum Sarawak</h2>
        {ask(1, f"Gel cards are the whole gap &mdash; {n(hus['cards26'])} units last year, "
                f"{n(hus['cards27'])} so far. Stock on hand, or a requisition still in process?")}
        {ask(2, "Cells, reagent and QC were ordered in every month without a break &mdash; "
                "the bench is as busy as last year.")}
        {ask(3, f"Sarawak as a state is {signed(pct(state27, state26))} on units while HUS is "
                f"{signed(sar['change'])}. Sibu and Miri have taken the growth.")}
        <div class="leave">
          <div class="label">Leave with</div>
          <p>A date for the next gel card requisition, and current card stock levels.</p>
        </div>
      </div>
      <div>
        <h2>Premier Integrated Labs &mdash; Kuching</h2>
        {ask(1, f"H. pylori went from {n(hp['kits26'])} kits to {n(hp['kits27'])}. Has the "
                "workload moved, or has demand simply fallen away?")}
        {ask(2, "Ordering is small-batch and uneven &mdash; would a standing quarterly "
                "schedule suit the lab better?")}
        {ask(3, f"All {pil['free_lines']} consumable lines invoice at RM 0. Worth confirming "
                "the placement terms are still what both sides expect.")}
        <div class="leave">
          <div class="label">Leave with</div>
          <p>The SNIBE contract document, so achievement can be tracked against a target.</p>
        </div>
      </div>
    </div>
  </div>
{foot(SOURCE_NOTE)}
</section>""")

    chips = "".join(
        f'<button data-go="{i}">{lab}</button>' for i, lab in enumerate([
            "HUS · Units", "HUS · Tender", "HUS · 6 months",
            "PIL · Tests", "PIL · 6 months", "Asks",
        ]))

    payload = json.dumps({"charts": charts}, ensure_ascii=False)

    return f"""<title>Kuching Visit Brief</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&display=swap">
<style>{CSS}</style>

<div id="app">
  <div id="stage-wrap"><div id="stage-box"><div id="stage">
{"".join(slides)}
  </div></div></div>
  <div id="bar">
    <button id="prev" aria-label="Previous slide">&larr;</button>
    {chips}
    <button id="next" aria-label="Next slide">&rarr;</button>
    <span class="count" id="count"></span>
  </div>
  <p id="notes"></p>
</div>
<div id="tip" role="status" aria-live="polite"></div>

<script>window.__DECK__ = {payload};</script>
<script>{JS}</script>
"""


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("src", help="SAP AR export .xlsx")
    ap.add_argument("-o", "--out", default=str(ROOT / "customer_visit_deck.html"))
    a = ap.parse_args()
    html = build_html(X.build(a.src))
    pathlib.Path(a.out).write_text(html, encoding="utf-8")
    print(f"wrote {a.out}  ({len(html):,} bytes)")
