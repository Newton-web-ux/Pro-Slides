/*!
 * ms-ppt.js — PowerPoint-style feature pack for Pro Slide
 * ------------------------------------------------------------------
 * Install: add ONE line to your editor file, after the main <script>:
 *     <script src="ms-ppt.js"></script>
 *
 * It hooks the editor's existing functions (render, renderCanvas, export,
 * presenter…) and adds new ribbon tabs/groups, so no other edits are needed.
 *
 * Adds: multi-select + marquee, align/distribute, slide layouts, drag-to-
 * reorder + hide slides, find & replace, comments, accessibility checker,
 * word count, header/footer/slide numbers, text-in-shapes, WordArt styles,
 * gradient fills, shadows, flips, hyperlinks, format painter, more shapes,
 * background formatting, transition timing, richer slideshow (laser, pen,
 * black/white screen, timer, jump-to-slide, rehearse timings), print/PDF
 * with notes pages, context menus, paste/drop images, shortcuts, status bar.
 */
(function () {
'use strict';
if (window.__msPptLoaded) return;
window.__msPptLoaded = true;
const VERSION = '1.0.0';

/* ====================================================================
   0. HELPERS
   ==================================================================== */
const g = id => document.getElementById(id);
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
function h(tag, props, ...kids) {
  const n = document.createElement(tag);
  if (props) for (const k in props) {
    const v = props[k];
    if (v == null || v === false) continue;
    if (k === 'class') n.className = v;
    else if (k === 'style') n.style.cssText = v;
    else if (k === 'html') n.innerHTML = v;
    else if (k.slice(0, 2) === 'on') n.addEventListener(k.slice(2), v);
    else n.setAttribute(k, v === true ? '' : v);
  }
  kids.flat(Infinity).forEach(c => {
    if (c == null || c === false) return;
    n.append(c.nodeType ? c : document.createTextNode(String(c)));
  });
  return n;
}
const LS = {
  get(k, d) { try { const v = localStorage.getItem('msppt_' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem('msppt_' + k, JSON.stringify(v)); } catch (e) { /* ignore */ } }
};
const inEditor = () => !!deck && g('editorView').style.display === 'flex';
const presenting = () => g('present').classList.contains('show');
const slide = () => deck.slides[current];
const byId = id => slide().elements.find(e => e.id === id);
const typing = () => {
  const a = document.activeElement;
  return !!a && (a.tagName === 'INPUT' || a.tagName === 'TEXTAREA' || a.tagName === 'SELECT' || a.isContentEditable);
};
const modalOpen = () => !!document.querySelector('.modal.show') || g('sorter').classList.contains('show');
const cloneJSON = x => JSON.parse(JSON.stringify(x));
const safeUrl = u => /^(https?:|mailto:)/i.test(u);

function lum(hex) {
  hex = String(hex || '').replace('#', '');
  if (hex.length === 3) hex = hex.split('').map(c => c + c).join('');
  if (!/^[0-9a-f]{6}$/i.test(hex)) return 0.5;
  const f = i => { const v = parseInt(hex.substr(i, 2), 16) / 255; return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
  return 0.2126 * f(0) + 0.7152 * f(2) + 0.0722 * f(4);
}
function contrastRatio(a, b) {
  const la = lum(a), lb = lum(b);
  return (Math.max(la, lb) + 0.05) / (Math.min(la, lb) + 0.05);
}

/* ====================================================================
   1. STYLES
   ==================================================================== */
const CSS = `
#canvas{user-select:none}
.ms-multi{outline:2px solid #F59E0B!important;outline-offset:0}
.ms-marq{position:fixed;z-index:120;border:1px solid #4F46E5;background:rgba(79,70,229,.12);pointer-events:none}
body.ms-painting #canvas .element{cursor:copy!important}
.ms-menu{position:fixed;z-index:300;min-width:210px;background:#fff;border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow-lg);padding:5px}
.ms-menu button{display:flex;justify-content:space-between;gap:22px;width:100%;border:none;background:transparent;text-align:left;padding:7px 10px;border-radius:6px;font-size:13px;color:var(--ink)}
.ms-menu button:hover:not(:disabled){background:var(--accent-soft);color:var(--accent-ink)}
.ms-menu button:disabled{opacity:.4;cursor:default}
.ms-menu button.danger{color:var(--danger)}
.ms-menu hr{border:none;border-top:1px solid var(--line-2);margin:4px 2px}
.ms-menu kbd{font:inherit;font-size:11px;color:var(--muted-2)}
.ms-status{flex:none;height:28px;display:flex;align-items:center;gap:14px;padding:0 14px;background:var(--paper);border-top:1px solid var(--line);font-size:11.5px;color:var(--muted)}
.ms-status button{border:none;background:transparent;color:var(--muted);font-size:11.5px;font-weight:600;padding:3px 7px;border-radius:5px}
.ms-status button:hover{background:var(--paper-2);color:var(--ink)}
.ms-status button.on{background:var(--accent-soft);color:var(--accent-ink)}
.ms-status input[type=range]{width:120px;accent-color:var(--accent)}
.ms-status .sp{flex:1}
.thumbwrap.ms-hidden .thumb{opacity:.42}
.thumbwrap.ms-hidden .thumb-num{text-decoration:line-through}
.thumbwrap.ms-over .thumb{box-shadow:0 0 0 3px var(--accent)}
.ms-badge{position:absolute;left:26px;bottom:6px;z-index:2;background:#F59E0B;color:#fff;font-size:9.5px;font-weight:800;border-radius:999px;padding:1px 6px;pointer-events:none}
.ms-pin{position:absolute;width:24px;height:24px;border-radius:50% 50% 50% 4px;background:#F59E0B;color:#fff;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center;transform:translate(-50%,-100%);cursor:grab;z-index:20;box-shadow:0 2px 6px rgba(0,0,0,.3)}
.ms-pin.done{background:#9CA3AF}
.ms-pin.pulse{animation:msPulse .9s ease 2}
@keyframes msPulse{50%{transform:translate(-50%,-100%) scale(1.5)}}
.ms-float{position:fixed;right:16px;top:118px;width:330px;max-height:calc(100vh - 160px);display:flex;flex-direction:column;background:#fff;border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow-lg);z-index:90}
.ms-fh{display:flex;align-items:center;gap:8px;padding:11px 14px;border-bottom:1px solid var(--line-2)}
.ms-fh b{flex:1;font-family:var(--font-display)}
.ms-cx{border:none;background:transparent;color:var(--muted);font-size:14px;width:26px;height:26px;border-radius:6px}
.ms-cx:hover{background:var(--paper-2);color:var(--ink)}
.ms-fb{padding:12px 14px;overflow:auto}
.ms-fb textarea,.ms-fb input[type=text]{width:100%;border:1px solid var(--line);border-radius:8px;padding:8px 9px;font-size:13px;outline:none;margin-bottom:8px}
.ms-fb textarea{min-height:64px;resize:vertical}
.ms-fb textarea:focus,.ms-fb input[type=text]:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft)}
.ms-row{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}
.ms-item{border:1px solid var(--line);border-radius:10px;padding:9px 10px;margin-bottom:8px;font-size:12.5px}
.ms-item.done{opacity:.55}
.ms-item .m{color:var(--muted);font-size:11px;margin-bottom:3px}
.ms-item .t{white-space:pre-wrap;word-break:break-word;margin-bottom:6px}
.ms-item .a{display:flex;gap:6px}
.ms-item .a button,.ms-link{border:none;background:var(--paper-2);border-radius:6px;padding:4px 9px;font-size:11.5px;font-weight:600;color:var(--ink)}
.ms-item .a button:hover{background:var(--accent-soft);color:var(--accent-ink)}
.ms-item .a button.danger{color:var(--danger)}
.ms-issue{display:flex;gap:8px;align-items:flex-start;border:1px solid var(--line);border-radius:9px;padding:8px 10px;margin-bottom:6px;font-size:12.5px;cursor:pointer}
.ms-issue:hover{background:var(--accent-soft)}
.ms-issue i{font-style:normal;font-size:11px;color:var(--muted);flex:none;width:52px}
.ms-shape-label{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;text-align:center;padding:6px;font-weight:600;line-height:1.2;white-space:pre-wrap;pointer-events:none;overflow:hidden;z-index:1}
.ms-label-edit{position:fixed;z-index:250;box-sizing:border-box;border:2px solid var(--accent);border-radius:4px;background:rgba(255,255,255,.96);padding:8px;font:600 16px/1.2 var(--font-body);text-align:center;resize:none;outline:none}
.ms-stat-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.ms-stat{border:1px solid var(--line);border-radius:12px;padding:14px;text-align:center}
.ms-stat b{display:block;font-family:var(--font-display);font-size:26px;color:var(--accent-ink)}
.ms-stat span{font-size:12px;color:var(--muted)}
.ms-sym{border:1px solid var(--line);background:#fff;border-radius:8px;height:40px;font-size:18px}
.ms-sym:hover{background:var(--accent-soft);border-color:#C9CBFB}
.ms-check{display:flex;align-items:center;gap:6px;font-size:12px;color:var(--ink);white-space:nowrap}
.ms-hint{font-size:11px;color:var(--muted);max-width:340px;line-height:1.5}
.ms-wa{border:1px solid var(--line);background:#fff;border-radius:7px;width:34px;height:34px;font-weight:800;font-size:15px}
.ms-wa:hover{border-color:var(--accent)}
/* slideshow */
.present{overflow:hidden}
#msPresentStage{position:relative;flex:none}
#msPresentStage .pslide{width:100%;height:100%}
.present .presentbar,.present .presenter-notes{z-index:9}
#msTimer{position:absolute;top:14px;right:16px;z-index:9;background:#111827cc;color:#fff;font:600 13px/1 var(--font-body);padding:7px 11px;border-radius:8px;display:none}
.present.ms-timer #msTimer{display:block}
#msLaser{display:none;position:fixed;width:20px;height:20px;margin:-10px 0 0 -10px;border-radius:50%;background:radial-gradient(circle,#fff 0 16%,#EF4444 30%,rgba(239,68,68,.35) 62%,transparent 72%);pointer-events:none;z-index:7}
.present.ms-laser #msLaser{display:block}
#msPen{position:absolute;inset:0;width:100%;height:100%;z-index:3;pointer-events:none}
.present.ms-pen #msPen{pointer-events:auto;cursor:crosshair}
#msBlank{position:absolute;inset:0;z-index:6;display:none;align-items:center;justify-content:center;color:#9CA3AF;font:500 16px var(--font-body)}
#msJump{position:absolute;left:16px;bottom:16px;z-index:9;background:#111827dd;color:#fff;font:600 13px/1 var(--font-body);padding:8px 12px;border-radius:8px;display:none}
.present.ms-idle,.present.ms-idle *{cursor:none!important}
.present.ms-laser,.present.ms-laser *{cursor:none!important}
/* print */
#msPrintRoot{display:none}
@media print{
  html,body{height:auto!important;overflow:visible!important;background:#fff!important}
  body>*:not(#msPrintRoot){display:none!important}
  #msPrintRoot{display:block!important}
  #msPrintRoot,#msPrintRoot *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .ms-pp{break-after:page;page-break-after:always;overflow:hidden}
  .ms-pp:last-child{break-after:auto;page-break-after:auto}
  .ms-pp.notes{padding:.75in;box-sizing:border-box;height:11in}
  .ms-pn{margin-top:.3in;font:13px/1.55 Inter,Arial,sans-serif;white-space:pre-wrap;color:#111}
  .ms-ps{position:relative;overflow:hidden}
}
`;
document.head.append(h('style', { id: 'msppt-style' }, CSS));

/* ====================================================================
   2. SMALL UI BUILDERS
   ==================================================================== */
function addTab(key, label, beforeKey) {
  const tabs = document.querySelector('.ribbon-tabs');
  const btn = h('button', { class: 'rtab', 'data-tab': key, onclick: () => switchRibbon(key) }, label);
  const ref = beforeKey && tabs.querySelector('[data-tab="' + beforeKey + '"]');
  if (ref) tabs.insertBefore(btn, ref); else tabs.append(btn);
  const panel = h('div', { class: 'ribbon-panel', 'data-panel': key });
  document.querySelector('.ribbon').append(panel);
  return panel;
}
const grp = (label, ...rows) => h('div', { class: 'rgroup' }, ...rows, h('div', { class: 'rgroup-label' }, label));
const row = (...c) => h('div', { class: 'rgroup-row' }, ...c);
const rb = (txt, title, fn, style) => h('button', { class: 'rbtn', title, onclick: fn, style: style || '' }, txt);
const lb = (icon, txt, title, fn, w) => h('button', { class: 'rbtn rbtn-lg', title, onclick: fn, style: w ? 'width:' + w + 'px' : '' }, icon, h('span', null, txt));
function sel(opts, fn, style, title) {
  const s = h('select', { class: 'rselect', style: style || '', title: title || '' });
  opts.forEach(([v, l]) => s.append(h('option', { value: v }, l)));
  s.addEventListener('change', () => fn(s.value));
  return s;
}
function colorIn(title, val, fn) {
  const i = h('input', { type: 'color', class: 'rcolor', title, value: val });
  i.addEventListener('change', () => fn(i.value));
  return i;
}
const pill = (txt, fn, primary) => h('button', { class: 'pill-btn' + (primary ? ' primary' : ''), onclick: fn }, txt);
const field = (label, input) => h('div', { class: 'field', style: 'margin-bottom:10px' }, h('label', null, label), input);
const lab = (input, txt) => h('label', { class: 'ms-check' }, input, txt);

function makeModal(id, title, sub, body, footer, width) {
  const old = g(id); if (old) old.remove();
  const m = h('div', { class: 'modal', id },
    h('div', { class: 'modalbox', style: 'position:relative;width:min(' + (width || 560) + 'px,92vw)' },
      h('button', { class: 'ed-iconbtn modal-close', style: 'color:var(--ink)', onclick: () => m.classList.remove('show') }, '✕'),
      h('h2', null, title), sub ? h('p', { class: 'sub' }, sub) : null,
      ...body,
      footer && footer.length ? h('div', { style: 'display:flex;gap:8px;justify-content:flex-end;margin-top:16px;flex-wrap:wrap' }, ...footer) : null));
  document.body.append(m);
  return m;
}
function floatPanel(id, title) {
  closeFloats(id);
  let p = g(id); if (p) p.remove();
  p = h('div', { id, class: 'ms-float' });
  const body = h('div', { class: 'ms-fb' });
  p.append(h('div', { class: 'ms-fh' }, h('b', null, title), h('button', { class: 'ms-cx', title: 'Close', onclick: () => p.remove() }, '✕')), body);
  document.body.append(p);
  return { p, body };
}
function closeFloats(except) { document.querySelectorAll('.ms-float').forEach(x => { if (x.id !== except) x.remove(); }); }

/* context menu */
let menuEl = null;
function closeMenu() { if (menuEl) { menuEl.remove(); menuEl = null; } }
function showMenu(x, y, items) {
  closeMenu();
  const m = h('div', { class: 'ms-menu' });
  items.forEach(it => {
    if (it === '-') { m.append(h('hr')); return; }
    const b = h('button', { class: it.danger ? 'danger' : '', onclick: () => { closeMenu(); it.run(); } }, h('span', null, it.label), it.key ? h('kbd', null, it.key) : null);
    if (it.disabled) b.disabled = true;
    m.append(b);
  });
  document.body.append(m);
  const r = m.getBoundingClientRect();
  m.style.left = Math.max(4, Math.min(x, innerWidth - r.width - 8)) + 'px';
  m.style.top = Math.max(4, Math.min(y, innerHeight - r.height - 8)) + 'px';
  menuEl = m;
}
document.addEventListener('mousedown', e => { if (menuEl && !menuEl.contains(e.target)) closeMenu(); }, true);
window.addEventListener('blur', closeMenu);
window.addEventListener('resize', closeMenu);

/* ====================================================================
   3. MULTI-SELECT, MARQUEE, GROUP OPERATIONS
   ==================================================================== */
const msSel = new Set();
let msClip = null, msClipN = 0, suppressClick = false, FP = null;

function groupIds() {
  if (!deck || !selected) return [];
  const ids = [];
  if (byId(selected)) ids.push(selected);
  msSel.forEach(i => { if (!ids.includes(i) && byId(i)) ids.push(i); });
  return ids;
}
const groupEls = () => groupIds().map(byId);
function setGroup(ids) {
  selected = ids[0] || null;
  msSel.clear();
  ids.slice(1).forEach(i => msSel.add(i));
  renderCanvas(); updateFormatPanel(); updateLayersPanel();
}
function selectAll() {
  const ids = slide().elements.filter(e => !e.hidden && !e.locked).map(e => e.id);
  editingId = null;
  setGroup(ids);
}
function needGroup() {
  const gs = groupEls();
  if (!gs.length) { toast('Select an object first'); return null; }
  return gs;
}
function applyToGroup(fn, types) {
  const gs = needGroup(); if (!gs) return;
  snapshot();
  gs.forEach(e => { if (!types || types.includes(e.type)) fn(e); });
  render(); scheduleSave();
}

function startGroupDrag(ev, ids, clickedId) {
  ev.preventDefault(); ev.stopPropagation();
  const els = ids.map(byId).filter(e => e && !e.locked);
  const canvas = g('canvas');
  const rect = canvas.getBoundingClientRect();
  const sx = ev.clientX, sy = ev.clientY;
  const orig = els.map(e => [e.x, e.y]);
  let moved = false;
  const mv = m => {
    if (!moved) { if (Math.hypot(m.clientX - sx, m.clientY - sy) < 4) return; moved = true; snapshot(); }
    let dx = (m.clientX - sx) / rect.width * 100, dy = (m.clientY - sy) / rect.height * 100;
    dx = clamp(dx, Math.max(...els.map(e => -e.x - 0 + (orig[els.indexOf(e)][0] - e.x) * 0 - orig[els.indexOf(e)][0] + e.x)), Infinity);
    const minDx = Math.max(...els.map((e, i) => -orig[i][0])), maxDx = Math.min(...els.map((e, i) => 100 - e.w - orig[i][0]));
    const minDy = Math.max(...els.map((e, i) => -orig[i][1])), maxDy = Math.min(...els.map((e, i) => 100 - e.h - orig[i][1]));
    dx = clamp(dx, minDx, maxDx); dy = clamp(dy, minDy, maxDy);
    els.forEach((e, i) => { e.x = orig[i][0] + dx; e.y = orig[i][1] + dy; });
    renderCanvas();
  };
  const up = () => {
    document.removeEventListener('mousemove', mv);
    document.removeEventListener('mouseup', up);
    if (moved) { renderThumbs(); updateFormatPanel(); scheduleSave(); }
    else { selected = clickedId; msSel.clear(); renderCanvas(); updateFormatPanel(); updateLayersPanel(); }
  };
  document.addEventListener('mousemove', mv);
  document.addEventListener('mouseup', up);
}

function alignOp(mode) {
  const gs = needGroup(); if (!gs) return;
  snapshot();
  let L = 0, R = 100, T = 0, B = 100;
  if (gs.length > 1) {
    L = Math.min(...gs.map(e => e.x)); R = Math.max(...gs.map(e => e.x + e.w));
    T = Math.min(...gs.map(e => e.y)); B = Math.max(...gs.map(e => e.y + e.h));
  }
  gs.forEach(e => {
    if (e.locked) return;
    if (mode === 'left') e.x = L;
    if (mode === 'right') e.x = R - e.w;
    if (mode === 'center') e.x = (L + R) / 2 - e.w / 2;
    if (mode === 'top') e.y = T;
    if (mode === 'bottom') e.y = B - e.h;
    if (mode === 'middle') e.y = (T + B) / 2 - e.h / 2;
  });
  render(); scheduleSave();
}
function distributeOp(axis) {
  const gs = groupEls().filter(e => e && !e.locked);
  if (gs.length < 3) return toast('Select 3 or more objects (Shift+click or drag a box)');
  snapshot();
  const k = axis === 'h' ? 'x' : 'y', sz = axis === 'h' ? 'w' : 'h';
  gs.sort((a, b) => a[k] - b[k]);
  const first = gs[0][k], last = gs[gs.length - 1][k] + gs[gs.length - 1][sz];
  const total = gs.reduce((s, e) => s + e[sz], 0);
  const gap = (last - first - total) / (gs.length - 1);
  let pos = first;
  gs.forEach(e => { e[k] = pos; pos += e[sz] + gap; });
  render(); scheduleSave();
}
function matchSize(which) {
  const gs = groupEls();
  if (gs.length < 2) return toast('Select 2 or more objects — the first one is the reference');
  snapshot();
  const ref = gs[0];
  gs.slice(1).forEach(e => {
    if (which === 'w' || which === 'both') e.w = Math.min(ref.w, 100 - e.x);
    if (which === 'h' || which === 'both') e.h = Math.min(ref.h, 100 - e.y);
  });
  render(); scheduleSave();
}
function flipOp(axis) { applyToGroup(e => { if (axis === 'x') e.flipX = !e.flipX; else e.flipY = !e.flipY; }); }
function rotateOp(deg) { applyToGroup(e => { e.rotation = (((e.rotation || 0) + deg) % 360 + 360) % 360; }); }

let nudging = false, nudgeTimer = null;
function nudge(dx, dy) {
  const gs = groupEls().filter(e => e && !e.locked);
  if (!gs.length) return;
  if (!nudging) { snapshot(); nudging = true; }
  clearTimeout(nudgeTimer);
  nudgeTimer = setTimeout(() => { nudging = false; renderThumbs(); updateFormatPanel(); scheduleSave(); }, 450);
  dx = clamp(dx, Math.max(...gs.map(e => -e.x)), Math.min(...gs.map(e => 100 - e.w - e.x)));
  dy = clamp(dy, Math.max(...gs.map(e => -e.y)), Math.min(...gs.map(e => 100 - e.h - e.y)));
  gs.forEach(e => { e.x += dx; e.y += dy; });
  renderCanvas();
}

/* Format painter */
const PAINT_KEYS = {
  text: ['fontFamily', 'fontSize', 'color', 'bold', 'italic', 'underline', 'align', 'boxFill', 'boxStroke', 'boxStrokeW', 'lineH', 'letterSp', 'valign', 'wordart', 'tshadow', 'opacity'],
  shape: ['fill', 'stroke', 'strokeWidth', 'grad', 'shadow', 'opacity', 'labelSize', 'labelColor'],
  icon: ['color', 'opacity', 'shadow'], chart: ['shadow', 'opacity'], image: ['shadow', 'opacity'],
  table: ['headerBg', 'headerColor', 'cellBg', 'altRowBg', 'useAltRows', 'textColor', 'borderColor', 'borderWidth', 'fontSize']
};
function startPainter() {
  const e = selected && byId(selected);
  if (!e) return toast('Select an object to copy its format');
  const f = { type: e.type };
  (PAINT_KEYS[e.type] || []).forEach(k => { f[k] = e[k] === undefined ? null : cloneJSON(e[k]); });
  FP = f;
  document.body.classList.add('ms-painting');
  toast('Format Painter: click the object to format (Esc to cancel)');
}
function cancelPainter() { FP = null; document.body.classList.remove('ms-painting'); }
function applyFormat(e, f) {
  if (e.type !== f.type) return false;
  Object.keys(f).forEach(k => {
    if (k === 'type') return;
    if (f[k] === null) delete e[k]; else e[k] = cloneJSON(f[k]);
  });
  return true;
}

/* canvas interaction hooks (installed once) */
function installCanvasHooks() {
  const canvas = g('canvas'), stage = g('stage');
  canvas.addEventListener('mousedown', ev => {
    if (ev.button !== 0 || !inEditor()) return;
    const node = ev.target.closest ? ev.target.closest('.element') : null;
    if (!node || !canvas.contains(node)) return;
    if (ev.target.classList && ev.target.classList.contains('handle')) return;
    const id = node.dataset.id;
    if (FP) {
      const e = byId(id);
      ev.preventDefault(); ev.stopPropagation();
      if (e) { snapshot(); toast(applyFormat(e, FP) ? 'Format applied' : 'Format Painter needs the same kind of object'); }
      cancelPainter(); render(); scheduleSave();
      return;
    }
    if (ev.shiftKey) {
      ev.preventDefault(); ev.stopPropagation();
      const ids = groupIds(), i = ids.indexOf(id);
      if (i >= 0) ids.splice(i, 1); else ids.push(id);
      editingId = null;
      setGroup(ids);
      return;
    }
    const ids = groupIds();
    if (ids.length > 1 && ids.includes(id)) { startGroupDrag(ev, ids, id); return; }
    if (msSel.size) msSel.clear();
  }, true);

  /* marquee selection on empty canvas */
  stage.addEventListener('mousedown', ev => {
    if (ev.button !== 0 || ev.target !== canvas || !inEditor() || drawTool) return;
    const r = canvas.getBoundingClientRect();
    const sx = ev.clientX, sy = ev.clientY, base = ev.shiftKey ? groupIds() : [];
    let box = null, moved = false;
    const rectOf = m => ({ x1: Math.min(sx, m.clientX), y1: Math.min(sy, m.clientY), x2: Math.max(sx, m.clientX), y2: Math.max(sy, m.clientY) });
    const mv = m => {
      if (!moved && Math.hypot(m.clientX - sx, m.clientY - sy) < 4) return;
      moved = true;
      if (!box) { box = h('div', { class: 'ms-marq' }); document.body.append(box); }
      const q = rectOf(m);
      box.style.cssText = 'left:' + q.x1 + 'px;top:' + q.y1 + 'px;width:' + (q.x2 - q.x1) + 'px;height:' + (q.y2 - q.y1) + 'px';
    };
    const up = u => {
      document.removeEventListener('mousemove', mv);
      document.removeEventListener('mouseup', up);
      if (box) box.remove();
      if (!moved) return;
      suppressClick = true;
      setTimeout(() => { suppressClick = false; }, 60);
      const q = rectOf(u);
      const X1 = (q.x1 - r.left) / r.width * 100, X2 = (q.x2 - r.left) / r.width * 100;
      const Y1 = (q.y1 - r.top) / r.height * 100, Y2 = (q.y2 - r.top) / r.height * 100;
      const hits = slide().elements.filter(e => !e.hidden && !e.locked && e.x < X2 && e.x + e.w > X1 && e.y < Y2 && e.y + e.h > Y1).map(e => e.id);
      const ids = base.concat(hits.filter(i => !base.includes(i)));
      editingId = null;
      setGroup(ids);
    };
    document.addEventListener('mousemove', mv);
    document.addEventListener('mouseup', up);
  });
  stage.addEventListener('click', ev => {
    if (suppressClick) { suppressClick = false; ev.stopPropagation(); ev.stopImmediatePropagation(); }
  }, true);

  /* Ctrl + wheel zoom */
  stage.addEventListener('wheel', ev => {
    if (!ev.ctrlKey) return;
    ev.preventDefault();
    setZoom(zoomLevel + (ev.deltaY < 0 ? 0.1 : -0.1));
  }, { passive: false });

  /* drag & drop images onto the canvas area */
  stage.addEventListener('dragover', ev => {
    if (ev.dataTransfer && [...ev.dataTransfer.types].includes('Files')) ev.preventDefault();
  });
  stage.addEventListener('drop', ev => {
    const files = ev.dataTransfer ? [...ev.dataTransfer.files].filter(f => f.type.startsWith('image/')) : [];
    if (!files.length) return;
    ev.preventDefault();
    files.forEach(addImageFile);
  });

  /* text inside shapes: double-click */
  canvas.addEventListener('dblclick', ev => {
    const node = ev.target.closest ? ev.target.closest('.element') : null;
    if (!node) return;
    const e = byId(node.dataset.id);
    if (e && e.type === 'shape' && !e.locked) { ev.preventDefault(); ev.stopPropagation(); editLabel(e, node); }
  }, true);

  /* context menus */
  canvas.addEventListener('contextmenu', ev => {
    if (!inEditor() || drawTool) return;
    ev.preventDefault();
    const node = ev.target.closest ? ev.target.closest('.element') : null;
    if (node) {
      const id = node.dataset.id;
      if (!groupIds().includes(id)) { editingId = null; msSel.clear(); selected = id; renderCanvas(); updateFormatPanel(); updateLayersPanel(); }
      elementMenu(ev.clientX, ev.clientY);
    } else {
      showMenu(ev.clientX, ev.clientY, [
        { label: 'Paste', key: 'Ctrl+V', run: () => pasteClipboard(), disabled: !clipboard && !msClip },
        { label: 'New slide', key: 'Ctrl+M', run: () => addSlide('content') },
        { label: 'Select all', key: 'Ctrl+A', run: selectAll },
        '-',
        { label: 'Format background…', run: openBackground },
        { label: 'Header & footer…', run: openHF },
        { label: 'Add comment…', run: () => openComments(true) }
      ]);
    }
  });
}

function elementMenu(x, y) {
  const gs = groupEls(); const e = gs[0]; if (!e) return;
  const multi = gs.length > 1;
  showMenu(x, y, [
    { label: 'Cut', key: 'Ctrl+X', run: cutSel },
    { label: 'Copy', key: 'Ctrl+C', run: () => copySelected() },
    { label: 'Paste', key: 'Ctrl+V', run: () => pasteClipboard(), disabled: !clipboard && !msClip },
    { label: 'Duplicate', key: 'Ctrl+D', run: () => duplicateSelected() },
    { label: 'Delete', key: 'Del', run: () => deleteSelected(), danger: true },
    '-',
    { label: 'Bring to front', run: () => bringToFront() },
    { label: 'Send to back', run: () => sendToBack() },
    { label: e.locked ? 'Unlock' : 'Lock', run: () => applyToGroup(x => { x.locked = !x.locked; }) },
    { label: 'Hide', run: () => applyToGroup(x => { x.hidden = true; }) },
    '-',
    { label: 'Flip horizontal', run: () => flipOp('x') },
    { label: 'Flip vertical', run: () => flipOp('y') },
    { label: 'Rotate 90°', run: () => rotateOp(90) },
    { label: 'Align to slide centre', run: () => { if (multi) alignOp('center'); else { alignOp('center'); alignOp('middle'); } }, disabled: false },
    '-',
    { label: 'Link…', run: setLink },
    { label: 'Alt text…', run: setAlt },
    { label: 'New comment…', run: () => openComments(true) },
    { label: 'Format painter', run: startPainter, disabled: multi }
  ]);
}
function cutSel() {
  copySelected();
  deleteSelected();
}

/* ====================================================================
   4. WRAPPERS AROUND THE EDITOR'S OWN FUNCTIONS
   ==================================================================== */
function numberSlides() {
  if (!deck) return;
  deck.slides.forEach((s, i) => s.elements.forEach(e => { if (e.msRole === 'num') e.text = String(i + 1); }));
}

const _render = render;
render = function () {
  numberSlides();
  _render.apply(this, arguments);
  try { afterRender(); } catch (err) { console.error('[ms-ppt]', err); }
};
const _renderCanvas = renderCanvas;
renderCanvas = function () {
  if (!selected) msSel.clear();
  numberSlides();
  _renderCanvas.apply(this, arguments);
  try { postCanvas(); } catch (err) { console.error('[ms-ppt]', err); }
};
const _renderThumbs = renderThumbs;
renderThumbs = function () {
  numberSlides();
  _renderThumbs.apply(this, arguments);
  try { decorateThumbs(); } catch (err) { console.error('[ms-ppt]', err); }
};
const _mini = renderMiniSlide;
renderMiniSlide = function (container, sl) {
  _mini.apply(this, arguments);
  try {
    if (!container || !sl) return;
    applyBg(container, sl, false);
    const kids = [...container.children].filter(c => c.classList.contains('mini-el'));
    (sl.elements || []).filter(e => !e.hidden).forEach((e, i) => { if (kids[i]) applyExtras(kids[i], kids[i], e, { mini: true }); });
  } catch (err) { /* thumbnails must never break */ }
};
const _styleShape = styleShapeDiv;
styleShapeDiv = function (d, e) {
  _styleShape.apply(this, arguments);
  const p = MS_POLY[e.shape];
  if (p) {
    d.style.clipPath = 'polygon(' + p + ')';
    d.style.border = 'none'; d.style.borderRadius = '0';
    d.style.background = e.fill === 'transparent' ? 'transparent' : e.fill;
  }
};
const _ufp = updateFormatPanel;
updateFormatPanel = function () {
  _ufp.apply(this, arguments);
  try {
    const n = groupIds().length;
    if (n > 1) g('inspFormat').insertAdjacentHTML('afterbegin',
      '<div class="empty-panel" style="padding:10px 6px 14px;color:var(--accent-ink);font-weight:700">' + n + ' objects selected<br><span style="font-weight:400;color:var(--muted)">Showing the first. Use Format → Arrange to align or distribute.</span></div>');
  } catch (err) { /* ignore */ }
};

const _copy = copySelected;
copySelected = function () {
  const gs = groupEls();
  if (gs.length > 1) { msClip = cloneJSON(gs); msClipN = 0; toast('Copied ' + gs.length + ' objects'); return; }
  msClip = null;
  return _copy.apply(this, arguments);
};
const _paste = pasteClipboard;
pasteClipboard = function () {
  if (msClip && msClip.length) {
    snapshot(); msClipN++;
    const ids = [];
    msClip.forEach(c => {
      const n = cloneJSON(c); n.id = uid();
      n.x = Math.min(100 - n.w, n.x + 3 * msClipN); n.y = Math.min(100 - n.h, n.y + 3 * msClipN);
      slide().elements.push(n); ids.push(n.id);
    });
    selected = ids[0]; msSel.clear(); ids.slice(1).forEach(i => msSel.add(i));
    render(); scheduleSave();
    return;
  }
  return _paste.apply(this, arguments);
};
const _dup = duplicateSelected;
duplicateSelected = function () {
  const gs = groupEls();
  if (gs.length > 1) {
    snapshot();
    const ids = [];
    gs.forEach(c => { const n = cloneJSON(c); n.id = uid(); n.x = Math.min(100 - n.w, n.x + 3); n.y = Math.min(100 - n.h, n.y + 3); slide().elements.push(n); ids.push(n.id); });
    selected = ids[0]; msSel.clear(); ids.slice(1).forEach(i => msSel.add(i));
    render(); scheduleSave();
    return;
  }
  return _dup.apply(this, arguments);
};
const _del = deleteSelected;
deleteSelected = function () {
  const ids = groupIds();
  if (ids.length > 1) {
    snapshot();
    const set = new Set(ids);
    slide().elements = slide().elements.filter(e => !set.has(e.id));
    selected = null; msSel.clear();
    render(); scheduleSave();
    return;
  }
  return _del.apply(this, arguments);
};
const _zoom = zoom;
zoom = function () { _zoom.apply(this, arguments); syncZoomUI(); };
const _fit = fitCanvas;
fitCanvas = function () { _fit.apply(this, arguments); syncZoomUI(); };
function setZoom(v) {
  v = clamp(v, 0.5, 2);
  zoomLevel = v;
  g('canvas').style.transform = 'scale(' + v + ')';
  g('zoomLabel').textContent = Math.round(v * 100) + '%';
  syncZoomUI();
}

/* ====================================================================
   5. RENDERING EXTRAS (gradients, shadows, WordArt, labels, flips…)
   ==================================================================== */
const POLY_BASE = {
  triangle: '50% 0%,0% 100%,100% 100%',
  rightArrow: '0% 25%,60% 25%,60% 0%,100% 50%,60% 100%,60% 75%,0% 75%',
  star: '50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%',
  diamond: '50% 0%,100% 50%,50% 100%,0% 50%',
  pentagon: '50% 0%,100% 38%,82% 100%,18% 100%,0% 38%',
  hexagon: '25% 0%,75% 0%,100% 50%,75% 100%,25% 100%,0% 50%',
  parallelogram: '20% 0%,100% 0%,80% 100%,0% 100%',
  trapezoid: '20% 0%,80% 0%,100% 100%,0% 100%',
  plus: '35% 0%,65% 0%,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0% 65%,0% 35%,35% 35%',
  chevron: '0% 0%,60% 0%,100% 50%,60% 100%,0% 100%,40% 50%',
  doubleArrow: '0% 50%,25% 0%,25% 30%,75% 30%,75% 0%,100% 50%,75% 100%,75% 70%,25% 70%,25% 100%',
  speechBubble: '0% 0%,100% 0%,100% 70%,35% 70%,20% 100%,20% 70%,0% 70%'
};
function starPoly(n, inner) {
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const r = i % 2 ? inner : 50, a = Math.PI * i / n - Math.PI / 2;
    pts.push((50 + r * Math.cos(a)).toFixed(1) + '% ' + (50 + r * Math.sin(a)).toFixed(1) + '%');
  }
  return pts.join(',');
}
const MS_POLY = {
  rtriangle: '0% 0%,0% 100%,100% 100%',
  octagon: '30% 0%,70% 0%,100% 30%,100% 70%,70% 100%,30% 100%,0% 70%,0% 30%',
  star4: starPoly(4, 20), star6: starPoly(6, 30), star8: starPoly(8, 36),
  arrowUp: '50% 0%,100% 40%,70% 40%,70% 100%,30% 100%,30% 40%,0% 40%',
  arrowDown: '30% 0%,70% 0%,70% 60%,100% 60%,50% 100%,0% 60%,30% 60%',
  arrowLeft: '0% 50%,40% 0%,40% 30%,100% 30%,100% 70%,40% 70%,40% 100%',
  bolt: '55% 0%,15% 55%,45% 55%,35% 100%,85% 40%,55% 40%',
  banner: '0% 20%,100% 20%,85% 50%,100% 80%,0% 80%,15% 50%'
};
const POLY_ALL = Object.assign({}, POLY_BASE, MS_POLY);

function bgLayers(s, grid) {
  const L = [], S = [], R = [];
  if (grid) {
    L.push('linear-gradient(#94a3b833 1px,transparent 1px)', 'linear-gradient(90deg,#94a3b833 1px,transparent 1px)');
    S.push('5% 5%', '5% 5%'); R.push('repeat', 'repeat');
  }
  if (s.bgImage) { L.push('url("' + s.bgImage + '")'); S.push('cover'); R.push('no-repeat'); }
  else if (s.bgGradient) {
    L.push('linear-gradient(' + (s.bgGradient.angle == null ? 135 : s.bgGradient.angle) + 'deg,' + s.bgGradient.c1 + ',' + s.bgGradient.c2 + ')');
    S.push('100% 100%'); R.push('no-repeat');
  }
  return { L, S, R };
}
function applyBg(node, s, grid) {
  if (!s.bgImage && !s.bgGradient) return;
  const { L, S, R } = bgLayers(s, grid);
  node.style.backgroundImage = L.join(',');
  node.style.backgroundSize = S.join(',');
  node.style.backgroundRepeat = R.join(',');
  node.style.backgroundPosition = 'center';
}

function applyWordArt(n, e) {
  const w = e.wordart; if (!w) return;
  const s = n.style;
  const grad = c => { s.background = c; s.webkitBackgroundClip = 'text'; s.backgroundClip = 'text'; s.webkitTextFillColor = 'transparent'; };
  if (w === 'grad') grad('linear-gradient(90deg,#4F46E5,#EC4899)');
  else if (w === 'gold') grad('linear-gradient(180deg,#FDE68A,#F59E0B 55%,#B45309)');
  else if (w === 'ice') grad('linear-gradient(180deg,#E0F2FE,#38BDF8 55%,#0369A1)');
  else if (w === 'neon') { s.color = '#fff'; s.textShadow = '0 0 6px #fff,0 0 14px #8B7CF8,0 0 28px #8B7CF8'; }
  else if (w === 'outline') { s.webkitTextFillColor = 'transparent'; s.webkitTextStroke = '1.5px ' + (e.color || '#12141C'); }
  else if (w === 'pop') s.textShadow = '3px 3px 0 rgba(0,0,0,.22)';
}
function textExtras(n, e) {
  const s = n.style;
  if (e.boxFill) s.background = e.boxFill;
  if (e.boxStroke && e.boxStrokeW) s.boxShadow = 'inset 0 0 0 ' + e.boxStrokeW + 'px ' + e.boxStroke;
  if (e.lineH) s.lineHeight = String(e.lineH);
  if (e.letterSp) s.letterSpacing = e.letterSp + 'em';
  if (e.valign && e.valign !== 'top' && !n.isContentEditable) {
    s.display = 'flex'; s.flexDirection = 'column';
    s.justifyContent = e.valign === 'middle' ? 'center' : 'flex-end';
  }
  if (e.tshadow) s.textShadow = '2px 3px 5px rgba(0,0,0,.35)';
  applyWordArt(n, e);
}
function autoLabelColor(e) {
  if (!e.fill || e.fill === 'transparent') return '#12141C';
  return lum(toHex(e.fill)) > 0.55 ? '#12141C' : '#FFFFFF';
}
function shapeExtras(n, e, opts) {
  if (e.grad && !COMPOUND_SHAPES.has(e.shape) && e.shape !== 'donut' && e.shape !== 'line') {
    n.style.background = 'linear-gradient(' + (e.grad.angle == null ? 90 : e.grad.angle) + 'deg,' + e.grad.c1 + ',' + e.grad.c2 + ')';
  }
  if (e.label) {
    const sc = opts && opts.mini ? 0.34 : 1;
    n.append(h('div', { class: 'ms-shape-label', style: 'font-size:' + ((e.labelSize || 20) * sc) + 'px;color:' + (e.labelColor || autoLabelColor(e)) }, e.label));
  }
}
function applyExtras(tNode, sNode, e, opts) {
  if (!tNode || !sNode) return;
  opts = opts || {};
  if ((e.flipX || e.flipY) && e.type !== 'text' && e.type !== 'table' && e.type !== 'math') {
    tNode.style.transform = 'rotate(' + (e.rotation || 0) + 'deg) scale(' + (e.flipX ? -1 : 1) + ',' + (e.flipY ? -1 : 1) + ')';
  }
  if (e.type === 'text') textExtras(sNode, e);
  else if (e.type === 'shape') shapeExtras(sNode, e, opts);
  if (e.shadow && e.type !== 'text' && !(e.type === 'shape' && COMPOUND_SHAPES.has(e.shape))) {
    const cp = sNode.style.clipPath;
    if (!cp || cp === 'none') sNode.style.filter = 'drop-shadow(3px 5px 6px rgba(15,23,42,.38))';
  }
}

/* ---- editor canvas post-processing ---- */
let msSpell = LS.get('spell', false), PINS = LS.get('pins', true), pinFlash = null;
function postCanvas() {
  const c = g('canvas'), s = slide();
  if (!s) return;
  applyBg(c, s, showGrid);
  s.elements.forEach(e => {
    if (e.hidden) return;
    const n = c.querySelector('[data-id="' + e.id + '"]');
    if (!n) return;
    applyExtras(n, n, e);
    if (e.link) n.title = e.link;
    if (e.alt) n.setAttribute('aria-label', e.alt);
  });
  const ids = groupIds();
  if (ids.length > 1) {
    ids.forEach(id => {
      const n = c.querySelector('[data-id="' + id + '"]');
      if (!n) return;
      n.classList.remove('selected'); n.classList.add('ms-multi');
      n.querySelectorAll('.handle').forEach(x => x.remove());
    });
  }
  if (PINS) renderPins(c, s);
  c.querySelectorAll('.text-el,.table-cell').forEach(n => { n.spellcheck = !!msSpell; });
}
function afterRender() {
  updateStatus();
  syncTransitionUI();
  if (g('msCPanel')) renderCommentPanel();
}

/* ---- thumbnails: hidden marker, comment badge, drag reorder, menu ---- */
let dragFrom = null;
function decorateThumbs() {
  [...g('thumbs').children].forEach((w, i) => {
    const s = deck.slides[i]; if (!s) return;
    if (s.hidden) w.classList.add('ms-hidden');
    const nc = (s.comments || []).filter(c => !c.done).length;
    if (nc) w.append(h('span', { class: 'ms-badge', title: nc + ' open comment(s)' }, '💬 ' + nc));
    w.draggable = true;
    w.addEventListener('dragstart', ev => { dragFrom = i; ev.dataTransfer.effectAllowed = 'move'; ev.dataTransfer.setData('text/plain', String(i)); });
    w.addEventListener('dragover', ev => { if (dragFrom == null) return; ev.preventDefault(); w.classList.add('ms-over'); });
    w.addEventListener('dragleave', () => w.classList.remove('ms-over'));
    w.addEventListener('drop', ev => {
      ev.preventDefault(); w.classList.remove('ms-over');
      if (dragFrom != null) { const from = dragFrom; dragFrom = null; moveSlide(from, i); }
    });
    w.addEventListener('dragend', () => { dragFrom = null; });
    w.addEventListener('contextmenu', ev => { ev.preventDefault(); thumbMenu(ev.clientX, ev.clientY, i); });
  });
}
function thumbMenu(x, y, i) {
  const s = deck.slides[i];
  showMenu(x, y, [
    { label: 'New slide', key: 'Ctrl+M', run: () => { current = i; addSlide('content'); } },
    { label: 'Duplicate slide', run: () => duplicateSlide(i) },
    { label: 'Delete slide', run: () => deleteSlide(i), danger: true },
    '-',
    { label: s.hidden ? 'Unhide slide' : 'Hide slide', run: () => toggleHide(i) },
    { label: 'Move up', run: () => moveSlide(i, i - 1), disabled: i === 0 },
    { label: 'Move down', run: () => moveSlide(i, i + 1), disabled: i === deck.slides.length - 1 },
    '-',
    { label: 'Format background…', run: () => { current = i; selected = null; render(); openBackground(); } }
  ]);
}
const _openSorter = openSorter;
openSorter = function () {
  _openSorter.apply(this, arguments);
  try {
    [...g('sorterGrid').children].forEach((card, i) => {
      card.draggable = true;
      if (deck.slides[i] && deck.slides[i].hidden) card.style.opacity = '.5';
      card.addEventListener('dragstart', ev => { dragFrom = i; ev.dataTransfer.setData('text/plain', String(i)); });
      card.addEventListener('dragover', ev => { if (dragFrom != null) ev.preventDefault(); });
      card.addEventListener('drop', ev => {
        ev.preventDefault();
        if (dragFrom == null) return;
        const from = dragFrom; dragFrom = null;
        moveSlide(from, i);
        _openSorter();
      });
    });
  } catch (err) { /* ignore */ }
};

/* ====================================================================
   6. LAYOUTS & SLIDE OPERATIONS
   ==================================================================== */
const T0 = () => ({ x: 8, y: 7, w: 84, h: 12, fs: 32, b: 1, title: 1, ph: 'Slide title' });
const LAYOUTS = {
  title: { name: 'Title Slide', spec: () => [{ x: 10, y: 30, w: 80, h: 16, fs: 44, b: 1, al: 'center', title: 1, ph: 'Presentation Title' }, { x: 14, y: 49, w: 72, h: 10, fs: 20, al: 'center', muted: 1, ph: 'A subtitle that sets the scene' }] },
  content: { name: 'Title and Content', spec: () => [T0(), { x: 10, y: 25, w: 80, h: 55, fs: 20, bullet: 1, ph: 'First important point\nSecond important point\nThird important point' }] },
  two: { name: 'Two Content', spec: () => [T0(), { rect: 1, x: 8, y: 24, w: 0.25, h: 60, fill: 'accent' }, { x: 10, y: 25, w: 38, h: 58, fs: 19, bullet: 1, ph: 'Left topic\nSupporting point\nSupporting point' }, { x: 54, y: 25, w: 38, h: 58, fs: 19, bullet: 1, ph: 'Right topic\nSupporting point\nSupporting point' }] },
  compare: { name: 'Comparison', spec: () => [T0(), { x: 8, y: 24, w: 40, h: 9, fs: 22, b: 1, ph: 'Option A' }, { x: 52, y: 24, w: 40, h: 9, fs: 22, b: 1, ph: 'Option B' }, { x: 8, y: 35, w: 40, h: 48, fs: 18, bullet: 1, ph: 'Point one\nPoint two' }, { x: 52, y: 35, w: 40, h: 48, fs: 18, bullet: 1, ph: 'Point one\nPoint two' }] },
  titleOnly: { name: 'Title Only', spec: () => [T0()] },
  section: { name: 'Section Header', bg: 'ink', spec: () => [{ x: 10, y: 42, w: 80, h: 16, fs: 40, b: 1, light: 1, title: 1, ph: 'Section title' }] },
  quote: { name: 'Quote', bg: 'soft', spec: () => [{ x: 12, y: 30, w: 76, h: 26, fs: 32, b: 1, al: 'center', title: 1, ph: '“A memorable quote goes here.”' }, { x: 26, y: 58, w: 48, h: 8, fs: 16, al: 'center', muted: 1, ph: '— Attribution' }] },
  caption: { name: 'Content with Caption', spec: () => [{ rect: 1, x: 46, y: 10, w: 48, h: 78, fill: 'soft' }, { x: 6, y: 10, w: 34, h: 18, fs: 28, b: 1, title: 1, ph: 'Slide title' }, { x: 6, y: 30, w: 34, h: 52, fs: 16, muted: 1, ph: 'Add a short description here.' }] },
  blank: { name: 'Blank', spec: () => [] }
};
function applyLayout(key) {
  const L = LAYOUTS[key]; if (!L || !deck) return;
  const t = themeOf(deck.theme), s = slide();
  snapshot();
  const old = s.elements.filter(e => e.type === 'text' && !e.msRole);
  const others = s.elements.filter(e => !(e.type === 'text' && !e.msRole));
  const used = new Set(), shapes = [], texts = [];
  L.spec().forEach(p => {
    if (p.rect) {
      const col = p.fill === 'accent' ? t.accent : t.soft;
      const r = shapeEl('rect', p.x, p.y, p.w, p.h, col); r.stroke = col; r.strokeWidth = 0;
      shapes.push(r); return;
    }
    const src = old.find(o => !used.has(o.id));
    if (src) used.add(src.id);
    texts.push(textEl(src ? src.text : p.ph, p.x, p.y, p.w, p.h, p.fs, {
      bold: !!p.b, align: p.al || 'left', bullet: !!p.bullet,
      color: p.light ? t.bg : (p.muted ? t.muted : t.ink), fontFamily: p.title ? t.display : 'Inter'
    }));
  });
  s.elements = [...shapes, ...others, ...texts, ...old.filter(o => !used.has(o.id))];
  const std = [t.bg, t.soft, t.ink].map(c => String(c).toLowerCase());
  if (std.includes(String(s.bg).toLowerCase())) s.bg = L.bg === 'ink' ? t.ink : L.bg === 'soft' ? t.soft : t.bg;
  selected = null; msSel.clear();
  render(); scheduleSave();
  toast('Layout: ' + L.name);
}
function toggleHide(i) {
  if (i == null) i = current;
  snapshot();
  deck.slides[i].hidden = !deck.slides[i].hidden;
  render(); scheduleSave();
  toast(deck.slides[i].hidden ? 'Slide hidden from show & export' : 'Slide unhidden');
}
function moveSlide(from, to) {
  if (to < 0 || to >= deck.slides.length || from === to) return;
  snapshot();
  const [s] = deck.slides.splice(from, 1);
  deck.slides.splice(to, 0, s);
  current = to; selected = null;
  render(); scheduleSave();
}

/* ====================================================================
   7. FORMAT OPERATIONS (fill, outline, text effects, links…)
   ==================================================================== */
function setFill(v) {
  applyToGroup(e => {
    if (e.type === 'shape') { e.fill = v; if (e.grad) e.grad.c1 = v; }
    else if (e.type === 'text') e.boxFill = v;
    else if (e.type === 'icon' || e.type === 'ink' || e.type === 'math') e.color = v;
  });
}
function noFill() {
  applyToGroup(e => { if (e.type === 'shape') { e.fill = 'transparent'; delete e.grad; } else if (e.type === 'text') delete e.boxFill; });
}
function setGradient(c2, angle) {
  applyToGroup(e => {
    const c1 = !e.fill || e.fill === 'transparent' ? '#FFFFFF' : e.fill;
    e.grad = { c1, c2, angle: +angle };
  }, ['shape']);
}
function clearGradient() { applyToGroup(e => { delete e.grad; }, ['shape']); }
function setOutlineColor(v) {
  applyToGroup(e => {
    if (e.type === 'shape') e.stroke = v;
    else if (e.type === 'text') { e.boxStroke = v; if (!e.boxStrokeW) e.boxStrokeW = 2; }
  });
}
function setOutlineWidth(w) {
  w = +w;
  applyToGroup(e => {
    if (e.type === 'shape') e.strokeWidth = w;
    else if (e.type === 'text') { e.boxStrokeW = w; if (!e.boxStroke) e.boxStroke = e.color || '#12141C'; }
  });
}
function toggleShadow() {
  const gs = needGroup(); if (!gs) return;
  const on = !gs[0].shadow;
  applyToGroup(e => { if (e.type === 'text') e.tshadow = on; else e.shadow = on; });
}
function setLabelSize(v) { applyToGroup(e => { e.labelSize = +v; }, ['shape']); }
function setLabelColor(v) { applyToGroup(e => { e.labelColor = v; }, ['shape']); }
function setWordArt(id) { applyToGroup(e => { if (id) e.wordart = id; else delete e.wordart; }, ['text']); }
function setLineSpacing(v) { applyToGroup(e => { e.lineH = +v || undefined; if (!e.lineH) delete e.lineH; }, ['text']); }
function setLetterSpacing(v) { applyToGroup(e => { e.letterSp = +v; if (!e.letterSp) delete e.letterSp; }, ['text']); }
function setVAlign(v) { applyToGroup(e => { e.valign = v; }, ['text']); }
const FONT_STEPS = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72, 80, 96];
function growFont(dir) {
  applyToGroup(e => {
    const cur = e.fontSize || 20;
    if (dir > 0) e.fontSize = FONT_STEPS.find(s => s > cur) || cur + 8;
    else { const smaller = FONT_STEPS.filter(s => s < cur); e.fontSize = smaller.length ? smaller[smaller.length - 1] : Math.max(6, cur - 2); }
  }, ['text']);
}
function changeCase(mode) {
  const f = { upper: s => s.toUpperCase(), lower: s => s.toLowerCase(), title: s => s.toLowerCase().replace(/\b\w/g, c => c.toUpperCase()), sentence: s => s.toLowerCase().replace(/(^\s*|[.!?]\s+)([a-z])/g, (m, a, b) => a + b.toUpperCase()) }[mode];
  if (f) applyToGroup(e => { e.text = f(e.text); }, ['text']);
}
function clearTextFormat() {
  const t = themeOf(deck.theme);
  applyToGroup(e => {
    e.bold = e.italic = e.underline = false; e.color = t.ink; e.fontFamily = 'Inter'; e.align = 'left';
    ['boxFill', 'boxStroke', 'boxStrokeW', 'lineH', 'letterSp', 'valign', 'wordart', 'tshadow'].forEach(k => delete e[k]);
  }, ['text']);
}
function toggleList(kind) {
  const gs = groupEls().filter(e => e.type === 'text');
  if (!gs.length) return toast('Select a text box');
  snapshot();
  gs.forEach(e => {
    const lines = e.text.split('\n');
    const all = lines.filter(l => l.trim()).every(l => kind === 'b' ? /^\s*•\s/.test(l) : /^\s*\d+[.)]\s/.test(l));
    let n = 0;
    e.text = lines.map(l => {
      if (!l.trim()) return l;
      const m = l.match(/^(\s*)(•\s+|\d+[.)]\s+)?(.*)$/);
      if (all) return m[1] + m[3];
      n++;
      return m[1] + (kind === 'b' ? '• ' : n + '. ') + m[3];
    }).join('\n');
  });
  render(); scheduleSave();
}
function indentText(dir) {
  const gs = groupEls().filter(e => e.type === 'text');
  if (!gs.length) return toast('Select a text box');
  snapshot();
  gs.forEach(e => { e.text = e.text.split('\n').map(l => dir > 0 ? '    ' + l : l.replace(/^( {1,4}|\t)/, '')).join('\n'); });
  render(); scheduleSave();
}
function setLink() {
  const e = selected && byId(selected);
  if (!e) return toast('Select an object first');
  let u = prompt('Link address (http://, https:// or mailto:). Leave empty to remove:', e.link || 'https://');
  if (u === null) return;
  u = u.trim();
  snapshot();
  if (!u || u === 'https://') delete e.link;
  else {
    if (!/^[a-z][a-z0-9+.-]*:/i.test(u)) u = 'https://' + u;
    if (!safeUrl(u)) { toast('Only http, https and mailto links are allowed'); return; }
    e.link = u;
  }
  render(); scheduleSave();
  toast(e.link ? 'Link set — it opens during the slide show' : 'Link removed');
}
function setAlt() {
  const e = selected && byId(selected);
  if (!e) return toast('Select an object first');
  const v = prompt('Alt text (describe this object for screen readers):', e.alt || '');
  if (v === null) return;
  snapshot();
  if (v.trim()) e.alt = v.trim(); else delete e.alt;
  render(); scheduleSave();
}
function editLabel(e, node) {
  const r = node.getBoundingClientRect();
  const ta = h('textarea', { class: 'ms-label-edit', style: 'left:' + r.left + 'px;top:' + r.top + 'px;width:' + Math.max(80, r.width) + 'px;height:' + Math.max(40, r.height) + 'px' });
  ta.value = e.label || '';
  document.body.append(ta);
  ta.focus(); ta.select();
  const done = commit => {
    if (!ta.isConnected) return;
    const v = ta.value;
    ta.remove();
    if (commit && v !== (e.label || '')) { snapshot(); if (v) e.label = v; else delete e.label; render(); scheduleSave(); }
  };
  ta.addEventListener('blur', () => done(true));
  ta.addEventListener('keydown', k => {
    k.stopPropagation();
    if (k.key === 'Escape') { done(false); }
    else if (k.key === 'Enter' && !k.shiftKey) { k.preventDefault(); done(true); }
  });
}
function addShapeMs(shape) {
  snapshot();
  const t = themeOf(deck.theme);
  const e = shapeEl(shape, 32, 28, 24, 24, t.accent);
  e.stroke = t.accent;
  slide().elements.push(e);
  selected = e.id; msSel.clear();
  render(); scheduleSave();
}

/* ====================================================================
   8. FIND & REPLACE
   ==================================================================== */
const FR = { find: '', repl: '', cs: false, word: false, pos: -1 };
function frRegex() {
  if (!FR.find) return null;
  let s = FR.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (FR.word) s = '\\b' + s + '\\b';
  return new RegExp(s, FR.cs ? 'g' : 'gi');
}
const countMatches = (str, re) => { re.lastIndex = 0; return (String(str).match(re) || []).length; };
function frCollect() {
  const out = [], re = frRegex();
  if (!re) return out;
  deck.slides.forEach((s, si) => s.elements.forEach(e => {
    if (e.type === 'text') { const n = countMatches(e.text, re); for (let k = 0; k < n; k++) out.push({ si, id: e.id, kind: 't', k }); }
    else if (e.type === 'table') e.cells.forEach((rw, r) => rw.forEach((v, c) => { const n = countMatches(v, re); for (let k = 0; k < n; k++) out.push({ si, id: e.id, kind: 'c', r, c, k }); }));
    else if (e.type === 'shape' && e.label) { const n = countMatches(e.label, re); for (let k = 0; k < n; k++) out.push({ si, id: e.id, kind: 'l', k }); }
  }));
  return out;
}
function frGo(m) {
  current = m.si; selected = m.id; msSel.clear(); editingId = null;
  render();
}
function frNext(info) {
  const ms = frCollect();
  if (!ms.length) { if (info) info.textContent = 'No matches found'; return; }
  FR.pos = (FR.pos + 1) % ms.length;
  frGo(ms[FR.pos]);
  if (info) info.textContent = 'Match ' + (FR.pos + 1) + ' of ' + ms.length;
}
function frReplaceOne(info) {
  const ms = frCollect();
  if (!ms.length) { if (info) info.textContent = 'No matches found'; return; }
  if (FR.pos < 0 || FR.pos >= ms.length) FR.pos = 0;
  const m = ms[FR.pos], re = frRegex();
  const e = deck.slides[m.si].elements.find(x => x.id === m.id);
  if (!e) return;
  const rep = (str, k) => { let i = -1; return String(str).replace(re, x => (++i === k ? FR.repl : x)); };
  snapshot();
  if (m.kind === 't') e.text = rep(e.text, m.k);
  else if (m.kind === 'l') e.label = rep(e.label, m.k);
  else e.cells[m.r][m.c] = rep(e.cells[m.r][m.c], m.k);
  scheduleSave();
  FR.pos -= 1;
  frNext(info);
}
function frReplaceAll(info) {
  const re = frRegex();
  if (!re) return;
  let n = 0;
  snapshot();
  const sub = str => String(str).replace(re, () => { n++; return FR.repl; });
  deck.slides.forEach(s => s.elements.forEach(e => {
    if (e.type === 'text') e.text = sub(e.text);
    else if (e.type === 'table') e.cells = e.cells.map(rw => rw.map(sub));
    else if (e.type === 'shape' && e.label) e.label = sub(e.label);
  }));
  render(); scheduleSave();
  if (info) info.textContent = 'Replaced ' + n + ' occurrence' + (n === 1 ? '' : 's');
  toast('Replaced ' + n + ' occurrence' + (n === 1 ? '' : 's'));
}
function openFind(withReplace) {
  if (!inEditor()) return;
  const { p, body } = floatPanel('msFindPanel', withReplace ? 'Find & Replace' : 'Find');
  const info = h('div', { style: 'font-size:12.5px;color:var(--muted);min-height:18px;margin:2px 0 10px' });
  const fi = h('input', { type: 'text', placeholder: 'Find what…' });
  const ri = h('input', { type: 'text', placeholder: 'Replace with…' });
  const cs = h('input', { type: 'checkbox' }), wd = h('input', { type: 'checkbox' });
  fi.value = FR.find; ri.value = FR.repl; cs.checked = FR.cs; wd.checked = FR.word;
  const sync = () => {
    FR.find = fi.value; FR.repl = ri.value; FR.cs = cs.checked; FR.word = wd.checked; FR.pos = -1;
    const n = FR.find ? frCollect().length : 0;
    info.textContent = FR.find ? n + ' match' + (n === 1 ? '' : 'es') + ' in this presentation' : '';
  };
  [fi, ri, cs, wd].forEach(i => i.addEventListener('input', sync));
  cs.addEventListener('change', sync); wd.addEventListener('change', sync);
  fi.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); sync(); frNext(info); } });
  body.append(fi, withReplace ? ri : null,
    h('div', { style: 'display:flex;gap:16px;margin-bottom:8px' }, lab(cs, 'Match case'), lab(wd, 'Whole word')),
    info,
    h('div', { class: 'ms-row' }, pill('Find next', () => { sync(); frNext(info); }, true),
      withReplace ? pill('Replace', () => { frReplaceOne(info); }) : null,
      withReplace ? pill('Replace all', () => { frReplaceAll(info); }) : null));
  sync();
  setTimeout(() => fi.focus(), 30);
  return p;
}

/* ====================================================================
   9. HEADER / FOOTER / SLIDE NUMBERS, DATE, SYMBOLS
   ==================================================================== */
function syncRole(s, role, want, make, text) {
  const existing = s.elements.filter(e => e.msRole === role);
  if (!want) { s.elements = s.elements.filter(e => e.msRole !== role); return; }
  if (!existing.length) s.elements.push(make());
  else if (text != null) existing.forEach(e => { e.text = text; });
}
function applyHF(o, all) {
  const t = themeOf(deck.theme);
  snapshot();
  deck.hf = o;
  const dateText = o.dateAuto ? new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }) : (o.dateText || '');
  (all ? deck.slides : [slide()]).forEach(s => {
    const si = deck.slides.indexOf(s), skip = o.skipTitle && si === 0;
    const mk = (txt, x, w, al, role) => () => textEl(txt, x, 91, w, 6, 14, { align: al, color: t.muted, msRole: role });
    syncRole(s, 'num', !!o.num && !skip, mk(String(si + 1), 88, 8, 'right', 'num'), String(si + 1));
    syncRole(s, 'footer', !!o.footer && !skip, mk(o.footer, 30, 40, 'center', 'footer'), o.footer);
    syncRole(s, 'date', !!o.date && !skip && !!dateText, mk(dateText, 5, 22, 'left', 'date'), dateText);
  });
  render(); scheduleSave();
  toast('Header & footer applied');
}
function openHF() {
  if (!inEditor()) return;
  const o = Object.assign({ num: true, footer: '', date: false, dateAuto: true, dateText: '', skipTitle: true }, deck.hf || {});
  const num = h('input', { type: 'checkbox' }), date = h('input', { type: 'checkbox' }), auto = h('input', { type: 'checkbox' });
  const skip = h('input', { type: 'checkbox' });
  const foot = h('input', { type: 'text', placeholder: 'e.g. Company name · Confidential' });
  const dt = h('input', { type: 'text', placeholder: 'Fixed date text' });
  num.checked = o.num; date.checked = o.date; auto.checked = o.dateAuto; skip.checked = o.skipTitle; foot.value = o.footer; dt.value = o.dateText;
  const read = () => ({ num: num.checked, footer: foot.value.trim(), date: date.checked, dateAuto: auto.checked, dateText: dt.value.trim(), skipTitle: skip.checked });
  const m = makeModal('msHF', 'Header & Footer', 'Adds slide number, footer text and date to the bottom of your slides.', [
    h('div', { style: 'display:grid;gap:10px' },
      lab(num, 'Slide number'),
      h('div', null, h('div', { style: 'margin-bottom:4px' }, lab(date, 'Date')), h('div', { style: 'display:flex;gap:10px;align-items:center;padding-left:22px' }, lab(auto, 'Update automatically'), dt)),
      field('Footer', foot),
      lab(skip, "Don't show on title slide"))
  ], [
    pill('Remove from all', () => { applyHF({ num: false, footer: '', date: false, dateAuto: true, dateText: '', skipTitle: true }, true); m.classList.remove('show'); }),
    pill('Apply to this slide', () => { applyHF(read(), false); m.classList.remove('show'); }),
    pill('Apply to all', () => { applyHF(read(), true); m.classList.remove('show'); }, true)
  ], 520);
  m.classList.add('show');
}
function insertDate() {
  snapshot();
  const e = textEl(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }), 60, 88, 35, 7, 16, { align: 'right', color: themeOf(deck.theme).muted });
  slide().elements.push(e); selected = e.id; msSel.clear();
  render(); scheduleSave();
}
const SYMBOLS = '© ® ™ ° ± × ÷ ≠ ≤ ≥ ≈ ∞ π Σ √ ∫ ∂ Δ Ω µ α β γ δ λ θ ← → ↑ ↓ ↔ ⇒ ✓ ✗ ★ ☆ ♥ ♦ ♣ ♠ ♪ € £ ¥ ₹ ¢ § ¶ † • … – — « » ¼ ½ ¾ ² ³ ☀ ☁ ☎ ✉ ⚠ ✔ ✈ ☂ ⚙'.split(' ');
function openSymbols() {
  if (!inEditor()) return;
  const grid = h('div', { style: 'display:grid;grid-template-columns:repeat(8,1fr);gap:7px' },
    SYMBOLS.map(sy => h('button', { class: 'ms-sym', title: sy, onclick: () => { insertSymbol(sy); } }, sy)));
  const m = makeModal('msSym', 'Insert a symbol', 'Appends to the selected text box, or creates a new one.', [grid], [pill('Close', () => m.classList.remove('show'))], 520);
  m.classList.add('show');
}
function insertSymbol(sy) {
  const e = selected && byId(selected);
  snapshot();
  if (e && e.type === 'text') { e.text += sy; }
  else {
    const n = textEl(sy, 42, 38, 16, 18, 60, { align: 'center', color: themeOf(deck.theme).ink });
    slide().elements.push(n); selected = n.id;
  }
  render(); scheduleSave();
}

/* ====================================================================
   10. COMMENTS
   ==================================================================== */
function commentAuthor() { try { return (currentUser && currentUser.name) || 'You'; } catch (e) { return 'You'; } }
let cDraft = '';
function allComments() {
  const out = [];
  deck.slides.forEach((s, si) => (s.comments || []).forEach(c => out.push({ si, c })));
  return out;
}
function addComment(text) {
  text = (text || '').trim();
  if (!text) return toast('Type a comment first');
  snapshot();
  const s = slide(); s.comments = s.comments || [];
  const sel = selected ? byId(selected) : null, n = s.comments.length;
  s.comments.push({
    id: uid(), author: commentAuthor(), text, ts: Date.now(), done: false,
    x: sel ? clamp(sel.x + sel.w, 3, 97) : clamp(94 - (n % 4) * 4, 3, 97),
    y: sel ? clamp(sel.y, 6, 95) : 8 + ((n * 6) % 60)
  });
  cDraft = '';
  render(); scheduleSave();
}
function findComment(si, id) { const s = deck.slides[si]; return s && (s.comments || []).find(c => c.id === id); }
function toggleResolve(si, id) { const c = findComment(si, id); if (!c) return; snapshot(); c.done = !c.done; render(); scheduleSave(); }
function deleteComment(si, id) {
  const s = deck.slides[si]; if (!s) return;
  snapshot(); s.comments = (s.comments || []).filter(c => c.id !== id);
  render(); scheduleSave();
}
function goComment(si, id) {
  current = si; selected = null; msSel.clear(); pinFlash = id;
  render();
  setTimeout(() => { pinFlash = null; }, 2200);
}
function openComments(focusNew) {
  if (!inEditor()) return;
  if (!g('msCPanel')) floatPanel('msCPanel', 'Comments');
  renderCommentPanel();
  if (focusNew) setTimeout(() => { const t = g('msCText'); if (t) t.focus(); }, 30);
}
function renderCommentPanel() {
  const p = g('msCPanel'); if (!p || !deck) return;
  const body = p.querySelector('.ms-fb'); body.innerHTML = '';
  const list = allComments();
  const ta = h('textarea', { id: 'msCText', placeholder: 'Add a comment to slide ' + (current + 1) + '…' });
  ta.value = cDraft;
  ta.addEventListener('input', () => { cDraft = ta.value; });
  body.append(ta, h('div', { class: 'ms-row' }, pill('Post comment', () => addComment(ta.value), true),
    h('span', { class: 'ms-hint', style: 'align-self:center' }, selected ? 'Pinned beside the selected object' : 'Pinned to the top-right of the slide')));
  if (!list.length) body.append(h('div', { class: 'empty-panel' }, 'No comments yet.'));
  list.forEach(({ si, c }) => body.append(h('div', { class: 'ms-item' + (c.done ? ' done' : '') },
    h('div', { class: 'm' }, h('b', null, c.author), ' · slide ' + (si + 1) + ' · ' + fmtTime(c.ts) + (c.done ? ' · resolved' : '')),
    h('div', { class: 't' }, c.text),
    h('div', { class: 'a' },
      h('button', { onclick: () => goComment(si, c.id) }, 'Go to'),
      h('button', { onclick: () => toggleResolve(si, c.id) }, c.done ? 'Reopen' : 'Resolve'),
      h('button', { class: 'danger', onclick: () => deleteComment(si, c.id) }, 'Delete')))));
}
function closeComments() { const p = g('msCPanel'); if (p) p.remove(); }
function renderPins(c, s) {
  (s.comments || []).forEach((cm, i) => {
    const pin = h('div', { class: 'ms-pin' + (cm.done ? ' done' : '') + (pinFlash === cm.id ? ' pulse' : ''), 'data-cid': cm.id, title: cm.author + ': ' + cm.text, style: 'left:' + cm.x + '%;top:' + cm.y + '%' }, String(i + 1));
    pin.addEventListener('mousedown', ev => {
      ev.stopPropagation(); ev.preventDefault();
      const r = c.getBoundingClientRect(); let moved = false;
      const mv = m => {
        moved = true;
        cm.x = clamp((m.clientX - r.left) / r.width * 100, 1, 99);
        cm.y = clamp((m.clientY - r.top) / r.height * 100, 4, 99);
        pin.style.left = cm.x + '%'; pin.style.top = cm.y + '%';
      };
      const up = () => {
        document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up);
        if (moved) scheduleSave(); else { openComments(false); pinFlash = cm.id; renderCanvas(); setTimeout(() => { pinFlash = null; }, 2200); }
      };
      document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
    });
    pin.addEventListener('click', ev => ev.stopPropagation());
    c.append(pin);
  });
}
function togglePins() { PINS = !PINS; LS.set('pins', PINS); renderCanvas(); updateStatus(); toast(PINS ? 'Comment pins shown' : 'Comment pins hidden'); }

/* ====================================================================
   11. REVIEW: word count, accessibility, spelling
   ==================================================================== */
const words = s => { const m = String(s || '').trim().match(/\S+/g); return m ? m.length : 0; };
function openWordCount() {
  if (!inEditor()) return;
  const calc = list => {
    let w = 0, c = 0, n = 0;
    list.forEach(s => s.elements.forEach(e => {
      const parts = e.type === 'text' ? [e.text] : e.type === 'table' ? e.cells.flat() : e.type === 'shape' && e.label ? [e.label] : [];
      parts.forEach(t => { w += words(t); c += String(t || '').replace(/\s/g, '').length; });
      n++;
    }));
    return { w, c, n, notes: list.reduce((a, s) => a + words(s.notes), 0) };
  };
  const d = calc(deck.slides), s = calc([slide()]);
  const box = (v, l) => h('div', { class: 'ms-stat' }, h('b', null, String(v)), h('span', null, l));
  const m = makeModal('msWC', 'Word count', 'Includes text boxes, tables and shape text.', [
    h('div', { class: 'ms-stat-grid' }, box(d.w, 'Words (deck)'), box(d.c, 'Characters, no spaces'), box(deck.slides.length, 'Slides'), box(d.n, 'Objects'), box(s.w, 'Words (this slide)'), box(d.notes, 'Words in speaker notes'))
  ], [pill('Close', () => m.classList.remove('show'), true)], 520);
  m.classList.add('show');
}
function a11yIssues() {
  const out = [];
  deck.slides.forEach((s, si) => {
    const bg = s.bgImage || s.bgGradient ? null : toHex(s.bg);
    if (!s.elements.some(e => !e.hidden && e.type === 'text' && (e.text || '').trim())) out.push({ si, msg: 'Slide has no text — add a title for screen readers' });
    s.elements.forEach(e => {
      if (e.hidden) return;
      if ((e.type === 'image' || e.type === 'chart') && !(e.alt || '').trim()) out.push({ si, id: e.id, msg: (e.type === 'image' ? 'Picture' : 'Chart') + ' is missing alt text', fix: true });
      if (e.type === 'text') {
        if (!(e.text || '').trim()) out.push({ si, id: e.id, msg: 'Empty text box' });
        if (e.fontSize < 12) out.push({ si, id: e.id, msg: 'Text is very small (' + e.fontSize + 'px)' });
        if (bg && /^#/.test(e.color || '')) { const cr = contrastRatio(e.color, bg); if (cr < 3) out.push({ si, id: e.id, msg: 'Low colour contrast (' + cr.toFixed(1) + ':1)' }); }
      }
    });
  });
  return out;
}
function openA11y() {
  if (!inEditor()) return;
  const { body } = floatPanel('msA11y', 'Accessibility checker');
  const issues = a11yIssues();
  if (!issues.length) { body.append(h('div', { class: 'empty-panel', style: 'color:var(--ok,#15803D)' }, '✓ No accessibility issues found.')); return; }
  body.append(h('div', { class: 'ms-hint', style: 'margin-bottom:10px' }, issues.length + ' issue' + (issues.length === 1 ? '' : 's') + ' — click one to jump to it.'));
  issues.forEach(it => body.append(h('div', {
    class: 'ms-issue', onclick: () => { current = it.si; selected = it.id || null; msSel.clear(); render(); if (it.fix) setAlt(); }
  }, h('i', null, 'Slide ' + (it.si + 1)), h('span', null, it.msg))));
}
function toggleSpell() { msSpell = !msSpell; LS.set('spell', msSpell); renderCanvas(); toast('Spell check ' + (msSpell ? 'on — misspellings are underlined while editing' : 'off')); }

/* ====================================================================
   12. BACKGROUND FORMATTING
   ==================================================================== */
function downscaleImage(file, maxW, cb) {
  const fr = new FileReader();
  fr.onload = () => {
    const im = new Image();
    im.onload = () => {
      const k = Math.min(1, maxW / im.width), c = document.createElement('canvas');
      c.width = Math.round(im.width * k); c.height = Math.round(im.height * k);
      c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
      cb(c.toDataURL('image/jpeg', 0.85));
    };
    im.onerror = () => toast('Could not read that image');
    im.src = fr.result;
  };
  fr.readAsDataURL(file);
}
function setSlideBgProps(s, p) {
  delete s.bgImage; delete s.bgGradient;
  if (p.type === 'solid') s.bg = p.color;
  else if (p.type === 'gradient') { s.bgGradient = { c1: p.c1, c2: p.c2, angle: +p.angle }; s.bg = p.c1; }
  else if (p.type === 'image' && p.image) s.bgImage = p.image;
}
function openBackground() {
  if (!inEditor()) return;
  const s = slide();
  const st = { type: s.bgImage ? 'image' : s.bgGradient ? 'gradient' : 'solid', color: toHex(s.bg), c1: s.bgGradient ? s.bgGradient.c1 : toHex(s.bg), c2: s.bgGradient ? s.bgGradient.c2 : '#EEF0FE', angle: s.bgGradient ? s.bgGradient.angle : 135, image: s.bgImage || null };
  const type = sel([['solid', 'Solid colour'], ['gradient', 'Gradient'], ['image', 'Picture']], v => { st.type = v; show(); }, 'width:100%;height:36px');
  type.value = st.type;
  const color = h('input', { type: 'color', value: st.color, oninput: () => { st.color = color.value; } });
  const c1 = h('input', { type: 'color', value: st.c1, oninput: () => { st.c1 = c1.value; } });
  const c2 = h('input', { type: 'color', value: st.c2, oninput: () => { st.c2 = c2.value; } });
  const angle = sel([['0', 'Bottom → top'], ['90', 'Left → right'], ['135', 'Diagonal ↘'], ['180', 'Top → bottom'], ['225', 'Diagonal ↙']], v => { st.angle = +v; }, 'width:100%;height:36px');
  angle.value = String(st.angle);
  const file = h('input', { type: 'file', accept: 'image/*' });
  const fname = h('div', { class: 'ms-hint', style: 'margin-top:6px' }, st.image ? 'A picture is set for this slide.' : 'Pick an image — it is resized to keep storage small.');
  file.addEventListener('change', () => { const f = file.files[0]; if (f) downscaleImage(f, 1920, d => { st.image = d; fname.textContent = 'Picture ready — click Apply.'; }); });
  const sBox = field('Colour', color), gBox = h('div', null, h('div', { class: 'row' }, field('Colour 1', c1), field('Colour 2', c2)), field('Direction', angle)), iBox = h('div', null, field('Picture', file), fname);
  function show() { sBox.style.display = st.type === 'solid' ? '' : 'none'; gBox.style.display = st.type === 'gradient' ? '' : 'none'; iBox.style.display = st.type === 'image' ? '' : 'none'; }
  const m = makeModal('msBG', 'Format background', 'Set a solid colour, gradient or picture for this slide.', [field('Fill type', type), sBox, gBox, iBox], [
    pill('Reset', () => { snapshot(); delete s.bgImage; delete s.bgGradient; s.bg = themeOf(deck.theme).bg; render(); scheduleSave(); m.classList.remove('show'); }),
    pill('Apply', () => { snapshot(); setSlideBgProps(s, st); render(); scheduleSave(); m.classList.remove('show'); }),
    pill('Apply to all', () => { snapshot(); deck.slides.forEach(x => setSlideBgProps(x, st)); render(); scheduleSave(); m.classList.remove('show'); }, true)
  ], 480);
  show();
  m.classList.add('show');
}
function applyBgToAll() {
  const s = slide();
  snapshot();
  deck.slides.forEach(x => {
    if (x === s) return;
    delete x.bgImage; delete x.bgGradient;
    x.bg = s.bg;
    if (s.bgGradient) x.bgGradient = cloneJSON(s.bgGradient);
    if (s.bgImage) x.bgImage = s.bgImage;
  });
  render(); scheduleSave(); toast('Background applied to all slides');
}

/* ====================================================================
   13. TRANSITION TIMING
   ==================================================================== */
let transDurIn, advIn;
function syncTransitionUI() {
  if (!transDurIn || !deck) return;
  const s = slide();
  if (!s) return;
  transDurIn.value = s.transDur || 0.4;
  advIn.value = s.advance || '';
}

/* ====================================================================
   14. SLIDE SHOW ENHANCEMENTS
   ==================================================================== */
const PS = Object.assign({ timer: true, click: true, loop: false, timings: true }, LS.get('show', {}));
const PSR = { laser: false, pen: null, penColor: '#EF4444', blank: null, end: false, digits: '', slideTs: 0, slideIdx: null, times: {}, rehearse: false, autoT: null, tick: null, idleT: null, drawing: false, startTs: 0 };
const savePS = () => LS.set('show', { timer: PS.timer, click: PS.click, loop: PS.loop, timings: PS.timings });
const firstVisible = () => { const i = deck.slides.findIndex(s => !s.hidden); return i < 0 ? 0 : i; };

function ensureOverlays() {
  const P = g('present');
  if (g('msTimer')) return;
  const stage = h('div', { id: 'msPresentStage' });
  const ps = g('presentSlide');
  ps.parentNode.insertBefore(stage, ps);
  stage.append(ps);
  P.append(
    h('div', { id: 'msTimer' }, '0:00'),
    h('div', { id: 'msLaser' }),
    h('canvas', { id: 'msPen' }),
    h('div', { id: 'msBlank', onclick: ev => { ev.stopPropagation(); if (PSR.end) exitPresentation(); else setBlank(null); } }),
    h('div', { id: 'msJump' }));
  P.addEventListener('mousemove', onPresentMove);
  P.addEventListener('click', onPresentClick);
  P.addEventListener('contextmenu', e => { e.preventDefault(); prevPresent(); });
  wirePen(g('msPen'));
  window.addEventListener('resize', () => { if (presenting()) { fitStage(); sizePen(); } });
}
function fitStage() {
  const st = g('msPresentStage'); if (!st) return;
  const { w, h: hh } = currentPageSize();
  const DW = 1040, DH = DW * hh / w;
  st.style.width = DW + 'px'; st.style.height = DH + 'px';
  st.style.transform = 'scale(' + Math.min(innerWidth / DW, innerHeight / DH) + ')';
}
function sizePen() { const c = g('msPen'); if (c) { c.width = innerWidth; c.height = innerHeight; } }
function clearPen() { const c = g('msPen'); if (c) c.getContext('2d').clearRect(0, 0, c.width, c.height); }
function wirePen(c) {
  let last = null;
  c.addEventListener('pointerdown', e => {
    if (!PSR.pen) return;
    e.preventDefault(); e.stopPropagation();
    try { c.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
    last = [e.clientX, e.clientY]; PSR.drawing = true;
  });
  c.addEventListener('pointermove', e => {
    if (!PSR.drawing || !last) return;
    const x = c.getContext('2d');
    x.lineCap = 'round'; x.lineJoin = 'round';
    if (PSR.pen === 'hl') { x.globalAlpha = 0.35; x.lineWidth = 22; x.strokeStyle = '#FDE047'; }
    else { x.globalAlpha = 1; x.lineWidth = 3.5; x.strokeStyle = PSR.penColor; }
    x.beginPath(); x.moveTo(last[0], last[1]); x.lineTo(e.clientX, e.clientY); x.stroke();
    last = [e.clientX, e.clientY];
  });
  const end = () => { last = null; PSR.drawing = false; };
  c.addEventListener('pointerup', end); c.addEventListener('pointercancel', end);
}
function onPresentMove(e) {
  const P = g('present');
  P.classList.remove('ms-idle');
  clearTimeout(PSR.idleT);
  PSR.idleT = setTimeout(() => P.classList.add('ms-idle'), 2500);
  if (PSR.laser) { const d = g('msLaser'); d.style.left = e.clientX + 'px'; d.style.top = e.clientY + 'px'; }
}
function onPresentClick(e) {
  if (e.target.closest('.presentbar') || e.target.closest('.presenter-notes')) return;
  if (PSR.pen) return;
  if (PSR.end) { exitPresentation(); return; }
  if (PSR.blank || !PS.click) return;
  nextPresent();
}
function flash(txt) {
  const j = g('msJump'); if (!j) return;
  j.textContent = txt; j.style.display = 'block';
  clearTimeout(flash.t); flash.t = setTimeout(() => { j.style.display = 'none'; }, 1300);
}
function setBlank(color, msg) {
  const b = g('msBlank'); if (!b) return;
  if (!color || PSR.blank === color && !msg) { PSR.blank = null; PSR.end = false; b.style.display = 'none'; b.textContent = ''; return; }
  PSR.blank = color; b.style.background = color; b.style.display = 'flex'; b.textContent = msg || '';
}
function toggleLaser() {
  PSR.laser = !PSR.laser;
  g('present').classList.toggle('ms-laser', PSR.laser);
  if (PSR.laser) { PSR.pen = null; g('present').classList.remove('ms-pen'); }
  flash(PSR.laser ? 'Laser pointer on (L)' : 'Laser pointer off');
}
function setPen(mode) {
  PSR.pen = PSR.pen === mode ? null : mode;
  g('present').classList.toggle('ms-pen', !!PSR.pen);
  if (PSR.pen) { PSR.laser = false; g('present').classList.remove('ms-laser'); }
  flash(PSR.pen ? (mode === 'hl' ? 'Highlighter (H) — E to erase' : 'Pen (P) — E to erase') : 'Pointer');
}
function updateTimer() {
  const t = g('msTimer'); if (!t) return;
  const sec = Math.floor((Date.now() - PSR.startTs) / 1000);
  t.textContent = Math.floor(sec / 60) + ':' + String(sec % 60).padStart(2, '0');
}
function resetShowState() {
  PSR.laser = false; PSR.pen = null; PSR.blank = null; PSR.end = false; PSR.digits = '';
  PSR.times = {}; PSR.slideIdx = null; PSR.drawing = false;
  clearTimeout(PSR.autoT);
  const P = g('present');
  if (P) P.classList.remove('ms-laser', 'ms-pen', 'ms-idle');
  const b = g('msBlank'); if (b) { b.style.display = 'none'; b.textContent = ''; }
  clearPen();
}
function recordRehearse() {
  if (!PSR.rehearse || PSR.slideIdx == null) return;
  PSR.times[PSR.slideIdx] = (PSR.times[PSR.slideIdx] || 0) + (Date.now() - PSR.slideTs) / 1000;
}
function afterShowPresent() {
  ensureOverlays();
  fitStage(); sizePen(); clearPen();
  setBlank(null);
  const s = deck.slides[presentIndex], p = g('presentSlide');
  applyBg(p, s, false);
  s.elements.forEach(e => {
    if (e.hidden) return;
    const wrap = p.querySelector('[data-id="' + e.id + '"]');
    if (!wrap) return;
    applyExtras(wrap, wrap.firstElementChild, e);
    if (e.alt) wrap.setAttribute('aria-label', e.alt);
    if (e.link && safeUrl(e.link)) {
      wrap.style.cursor = 'pointer';
      wrap.addEventListener('click', ev => { ev.stopPropagation(); window.open(e.link, '_blank', 'noopener'); });
    }
  });
  p.style.animationDuration = (s.transDur || 0.4) + 's';
  g('present').classList.toggle('ms-timer', !!PS.timer);
  PSR.slideIdx = presentIndex; PSR.slideTs = Date.now();
  clearTimeout(PSR.autoT);
  if (PS.timings && !PSR.rehearse && s.advance > 0) {
    PSR.autoT = setTimeout(() => { if (presenting() && !PSR.blank && !PSR.pen) nextPresent(); }, s.advance * 1000);
  }
}
const _showPresent = showPresent;
showPresent = function () {
  recordRehearse();
  _showPresent.apply(this, arguments);
  try { afterShowPresent(); } catch (err) { console.error('[ms-ppt]', err); }
};
const _startPresentation = startPresentation;
startPresentation = function () {
  ensureOverlays();
  resetShowState();
  _startPresentation.apply(this, arguments);
  PSR.startTs = Date.now();
  clearInterval(PSR.tick); PSR.tick = setInterval(updateTimer, 1000); updateTimer();
};
const _nextPresent = nextPresent;
nextPresent = function () {
  if (PSR.end) { exitPresentation(); return; }
  if (presentStep + 1 < presentGroups.length) return _nextPresent.apply(this, arguments);
  let i = presentIndex + 1;
  while (i < deck.slides.length && deck.slides[i].hidden) i++;
  if (i >= deck.slides.length) {
    if (PS.loop) { i = firstVisible(); }
    else { PSR.end = true; setBlank('#000', 'End of slide show — click to exit'); return; }
  }
  presentIndex = i;
  showPresent();
};
const _prevPresent = prevPresent;
prevPresent = function () {
  if (PSR.end) { setBlank(null); return; }
  let i = presentIndex - 1;
  while (i >= 0 && deck.slides[i].hidden) i--;
  if (i < 0) return;
  presentIndex = i;
  showPresent();
};
const _exitPresentation = exitPresentation;
exitPresentation = function () {
  recordRehearse();
  const rehearse = PSR.rehearse, times = PSR.times;
  clearInterval(PSR.tick); PSR.tick = null;
  resetShowState();
  _exitPresentation.apply(this, arguments);
  if (rehearse) {
    PSR.rehearse = false;
    setTimeout(() => {
      const total = Object.values(times).reduce((a, b) => a + b, 0);
      if (!total) return;
      if (confirm('Keep the new slide timings? Total time: ' + Math.floor(total / 60) + 'm ' + Math.round(total % 60) + 's')) {
        snapshot();
        deck.slides.forEach((s, i) => { if (times[i] != null) s.advance = Math.max(1, Math.round(times[i])); });
        PS.timings = true; savePS();
        render(); scheduleSave();
      }
    }, 250);
  }
};
function msStart(i) {
  if (!inEditor()) return toast('Open a presentation first');
  const keep = current;
  current = i;
  try { startPresentation(); } finally { current = keep; }
}
function startRehearse() {
  if (!inEditor()) return;
  PSR.rehearse = true;
  msStart(firstVisible());
  toast('Rehearsing — advance at your real pace, then exit to save timings');
}
function jumpTo(idx) {
  idx = clamp(idx, 0, deck.slides.length - 1);
  presentIndex = idx;
  showPresent();
}
function onPresentKey(e) {
  const k = e.key;
  if (PSR.blank && !PSR.end && !/^(b|w|\.|,)$/i.test(k) && k.length === 1 + 0 && false) return;
  if (/^[0-9]$/.test(k)) { PSR.digits += k; flash('Go to slide ' + PSR.digits + ' ↵'); return; }
  if (k === 'Enter' && PSR.digits) { e.preventDefault(); const n = +PSR.digits; PSR.digits = ''; jumpTo(n - 1); return; }
  if (k === 'Backspace' && PSR.digits) { PSR.digits = PSR.digits.slice(0, -1); flash('Go to slide ' + (PSR.digits || '…')); return; }
  const lk = k.toLowerCase();
  if (PSR.blank && !PSR.end && lk !== 'b' && lk !== 'w' && lk !== '.' && lk !== ',' && k !== 'Escape') { setBlank(null); return; }
  switch (lk) {
    case 'b': case '.': setBlank('#000'); break;
    case 'w': case ',': setBlank('#fff'); break;
    case 'l': toggleLaser(); break;
    case 'p': setPen('pen'); break;
    case 'h': setPen('hl'); break;
    case 'e': clearPen(); flash('Annotations erased'); break;
    case 't': PS.timer = !PS.timer; savePS(); g('present').classList.toggle('ms-timer', PS.timer); break;
    case 'home': e.preventDefault(); jumpTo(firstVisible()); break;
    case 'end': { e.preventDefault(); let i = deck.slides.length - 1; while (i > 0 && deck.slides[i].hidden) i--; jumpTo(i); break; }
    case 'pagedown': case 'enter': e.preventDefault(); nextPresent(); break;
    case 'pageup': case 'backspace': e.preventDefault(); prevPresent(); break;
    case 'n': nextPresent(); break;
    default: break;
  }
}

/* ====================================================================
   15. PRINT / PDF
   ==================================================================== */
function printDeck(mode) {
  const pages = deck.slides.filter(s => !s.hidden);
  if (!pages.length) return toast('No visible slides to print');
  const { w, h: hh } = currentPageSize();
  const DW = 1040, DH = DW * hh / w, notes = mode === 'notes';
  const pageW = notes ? 8.5 : w, pageH = notes ? 11 : hh, slideW = notes ? 7 : w;
  const scale = slideW * 96 / DW;
  const oldRoot = g('msPrintRoot'); if (oldRoot) oldRoot.remove();
  const oldStyle = g('msPrintStyle'); if (oldStyle) oldStyle.remove();
  const root = h('div', { id: 'msPrintRoot' });
  pages.forEach(s => {
    const inner = h('div', { style: 'width:' + DW + 'px;height:' + DH + 'px;transform:scale(' + scale + ');transform-origin:0 0;background:' + s.bg + ';position:relative;overflow:hidden' });
    s.elements.forEach(e => { if (!e.hidden) presentElement(inner, e); });
    s.elements.forEach(e => {
      if (e.hidden) return;
      const wrap = inner.querySelector('[data-id="' + e.id + '"]');
      if (!wrap) return;
      const d = wrap.firstElementChild;
      if (d) { d.style.opacity = e.opacity == null ? 1 : e.opacity; d.style.animation = 'none'; }
      applyExtras(wrap, d, e);
    });
    applyBg(inner, s, false);
    const box = h('div', { class: 'ms-ps', style: 'width:' + slideW + 'in;height:' + (slideW * hh / w) + 'in' }, inner);
    const page = h('div', { class: 'ms-pp' + (notes ? ' notes' : '') }, box);
    if (notes) page.append(h('div', { class: 'ms-pn' }, (s.notes || '').trim() || '—'));
    root.append(page);
  });
  document.body.append(root);
  const st = h('style', { id: 'msPrintStyle' });
  st.textContent = '@page{size:' + pageW + 'in ' + pageH + 'in;margin:0}';
  document.head.append(st);
  const cleanup = () => { root.remove(); st.remove(); window.removeEventListener('afterprint', cleanup); };
  window.addEventListener('afterprint', cleanup);
  setTimeout(() => window.print(), 350);
}
function openPrint() {
  if (!inEditor()) return;
  const m = makeModal('msPrint', 'Print / Save as PDF', 'In the print dialog choose “Save as PDF” as the destination. Hidden slides are skipped.', [
    h('div', { style: 'display:grid;gap:10px' },
      h('button', { class: 'pill-btn primary', style: 'padding:14px', onclick: () => { m.classList.remove('show'); printDeck('slides'); } }, '🖨  Slides — one per page, exact slide size'),
      h('button', { class: 'pill-btn', style: 'padding:14px', onclick: () => { m.classList.remove('show'); printDeck('notes'); } }, '🗒  Notes pages — slide with speaker notes below'))
  ], [pill('Cancel', () => m.classList.remove('show'))], 500);
  m.classList.add('show');
}

/* ====================================================================
   16. EXPORT PATCH (.pptx) — keeps new features in the exported file
   ==================================================================== */
const polyPts = s => s.split(',').map(p => p.trim().split(/\s+/).map(v => parseFloat(v)).join(',')).join(' ');
const svgUri = svg => 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
function gradDef(id, gr) {
  const a = (gr.angle == null ? 90 : gr.angle) * Math.PI / 180, s = Math.sin(a), c = Math.cos(a), f = n => n.toFixed(4);
  return '<linearGradient id="' + id + '" x1="' + f(0.5 - 0.5 * s) + '" y1="' + f(0.5 + 0.5 * c) + '" x2="' + f(0.5 + 0.5 * s) + '" y2="' + f(0.5 - 0.5 * c) + '"><stop offset="0" stop-color="' + gr.c1 + '"/><stop offset="1" stop-color="' + gr.c2 + '"/></linearGradient>';
}
function coverImage(src, w, hh) {
  return new Promise(res => {
    const im = new Image();
    im.onload = () => {
      const W = 1600, H = Math.round(W * hh / w), c = document.createElement('canvas');
      c.width = W; c.height = H;
      const x = c.getContext('2d'), k = Math.max(W / im.width, H / im.height), dw = im.width * k, dh = im.height * k;
      x.drawImage(im, (W - dw) / 2, (H - dh) / 2, dw, dh);
      res(c.toDataURL('image/jpeg', 0.88));
    };
    im.onerror = () => res(null);
    im.src = src;
  });
}
function shapeSvg(e) {
  let fill = !e.fill || e.fill === 'transparent' ? 'none' : e.fill, defs = '';
  if (e.grad) { defs = gradDef('gf', e.grad); fill = 'url(#gf)'; }
  const poly = POLY_ALL[e.shape];
  const stroke = poly ? '' : ' stroke="' + e.stroke + '" stroke-width="' + (e.strokeWidth || 0) + '" vector-effect="non-scaling-stroke"';
  let geo;
  if (poly) geo = '<polygon points="' + polyPts(poly) + '" fill="' + fill + '"/>';
  else if (e.shape === 'ellipse') geo = '<ellipse cx="50" cy="50" rx="50" ry="50" fill="' + fill + '"' + stroke + '/>';
  else if (e.shape === 'roundRect') geo = '<rect width="100" height="100" rx="14" fill="' + fill + '"' + stroke + '/>';
  else geo = '<rect width="100" height="100" rx="2" fill="' + fill + '"' + stroke + '/>';
  const tr = (e.flipX || e.flipY) ? ' transform="translate(' + (e.flipX ? 100 : 0) + ' ' + (e.flipY ? 100 : 0) + ') scale(' + (e.flipX ? -1 : 1) + ' ' + (e.flipY ? -1 : 1) + ')"' : '';
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none" width="400" height="400"><defs>' + defs + '</defs><g' + tr + '>' + geo + '</g></svg>';
}
const imgEl = (src, x, y, w, hh, rot, op) => ({ id: uid(), type: 'image', src, x, y, w, h: hh, opacity: op == null ? 1 : op, rotation: rot || 0, locked: false, hidden: false, anim: defaultAnim() });
const boxKey = (x, y, w, hh) => [x, y, w, hh].map(n => n.toFixed(3)).join('|');

async function buildExportDeck(real) {
  const page = currentPageSize();
  const d = cloneJSON(real);
  d.slides = d.slides.filter(s => !s.hidden);
  const extras = [];
  for (const s of d.slides) {
    const ex = {}; extras.push(ex);
    const out = [];
    if (s.bgImage) { const du = await coverImage(s.bgImage, page.w, page.h); if (du) out.push(imgEl(du, 0, 0, 100, 100)); }
    else if (s.bgGradient) {
      out.push(imgEl(svgUri('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" preserveAspectRatio="none" width="400" height="400"><defs>' + gradDef('bg', s.bgGradient) + '</defs><rect width="100" height="100" fill="url(#bg)"/></svg>'), 0, 0, 100, 100));
      s.bg = s.bgGradient.c1;
    }
    for (const e of s.elements) {
      if (e.hidden) continue;
      const px = e.x / 100 * page.w, py = e.y / 100 * page.h, pw = e.w / 100 * page.w, ph = e.h / 100 * page.h;
      if (e.type === 'text') {
        if (e.boxFill || (e.boxStroke && e.boxStrokeW)) {
          const r = shapeEl('rect', e.x, e.y, e.w, e.h, e.boxFill || '#FFFFFF');
          r.stroke = e.boxStroke || '#FFFFFF'; r.strokeWidth = e.boxStroke ? (e.boxStrokeW || 0) : 0;
          r.opacity = e.boxFill ? (e.opacity == null ? 1 : e.opacity) : 0; r.rotation = e.rotation || 0;
          out.push(r);
        }
        out.push(e);
        const o = {};
        if (e.valign && e.valign !== 'top') o.valign = e.valign === 'middle' ? 'middle' : 'bottom';
        if (e.lineH) o.lineSpacingMultiple = e.lineH;
        if (e.letterSp) o.charSpacing = +(e.letterSp * (e.fontSize || 20) * 0.72).toFixed(2);
        if (e.link && safeUrl(e.link)) o.hyperlink = { url: e.link };
        if (e.tshadow) o.shadow = { type: 'outer', angle: 45, blur: 3, offset: 2, color: '000000', opacity: 0.35 };
        if (Object.keys(o).length) ex['addText:' + boxKey(px, py, pw, ph)] = o;
      } else if (e.type === 'shape') {
        const asSvg = MS_POLY[e.shape] || (e.grad && e.shape !== 'line' && !COMPOUND_SHAPES.has(e.shape) && e.shape !== 'donut' && (POLY_ALL[e.shape] || ['rect', 'roundRect', 'ellipse'].includes(e.shape)));
        if (asSvg) out.push(imgEl(svgUri(shapeSvg(e)), e.x, e.y, e.w, e.h, e.rotation, e.opacity));
        else {
          if (e.fill === 'transparent' && e.shape !== 'line') { e.fill = '#FFFFFF'; e.opacity = 0; }
          out.push(e);
          const o = {};
          if (e.shadow && !COMPOUND_SHAPES.has(e.shape)) o.shadow = { type: 'outer', angle: 45, blur: 6, offset: 4, color: '000000', opacity: 0.35 };
          if (e.flipX) o.flipH = true;
          if (e.flipY) o.flipV = true;
          if (Object.keys(o).length) ex['addShape:' + boxKey(px, py, pw, ph)] = o;
        }
        if (e.label) {
          const lt = textEl(e.label, e.x, e.y, e.w, e.h, e.labelSize || 20, { align: 'center', bold: true, color: e.labelColor || autoLabelColor(e), rotation: e.rotation || 0 });
          out.push(lt);
          ex['addText:' + boxKey(px, py, pw, ph)] = Object.assign({}, ex['addText:' + boxKey(px, py, pw, ph)], { valign: 'middle' });
        }
      } else out.push(e);
    }
    s.elements = out;
  }
  return { deck: d, extras };
}
function patchPptx(Base, extras) {
  const P = function () {
    const inst = new Base(...arguments);
    const origAdd = inst.addSlide.bind(inst);
    let idx = -1;
    inst.addSlide = function () {
      const sl = origAdd.apply(null, arguments);
      idx++;
      const ex = extras[idx] || {};
      ['addText', 'addShape'].forEach(fn => {
        const o = sl[fn];
        if (typeof o !== 'function') return;
        sl[fn] = function () {
          const args = [...arguments], opts = args[1];
          if (opts && typeof opts === 'object') {
            const add = ex[fn + ':' + boxKey(opts.x || 0, opts.y || 0, opts.w || 0, opts.h || 0)];
            if (add) args[1] = Object.assign({}, opts, add);
          }
          return o.apply(sl, args);
        };
      });
      return sl;
    };
    return inst;
  };
  P.prototype = Base.prototype;
  Object.keys(Base).forEach(k => { try { P[k] = Base[k]; } catch (e) { /* ignore */ } });
  return P;
}
const _export = exportPPTX;
exportPPTX = async function () {
  const real = deck, Orig = window.PptxGenJS;
  let built = null;
  try { built = await buildExportDeck(real); } catch (err) { console.warn('[ms-ppt] export prep failed, using plain export', err); }
  if (!built || !Orig) return _export.apply(this, arguments);
  try {
    deck = built.deck;
    window.PptxGenJS = patchPptx(Orig, built.extras);
    return await _export.apply(this, arguments);
  } finally {
    deck = real;
    window.PptxGenJS = Orig;
  }
};

/* ====================================================================
   17. CLIPBOARD IMAGES & TEXT, KEYBOARD SHORTCUTS
   ==================================================================== */
function addImageFile(f) {
  if (!f || !inEditor()) return;
  const fr = new FileReader();
  fr.onload = () => {
    snapshot();
    const e = { id: uid(), type: 'image', src: fr.result, x: 20, y: 20, w: 55, h: 45, opacity: 1, rotation: 0, locked: false, hidden: false, anim: defaultAnim() };
    slide().elements.push(e); selected = e.id; msSel.clear();
    render(); scheduleSave();
  };
  fr.readAsDataURL(f);
}
document.addEventListener('paste', ev => {
  if (!inEditor() || typing() || modalOpen()) return;
  const cd = ev.clipboardData;
  const imgItem = cd && [...cd.items].find(i => i.type.indexOf('image/') === 0);
  if (imgItem) { ev.preventDefault(); addImageFile(imgItem.getAsFile()); return; }
  if (clipboard || msClip) { ev.preventDefault(); pasteClipboard(); return; }
  const txt = cd && cd.getData('text/plain');
  if (txt && txt.trim()) {
    ev.preventDefault();
    snapshot();
    const e = textEl(txt.trim().slice(0, 4000), 10, 25, 80, 40, 20, { color: themeOf(deck.theme).ink });
    slide().elements.push(e); selected = e.id; msSel.clear();
    render(); scheduleSave();
  }
});

function keepEditing(fn) {
  const id = editingId;
  fn();
  if (id) requestAnimationFrame(() => { enterTextEdit(id); });
}
window.addEventListener('keydown', ev => {
  if (presenting()) { onPresentKey(ev); return; }
  if (!inEditor() || modalOpen()) return;
  const k = ev.key, mod = ev.ctrlKey || ev.metaKey, lk = k.length === 1 ? k.toLowerCase() : k;
  if (k === 'Escape') { msSel.clear(); closeMenu(); cancelPainter(); return; }
  if (k === 'F5') { ev.preventDefault(); ev.stopImmediatePropagation(); msStart(ev.shiftKey ? current : firstVisible()); return; }
  if (mod && !ev.shiftKey && !ev.altKey && lk === 'f') { ev.preventDefault(); ev.stopImmediatePropagation(); openFind(false); return; }
  if (mod && !ev.shiftKey && !ev.altKey && lk === 'h') { ev.preventDefault(); ev.stopImmediatePropagation(); openFind(true); return; }
  if (mod && lk === 'p' && !ev.shiftKey) { ev.preventDefault(); ev.stopImmediatePropagation(); openPrint(); return; }
  if (mod && !ev.shiftKey && (lk === 'b' || lk === 'i' || lk === 'u') && (!typing() || editingId)) {
    const e = selected && byId(selected);
    if (e && e.type === 'text') {
      ev.preventDefault(); ev.stopImmediatePropagation();
      keepEditing(() => toggleText({ b: 'bold', i: 'italic', u: 'underline' }[lk]));
    }
    return;
  }
  if (typing()) return;
  if (mod && lk === 'v') { ev.stopImmediatePropagation(); return; } /* the 'paste' event handles it */
  if (mod && !ev.shiftKey && lk === 'a') { ev.preventDefault(); selectAll(); return; }
  if (mod && !ev.shiftKey && lk === 'x') { ev.preventDefault(); ev.stopImmediatePropagation(); cutSel(); return; }
  if (mod && !ev.shiftKey && lk === 'm') { ev.preventDefault(); ev.stopImmediatePropagation(); addSlide('content'); return; }
  if (mod && ev.shiftKey && (k === '>' || k === '.')) { ev.preventDefault(); growFont(1); return; }
  if (mod && ev.shiftKey && (k === '<' || k === ',')) { ev.preventDefault(); growFont(-1); return; }
  if (mod && (k === ']' || k === '[')) {
    ev.preventDefault();
    if (k === ']') (ev.shiftKey ? bringToFront : bringForward)(); else (ev.shiftKey ? sendToBack : sendBackward)();
    return;
  }
  if (drawTool) return;
  if (k === 'Backspace' && selected) { ev.preventDefault(); ev.stopImmediatePropagation(); deleteSelected(); return; }
  if (k.indexOf('Arrow') === 0 && selected) {
    ev.preventDefault(); ev.stopImmediatePropagation();
    const step = ev.shiftKey ? 2 : (mod || ev.altKey) ? 0.1 : 0.5;
    nudge(k === 'ArrowLeft' ? -step : k === 'ArrowRight' ? step : 0, k === 'ArrowUp' ? -step : k === 'ArrowDown' ? step : 0);
    return;
  }
  if (k === 'Enter' && selected) {
    const e = byId(selected);
    if (e && !e.locked) {
      if (e.type === 'text') { ev.preventDefault(); enterTextEdit(selected); }
      else if (e.type === 'table') { ev.preventDefault(); enterTableEdit(selected); }
      else if (e.type === 'math') { ev.preventDefault(); openMathEditor(selected); }
    }
    return;
  }
  if (k === 'Tab' && document.activeElement === document.body) {
    const vis = slide().elements.filter(e => !e.hidden);
    if (!vis.length) return;
    ev.preventDefault();
    const i = vis.findIndex(e => e.id === selected), n = vis.length;
    const next = vis[(i + (ev.shiftKey ? -1 : 1) + n) % n];
    editingId = null; msSel.clear(); selected = next.id;
    render();
    return;
  }
  if ((k === 'PageDown' || k === 'PageUp') && !selected) {
    ev.preventDefault();
    const to = clamp(current + (k === 'PageDown' ? 1 : -1), 0, deck.slides.length - 1);
    if (to !== current) { current = to; render(); }
  }
}, true);

/* ====================================================================
   18. STATUS BAR
   ==================================================================== */
let stZoom, stZoomLabel;
function syncZoomUI() { if (stZoom) { stZoom.value = Math.round(zoomLevel * 100); stZoomLabel.textContent = Math.round(zoomLevel * 100) + '%'; } }
let stInfo, stNotes, stComments;
function buildStatusBar() {
  stInfo = h('span', null, 'Slide 1 of 1');
  stNotes = h('button', { title: 'Show / hide speaker notes', onclick: () => toggleNotes() }, '🗒 Notes');
  stComments = h('button', { title: 'Open the comments panel', onclick: () => openComments(false) }, '💬 Comments');
  stZoom = h('input', { type: 'range', min: '50', max: '200', step: '5', value: '100', title: 'Zoom' });
  stZoom.addEventListener('input', () => setZoom(+stZoom.value / 100));
  stZoomLabel = h('span', { style: 'width:38px;text-align:right' }, '100%');
  g('editorView').append(h('div', { class: 'ms-status' },
    stInfo, stNotes, stComments, h('span', { class: 'sp' }),
    h('span', { class: 'ms-hint', style: 'max-width:none' }, 'Shift+click or drag a box to select several · Ctrl+scroll to zoom · F5 to present'),
    h('button', { title: 'Zoom out', onclick: () => setZoom(zoomLevel - 0.1) }, '−'), stZoom,
    h('button', { title: 'Zoom in', onclick: () => setZoom(zoomLevel + 0.1) }, '+'), stZoomLabel,
    h('button', { title: 'Fit to window', onclick: () => fitCanvas() }, '⤢ Fit')));
}
function updateStatus() {
  if (!stInfo || !deck) return;
  const open = allComments().filter(x => !x.c.done).length;
  stInfo.textContent = 'Slide ' + (current + 1) + ' of ' + deck.slides.length + (slide().hidden ? ' (hidden)' : '');
  stNotes.classList.toggle('on', !!notesOpen);
  stComments.textContent = '💬 Comments' + (open ? ' (' + open + ')' : '');
  syncZoomUI();
}

/* ====================================================================
   19. RIBBON UI
   ==================================================================== */
const WORDART = [
  ['grad', 'Gradient', 'background:linear-gradient(90deg,#4F46E5,#EC4899);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent'],
  ['gold', 'Gold', 'background:linear-gradient(180deg,#FDE68A,#F59E0B 55%,#B45309);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent'],
  ['ice', 'Ice', 'background:linear-gradient(180deg,#38BDF8,#0369A1);-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent'],
  ['neon', 'Neon glow', 'background:#12141C;color:#fff;border-radius:6px;text-shadow:0 0 6px #fff,0 0 12px #8B7CF8'],
  ['outline', 'Outline', '-webkit-text-fill-color:transparent;-webkit-text-stroke:1.2px #12141C'],
  ['pop', 'Shadow pop', 'color:#4F46E5;text-shadow:2px 2px 0 rgba(0,0,0,.25)']
];
const MS_SHAPE_BTNS = [['rtriangle', '◺', 'Right triangle'], ['octagon', '⯃', 'Octagon'], ['star4', '✦', '4-point star'], ['star6', '✶', '6-point star'], ['star8', '✷', '8-point star'], ['arrowUp', '⬆', 'Up arrow'], ['arrowDown', '⬇', 'Down arrow'], ['arrowLeft', '⬅', 'Left arrow'], ['bolt', '⚡', 'Lightning bolt'], ['banner', '▰', 'Banner']];

function buildHome() {
  const p = document.querySelector('[data-panel="home"]');
  const layoutSel = sel([['', 'Layout…']].concat(Object.keys(LAYOUTS).map(k => [k, LAYOUTS[k].name])), v => { if (v) { applyLayout(v); layoutSel.value = ''; } }, 'width:150px', 'Apply a layout to this slide (keeps your text)');
  p.append(
    grp('Editing', row(lb('🔍', 'Find', 'Find (Ctrl+F)', () => openFind(false), 50), lb('⇄', 'Replace', 'Find & Replace (Ctrl+H)', () => openFind(true), 56), lb('▣', 'Select all', 'Select all (Ctrl+A)', selectAll, 56))),
    grp('Painter', row(lb('🖌', 'Format', 'Format Painter — copy one object\'s style to another', startPainter, 56))),
    grp('List', row(rb('•≡', 'Bullets', () => toggleList('b')), rb('1.', 'Numbering', () => toggleList('n'), 'font-weight:700'), rb('→≡', 'Increase indent', () => indentText(1)), rb('←≡', 'Decrease indent', () => indentText(-1)))),
    grp('Layout', row(layoutSel)),
    grp('Slide tools', row(rb('👁', 'Hide / unhide slide', () => toggleHide()), rb('⬆', 'Move slide up', () => moveSlide(current, current - 1)), rb('⬇', 'Move slide down', () => moveSlide(current, current + 1)), rb('🎨', 'Format background…', openBackground)))
  );
}
function buildInsert() {
  const p = document.querySelector('[data-panel="insert"]');
  const last = p.lastElementChild; if (last) last.style.borderRight = '';
  p.append(
    grp('More shapes', row(...MS_SHAPE_BTNS.slice(0, 5).map(([k, ic, t]) => rb(ic, t, () => addShapeMs(k)))), row(...MS_SHAPE_BTNS.slice(5).map(([k, ic, t]) => rb(ic, t, () => addShapeMs(k))))),
    grp('Text', row(lb('Ω', 'Symbol', 'Insert a symbol', openSymbols, 52), lb('📅', 'Date', 'Insert today\'s date', insertDate, 52), lb('#', 'Header & footer', 'Slide numbers, footer, date', openHF, 64))),
    grp('Links & notes', row(lb('🔗', 'Link', 'Add a hyperlink to the selected object', setLink, 52), lb('💬', 'Comment', 'Add a comment', () => openComments(true), 62)))
  );
}
function buildFormat() {
  const p = addTab('format', 'Format', 'design');
  const c2 = colorIn('Gradient end colour', '#EC4899', () => {});
  const ang = sel([['90', '→'], ['0', '↑'], ['135', '↘'], ['180', '↓'], ['225', '↙']], () => {}, 'width:54px', 'Gradient direction');
  const lsel = sel([['10', 'Size 10'], ['14', '14'], ['18', '18'], ['20', '20'], ['24', '24'], ['32', '32'], ['40', '40']], v => setLabelSize(v), 'width:76px', 'Text size inside a shape (double-click a shape to type)');
  p.append(
    grp('Shape fill & outline',
      row(colorIn('Fill colour', '#4F46E5', setFill), rb('∅', 'No fill', noFill), c2, ang, rb('◐', 'Apply gradient', () => setGradient(c2.value, ang.value)), rb('■', 'Solid fill (remove gradient)', clearGradient)),
      row(colorIn('Outline colour', '#3730A3', setOutlineColor), sel([['0', 'No line'], ['1', '1 px'], ['2', '2 px'], ['3', '3 px'], ['4', '4 px'], ['6', '6 px'], ['8', '8 px']], setOutlineWidth, 'width:76px', 'Outline weight'), rb('▣', 'Toggle shadow', toggleShadow),
        lsel, colorIn('Text colour inside shape', '#FFFFFF', setLabelColor))),
    grp('Arrange',
      row(...[['⇤', 'left', 'Align left'], ['⇔', 'center', 'Align centre'], ['⇥', 'right', 'Align right'], ['⤒', 'top', 'Align top'], ['⇕', 'middle', 'Align middle'], ['⤓', 'bottom', 'Align bottom']].map(([i, m, t]) => rb(i, t, () => alignOp(m)))),
      row(rb('⇹', 'Distribute horizontally', () => distributeOp('h')), rb('⇳', 'Distribute vertically', () => distributeOp('v')), rb('⇋', 'Flip horizontal', () => flipOp('x')), rb('⇵', 'Flip vertical', () => flipOp('y')), rb('↺', 'Rotate left 90°', () => rotateOp(-90)), rb('↻', 'Rotate right 90°', () => rotateOp(90))),
      row(rb('W=', 'Match width of first object', () => matchSize('w'), 'font-size:11px;font-weight:700'), rb('H=', 'Match height of first object', () => matchSize('h'), 'font-size:11px;font-weight:700'), rb('W=H', 'Match both', () => matchSize('both'), 'font-size:10px;font-weight:700'))),
    grp('WordArt styles',
      row(...WORDART.map(([id, name, css]) => h('button', { class: 'ms-wa', title: name, style: css, onclick: () => setWordArt(id) }, 'A')), rb('✕', 'Clear WordArt', () => setWordArt(null)))),
    grp('Text effects',
      row(sel([['', 'Line spacing'], ['1', '1.0'], ['1.15', '1.15'], ['1.5', '1.5'], ['2', '2.0']], setLineSpacing, 'width:104px', 'Line spacing'),
        sel([['0', 'Letter spacing'], ['-0.03', 'Tight'], ['0', 'Normal'], ['0.06', 'Loose'], ['0.14', 'Very loose']], setLetterSpacing, 'width:110px', 'Letter spacing')),
      row(rb('⤒', 'Text at top', () => setVAlign('top')), rb('↕', 'Text in middle', () => setVAlign('middle')), rb('⤓', 'Text at bottom', () => setVAlign('bottom')),
        rb('A▲', 'Increase font size (Ctrl+Shift+>)', () => growFont(1), 'font-size:11px;font-weight:700'), rb('A▼', 'Decrease font size (Ctrl+Shift+<)', () => growFont(-1), 'font-size:11px;font-weight:700'))),
    grp('Text',
      row(sel([['', 'Change case'], ['upper', 'UPPERCASE'], ['lower', 'lowercase'], ['title', 'Title Case'], ['sentence', 'Sentence case']], v => v && changeCase(v), 'width:116px', 'Change case'),
        rb('Tx', 'Clear text formatting', clearTextFormat, 'font-weight:700;font-size:12px'))),
    grp('Accessibility', row(lb('🔗', 'Link', 'Hyperlink', setLink, 52), lb('🖼', 'Alt text', 'Alt text for screen readers', setAlt, 56)))
  );
}
function buildDesign() {
  const p = document.querySelector('[data-panel="design"]');
  const last = p.lastElementChild; if (last) last.style.borderRight = '';
  p.append(grp('Background', row(lb('🎨', 'Format', 'Solid, gradient or picture background', openBackground, 56), lb('⧉', 'Apply to all', 'Copy this slide\'s background to every slide', applyBgToAll, 62))));
}
function buildTransitions() {
  const p = document.querySelector('[data-panel="transitions"]');
  const last = p.lastElementChild; if (last) last.style.borderRight = '';
  transDurIn = h('input', { type: 'number', min: '0.1', max: '3', step: '0.1', style: 'width:100%' });
  advIn = h('input', { type: 'number', min: '0', max: '600', step: '1', placeholder: 'off', style: 'width:100%' });
  transDurIn.addEventListener('change', () => { snapshot(); slide().transDur = clamp(+transDurIn.value || 0.4, 0.1, 3); scheduleSave(); });
  advIn.addEventListener('change', () => { snapshot(); const v = Math.max(0, +advIn.value || 0); if (v) slide().advance = v; else delete slide().advance; scheduleSave(); });
  p.append(
    grp('Timing', row(h('div', { class: 'field', style: 'width:84px' }, h('label', null, 'Duration (s)'), transDurIn), h('div', { class: 'field', style: 'width:110px' }, h('label', null, 'Advance after (s)'), advIn))),
    grp('Apply', row(lb('⧉', 'To all', 'Use this slide\'s transition and timing on every slide', () => {
      snapshot(); const s = slide();
      deck.slides.forEach(x => { x.transition = s.transition; if (s.transDur) x.transDur = s.transDur; else delete x.transDur; if (s.advance) x.advance = s.advance; else delete x.advance; });
      render(); scheduleSave(); toast('Transition applied to all slides');
    }, 56)))
  );
}
function buildShowTab() {
  const p = addTab('show', 'Slide Show', 'view');
  const chk = (label, key) => { const i = h('input', { type: 'checkbox' }); i.checked = !!PS[key]; i.addEventListener('change', () => { PS[key] = i.checked; savePS(); }); return lab(i, label); };
  p.append(
    grp('Start', row(lb('▶', 'From start', 'Start from the first slide (F5)', () => msStart(firstVisible()), 58), lb('▷', 'From current', 'Start from this slide (Shift+F5)', () => msStart(current), 62))),
    grp('Set up', row(chk('Use timings', 'timings'), chk('Loop until Esc', 'loop')), row(chk('Click to advance', 'click'), chk('Show timer', 'timer'))),
    grp('Rehearse', row(lb('⏱', 'Rehearse', 'Run the show and record time per slide', startRehearse, 60), lb('👁', 'Hide slide', 'Hide / unhide this slide', () => toggleHide(), 56))),
    grp('During the show', h('div', { class: 'ms-hint' }, 'L laser · P pen · H highlighter · E erase', h('br'), 'B / W black / white screen · T timer', h('br'), 'Type a number + Enter to jump · Right-click = back'))
  );
}
function buildReview() {
  const p = addTab('review', 'Review', 'view');
  p.append(
    grp('Proofing', row(lb('abc', 'Spelling', 'Underline misspellings while editing text', toggleSpell, 56), lb('♿', 'Accessibility', 'Check for accessibility issues', openA11y, 66), lb('123', 'Word count', 'Count words and characters', openWordCount, 62))),
    grp('Comments', row(lb('💬', 'New comment', 'Add a comment', () => openComments(true), 66), lb('📌', 'Show pins', 'Show / hide comment pins on the slide', togglePins, 56), lb('☰', 'All comments', 'Open the comments panel', () => openComments(false), 66))),
    grp('Find', row(lb('🔍', 'Find', 'Ctrl+F', () => openFind(false), 50), lb('⇄', 'Replace', 'Ctrl+H', () => openFind(true), 56)))
  );
}
function buildTopbar() {
  const primary = document.querySelector('.ed-topbar .ed-btn.primary');
  const b = h('button', { class: 'ed-btn', title: 'Print or save as PDF (Ctrl+P)', onclick: openPrint }, '⎙ Print / PDF');
  if (primary) primary.parentNode.insertBefore(b, primary); else document.querySelector('.ed-topbar').append(b);
}

/* ====================================================================
   20. INIT
   ==================================================================== */
function init() {
  try {
    buildHome(); buildInsert(); buildFormat(); buildDesign(); buildTransitions();
    buildShowTab(); buildReview(); buildTopbar(); buildStatusBar();
    installCanvasHooks();
    ensureOverlays();
    if (deck && inEditor()) render();
    console.info('[ms-ppt] loaded v' + VERSION);
  } catch (err) {
    console.error('[ms-ppt] failed to initialise', err);
  }
}
init();
})();
