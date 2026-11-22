/* PowerPoint extras 3: animations (timeline, pane, triggers, motion paths), transitions, slide show engine (presenter view, pen/laser, timings, custom shows), pptx export hooks — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, el = OA.el, PX = window.PPTX || (window.PPTX = {});
  const xe = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const clone = o => JSON.parse(JSON.stringify(o));

  /* ================= animation model ================= */
  const EFF = {
    entr: [['appear', 'Appear'], ['fade', 'Fade'], ['fly', 'Fly In'], ['float', 'Float In'], ['wipe', 'Wipe'], ['split', 'Split'], ['zoom', 'Zoom'], ['grow', 'Grow & Turn'], ['spin', 'Spinner']],
    emph: [['pulse', 'Pulse'], ['grow', 'Grow/Shrink'], ['spin', 'Spin'], ['teeter', 'Teeter'], ['blink', 'Blink']],
    exit: [['disappear', 'Disappear'], ['fade', 'Fade'], ['fly', 'Fly Out'], ['wipe', 'Wipe'], ['zoom', 'Zoom'], ['shrink', 'Shrink & Turn'], ['spin', 'Spinner']],
    path: [['line', 'Lines'], ['arc', 'Arcs'], ['circle', 'Circle'], ['diamond', 'Diamond'], ['zigzag', 'Zigzag']]
  };
  const CLS = { entr: 'Entrance', emph: 'Emphasis', exit: 'Exit', path: 'Motion Path' };
  const DIRS = [['b', 'From Bottom'], ['l', 'From Left'], ['r', 'From Right'], ['t', 'From Top']];
  PX.EFF = EFF;
  const getAnims = s => {
    if (s.anim) return s.anim;
    const a = []; (s.objs || []).forEach(o => { if (o.anim && o.anim !== 'none') a.push({ oid: o.id, cls: 'entr', eff: o.anim, trig: 'click', dur: .6, delay: 0, dir: 'b', rep: 1 }); }); return a;
  };
  PX.getAnims = getAnims;
  const effName = a => (EFF[a.cls].find(e => e[0] === a.eff) || [0, a.eff])[1];

  /* keyframes for one animation; base = node's own transform (rotation) */
  function frames(a, node, deck) {
    const base = node.style.transform || ''; const T = t => (base + ' ' + t).trim(); const W = deck.w, H = deck.h; const d = a.dir || 'b';
    const off = { b: [0, H], t: [0, -H], l: [-W, 0], r: [W, 0] }[d];
    const clipIn = { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }[d];
    const full = 'inset(0 0 0 0)';
    const c = a.cls, e = a.eff;
    if (c === 'entr' || c === 'exit') {
      let k;
      if (e === 'appear' || e === 'disappear') k = null;
      else if (e === 'fade') k = [{ opacity: 0 }, { opacity: 1 }];
      else if (e === 'fly') k = [{ transform: T(`translate(${off[0]}px,${off[1]}px)`) }, { transform: T('translate(0,0)') }];
      else if (e === 'float') k = [{ opacity: 0, transform: T('translateY(40px)') }, { opacity: 1, transform: T('translateY(0)') }];
      else if (e === 'wipe') k = [{ clipPath: clipIn }, { clipPath: full }];
      else if (e === 'split') k = [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: full }];
      else if (e === 'zoom') k = [{ opacity: 0, transform: T('scale(0)') }, { opacity: 1, transform: T('scale(1)') }];
      else if (e === 'grow' || e === 'shrink') k = [{ opacity: 0, transform: T('scale(0) rotate(-90deg)') }, { opacity: 1, transform: T('scale(1) rotate(0deg)') }];
      else if (e === 'spin') k = [{ opacity: 0, transform: T('rotate(-360deg) scale(.1)') }, { opacity: 1, transform: T('rotate(0deg) scale(1)') }];
      else k = [{ opacity: 0 }, { opacity: 1 }];
      if (k && c === 'exit') k = k.slice().reverse();
      return k;
    }
    if (c === 'emph') {
      if (e === 'pulse') return [{ transform: T('scale(1)') }, { transform: T('scale(1.12)') }, { transform: T('scale(1)') }];
      if (e === 'grow') return [{ transform: T('scale(1)') }, { transform: T('scale(1.5)') }];
      if (e === 'spin') return [{ transform: T('rotate(0deg)') }, { transform: T('rotate(360deg)') }];
      if (e === 'teeter') return [{ transform: T('rotate(0deg)') }, { transform: T('rotate(-8deg)'), offset: .25 }, { transform: T('rotate(8deg)'), offset: .75 }, { transform: T('rotate(0deg)') }];
      return [{ opacity: 1 }, { opacity: 0, offset: .5 }, { opacity: 1 }];
    }
    /* motion paths */
    const dx = W * .25, dy = H * .25; const P2 = pts => pts.map(p => ({ transform: T(`translate(${p[0]}px,${p[1]}px)`) }));
    if (e === 'line') { const v = { b: [0, dy], t: [0, -dy], l: [-dx, 0], r: [dx, 0] }[d]; return P2([[0, 0], v]); }
    if (e === 'arc') { const v = { b: [0, 1], t: [0, -1], l: [-1, 0], r: [1, 0] }[d]; const pts = []; for (let i = 0; i <= 10; i++) { const t = Math.PI * i / 10; const along = (1 - Math.cos(t)) / 2 * dx * 2, perp = Math.sin(t) * dx * .6; pts.push(v[0] ? [v[0] * along, -perp] : [perp, v[1] * along]); } return P2(pts); }
    if (e === 'circle') { const r = dx * .5, pts = []; for (let i = 0; i <= 16; i++) { const t = 2 * Math.PI * i / 16; pts.push([r * Math.sin(t), r * (1 - Math.cos(t))]); } return P2(pts); }
    if (e === 'diamond') { const r = dx * .5; return P2([[0, 0], [r, r], [0, 2 * r], [-r, r], [0, 0]]); }
    return P2([[0, 0], [dx * .3, -dy * .3], [dx * .6, dy * .3], [dx, 0]]);
  }

  /* ---- timeline runner (shared by editor preview and slide show) ---- */
  function buildSteps(s) {
    const steps = [], trig = {}; let cur = null, prevBase = 0, prevEnd = 0;
    getAnims(s).forEach(a => {
      const it = { a, t: 0 }; const dur = (a.dur || .5) * 1000, rep = Math.max(1, a.rep || 1);
      if (a.trigOid) { (trig[a.trigOid] = trig[a.trigOid] || []).push(Object.assign(it, { t: (a.delay || 0) * 1000 })); return; }
      let base = 0;
      if (a.trig === 'click' || !cur) { cur = { items: [], auto: a.trig !== 'click' && !steps.length }; steps.push(cur); base = 0; }
      else if (a.trig === 'with') base = prevBase; else base = prevEnd;
      it.t = base + (a.delay || 0) * 1000; prevBase = base; prevEnd = it.t + dur * rep; cur.items.push(it);
    });
    return { steps, trig };
  }
  PX.buildSteps = buildSteps;
  function Runner(slide, root, deck) {
    const B = buildSteps(slide); const steps = B.steps; let si = 0; const timers = new Set(), running = new Set(); const started = new Set();
    const q = oid => root.querySelector('.p-obj[data-id="' + oid + '"]');
    function init() {
      const seen = new Set(); getAnims(slide).forEach(a => { if (seen.has(a.oid)) return; seen.add(a.oid); const n = q(a.oid); if (n) { n.getAnimations().forEach(x => x.cancel()); n.style.visibility = a.cls === 'entr' ? 'hidden' : ''; } });
    }
    const endState = it => { const n = q(it.a.oid); if (n && it.a.cls === 'exit') n.style.visibility = 'hidden'; };
    function start(it) {
      if (started.has(it)) return null; started.add(it); const a = it.a; const n = q(a.oid); if (!n) return null;
      if (a.cls === 'entr') n.style.visibility = 'visible'; else if (a.cls !== 'exit') n.style.visibility = 'visible';
      const k = frames(a, n, deck); if (!k) { endState(it); return null; }
      const an = n.animate(k, { duration: (a.dur || .5) * 1000, iterations: Math.max(1, a.rep || 1), fill: a.cls === 'emph' && a.eff === 'blink' ? 'none' : 'both', easing: a.cls === 'path' ? 'ease-in-out' : 'ease-out' });
      running.add(an); an.onfinish = () => { running.delete(an); endState(it); }; an.oncancel = () => running.delete(an); it._an = an; return an;
    }
    function sched(items) { items.forEach(it => { if (it.t <= 0) { start(it); return; } const tm = setTimeout(() => { timers.delete(tm); start(it); }, it.t); timers.add(tm); }); }
    function ff() { timers.forEach(clearTimeout); timers.clear(); steps.slice(0, si).forEach(st => st.items.forEach(it => { start(it); })); Array.from(Object.values(B.trig)).flat().forEach(it => { if (it._an) { } }); steps.slice(0, si).forEach(st => st.items.forEach(it => { if (it._an) { try { it._an.finish(); } catch (e) { } } endState(it); })); }
    const total = st => Math.max(0, ...st.items.map(it => it.t + (it.a.dur || .5) * 1000 * Math.max(1, it.a.rep || 1)));
    return {
      steps, trig: B.trig, init, get si() { return si; }, get left() { return steps.length - si; },
      busy: () => timers.size > 0 || running.size > 0,
      next() { if (this.busy()) ff(); if (si >= steps.length) return false; const st = steps[si++]; sched(st.items); this.lastMs = total(st); return true; },
      ff,
      goTo(k) { timers.forEach(clearTimeout); timers.clear(); running.forEach(x => x.cancel()); running.clear(); started.clear(); steps.forEach(st => st.items.forEach(it => { delete it._an; })); init(); si = 0; while (si < k && si < steps.length) { const st = steps[si++]; st.items.forEach(it => { start(it); if (it._an) { try { it._an.finish(); } catch (e) { } } endState(it); }); } },
      fire(oid) { (B.trig[oid] || []).forEach(it => { started.delete(it); const tm = setTimeout(() => { timers.delete(tm); const n = q(it.a.oid); start(it); }, it.t); timers.add(tm); }); },
      total
    };
  }
  PX.Runner = Runner;

  /* ================= slide transitions ================= */
  const VEC = { l: [-1, 0], r: [1, 0], u: [0, -1], d: [0, 1] };
  function runTrans(oldW, newW, tr, deck, oldS, newS, host) {
    let t = tr.type; const all = ['fade', 'push', 'wipe', 'zoom', 'split', 'cover', 'uncover', 'dissolve', 'blinds', 'circle', 'diamond', 'curtains', 'flip', 'cube', 'vortex', 'flash', 'fall', 'reveal'];
    if (t === 'random') t = all[Math.floor(Math.random() * all.length)];
    const dur = Math.max(100, (tr.dur || 1) * 1000), d = tr.dir || 'r'; const v = VEC[d] || VEC.r; const anims = [];
    const A = (e, kf, o) => { const a = e.animate(kf, Object.assign({ duration: dur, fill: 'both', easing: 'ease-in-out' }, o || {})); anims.push(a); return a; };
    const tf = (p, m) => `translate(${p[0] * m * 100}%,${p[1] * m * 100}%)`;
    newW.style.zIndex = 2; oldW && (oldW.style.zIndex = 1);
    const clip = { l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)', u: 'inset(0 0 100% 0)', d: 'inset(100% 0 0 0)' }[d];
    switch (t) {
      case 'fade': A(newW, [{ opacity: 0 }, { opacity: 1 }]); break;
      case 'push': A(newW, [{ transform: tf(v, 1) }, { transform: 'none' }]); oldW && A(oldW, [{ transform: 'none' }, { transform: tf(v, -1) }]); break;
      case 'cover': A(newW, [{ transform: tf(v, 1) }, { transform: 'none' }]); break;
      case 'uncover': case 'reveal': newW.style.zIndex = 1; if (oldW) { oldW.style.zIndex = 3; A(oldW, [{ transform: 'none', opacity: 1 }, { transform: tf(v, 1), opacity: t === 'reveal' ? 0 : 1 }]); } break;
      case 'wipe': A(newW, [{ clipPath: clip }, { clipPath: 'inset(0 0 0 0)' }]); break;
      case 'zoom': A(newW, [{ opacity: 0, transform: 'scale(.2)' }, { opacity: 1, transform: 'none' }]); oldW && A(oldW, [{ opacity: 1 }, { opacity: 0 }]); break;
      case 'split': A(newW, [{ clipPath: 'inset(50% 0 50% 0)' }, { clipPath: 'inset(0 0 0 0)' }]); break;
      case 'curtains': A(newW, [{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0 0 0)' }], { easing: 'ease-in' }); break;
      case 'dissolve': A(newW, [{ opacity: 0 }, { opacity: 1 }], { easing: 'steps(8,end)' }); break;
      case 'blinds': { const m = p => `repeating-linear-gradient(to bottom,#000 0,#000 ${p}%,transparent ${p}%,transparent 10%)`; const kf = [0, 25, 50, 75, 100].map(p => ({ webkitMaskImage: m(p / 10 * 10 / 10 * 1), maskImage: m(p / 10 * 10 / 10 * 1) })); kf.forEach((f, i) => { const pp = [0.01, 2.5, 5, 7.5, 10][i]; f.webkitMaskImage = f.maskImage = `repeating-linear-gradient(to bottom,#000 0,#000 ${pp}%,transparent ${pp}%,transparent 10%)`; }); A(newW, kf, { easing: 'linear' }); break; }
      case 'circle': A(newW, [{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }]); break;
      case 'diamond': A(newW, [{ clipPath: 'polygon(50% 50%,50% 50%,50% 50%,50% 50%)' }, { clipPath: 'polygon(50% -50%,150% 50%,50% 150%,-50% 50%)' }]); break;
      case 'flip': host.style.perspective = '1800px'; newW.style.backfaceVisibility = 'hidden'; A(newW, [{ transform: 'rotateY(-90deg)' }, { transform: 'rotateY(0)' }]); oldW && A(oldW, [{ transform: 'rotateY(0)' }, { transform: 'rotateY(90deg)' }]); break;
      case 'cube': host.style.perspective = '1800px'; newW.style.transformOrigin = '0 50%'; A(newW, [{ transform: 'translateX(100%) rotateY(90deg)' }, { transform: 'none' }]); if (oldW) { oldW.style.transformOrigin = '100% 50%'; A(oldW, [{ transform: 'none' }, { transform: 'translateX(-100%) rotateY(-90deg)' }]); } break;
      case 'vortex': A(newW, [{ transform: 'rotate(-360deg) scale(0)' }, { transform: 'none' }]); break;
      case 'flash': A(newW, [{ filter: 'brightness(10)', opacity: 0 }, { filter: 'brightness(10)', opacity: 1, offset: .3 }, { filter: 'none', opacity: 1 }]); break;
      case 'fall': newW.style.zIndex = 1; if (oldW) { oldW.style.zIndex = 3; oldW.style.transformOrigin = '0 100%'; A(oldW, [{ transform: 'none' }, { transform: 'rotate(40deg) translateY(130%)' }], { easing: 'ease-in' }); } break;
      case 'morph': {
        A(newW, [{ opacity: 0 }, { opacity: 1 }]);
        if (oldS && newS) newS.objs.forEach(no => { const oo = oldS.objs.find(x => (no.name && x.name === no.name) || (x.type === no.type && no.type === 'text' && x.text && x.text === no.text) || (x.type === 'shape' && no.type === 'shape' && x.kind === no.kind && x.name === no.name && no.name)); if (!oo) return; const n = newW.querySelector('.p-obj[data-id="' + no.id + '"]'); if (n) A(n, [{ left: oo.x + 'px', top: oo.y + 'px', width: oo.w + 'px', height: oo.h + 'px' }, { left: no.x + 'px', top: no.y + 'px', width: no.w + 'px', height: no.h + 'px' }]); });
        break;
      }
      default: break;
    }
    return Promise.all(anims.map(a => a.finished.catch(() => { }))).then(() => { anims.forEach(a => { try { a.cancel(); } catch (e) { } }); host.style.perspective = ''; });
  }
  PX.runTrans = runTrans;

  /* ================= slide show engine ================= */
  const CUSTOM = deck => (deck.customShows = deck.customShows || []);
  OA.pptPresent = function (P, from, opts) {
    opts = opts || {}; const deck = P.deck; P.finishEdit && P.finishEdit();
    if (deck._mm && PX.exitMaster) PX.exitMaster(true);
    const D = P.deck; const cfg = Object.assign({ type: 'presenter', loop: false, noAnim: false, useTimings: true, from: 1, to: 0, custom: '' }, D.show || {}, opts.cfg || {});
    const N = D.slides.length; let seq = [];
    if (opts.custom || cfg.custom) { const cs = CUSTOM(D).find(c => c.name === (opts.custom || cfg.custom)); if (cs) seq = cs.ids.map(id => D.slides.findIndex(s => s.id === id)).filter(i => i >= 0); }
    if (!seq.length) { const a = Math.max(1, cfg.from || 1) - 1, b = (cfg.to && cfg.to > 0 ? Math.min(cfg.to, N) : N) - 1; for (let i = a; i <= b; i++) seq.push(i); }
    if (!seq.length) seq = D.slides.map((_, i) => i);
    let i = from != null && from >= 0 && from < N ? from : seq[0]; if (from == null) i = seq[0];
    const kiosk = cfg.type === 'kiosk', loop = cfg.loop || kiosk;
    const ov = el('div', { class: 'p-show ps2' }); const host = el('div', { class: 'ps-host' }); ov.appendChild(host);
    let sw = null, runner = null, tool = 'ptr', penC = '#e53935', pv = !!(opts.pv || cfg.pv), ended = false, blank = null, busy = null, autoT = null, digits = '', t0 = Date.now(), slideT0 = Date.now(); const ink = {}; const rec = {};
    const ctr = el('div', { class: 'ctr' }); ov.appendChild(ctr);
    /* toolbar */
    const tb = el('div', { class: 'ps-tb' });
    const tbtn = (txt, title, fn, cls) => { const b = el('button', { type: 'button', class: 'ps-b ' + (cls || ''), title, text: txt, onclick: e => { e.stopPropagation(); fn(); } }); tb.appendChild(b); return b; };
    tbtn('◀', 'Previous', () => prev()); tbtn('▶', 'Next', () => next());
    const toolBtns = {}; [['ptr', '⌖', 'Pointer (Ctrl+A)'], ['laser', '●', 'Laser pointer (Ctrl+L)'], ['pen', '✎', 'Pen (Ctrl+P)'], ['hl', '▮', 'Highlighter (Ctrl+I)'], ['eraser', '⌫', 'Eraser (Ctrl+E)']].forEach(([k, g, t]) => { toolBtns[k] = tbtn(g, t, () => setTool(k), 'tool-' + k); });
    tbtn('✕', 'Erase all ink on slide (E)', () => eraseAll());
    const cols = el('span', { class: 'ps-cols' }); ['#e53935', '#1e88e5', '#43a047', '#111111', '#fdd835', '#ffffff'].forEach(c => { const s = el('i', { style: 'background:' + c, title: c }); s.onclick = e => { e.stopPropagation(); penC = c; if (tool !== 'pen' && tool !== 'hl') setTool('pen'); }; cols.appendChild(s); }); tb.appendChild(cols);
    tbtn('▦', 'See all slides (G)', () => grid()); tbtn('🗒', 'Presenter view (V)', () => setPV(!pv), 'pvb'); tbtn('■', 'Black screen (B)', () => setBlank('black')); tbtn('□', 'White screen (W)', () => setBlank('white')); tbtn('?', 'Keyboard shortcuts', () => help());
    ov.appendChild(tb);
    const ex = el('button', { class: 'exit', type: 'button', text: 'Exit (Esc)', onclick: e => { e.stopPropagation(); close(); } }); ov.appendChild(ex);
    const laser = el('div', { class: 'ps-laser' }); ov.appendChild(laser);
    document.body.appendChild(ov); window.__pptShow = api();
    try { if (!opts.noFs && ov.requestFullscreen) ov.requestFullscreen().catch(() => { }); } catch (e) { }

    function setTool(k) { tool = k; Object.keys(toolBtns).forEach(x => toolBtns[x].classList.toggle('on', x === k)); ov.dataset.tool = k; laser.style.display = k === 'laser' ? 'block' : 'none'; if (sw) sw.cv.style.pointerEvents = ['pen', 'hl', 'eraser'].includes(k) ? 'auto' : 'none'; }
    const area = () => pv ? { x: 0, y: 0, w: innerWidth * .62, h: innerHeight } : { x: 0, y: 0, w: innerWidth, h: innerHeight };
    function place(w) { const a = area(); const k = Math.min(a.w / D.w, a.h / D.h); const W = D.w * k, H = D.h * k; w.style.cssText += `;left:${a.x + (a.w - W) / 2}px;top:${a.y + (a.h - H) / 2}px;width:${W}px;height:${H}px`; w._k = k; const s = w.firstChild; s.style.transform = `scale(${k})`; s.style.setProperty('--sc', k); }
    function mk(idx) {
      const s = D.slides[idx]; const w = el('div', { class: 'ps-sl' }); const sl = P.renderSlide(s, D, { anim: true }); sl.style.position = 'absolute'; sl.style.left = '0'; sl.style.top = '0'; sl.style.transformOrigin = '0 0'; w.appendChild(sl);
      const cv = el('canvas', { class: 'ps-ink' }); cv.width = D.w; cv.height = D.h; cv.style.cssText = `position:absolute;left:0;top:0;width:${D.w}px;height:${D.h}px;transform-origin:0 0;pointer-events:none;z-index:50`; sl.appendChild(cv); w.cv = cv; w.sl = sl; w.idx = idx; place(w); return w;
    }
    function redrawInk(w) { const c = w.cv.getContext('2d'); c.clearRect(0, 0, D.w, D.h); (ink[w.idx] || []).forEach(st => drawStroke(c, st)); }
    function drawStroke(c, st) { c.save(); c.lineCap = c.lineJoin = 'round'; c.strokeStyle = st.c; c.lineWidth = st.w; c.globalAlpha = st.t === 'hl' ? .4 : 1; c.beginPath(); st.pts.forEach((p, j) => j ? c.lineTo(p[0], p[1]) : c.moveTo(p[0], p[1])); if (st.pts.length === 1) c.lineTo(st.pts[0][0] + .1, st.pts[0][1]); c.stroke(); c.restore(); }
    const pt = e => { const r = sw.sl.getBoundingClientRect(); return [(e.clientX - r.left) / sw._k, (e.clientY - r.top) / sw._k]; };
    function bindInk(w) {
      let cur = null; w.cv.addEventListener('pointerdown', e => { if (!['pen', 'hl', 'eraser'].includes(tool)) return; e.preventDefault(); w.cv.setPointerCapture(e.pointerId); const p = pt(e); if (tool === 'eraser') { cur = 'e'; erase(p); } else { cur = { t: tool, c: penC, w: tool === 'hl' ? 22 : 4, pts: [p] }; (ink[w.idx] = ink[w.idx] || []).push(cur); redrawInk(w); } });
      w.cv.addEventListener('pointermove', e => { if (!cur) return; const p = pt(e); if (cur === 'e') erase(p); else { cur.pts.push(p); redrawInk(w); } });
      const up = () => { cur = null; }; w.cv.addEventListener('pointerup', up); w.cv.addEventListener('pointercancel', up);
      function erase(p) { const a = ink[w.idx]; if (!a) return; const n = a.length; ink[w.idx] = a.filter(st => !st.pts.some(q => Math.hypot(q[0] - p[0], q[1] - p[1]) < 14)); if (ink[w.idx].length !== n) redrawInk(w); }
    }
    function eraseAll() { if (!sw) return; ink[sw.idx] = []; redrawInk(sw); }
    /* navigation */
    const hidden = j => !!(D.slides[j] && D.slides[j].hide);
    function seqNext(j) { const p = seq.indexOf(j); for (let k = p + 1; k < seq.length; k++) if (!hidden(seq[k])) return seq[k]; return -1; }
    function seqPrev(j) { const p = seq.indexOf(j); for (let k = p - 1; k >= 0; k--) if (!hidden(seq[k])) return seq[k]; return -1; }
    function clearAuto() { if (autoT) { clearTimeout(autoT); autoT = null; } }
    function recStop() { if (opts.rehearse && sw) { rec[sw.idx] = (rec[sw.idx] || 0) + (Date.now() - slideT0) / 1000; } slideT0 = Date.now(); }
    function armAuto() {
      clearAuto(); const s = D.slides[i]; const tm = s.trans && s.trans.advTm; if (opts.rehearse || !tm || !(cfg.useTimings || kiosk)) return;
      autoT = setTimeout(() => { autoT = null; if (runner && runner.left > 0) { /* animations still pending: play them out */ } advance(); }, tm * 1000);
    }
    function showSlide(idx, dirn, instantEnd) {
      recStop(); clearAuto(); const old = sw; const oldS = old ? D.slides[old.idx] : null; i = idx; const s = D.slides[i];
      sw = mk(i); bindInk(sw); redrawInk(sw); host.appendChild(sw); setTool(tool);
      runner = Runner(s, sw.sl, D); if (!cfg.noAnim) { runner.init(); if (instantEnd) runner.goTo(runner.steps.length); } else { runner.noAnim = true; }
      if (cfg.noAnim) runner.steps.length = 0;
      /* object triggers */
      if (!cfg.noAnim) Object.keys(runner.trig).forEach(oid => { const n = sw.sl.querySelector('.p-obj[data-id="' + oid + '"]'); if (n) { n.style.cursor = 'pointer'; n.addEventListener('click', e => { e.stopPropagation(); runner.fire(oid); }); } });
      sw.sl.querySelectorAll('video,audio').forEach(m => { if (m.dataset.auto) { const pp = m.play(); pp && pp.catch && pp.catch(() => { }); } });
      const tr = s.trans && s.trans.type && s.trans.type !== 'none' ? s.trans : null;
      const fin = () => { if (old && old.parentNode) old.remove(); busy = null; if (!cfg.noAnim && runner.steps[0] && runner.steps[0].auto) runner.next(); armAuto(); upd(); };
      if (old) { if (tr && dirn !== 0 && !cfg.noAnim) { busy = runTrans(old, sw, tr, D, oldS, s, host).then(fin); } else { old.remove(); fin(); } } else fin();
      slideT0 = Date.now(); upd();
    }
    function advance() { if (ended) return close(); const n = seqNext(i); if (n >= 0) showSlide(n, 1); else if (loop) showSlide(seq.find(j => !hidden(j)) != null ? seq.find(j => !hidden(j)) : seq[0], 1); else { recStop(); endScreen(); } }
    function endScreen() { ended = true; clearAuto(); if (sw) sw.style.visibility = 'hidden'; blank = el('div', { class: 'ps-end', text: 'End of slide show, click to exit.' }); ov.appendChild(blank); }
    function finishBusy() { if (busy) { host.getAnimations({ subtree: true }).forEach(a => { try { a.finish(); } catch (e) { } }); } }
    function next() { if (blank && !ended) { setBlank(null); return; } if (ended) return close(); finishBusy(); clearAuto(); if (runner && runner.next()) { upd(); return; } advance(); }
    function prev() { if (ended) { ended = false; blank && blank.remove(); blank = null; sw.style.visibility = ''; upd(); return; } finishBusy(); clearAuto(); if (runner && runner.si > 0) { runner.goTo(runner.si - 1); upd(); return; } const p = seqPrev(i); if (p >= 0) showSlide(p, -1, true); }
    function go(idx, dirn) { if (idx < 0 || idx >= N) return; ended = false; blank && blank.remove(); blank = null; finishBusy(); showSlide(idx, dirn == null ? 1 : dirn, false); }
    function setBlank(c) { if (blank && blank.dataset.c === c) { blank.remove(); blank = null; return; } blank && blank.remove(); blank = null; if (!c) return; blank = el('div', { class: 'ps-blank' }); blank.dataset.c = c; blank.style.background = c === 'white' ? '#fff' : '#000'; ov.insertBefore(blank, ctr); }
    function doLink(l) {
      if (!l) return; const t = l.type; if (t === 'url') { let u = l.v || ''; if (!/^[a-z]+:/i.test(u)) u = 'https://' + u; try { window.open(u, '_blank', 'noopener'); } catch (e) { } }
      else if (t === 'slide') go((parseInt(l.v) || 1) - 1, 1); else if (t === 'next') advance(); else if (t === 'prev') { const p = seqPrev(i); if (p >= 0) go(p, -1); } else if (t === 'first') go(0, 1); else if (t === 'last') go(N - 1, 1); else if (t === 'end') close();
    }
    /* presenter view */
    let pvEl = null;
    function setPV(on) { pv = on; pvEl && pvEl.remove(); pvEl = null; ov.classList.toggle('pv', pv); if (pv) { pvEl = el('div', { class: 'ps-pv' }); ov.appendChild(pvEl); buildPV(); } if (sw) { place(sw); } upd(); }
    let notesFs = 18, pvTimer = null;
    function buildPV() {
      pvEl.innerHTML = ''; const nx = el('div', { class: 'pv-next' }); const nt = el('div', { class: 'pv-notes' }); const hd = el('div', { class: 'pv-hd' }); const clock = el('span', { class: 'pv-clock' }); const timer = el('span', { class: 'pv-timer' }); hd.append(el('b', { text: 'Presenter View' }), timer, clock);
      const btns = el('div', { class: 'pv-btns' }); [['◀ Prev', () => prev()], ['Next ▶', () => next()], ['A−', () => { notesFs = Math.max(10, notesFs - 2); nt.style.fontSize = notesFs + 'px'; }], ['A+', () => { notesFs = Math.min(48, notesFs + 2); nt.style.fontSize = notesFs + 'px'; }], ['End show', () => close()]].forEach(([t, f]) => btns.appendChild(el('button', { type: 'button', class: 'ps-b', text: t, onclick: e => { e.stopPropagation(); f(); } })));
      nt.style.fontSize = notesFs + 'px'; pvEl.append(hd, el('div', { class: 'pv-lbl', text: 'Next slide' }), nx, el('div', { class: 'pv-lbl', text: 'Notes' }), nt, btns); pvEl._nx = nx; pvEl._nt = nt; pvEl._timer = timer; pvEl._clock = clock;
      const tick = () => { const sec = Math.floor((Date.now() - t0) / 1000); timer.textContent = String(Math.floor(sec / 60)).padStart(2, '0') + ':' + String(sec % 60).padStart(2, '0'); clock.textContent = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); }; tick(); clearInterval(pvTimer); pvTimer = setInterval(tick, 1000);
    }
    function upd() {
      const p = seq.indexOf(i); ctr.textContent = (i + 1) + ' / ' + N + (runner && runner.left > 0 ? '  •  ' + runner.left + ' step' + (runner.left > 1 ? 's' : '') + ' left' : '');
      ov.querySelector('.pvb') && ov.querySelector('.pvb').classList.toggle('on', pv);
      if (pvEl && pvEl._nx) { const n = seqNext(i); pvEl._nx.innerHTML = ''; if (n >= 0) { const w = 330, k = w / D.w; const sl = P.renderSlide(D.slides[n], D, {}); sl.style.transform = `scale(${k})`; sl.style.transformOrigin = '0 0'; pvEl._nx.style.width = w + 'px'; pvEl._nx.style.height = D.h * k + 'px'; pvEl._nx.appendChild(sl); } else pvEl._nx.textContent = 'End of slide show'; pvEl._nt.textContent = D.slides[i].notes || '(No notes for this slide)'; }
      if (opts.rehearse) { rh.textContent = '⏺ Rehearsing — slide ' + (Math.floor((Date.now() - slideT0) / 1000)) + 's · total ' + Math.floor((Date.now() - t0) / 1000) + 's'; }
    }
    let rh = el('div', { class: 'ps-rehearse' }); if (opts.rehearse) { ov.appendChild(rh); setInterval(upd, 1000); }
    function grid() {
      const g = el('div', { class: 'ps-grid' }); g.onclick = e => { if (e.target === g) g.remove(); };
      seq.forEach(j => { const w = 190, k = w / D.w; const c = el('div', { class: 'gi' + (j === i ? ' cur' : '') + (hidden(j) ? ' hid' : '') }); const sl = P.renderSlide(D.slides[j], D, {}); sl.style.transform = `scale(${k})`; sl.style.transformOrigin = '0 0'; const b = el('div', { class: 'gb' }); b.style.width = w + 'px'; b.style.height = D.h * k + 'px'; b.appendChild(sl); c.append(b, el('span', { text: String(j + 1) })); c.onclick = e => { e.stopPropagation(); g.remove(); go(j, 0); }; g.appendChild(c); });
      ov.appendChild(g);
    }
    function help() {
      const h = el('div', { class: 'ps-help' }); h.innerHTML = '<b>Slide show shortcuts</b><br>→ ↓ Space Enter N : next · ← ↑ Backspace P : previous<br>Number + Enter : go to slide · Home / End<br>B or . : black · W or , : white<br>Ctrl+P pen · Ctrl+I highlighter · Ctrl+L laser · Ctrl+E eraser · Ctrl+A pointer · E erase ink<br>G : all slides · V : presenter view · Esc : exit<br><i>Click to close</i>'; h.onclick = () => h.remove(); ov.appendChild(h);
    }
    async function close() {
      if (ov._closed) return; ov._closed = true; clearAuto(); clearInterval(pvTimer); document.removeEventListener('keydown', onKey, true); removeEventListener('resize', onResize); recStop(); const last = i; const hadInk = Object.values(ink).some(a => a.length); ov.remove(); window.__pptShow = null; try { if (document.fullscreenElement) await document.exitFullscreen(); } catch (e) { }
      P.gotoSlide(Math.min(last, D.slides.length - 1));
      if (opts.rehearse) { const ok = await OA.confirm('Slide timings save karein? Total ' + Math.round(Object.values(rec).reduce((a, b) => a + b, 0)) + ' seconds.', { title: 'Rehearse Timings', ok: 'Yes', cancel: 'No' }); if (ok) { P.mutate(() => { Object.keys(rec).forEach(k => { const s = D.slides[k]; s.trans = Object.assign({ type: 'none', dur: 1 }, s.trans, { advTm: Math.max(1, Math.round(rec[k])) }); }); D.show = Object.assign({}, D.show, { useTimings: true }); }); } }
      if (hadInk && !kiosk) { const ok = await OA.confirm('Slide show mein ki gayi ink annotations rakhni hain?', { title: 'Keep ink', ok: 'Keep', cancel: 'Discard' }); if (ok) { P.mutate(() => { Object.keys(ink).forEach(k => { const st = ink[k]; if (!st.length) return; const c = document.createElement('canvas'); c.width = D.w; c.height = D.h; const x = c.getContext('2d'); st.forEach(s => drawStroke(x, s)); D.slides[k].objs.push({ id: P.uid(), type: 'image', x: 0, y: 0, w: D.w, h: D.h, src: c.toDataURL('image/png'), name: 'Ink Annotation' }); }); }); } }
    }
    const onKey = e => {
      const k = e.key; const stop = () => { e.preventDefault(); e.stopPropagation(); };
      if (e.ctrlKey || e.metaKey) { const m = { p: 'pen', i: 'hl', l: 'laser', e: 'eraser', a: 'ptr' }[k.toLowerCase()]; if (m) { stop(); setTool(m); } return; }
      if (k === 'Escape') { stop(); const g = ov.querySelector('.ps-grid,.ps-help'); if (g) g.remove(); else if (blank && !ended) setBlank(null); else close(); return; }
      if (/^[0-9]$/.test(k)) { stop(); digits += k; ctr.textContent = 'Go to: ' + digits; return; }
      if (k === 'Enter' && digits) { stop(); const n = parseInt(digits) - 1; digits = ''; go(n, 0); return; }
      if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n', 'N'].includes(k)) { stop(); if (kiosk) return; next(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) { stop(); if (kiosk) return; prev(); }
      else if (k === 'Home') { stop(); go(seq[0], 0); } else if (k === 'End') { stop(); go(seq[seq.length - 1], 0); }
      else if (k === 'b' || k === 'B' || k === '.') { stop(); setBlank('black'); } else if (k === 'w' || k === 'W' || k === ',') { stop(); setBlank('white'); }
      else if (k === 'e' || k === 'E') { stop(); eraseAll(); } else if (k === 'g' || k === 'G') { stop(); grid(); } else if (k === 'v' || k === 'V') { stop(); setPV(!pv); } else if (k === 'F1' || k === '?') { stop(); help(); }
    };
    const onResize = () => { if (sw) place(sw); };
    document.addEventListener('keydown', onKey, true); addEventListener('resize', onResize);
    let mvT = null; ov.addEventListener('mousemove', e => { ov.classList.add('mv'); clearTimeout(mvT); mvT = setTimeout(() => ov.classList.remove('mv'), 2200); if (tool === 'laser') { laser.style.left = e.clientX + 'px'; laser.style.top = e.clientY + 'px'; } });
    ov.addEventListener('click', e => {
      if (e.target.closest('.ps-tb,.ps-pv,.ps-grid,.ps-help,.exit')) return; const lk = e.target.closest('[data-link]'); if (lk) { try { doLink(JSON.parse(lk.dataset.link)); } catch (x) { } return; }
      if (e.target.closest('video,audio')) return; if (['pen', 'hl', 'eraser'].includes(tool) && !ended) return; if (kiosk) return; if (D.slides[i].trans && D.slides[i].trans.advClick === false && !ended && !blank) return;
      if (e.clientX < innerWidth * .2 && !pv && !ended) prev(); else next();
    });
    let tx0 = null; ov.addEventListener('touchstart', e => { tx0 = e.touches[0].clientX; }, { passive: true }); ov.addEventListener('touchend', e => { if (tx0 === null) return; const dx = e.changedTouches[0].clientX - tx0; tx0 = null; if (Math.abs(dx) > 60) { e.preventDefault(); dx < 0 ? next() : prev(); } });
    setTool('ptr'); if (pv) setPV(true); showSlide(i, 0); ov.tabIndex = 0; ov.focus();
    function api() { return { next, prev, go, close, get i() { return i; }, get runner() { return runner; }, setTool, setPV, grid, setBlank, advance, get ink() { return ink; }, get seq() { return seq; }, get ended() { return ended; }, get tool() { return tool; }, get ov() { return ov; } }; }
    return window.__pptShow;
  };

  /* ================= export hooks ================= */
  const hex6 = c => String(c || '000000').replace('#', '').slice(0, 6).toUpperCase().padEnd(6, '0');
  OA.pptLinkXml = function (link, rels, nextRid, deck) {
    if (!link) return ''; const t = link.type;
    if (t === 'url') { let u = link.v || ''; if (!u) return ''; if (!/^[a-z]+:/i.test(u)) u = 'https://' + u; const r = nextRid(); rels.push(`<Relationship Id="${r}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xe(u)}" TargetMode="External"/>`); return `<a:hlinkClick r:id="${r}"/>`; }
    if (t === 'slide') { const n = Math.max(1, Math.min(deck.slides.length, parseInt(link.v) || 1)); const r = nextRid(); rels.push(`<Relationship Id="${r}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${n}.xml"/>`); return `<a:hlinkClick r:id="${r}" action="ppaction://hlinksldjump"/>`; }
    const j = { next: 'nextslide', prev: 'previousslide', first: 'firstslide', last: 'lastslide', end: 'endshow' }[t]; return j ? `<a:hlinkClick r:id="" action="ppaction://hlinkshowjump?jump=${j}"/>` : '';
  };
  const TX = { fade: '<p:fade/>', push: d => `<p:push dir="${d}"/>`, wipe: d => `<p:wipe dir="${d}"/>`, zoom: '<p:zoom dir="in"/>', split: '<p:split orient="horz" dir="out"/>', cover: d => `<p:cover dir="${d}"/>`, uncover: d => `<p:pull dir="${d}"/>`, dissolve: '<p:dissolve/>', blinds: '<p:blinds dir="horz"/>', circle: '<p:circle/>', diamond: '<p:diamond/>', curtains: '<p:split orient="vert" dir="out"/>', flip: '<p:fade/>', cube: d => `<p:cover dir="${d}"/>`, vortex: '<p:wedge/>', flash: '<p:fade thruBlk="1"/>', fall: '<p:pull dir="d"/>', reveal: d => `<p:pull dir="${d}"/>`, morph: '<p:fade/>', random: '<p:random/>' };
  const DMAP = { l: 'r', r: 'l', u: 'd', d: 'u' };
  OA.pptTransXml = function (tr) {
    if (!tr) return ''; const attrs = (tr.advClick === false ? ' advClick="0"' : '') + (tr.advTm > 0 ? ` advTm="${Math.round(tr.advTm * 1000)}"` : '');
    const x = tr.type && tr.type !== 'none' ? TX[tr.type] : null; const spd = (tr.dur || 1) <= .6 ? 'fast' : (tr.dur || 1) >= 1.6 ? 'slow' : 'med';
    if (!x) return attrs ? `<p:transition spd="med"${attrs}/>` : '';
    const inner = typeof x === 'function' ? x(DMAP[tr.dir || 'r'] || 'l') : x; return `<p:transition spd="${spd}"${attrs}>${inner}</p:transition>`;
  };
  const PRESET = { entr: { appear: [1, 0], fade: [10, 0], fly: [2, 4], float: [42, 0], wipe: [22, 8], split: [16, 21], zoom: [53, 16], grow: [31, 0], spin: [31, 0] }, exit: { disappear: [1, 0], fade: [10, 0], fly: [2, 4], wipe: [22, 8], zoom: [53, 32], shrink: [31, 0], spin: [31, 0] }, emph: { pulse: [26, 0], grow: [6, 0], spin: [8, 0], teeter: [32, 0], blink: [35, 0] }, path: { line: [0, 0], arc: [0, 0], circle: [0, 0], diamond: [0, 0], zigzag: [0, 0] } };
  const FLYSUB = { b: 4, l: 8, r: 2, t: 1 };
  const WIPEF = { b: 'wipe(down)', l: 'wipe(left)', r: 'wipe(right)', t: 'wipe(up)' };
  OA.pptTiming = function (s, idMap) {
    const B = buildSteps(s); let id = 2; const nid = () => ++id; const steps = B.steps.filter(st => st.items.some(it => idMap[it.a.oid])); if (!steps.length) return '';
    const tgt = sp => `<p:tgtEl><p:spTgt spid="${sp}"/></p:tgtEl>`;
    const setVis = (sp, v, delay) => `<p:set><p:cBhvr><p:cTn id="${nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${delay || 0}"/></p:stCondLst></p:cTn>${tgt(sp)}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${v}"/></p:to></p:set>`;
    const animEff = (sp, f, tr, ms) => `<p:animEffect transition="${tr}" filter="${f}"><p:cBhvr><p:cTn id="${nid()}" dur="${ms}"/>${tgt(sp)}</p:cBhvr></p:animEffect>`;
    const anim = (sp, attr, a, b, ms) => `<p:anim calcmode="lin" valueType="num"><p:cBhvr><p:cTn id="${nid()}" dur="${ms}" fill="hold"/>${tgt(sp)}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:tavLst><p:tav tm="0"><p:val><p:strVal val="${a}"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="${b}"/></p:val></p:tav></p:tavLst></p:anim>`;
    const effXml = a => {
      const sp = idMap[a.oid], ms = Math.round((a.dur || .5) * 1000), c = a.cls, e = a.eff, d = a.dir || 'b'; const out = c === 'exit';
      if (c === 'entr' || c === 'exit') {
        const vis = out ? '' : setVis(sp, 'visible'); const end = out ? setVis(sp, 'hidden', ms) : '';
        let body = '';
        if (e === 'appear' || e === 'disappear') body = '';
        else if (e === 'fade') body = animEff(sp, 'fade', out ? 'out' : 'in', ms);
        else if (e === 'wipe') body = animEff(sp, WIPEF[d], out ? 'out' : 'in', ms);
        else if (e === 'split') body = animEff(sp, 'barn(inVertical)', out ? 'out' : 'in', ms);
        else if (e === 'fly') { const ax = d === 'l' || d === 'r' ? 'ppt_x' : 'ppt_y'; const far = { l: '0-#ppt_w/2', r: '1+#ppt_w/2', t: '0-#ppt_h/2', b: '1+#ppt_h/2' }[d]; const home = '#' + ax; body = out ? anim(sp, ax, home, far, ms) : anim(sp, ax, far, home, ms); }
        else if (e === 'float') body = animEff(sp, 'fade', out ? 'out' : 'in', ms);
        else { const f = out ? ['#ppt_w', '0'] : ['0', '#ppt_w']; const g = out ? ['#ppt_h', '0'] : ['0', '#ppt_h']; body = anim(sp, 'ppt_w', f[0], f[1], ms) + anim(sp, 'ppt_h', g[0], g[1], ms) + animEff(sp, 'fade', out ? 'out' : 'in', ms); }
        return vis + body + end;
      }
      if (c === 'emph') {
        const B2 = x => `<p:cBhvr><p:cTn id="${nid()}" dur="${ms}" fill="hold"/>${tgt(sp)}</p:cBhvr>`;
        if (e === 'spin') return `<p:animRot by="21600000">${B2()}</p:animRot>`;
        if (e === 'teeter') return `<p:animRot by="300000"><p:cBhvr><p:cTn id="${nid()}" dur="${Math.max(100, Math.round(ms / 4))}" autoRev="1" fill="hold"/>${tgt(sp)}</p:cBhvr></p:animRot>`;
        if (e === 'blink') return animEff(sp, 'fade', 'out', ms);
        const by = e === 'grow' ? 150000 : 110000; return `<p:animScale><p:cBhvr><p:cTn id="${nid()}" dur="${ms}" ${e === 'pulse' ? 'autoRev="1" ' : ''}fill="hold"/>${tgt(sp)}</p:cBhvr><p:by x="${by}" y="${by}"/></p:animScale>`;
      }
      const pathOf = { line: { b: 'M 0 0 L 0 0.25 E', t: 'M 0 0 L 0 -0.25 E', l: 'M 0 0 L -0.25 0 E', r: 'M 0 0 L 0.25 0 E' }[d], arc: 'M 0 0 C 0.1 -0.15 0.2 -0.15 0.25 0 E', circle: 'M 0 0 C 0.1 0 0.1 0.2 0 0.2 C -0.1 0.2 -0.1 0 0 0 Z', diamond: 'M 0 0 L 0.1 0.1 L 0 0.2 L -0.1 0.1 L 0 0 Z', zigzag: 'M 0 0 L 0.07 -0.1 L 0.15 0.1 L 0.25 0 E' }[e] || 'M 0 0 L 0.25 0 E';
      return `<p:animMotion origin="layout" path="${pathOf}" pathEditMode="relative" ptsTypes=""><p:cBhvr><p:cTn id="${nid()}" dur="${ms}" fill="hold"/>${tgt(sp)}<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>`;
    };
    const eff = (it, nt) => {
      const a = it.a, sp = idMap[a.oid]; if (!sp) return ''; const pr = (PRESET[a.cls][a.eff] || [0, 0]); const sub = a.eff === 'fly' || a.eff === 'wipe' ? ({ fly: FLYSUB[a.dir || 'b'], wipe: { b: 4, l: 8, r: 2, t: 1 }[a.dir || 'b'] }[a.eff]) : pr[1];
      const rep = (a.rep || 1) > 1 ? ` repeatCount="${a.rep * 1000}"` : '';
      const idc = nid(); const inner = effXml(a);
      return `<p:par><p:cTn id="${idc}" presetID="${pr[0]}" presetClass="${a.cls}" presetSubtype="${sub || 0}" fill="hold" grpId="0" nodeType="${nt}"${rep}><p:stCondLst><p:cond delay="${Math.round(it.t - (it._base || 0))}"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par>`;
    };
    const stepXml = st => {
      const outer = nid(); const mid = nid(); let inner = ''; let first = true;
      st.items.forEach(it => { const a = it.a; it._base = 0; const nt = first && a.trig === 'click' ? 'clickEffect' : a.trig === 'after' ? 'afterEffect' : a.trig === 'with' ? 'withEffect' : 'clickEffect'; first = false; inner += eff(it, nt); });
      return `<p:par><p:cTn id="${outer}" fill="hold"><p:stCondLst><p:cond delay="${st.auto ? '0' : 'indefinite'}"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${mid}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`;
    };
    const body = steps.map(stepXml).join('');
    return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${body}</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>`;
  };

  /* rasterise helpers + deck flattening for export */
  const loadImg = src => new Promise(res => { const im = new Image(); im.onload = () => res(im); im.onerror = () => res(null); im.src = src; });
  async function svgToPng(svg, w, h) {
    const im = await loadImg('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg)); if (!im) return null;
    const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w)); c.height = Math.max(1, Math.round(h)); c.getContext('2d').drawImage(im, 0, 0, c.width, c.height); return c.toDataURL('image/png');
  }
  async function nodeToPng(o, deck, P) {
    const host = el('div'); const n = window.__pptObjNode(Object.assign({}, o, { x: 0, y: 0, rot: 0 }), deck, {}); n.style.left = '0'; n.style.top = '0'; host.appendChild(n);
    const css = Array.from(document.styleSheets).map(ss => { try { return Array.from(ss.cssRules).map(r => r.cssText).join(''); } catch (e) { return ''; } }).join('');
    const k = 2; const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${o.w * k}" height="${o.h * k}"><foreignObject width="${o.w}" height="${o.h}" transform="scale(${k})"><div xmlns="http://www.w3.org/1999/xhtml" style="position:relative;width:${o.w}px;height:${o.h}px"><style>${css.replace(/</g, '&lt;')}</style>${host.innerHTML.replace(/<(video|audio)[^>]*>(<\/\1>)?/g, '')}</div></foreignObject></svg>`;
    return svgToPng(svg, o.w * k, o.h * k);
  }
  async function bakeImage(o) {
    const im = await loadImg(o.src); if (!im) return null; const cr = o.crop || {}; const sw = im.naturalWidth, sh = im.naturalHeight;
    const sx = (cr.l || 0) / 100 * sw, sy = (cr.t || 0) / 100 * sh, sW = sw * (1 - ((cr.l || 0) + (cr.r || 0)) / 100), sH = sh * (1 - ((cr.t || 0) + (cr.b || 0)) / 100);
    const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(sW)); c.height = Math.max(1, Math.round(sH)); const x = c.getContext('2d'); const f = o.flt || {};
    const fs = []; if (f.bright != null && f.bright !== 100) fs.push(`brightness(${f.bright}%)`); if (f.contr != null && f.contr !== 100) fs.push(`contrast(${f.contr}%)`); if (f.sat != null && f.sat !== 100) fs.push(`saturate(${f.sat}%)`); if (f.gray) fs.push(`grayscale(${f.gray}%)`); if (f.sepia) fs.push(`sepia(${f.sepia}%)`); if (f.blur) fs.push(`blur(${f.blur}px)`); if (f.hue) fs.push(`hue-rotate(${f.hue}deg)`); if (f.inv) fs.push(`invert(${f.inv}%)`);
    if (fs.length) x.filter = fs.join(' '); x.drawImage(im, sx, sy, sW, sH, 0, 0, c.width, c.height);
    if (o.round) { x.filter = 'none'; x.globalCompositeOperation = 'destination-in'; const r = Math.min(c.width, c.height) * Math.min(.5, o.round / 100 || .15); x.beginPath(); x.roundRect ? x.roundRect(0, 0, c.width, c.height, r) : x.rect(0, 0, c.width, c.height); x.fill(); }
    try { return c.toDataURL('image/png'); } catch (e) { return null; }
  }
  OA.pptPrepare = async function (deck0) {
    const P = PX.P; const src = deck0._mm && deck0._real ? Object.assign({}, deck0, { slides: deck0._real }) : deck0; const deck = clone(Object.assign({}, src, { _real: undefined }));
    const flatten = async (o, out, dx, dy, kx, ky, keepId) => {
      const x = Object.assign({}, o); x.x = dx + o.x * kx; x.y = dy + o.y * ky; x.w = o.w * kx; x.h = o.h * ky; if (!keepId) x.id = P.uid(); if (x.hidden) return;
      if (x.type === 'group') { for (const k of (x.kids || [])) await flatten(k, out, x.x, x.y, x.w / (x.w0 || x.w || 1), x.h / (x.h0 || x.h || 1), false); if (keepId) { /* group animations: map to first kid is lost */ } return; }
      if (['chart', 'smartart', 'wordart', 'icon'].includes(x.type) && PX.svgOf) { const png = await svgToPng(PX.svgOf(x, Math.round(x.w * 2), Math.round(x.h * 2)).replace('<svg ', `<svg xmlns="http://www.w3.org/2000/svg" `), x.w * 2, x.h * 2); if (png) { out.push(Object.assign({}, x, { type: 'image', src: png, alt: x.alt || x.type })); return; } }
      if (x.type === 'equation') { const png = await nodeToPng(x, deck, P); if (png) { out.push(Object.assign({}, x, { type: 'image', src: png, alt: x.text })); return; } }
      if (x.type === 'media') { if (x.poster) { out.push(Object.assign({}, x, { type: 'image', src: x.poster })); return; } const png = await svgToPng(`<svg xmlns="http://www.w3.org/2000/svg" width="${x.w}" height="${x.h}"><rect width="100%" height="100%" fill="#222"/><text x="50%" y="55%" fill="#fff" font-size="${Math.min(x.w, x.h) * .3}" text-anchor="middle" font-family="sans-serif">${x.kind === 'audio' ? '♪' : '▶'}</text></svg>`, x.w, x.h); if (png) { out.push(Object.assign({}, x, { type: 'image', src: png })); return; } }
      if (x.type === 'image' && (x.crop || x.flt || x.round) && x.src) { const b = await bakeImage(x); if (b) { x.src = b; delete x.crop; delete x.flt; delete x.round; } }
      if (x.html && PX.htmlRuns) { x.runs = PX.htmlRuns(x.html); if (!x.runs.length) delete x.runs; }
      out.push(x);
    };
    for (let si = 0; si < deck.slides.length; si++) {
      const s = deck.slides[si]; const out = [];
      if (s.bgImg) { out.push({ id: P.uid(), type: 'image', x: 0, y: 0, w: deck.w, h: deck.h, src: s.bgImg, name: 'Background' }); s.bgImg = null; }
      else if (!s.bg && deck.master && deck.master.bg) s.bg = deck.master.bg;
      if (!s.nomaster) {
        const layers = [].concat((deck.master && deck.master.objs) || [], (deck.lay && deck.lay[s.layout]) || []); for (const o of layers) await flatten(o, out, 0, 0, 1, 1, false);
        const hf = deck.hf; if (hf && !(hf.notTitle && s.layout === 'title')) {
          const base = { type: 'text', sz: 12, color: (deck.theme && deck.theme.fg) || '#888888', y: deck.h * .93, h: deck.h * .05, va: 'middle' }; const add = (t, x, w, ha) => out.push(Object.assign({ id: P.uid(), text: t, x, w, ha: ha || 'left' }, base));
          if (hf.date && hf.date.on) add(hf.date.auto ? new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }) : hf.date.text || '', deck.w * .04, deck.w * .25);
          if (hf.footer && hf.footer.on) add(hf.footer.text || '', deck.w * .3, deck.w * .4, 'center');
          if (hf.num && hf.num.on) add(String(si + (deck.firstNum || 1)), deck.w * .85, deck.w * .11, 'right');
        }
      }
      for (const o of (s.objs || [])) await flatten(o, out, 0, 0, 1, 1, true);
      s.objs = out;
    }
    return deck;
  };

  /* ================= editor UI ================= */
  OA.pptExtras = OA.pptExtras || [];
  OA.pptExtras.push(function (P, spec) {
    const { B, L, toast, uid } = P; const tab = id => spec.tabs.find(t => t.id === id); const D = () => P.deck;
    let selIdx = -1, paneOn = false, pane = null, painter = null;
    const own = s => { if (!s.anim) { s.anim = getAnims(s); (s.objs || []).forEach(o => { delete o.anim; }); } return s.anim; };
    const curA = () => { const s = P.S(); if (!s) return null; const a = getAnims(s); if (selIdx >= 0 && a[selIdx]) return a[selIdx]; const o = P.SEL(); return o ? a.find(x => x.oid === o.id) || null : null; };
    const oname = o => o ? (o.name || (o.type === 'text' || o.type === 'shape' ? (String(o.text || '').slice(0, 18) || o.type) : o.type) + ' ' + String(o.id).slice(-3)) : '(missing)';
    const addAnim = (cls, eff) => {
      const o = P.SEL(); if (!o) return toast('Pehle object select karein'); P.mutate(() => { const a = own(P.S()); const n = { oid: o.id, cls, eff, trig: 'click', dur: cls === 'path' ? 1 : .6, delay: 0, dir: 'b', rep: 1 }; a.push(n); selIdx = a.length - 1; });
    };
    const setMain = v => {
      const o = P.SEL(); if (!o) return toast('Pehle object select karein'); P.mutate(() => { const a = own(P.S()); const k = a.findIndex(x => x.oid === o.id); if (v === 'none') { for (let j = a.length - 1; j >= 0; j--) if (a[j].oid === o.id) a.splice(j, 1); selIdx = -1; return; } const [cls, eff] = v.split(':'); if (k >= 0) { a[k].cls = cls; a[k].eff = eff; selIdx = k; } else { a.push({ oid: o.id, cls, eff, trig: 'click', dur: cls === 'path' ? 1 : .6, delay: 0, dir: 'b', rep: 1 }); selIdx = a.length - 1; } });
    };
    const edit = fn => { const a = curA(); if (!a) return toast('Pehle animation wala object select karein'); P.mutate(() => { own(P.S()); const real = P.S().anim.find(x => x === a) || P.S().anim.find(x => x.oid === a.oid); real && fn(real); }); };
    const move = dlt => { const a = curA(); if (!a) return; P.mutate(() => { const l = own(P.S()); const k = l.indexOf(a) >= 0 ? l.indexOf(a) : l.findIndex(x => x.oid === a.oid); const j = k + dlt; if (k < 0 || j < 0 || j >= l.length) return; [l[k], l[j]] = [l[j], l[k]]; selIdx = j; }); };
    const galleryOpts = [['none', 'None']].concat(Object.keys(EFF).flatMap(c => EFF[c].map(e => [c + ':' + e[0], CLS[c] + ' – ' + e[1]])));
    const addMenu = Object.keys(EFF).flatMap(c => EFF[c].map(e => ({ label: CLS[c] + ' – ' + e[1], onClick: () => addAnim(c, e[0]) })));
    /* editor preview */
    function preview() {
      P.finishEdit(); const s = P.S(); const rn = Runner(s, P.stage, D()); rn.init(); let guard = 0; const go = () => { if (guard++ > 60) return; if (!rn.next()) { setTimeout(() => P.renderStage(), 500); return; } setTimeout(go, rn.lastMs + 250); }; if (!rn.steps.length) return toast('Is slide par koi animation nahi hai'); go();
    }
    PX.previewAnim = preview;
    /* animation pane */
    function buildPane() {
      pane && pane.remove(); pane = null; if (!paneOn) return; const s = P.S(); const a = getAnims(s); const list = el('div', { class: 'ap-list' });
      a.forEach((x, k) => { const o = P.O(x.oid); const row = el('div', { class: 'ap-row' + (k === selIdx ? ' sel' : '') }); const ic = el('i', { class: 'c-' + x.cls, text: { entr: '★', emph: '✦', exit: '✖', path: '➤' }[x.cls] }); row.append(el('b', { text: String(k + 1) }), ic, el('span', { class: 'ap-n', text: oname(o) + ' · ' + effName(x) }), el('small', { text: { click: '🖱', with: '⊕', after: '⏭' }[x.trig] || '' }));
        row.onclick = () => { selIdx = k; if (o) P.select(o.id); buildPane(); }; list.appendChild(row); });
      if (!a.length) list.appendChild(el('div', { class: 'ap-empty', text: 'No animations on this slide. Select an object and choose an effect.' }));
      const bar = el('div', { class: 'p-pbar' }, [el('b', { text: 'Animation Pane' }), el('button', { class: 'btn', type: 'button', text: '▶ Play', onclick: preview }), el('button', { class: 'btn', type: 'button', text: '✕', onclick: () => { paneOn = false; buildPane(); P.renderStage(); } })]);
      pane = el('div', { class: 'p-anipane' }, [bar, list]); P.root.appendChild(pane);
    }
    OA.pptSync = OA.pptSync || []; OA.pptSync.push(function (P2, rb) {
      const c = rb.ctl; const a = curA(); const o = P.SEL();
      if (c.agal) { const m = o && getAnims(P.S()).find(x => x.oid === o.id); c.agal.value = m ? m.cls + ':' + m.eff : 'none'; }
      if (c.astart) { c.astart.value = a ? a.trig : 'click'; c.adur.value = a ? a.dur : .5; c.adel.value = a ? a.delay || 0 : 0; c.adir.value = a ? a.dir || 'b' : 'b'; c.arep.value = a ? String(a.rep || 1) : '1'; }
      const s = P.S(); if (c.tdir && s) { c.tdir.value = (s.trans && s.trans.dir) || 'r'; c.tadv.value = (s.trans && s.trans.advTm) || 0; rb.setOn('tclick', !(s.trans && s.trans.advClick === false)); }
      if (paneOn) buildPane();
    });
    OA.pptObjHooks = OA.pptObjHooks || []; OA.pptObjHooks.push(function (n, o, deck, opts) {
      if (opts && opts.anim) return; if (!paneOn || deck._mm) return; const s = deck.slides[P.cur]; if (!s) return; const a = getAnims(s); const idx = a.map((x, k) => x.oid === o.id ? k + 1 : 0).filter(Boolean); if (idx.length) { const b = document.createElement('span'); b.className = 'p-abadge'; b.textContent = idx.join(','); n.appendChild(b); }
    });
    /* tabs */
    const at = tab('anim'); if (at) at.groups = [
      { label: 'Preview', items: [L('apv', 'Preview', 'play', preview)] },
      { label: 'Animation', items: [{ type: 'select', id: 'agal', options: galleryOpts, width: 170, title: 'Animation of selected object', onChange: setMain }, B('aadd', 'Add Animation', null, null, { glyph: '＋✦', label: 'Add Animation', menu: addMenu })] },
      { label: 'Effect Options', items: [{ type: 'select', id: 'adir', options: DIRS, width: 100, title: 'Direction', onChange: v => edit(a => { a.dir = v; }) }, { type: 'select', id: 'arep', options: [['1', 'Repeat: 1'], ['2', 'Repeat: 2'], ['3', 'Repeat: 3'], ['5', 'Repeat: 5'], ['10', 'Repeat: 10']], width: 100, title: 'Repeat', onChange: v => edit(a => { a.rep = +v; }) }] },
      { label: 'Advanced Animation', items: [B('apane', 'Animation Pane', null, () => { paneOn = !paneOn; buildPane(); P.renderStage(); }, { glyph: '☰', label: 'Animation Pane' }), B('atrig', 'Trigger', null, async () => {
        const a = curA(); if (!a) return toast('Pehle animation wala object select karein'); const objs = P.S().objs; const v = await OA.form('Trigger', [{ id: 't', label: 'Start effect', type: 'select', value: a.trigOid || '', options: [['', 'As part of the click sequence']].concat(objs.map(o => [o.id, 'On click of: ' + oname(o)])) }]); if (v) edit(x => { if (v.t) x.trigOid = v.t; else delete x.trigOid; });
      }, { glyph: '⚡', label: 'Trigger' }), B('apaint', 'Animation Painter', null, () => { const a = curA(); if (!a) return toast('Pehle animation wala object select karein'); painter = clone(a); toast('Ab dusre object par click karke "Paste" dabayein'); }, { glyph: '🖌', label: 'Anim Painter' }), B('apaste', 'Apply painted animation to selected', null, () => { if (!painter) return toast('Pehle Animation Painter se copy karein'); const o = P.SEL(); if (!o) return; P.mutate(() => { const l = own(P.S()); const n = Object.assign(clone(painter), { oid: o.id }); delete n.trigOid; l.push(n); selIdx = l.length - 1; }); }, { glyph: '⎘', label: 'Paint' }), B('arm', 'Remove animation', null, () => { const o = P.SEL(); if (!o) return; setMain('none'); }, { glyph: '🗑', label: 'Remove' })] },
      { label: 'Timing', items: [{ type: 'select', id: 'astart', options: [['click', 'Start: On Click'], ['with', 'Start: With Previous'], ['after', 'Start: After Previous']], width: 150, title: 'Start', onChange: v => edit(a => { a.trig = v; }) }, { type: 'number', id: 'adur', title: 'Duration', label: 'Duration', value: .5, min: .1, step: .1, width: 56, onChange: v => v > 0 && edit(a => { a.dur = v; }) }, { type: 'number', id: 'adel', title: 'Delay', label: 'Delay', value: 0, min: 0, step: .1, width: 56, onChange: v => v != null && v >= 0 && edit(a => { a.delay = v; }) }, B('aup', 'Move Earlier', null, () => move(-1), { glyph: '▲', label: 'Earlier' }), B('adn', 'Move Later', null, () => move(1), { glyph: '▼', label: 'Later' })] }];
    /* transitions tab */
    const tt = tab('trans'); const mt = fn => P.mutate(() => { const s = P.S(); s.trans = Object.assign({ type: 'none', dur: 1 }, s.trans); fn(s.trans); });
    if (tt) tt.groups.push({ label: 'Effect / Advance', items: [{ type: 'select', id: 'tdir', options: [['r', 'From Right'], ['l', 'From Left'], ['u', 'From Top'], ['d', 'From Bottom']], width: 100, title: 'Effect options', onChange: v => mt(t => { t.dir = v; }) }, B('tclick', 'On Mouse Click', null, () => mt(t => { t.advClick = t.advClick === false ? true : false; }), { glyph: '🖱', label: 'On Click' }), { type: 'number', id: 'tadv', title: 'Advance after (s)', label: 'After (s)', value: 0, min: 0, step: 1, width: 56, onChange: v => mt(t => { if (v > 0) t.advTm = v; else delete t.advTm; }) }] });
    /* slide show tab */
    const st = tab('show'); const setup = async () => {
      const d = D(); const c = Object.assign({ type: 'presenter', loop: false, noAnim: false, useTimings: true, from: 1, to: 0 }, d.show || {});
      const v = await OA.form('Set Up Show', [{ id: 'type', label: 'Show type', type: 'select', value: c.type, options: [['presenter', 'Presented by a speaker (full screen)'], ['kiosk', 'Browsed at a kiosk (full screen, loops)']] }, { id: 'loop', label: 'Loop continuously until Esc', type: 'checkbox', value: c.loop }, { id: 'noAnim', label: 'Show without animation', type: 'checkbox', value: c.noAnim }, { id: 'useTimings', label: 'Use timings, if present', type: 'checkbox', value: c.useTimings }, { id: 'from', label: 'Slides from', type: 'number', value: c.from || 1, min: 1 }, { id: 'to', label: 'to (0 = last)', type: 'number', value: c.to || 0, min: 0 }, { id: 'pv', label: 'Start in Presenter View', type: 'checkbox', value: !!c.pv }]);
      if (v) P.mutate(() => { d.show = Object.assign({}, d.show, v); });
    };
    const customDlg = async () => {
      const d = D(); const cs = CUSTOM(d); const cur = cs.map(c => c.name + ': ' + c.ids.map(id => d.slides.findIndex(s => s.id === id) + 1).join(',')).join('\n');
      const v = await OA.form('Custom Shows', [{ id: 'name', label: 'Show name (existing name = replace)', value: 'Custom Show ' + (cs.length + 1) }, { id: 'nums', label: 'Slide numbers (comma-separated, e.g. 1,3,5 — blank = delete this show)', value: '' }, { id: 'info', label: 'Existing: ' + (cur ? cur.replace(/\n/g, ' | ') : 'none'), type: 'checkbox', value: false }]);
      if (!v || !v.name) return; P.mutate(() => { const k = cs.findIndex(c => c.name === v.name); if (!v.nums.trim()) { if (k >= 0) cs.splice(k, 1); return; } const ids = v.nums.split(/[ ,]+/).map(n => parseInt(n) - 1).filter(n => n >= 0 && n < d.slides.length).map(n => d.slides[n].id); if (!ids.length) return; const o = { name: v.name, ids }; if (k >= 0) cs[k] = o; else cs.push(o); });
    };
    const customMenu = () => { const cs = CUSTOM(D()); return cs.length ? cs.map(c => ({ label: c.name, onClick: () => OA.pptPresent(P, null, { custom: c.name }) })).concat([{ label: 'Custom Shows… (edit)', onClick: customDlg }]) : [{ label: 'Custom Shows… (create)', onClick: customDlg }]; };
    if (st) {
      st.groups[0].items.push(B('fpv', 'Presenter View', null, () => OA.pptPresent(P, 0, { pv: true }), { glyph: '🗒', label: 'Presenter View', size: 'lg' }));
      st.groups.push({ label: 'Set Up', items: [L('sus', 'Set Up Slide Show', 'settings', setup), B('rht', 'Rehearse Timings', null, () => OA.pptPresent(P, 0, { rehearse: true }), { glyph: '⏺', label: 'Rehearse', size: 'lg' }), B('hs', 'Hide Slide', null, () => P.mutate(() => { P.S().hide = !P.S().hide; }), { glyph: '🙈', label: 'Hide Slide' }), B('cs', 'Custom Slide Show', null, null, { glyph: '☰▶', label: 'Custom Show', menu: customMenu })] });
    }
    PX.setupShow = setup; PX.customDlg = customDlg;
  });
})();
