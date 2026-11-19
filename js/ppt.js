/* PowerPoint app UI — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, el = OA.el, IO = OA.pptIO;
  const uid = () => 's' + Math.random().toString(36).slice(2, 9);
  const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Georgia', 'Tahoma', 'Trebuchet MS', 'Impact', 'Segoe UI', 'Cambria', 'Consolas'];
  const THEMES = [
    { name: 'Office', bg: '#ffffff', fg: '#222222', accent: '#4472c4', accent2: '#ed7d31', font: 'Calibri' },
    { name: 'Dark', bg: '#1f2430', fg: '#f2f2f2', accent: '#ff9f43', accent2: '#54a0ff', font: 'Segoe UI' },
    { name: 'Ion Green', bg: '#f1faf4', fg: '#14352a', accent: '#217346', accent2: '#f4b400', font: 'Calibri' },
    { name: 'Sunset', bg: '#fff4ec', fg: '#3b1d12', accent: '#e4572e', accent2: '#f3a712', font: 'Georgia' },
    { name: 'Ocean', bg: '#eaf4fb', fg: '#0b2a43', accent: '#1b6ca8', accent2: '#2ec4b6', font: 'Trebuchet MS' },
    { name: 'Midnight', bg: '#0b1020', fg: '#e8ecff', accent: '#7c83ff', accent2: '#ff6b9d', font: 'Segoe UI' },
    { name: 'Paper', bg: '#faf7f0', fg: '#2d2a26', accent: '#8c5a3c', accent2: '#3c7a6b', font: 'Cambria' },
    { name: 'Mono', bg: '#ffffff', fg: '#111111', accent: '#555555', accent2: '#999999', font: 'Arial' }
  ];
  const SHAPES = [['rect', 'Rectangle'], ['round', 'Rounded Rectangle'], ['ellipse', 'Oval'], ['triangle', 'Triangle'], ['diamond', 'Diamond'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'], ['star', 'Star'], ['heart', 'Heart'], ['cloud', 'Cloud'], ['arrow', 'Right Arrow'], ['rarrow', 'Left Arrow'], ['callout', 'Callout'], ['line', 'Line']];
  const GEO = {
    rect: '<rect x="0" y="0" width="100" height="100"/>', round: '<rect x="0" y="0" width="100" height="100" rx="14" ry="14"/>', ellipse: '<ellipse cx="50" cy="50" rx="50" ry="50"/>',
    triangle: '<polygon points="50,0 100,100 0,100"/>', diamond: '<polygon points="50,0 100,50 50,100 0,50"/>', star: '<polygon points="50,0 61,35 98,35 68,57 79,91 50,70 21,91 32,57 2,35 39,35"/>',
    arrow: '<polygon points="0,30 60,30 60,10 100,50 60,90 60,70 0,70"/>', rarrow: '<polygon points="100,30 40,30 40,10 0,50 40,90 40,70 100,70"/>', hexagon: '<polygon points="25,0 75,0 100,50 75,100 25,100 0,50"/>',
    pentagon: '<polygon points="50,0 100,38 82,100 18,100 0,38"/>', heart: '<path d="M50 92 C8 60 0 35 20 15 C35 2 50 15 50 27 C50 15 65 2 80 15 C100 35 92 60 50 92 Z"/>',
    cloud: '<path d="M25 85 C2 85 2 52 25 52 C18 22 52 12 62 36 C76 18 104 34 92 56 C108 72 92 88 76 85 Z"/>', callout: '<polygon points="0,0 100,0 100,72 42,72 20,100 26,72 0,72"/>', line: '<line x1="0" y1="50" x2="100" y2="50"/>'
  };
  const T = deck => deck.theme;

  /* ---------- slide layouts */
  const LAYOUTS = [['title', 'Title Slide'], ['content', 'Title and Content'], ['section', 'Section Header'], ['two', 'Two Content'], ['titleonly', 'Title Only'], ['blank', 'Blank']];
  function txt(deck, o) { const t = T(deck); return Object.assign({ id: uid(), type: 'text', text: '', font: t.font, sz: 24, color: t.fg, ha: 'left', va: 'top' }, o); }
  function layoutObjs(deck, lay, h) {
    const t = T(deck), W = deck.w, H = deck.h || 540; const k = W / 960;
    const title = (o) => txt(deck, Object.assign({ x: 60 * k, y: 36 * k, w: W - 120 * k, h: 90 * k, text: 'Click to add title', sz: 40, b: true, va: 'middle', color: t.accent }, o));
    const body = (o) => txt(deck, Object.assign({ x: 60 * k, y: 140 * k, w: W - 120 * k, h: H - 190 * k, text: 'Click to add text\nSecond level point\nThird level point', sz: 26, bullet: true }, o));
    switch (lay) {
      case 'title': return [txt(deck, { x: 90 * k, y: H * 0.28, w: W - 180 * k, h: 120 * k, text: 'Presentation Title', sz: 54, b: true, ha: 'center', va: 'middle', color: t.accent }), txt(deck, { x: 140 * k, y: H * 0.28 + 130 * k, w: W - 280 * k, h: 70 * k, text: 'Subtitle or author name', sz: 26, ha: 'center', va: 'top' })];
      case 'content': return [title(), body()];
      case 'section': return [txt(deck, { x: 80 * k, y: H * 0.34, w: W - 160 * k, h: 100 * k, text: 'Section Title', sz: 48, b: true, va: 'bottom', color: t.accent }), txt(deck, { x: 80 * k, y: H * 0.34 + 110 * k, w: W - 160 * k, h: 60 * k, text: 'Short description', sz: 24 })];
      case 'two': return [title(), body({ w: (W - 150 * k) / 2 }), body({ x: 60 * k + (W - 150 * k) / 2 + 30 * k, w: (W - 150 * k) / 2 })];
      case 'titleonly': return [title()];
      default: return [];
    }
  }
  const newSlide = (deck, lay) => ({ id: uid(), layout: lay || 'blank', bg: null, trans: { type: 'none', dur: 1 }, notes: '', objs: layoutObjs(deck, lay || 'blank') });
  const newDeck = tpl => {
    const d = { w: 960, h: 540, title: 'Presentation1', theme: Object.assign({}, THEMES[0]), slides: [] };
    if (tpl === 'pitch') { d.theme = Object.assign({}, THEMES[1]); const s1 = newSlide(d, 'title'); s1.objs[0].text = 'Startup Pitch Deck'; s1.objs[1].text = 'One line that explains what you do'; const mk = (t, items) => { const s = newSlide(d, 'content'); s.objs[0].text = t; s.objs[1].text = items; return s; };
      d.slides = [s1, mk('The Problem', 'Describe the pain your customers feel\nWho has this problem?\nHow big is it?'), mk('Our Solution', 'What we built\nWhy it is 10x better\nHow it works in 3 steps'), mk('Market Opportunity', 'Total market size\nTarget segment\nGrowth rate'), mk('Business Model', 'How we make money\nPricing\nUnit economics'), mk('The Ask', 'Funding required\nUse of funds\nKey milestones')]; }
    else if (tpl === 'update') { d.theme = Object.assign({}, THEMES[2]); const s1 = newSlide(d, 'title'); s1.objs[0].text = 'Project Status Update'; s1.objs[1].text = 'Prepared by: [Name]  •  [Date]'; const s2 = newSlide(d, 'content'); s2.objs[0].text = 'Highlights'; s2.objs[1].text = 'Milestone 1 completed\nBudget on track\nTeam grew by 2 members'; const s3 = newSlide(d, 'titleonly'); s3.objs[0].text = 'Progress'; const tb = { id: uid(), type: 'table', x: 60, y: 150, w: 840, h: 220, sz: 18, font: d.theme.font, hc: d.theme.accent, band: '#dff0e6', rows: [['Task', 'Owner', 'Status', 'Due'], ['Design', 'Asha', 'Done', 'Nov 10'], ['Build', 'Rohan', 'In progress', 'Dec 1'], ['Launch', 'Team', 'Planned', 'Dec 20']] }; s3.objs.push(tb); const s4 = newSlide(d, 'content'); s4.objs[0].text = 'Next Steps'; s4.objs[1].text = 'Finalize requirements\nStart user testing\nPlan release'; d.slides = [s1, s2, s3, s4]; }
    else if (tpl === 'lesson') { d.theme = Object.assign({}, THEMES[4]); const s1 = newSlide(d, 'title'); s1.objs[0].text = 'Lesson Title'; s1.objs[1].text = 'Subject • Class • Teacher'; const s2 = newSlide(d, 'content'); s2.objs[0].text = 'Learning Objectives'; s2.objs[1].text = 'Understand the key concept\nApply it to examples\nSolve practice problems'; const s3 = newSlide(d, 'two'); s3.objs[0].text = 'Compare'; s3.objs[1].text = 'Point A\nPoint B'; s3.objs[2].text = 'Point C\nPoint D'; const s4 = newSlide(d, 'section'); s4.objs[0].text = 'Questions?'; s4.objs[1].text = 'Thank you'; d.slides = [s1, s2, s3, s4]; }
    else d.slides = [newSlide(d, 'title')];
    return d;
  }

  /* ---------- rendering */
  function textStyle(o) { return `font-family:'${o.font || 'Calibri'}',Calibri,Carlito,sans-serif;font-size:${o.sz || 18}px;font-weight:${o.b ? 700 : 400};font-style:${o.i ? 'italic' : 'normal'};text-decoration:${o.u ? 'underline' : 'none'};color:${o.color || '#000'};text-align:${o.ha || 'left'};line-height:1.2;`; }
  function objNode(o, deck, opts) {
    const n = el('div', { class: 'p-obj', 'data-id': o.id }); n.style.cssText = `left:${o.x}px;top:${o.y}px;width:${o.w}px;height:${o.h}px;` + (o.rot ? `transform:rotate(${o.rot}deg);` : '');
    if (o.type === 'image') { const im = el('img', { src: o.src, alt: o.alt || '', draggable: 'false' }); n.appendChild(im); if (o.line) n.style.outline = (o.lw || 1) + 'px solid ' + o.line; }
    else if (o.type === 'table') {
      const tb = el('table'); const rows = o.rows || [['']]; const th = T(deck); rows.forEach((r, ri) => { const tr = el('tr'); r.forEach(c => { const hd = ri === 0 && o.header !== false; const td = el('td', { text: c }); td.style.cssText = textStyle(Object.assign({}, o, { b: hd, color: hd ? '#fff' : (o.color || th.fg), ha: 'left' })) + `background:${hd ? (o.hc || th.accent) : (ri % 2 === 0 ? (o.band || '#e9edf7') : '#fff')};` + (hd ? '' : 'color:#222;'); tr.appendChild(td); }); tb.appendChild(tr); }); n.appendChild(tb);
    } else {
      if (o.type === 'shape') { const ns = 'http://www.w3.org/2000/svg'; const k = o.kind || 'rect'; const s = el('div'); s.innerHTML = `<svg class="shp" viewBox="0 0 100 100" preserveAspectRatio="none" xmlns="${ns}"><g fill="${k === 'line' ? 'none' : (o.fill || 'none')}" stroke="${o.line || (k === 'line' ? (o.fill || '#000') : 'none')}" stroke-width="${o.lw || 1}" vector-effect="non-scaling-stroke">${(GEO[k] || GEO.rect).replace(/<(rect|ellipse|polygon|path|line)/g, '<$1 vector-effect="non-scaling-stroke"')}</g></svg>`; n.appendChild(s.firstChild); }
      else { if (o.fill) n.style.background = o.fill; if (o.line) n.style.border = (o.lw || 1) + 'px solid ' + o.line; }
      const hasText = o.text !== undefined && (o.type === 'text' || o.text !== '');
      if (hasText || opts.edit) {
        const box = el('div', { class: 'p-t' }); const jc = { top: 'flex-start', middle: 'center', bottom: 'flex-end' }[o.va || (o.type === 'shape' ? 'middle' : 'top')]; box.style.justifyContent = jc || 'flex-start';
        const tx = el('div', { class: 'tx' + (o.bullet ? ' bul' : '') }); tx.style.cssText = textStyle(Object.assign({}, o, { ha: o.ha || (o.type === 'shape' ? 'center' : 'left') }));
        String(o.text || '').split('\n').forEach(l => { tx.appendChild(el('div', { text: l === '' ? '​' : l })); }); box.appendChild(tx); n.appendChild(box);
      }
    }
    if (o.anim && o.anim !== 'none' && opts.anim) n.dataset.anim = o.anim;
    return n;
  }
  function renderSlide(slide, deck, opts) {
    opts = opts || {}; const s = el('div', { class: 'p-slide' }); s.style.cssText = `width:${deck.w}px;height:${deck.h}px;background:${slide.bg || T(deck).bg};`;
    slide.objs.forEach(o => s.appendChild(objNode(o, deck, opts))); return s;
  }

  OA.registerApp({
    id: 'ppt', name: 'PowerPoint', letter: 'P', color: '#b7472a', accept: ['pptx'], defaultName: 'Presentation1',
    desc: 'Design slides with shapes, images, tables, themes, transitions, animations and a full-screen slide show. Export to .pptx.',
    templates: [{ id: 'pitch', label: 'Pitch Deck' }, { id: 'update', label: 'Status Update' }, { id: 'lesson', label: 'Lesson' }],
    create(root) {
      const a = OA.apps.ppt; const sh = OA.appShell(root, a); root.classList.add('app-ppt'); const toast = OA.toast;
      let deck = newDeck(), cur = 0, selId = null, undoS = [], redoS = [], clip = null, zoomMul = 1, scale = 1, dirty = false, editing = null, rb, bs, showGrid = false, showNotes = true, thumbT = 0;
      const S = () => deck.slides[cur], O = id => S().objs.find(o => o.id === id), SEL = () => selId ? O(selId) : null;
      const baseName = () => (sh.name.value.trim() || 'Presentation1').replace(/\.pptx?$/i, '');

      /* ---------- layout skeleton */
      const thumbs = el('div', { class: 'p-thumbs' }), main = el('div', { class: 'p-main' }), box = el('div', { class: 'p-stagebox', tabindex: '0' }), wrap = el('div', { class: 'p-stagewrap' }), stage = el('div', { class: 'p-stage' }), gridEl = el('div', { class: 'p-grid' });
      const notes = el('textarea', { class: 'p-notes', placeholder: 'Click to add notes' }); wrap.appendChild(stage); box.appendChild(wrap); main.append(box, notes); sh.content.append(thumbs, main);
      const sbSlide = el('span'), sbInfo = el('span', { class: 'sb-hide-s', text: 'English (India)' }), zoomLbl = el('button', { type: 'button', text: 'Fit', onclick: () => { zoomMul = 1; fit(); } });
      const zoomRange = el('input', { type: 'range', min: 30, max: 250, value: 100, 'aria-label': 'Zoom' }); zoomRange.addEventListener('input', () => { zoomMul = zoomRange.value / 100; fit(); });
      sh.status.append(sbSlide, sbInfo, el('span', { class: 'sb-fill' }), el('span', { class: 'sb-dev sb-hide-s', text: 'Developer: ' + OA.DEV }), el('button', { type: 'button', text: '▤ Notes', title: 'Show / hide notes', onclick: () => { showNotes = !showNotes; notes.hidden = !showNotes; fit(); } }), el('button', { type: 'button', html: OA.icon('play', 14) + ' Show', title: 'Slide Show (F5)', onclick: () => present(cur) }), el('button', { type: 'button', html: OA.icon('zoomOut', 15), onclick: () => { zoomMul = Math.max(0.3, zoomMul - 0.1); fit(); } }), zoomRange, el('button', { type: 'button', html: OA.icon('zoomIn', 15), onclick: () => { zoomMul = Math.min(2.5, zoomMul + 0.1); fit(); } }), zoomLbl);

      /* ---------- history */
      const snap = () => { undoS.push(JSON.stringify(deck)); if (undoS.length > 60) undoS.shift(); redoS = []; dirty = true; };
      function restore(j) { deck = JSON.parse(j); cur = Math.min(cur, deck.slides.length - 1); selId = null; editing = null; renderAll(); }
      const undo = () => { if (!undoS.length) return; finishEdit(true); redoS.push(JSON.stringify(deck)); restore(undoS.pop()); };
      const redo = () => { if (!redoS.length) return; undoS.push(JSON.stringify(deck)); restore(redoS.pop()); };
      const mutate = fn => { finishEdit(); snap(); fn(); renderAll(); };

      /* ---------- stage rendering */
      function fit() {
        const pad = 28; const bw = box.clientWidth - pad, bh = box.clientHeight - pad; if (bw < 50 || bh < 50) return; scale = Math.min(bw / deck.w, bh / deck.h) * zoomMul;
        wrap.style.width = deck.w * scale + 'px'; wrap.style.height = deck.h * scale + 'px'; stage.style.transform = `scale(${scale})`; stage.style.width = deck.w + 'px'; stage.style.height = deck.h + 'px'; zoomLbl.textContent = zoomMul === 1 ? 'Fit' : Math.round(zoomMul * 100) + '%'; zoomRange.value = Math.round(zoomMul * 100); box.style.placeItems = zoomMul > 1.001 ? 'start' : 'center';
        drawSel();
      }
      function renderStage() {
        stage.innerHTML = ''; const sl = renderSlide(S(), deck, {}); stage.appendChild(sl); if (showGrid) { gridEl.style.cssText = ''; stage.appendChild(gridEl); }
        if (editing) { const n = nodeOf(editing); n && beginEditDom(n); } drawSel(); if (document.activeElement !== notes) notes.value = S().notes || '';
        sbSlide.textContent = 'Slide ' + (cur + 1) + ' of ' + deck.slides.length;
      }
      const nodeOf = id => stage.querySelector('.p-obj[data-id="' + id + '"]');
      function renderThumbs() {
        const keep = thumbs.scrollTop; thumbs.innerHTML = '';
        deck.slides.forEach((s, i) => {
          const w = 150, k = w / deck.w; const bx = el('div', { class: 'box' }); bx.style.width = w + 'px'; bx.style.height = deck.h * k + 'px'; const sl = renderSlide(s, deck, {}); sl.style.transform = `scale(${k})`; sl.style.transformOrigin = '0 0'; bx.appendChild(sl);
          const th = el('div', { class: 'p-th' + (i === cur ? ' on' : ''), draggable: 'true', 'data-i': i }, [el('span', { class: 'n', text: i + 1 }), bx]);
          th.addEventListener('click', () => gotoSlide(i)); th.addEventListener('contextmenu', e => { e.preventDefault(); gotoSlide(i); thumbMenu(e); });
          th.addEventListener('dragstart', e => { e.dataTransfer.setData('text/plain', String(i)); e.dataTransfer.effectAllowed = 'move'; }); th.addEventListener('dragover', e => { e.preventDefault(); th.classList.add('over'); }); th.addEventListener('dragleave', () => th.classList.remove('over'));
          th.addEventListener('drop', e => { e.preventDefault(); th.classList.remove('over'); const f = +e.dataTransfer.getData('text/plain'); if (!isNaN(f) && f !== i) moveSlide(f, i); });
          thumbs.appendChild(th);
        }); thumbs.scrollTop = keep; const on = thumbs.querySelector('.on'); on && on.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
      const thumbsSoon = () => { clearTimeout(thumbT); thumbT = setTimeout(renderThumbs, 120); };
      function renderAll() { renderStage(); renderThumbs(); fit(); syncRibbon(); }
      function gotoSlide(i) { if (i < 0 || i >= deck.slides.length || i === cur) { return; } finishEdit(); cur = i; selId = null; renderStage(); renderThumbs(); fit(); syncRibbon(); }

      /* ---------- selection overlay */
      let selEl = null;
      function drawSel() {
        if (selEl) { selEl.remove(); selEl = null; } const o = SEL(); if (!o || editing === o.id) { if (!o) return; } if (!o) return;
        const b = el('div', { class: 'p-selbox' }); b.style.cssText = `left:${o.x}px;top:${o.y}px;width:${o.w}px;height:${o.h}px;border-width:${1.5 / scale}px;` + (o.rot ? `transform:rotate(${o.rot}deg);` : '');
        const hs = 11 / scale; const pts = [['nw', 0, 0], ['n', .5, 0], ['ne', 1, 0], ['e', 1, .5], ['se', 1, 1], ['s', .5, 1], ['sw', 0, 1], ['w', 0, .5]];
        if (editing !== o.id) pts.forEach(([k, px, py]) => { const h = el('div', { class: 'p-h', 'data-h': k }); h.style.cssText = `left:${px * 100}%;top:${py * 100}%;width:${hs}px;height:${hs}px;margin:${-hs / 2}px 0 0 ${-hs / 2}px;border-width:${1.5 / scale}px;cursor:${{ nw: 'nwse', se: 'nwse', ne: 'nesw', sw: 'nesw', n: 'ns', s: 'ns', e: 'ew', w: 'ew' }[k]}-resize`; h.addEventListener('pointerdown', e => startResize(e, k)); b.appendChild(h); });
        if (editing !== o.id) { const r = el('div', { class: 'p-h rot' }); r.style.cssText = `left:50%;top:${-26 / scale}px;width:${hs}px;height:${hs}px;margin:${-hs / 2}px 0 0 ${-hs / 2}px;border-width:${1.5 / scale}px`; r.addEventListener('pointerdown', startRotate); b.appendChild(r); }
        stage.appendChild(b); selEl = b;
      }
      function select(id) { if (selId === id) return; finishEdit(); selId = id; drawSel(); syncRibbon(); }
      const stagePt = e => { const r = wrap.getBoundingClientRect(); return { x: (e.clientX - r.left) / scale, y: (e.clientY - r.top) / scale }; };
      let drag = null;
      const moveNode = o => { const n = nodeOf(o.id); if (n) { n.style.left = o.x + 'px'; n.style.top = o.y + 'px'; n.style.width = o.w + 'px'; n.style.height = o.h + 'px'; n.style.transform = o.rot ? `rotate(${o.rot}deg)` : ''; } if (selEl) { selEl.style.left = o.x + 'px'; selEl.style.top = o.y + 'px'; selEl.style.width = o.w + 'px'; selEl.style.height = o.h + 'px'; selEl.style.transform = o.rot ? `rotate(${o.rot}deg)` : ''; } };
      const snapV = (v, targets) => { for (const t of targets) if (Math.abs(v - t) < 5 / Math.max(scale, .3)) return t; return v; };
      wrap.addEventListener('pointerdown', e => {
        if (e.target.closest('.p-h')) return; const n = e.target.closest('.p-obj'); if (editing && n && n.dataset.id === editing) return;
        box.focus({ preventScroll: true }); if (!n) { select(null); return; }
        const id = n.dataset.id; if (selId !== id) { finishEdit(); selId = id; drawSel(); syncRibbon(); } const o = O(id); if (!o) return; e.preventDefault();
        const p = stagePt(e); drag = { mode: 'move', id, sx: p.x, sy: p.y, ox: o.x, oy: o.y, moved: false, t: Date.now() }; wrap.setPointerCapture(e.pointerId);
      });
      wrap.addEventListener('pointermove', e => {
        if (!drag) return; const o = O(drag.id); if (!o) return; const p = stagePt(e);
        if (drag.mode === 'move') { const dx = p.x - drag.sx, dy = p.y - drag.sy; if (!drag.moved && Math.abs(dx) + Math.abs(dy) < 3 / scale) return; if (!drag.moved) { snap(); drag.moved = true; } let nx = drag.ox + dx, ny = drag.oy + dy; if (!e.altKey) { nx = snapV(nx, [0, deck.w - o.w, (deck.w - o.w) / 2]); ny = snapV(ny, [0, deck.h - o.h, (deck.h - o.h) / 2]); } o.x = Math.round(nx * 10) / 10; o.y = Math.round(ny * 10) / 10; moveNode(o); }
        else if (drag.mode === 'resize') { if (!drag.moved) { snap(); drag.moved = true; } resizeTo(o, p, e); moveNode(o); }
        else if (drag.mode === 'rotate') { if (!drag.moved) { snap(); drag.moved = true; } const cx = o.x + o.w / 2, cy = o.y + o.h / 2; let ang = Math.atan2(p.y - cy, p.x - cx) * 180 / Math.PI + 90; if (e.shiftKey) ang = Math.round(ang / 15) * 15; ang = ((Math.round(ang) % 360) + 360) % 360; o.rot = ang === 0 ? 0 : ang; moveNode(o); }
      });
      const endDrag = e => { if (!drag) return; const d = drag; drag = null; if (d.moved) { thumbsSoon(); syncRibbon(); } };
      wrap.addEventListener('pointerup', endDrag); wrap.addEventListener('pointercancel', endDrag);
      wrap.addEventListener('dblclick', e => { const n = e.target.closest('.p-obj'); if (!n) return; const o = O(n.dataset.id); if (o && (o.type === 'text' || o.type === 'shape')) startEdit(o.id); else if (o && o.type === 'table') editTable(o); });
      function startResize(e, k) { e.stopPropagation(); e.preventDefault(); const o = SEL(); if (!o) return; const p = stagePt(e); drag = { mode: 'resize', id: o.id, h: k, sx: p.x, sy: p.y, o0: { x: o.x, y: o.y, w: o.w, h: o.h, rot: o.rot || 0 }, moved: false }; wrap.setPointerCapture(e.pointerId); }
      function startRotate(e) { e.stopPropagation(); e.preventDefault(); const o = SEL(); if (!o) return; drag = { mode: 'rotate', id: o.id, moved: false }; wrap.setPointerCapture(e.pointerId); }
      function resizeTo(o, p, e) {
        const d = drag, o0 = d.o0, th = (o0.rot || 0) * Math.PI / 180, c = Math.cos(th), s = Math.sin(th); const dx = p.x - d.sx, dy = p.y - d.sy; const lx = dx * c + dy * s, ly = -dx * s + dy * c;
        const sx = d.h.includes('e') ? 1 : d.h.includes('w') ? -1 : 0, sy = d.h.includes('s') ? 1 : d.h.includes('n') ? -1 : 0; let w = o0.w + sx * lx, h = o0.h + sy * ly;
        if ((o.type === 'image' || e.shiftKey) && sx && sy) { const r = o0.w / o0.h; if (Math.abs(w / o0.w) > Math.abs(h / o0.h)) h = w / r; else w = h * r; }
        w = Math.max(10, w); h = Math.max(10, h); const wd = w - o0.w, hd = h - o0.h; const shx = sx * wd / 2, shy = sy * hd / 2; const gx = shx * c - shy * s, gy = shx * s + shy * c;
        const cx = o0.x + o0.w / 2 + gx, cy = o0.y + o0.h / 2 + gy; o.w = Math.round(w * 10) / 10; o.h = Math.round(h * 10) / 10; o.x = Math.round((cx - w / 2) * 10) / 10; o.y = Math.round((cy - h / 2) * 10) / 10;
      }

      /* ---------- text editing */
      function beginEditDom(n) {
        const o = O(n.dataset.id); if (!o) return; let tx = n.querySelector('.tx'); if (!tx) { const bx = el('div', { class: 'p-t' }); tx = el('div', { class: 'tx' }); tx.style.cssText = textStyle(Object.assign({}, o, { ha: o.ha || 'center' })); bx.style.justifyContent = 'center'; bx.appendChild(tx); n.appendChild(bx); }
        tx.classList.remove('bul'); tx.textContent = o.text || ''; tx.setAttribute('contenteditable', 'plaintext-only'); if (tx.contentEditable !== 'plaintext-only') tx.setAttribute('contenteditable', 'true'); n.classList.add('editing'); tx.focus({ preventScroll: true });
        const r = document.createRange(); r.selectNodeContents(tx); if (!/^(Click to add|Presentation Title|Subtitle or author|Section Title|Short description|Text box|Lesson Title|Header )/.test(o.text || '')) r.collapse(false); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r);
        tx.addEventListener('keydown', ev => { if (ev.key === 'Escape') { ev.preventDefault(); finishEdit(); box.focus(); } ev.stopPropagation(); }); tx.addEventListener('blur', () => setTimeout(() => { if (editing === o.id && document.activeElement !== tx && !document.activeElement.closest('.ribbon, .dlg, .popup')) finishEdit(); }, 150));
        tx.addEventListener('input', () => { o.text = tx.innerText.replace(/​/g, '').replace(/\n$/, ''); });
        tx.addEventListener('paste', ev => ev.stopPropagation()); tx.addEventListener('copy', ev => ev.stopPropagation()); tx.addEventListener('cut', ev => ev.stopPropagation());
      }
      function startEdit(id) { if (editing === id) return; finishEdit(); const o = O(id); if (!o) return; snap(); editing = id; selId = id; renderStage(); syncRibbon(); }
      function finishEdit(silent) {
        if (!editing) return; const id = editing; editing = null; const o = O(id); const n = nodeOf(id); const tx = n && n.querySelector('.tx[contenteditable]'); if (o && tx) o.text = tx.innerText.replace(/​/g, '').replace(/\n$/, '');
        if (!silent) { renderStage(); thumbsSoon(); }
      }
      notes.addEventListener('input', () => { S().notes = notes.value; dirty = true; });
      notes.addEventListener('focus', () => finishEdit()); notes.addEventListener('change', () => { });

      /* ---------- slide & object operations */
      const cloneObj = o => Object.assign(JSON.parse(JSON.stringify(o)), { id: uid() });
      function addSlide(lay, at) { finishEdit(); snap(); const i = at === undefined ? cur + 1 : at; deck.slides.splice(i, 0, newSlide(deck, lay || 'content')); cur = i; selId = null; renderAll(); }
      function dupSlide() { finishEdit(); snap(); const s = JSON.parse(JSON.stringify(S())); s.id = uid(); s.objs.forEach(o => o.id = uid()); deck.slides.splice(cur + 1, 0, s); cur++; selId = null; renderAll(); }
      function delSlide() { if (deck.slides.length < 2) { const s = S(); finishEdit(); snap(); s.objs = []; s.bg = null; renderAll(); return toast('Aakhri slide delete nahi hoti, saaf kar di gayi'); } finishEdit(); snap(); deck.slides.splice(cur, 1); cur = Math.min(cur, deck.slides.length - 1); selId = null; renderAll(); }
      function moveSlide(f, t) { finishEdit(); snap(); const [s] = deck.slides.splice(f, 1); deck.slides.splice(t, 0, s); cur = t; renderAll(); }
      function setLayout(lay) { mutate(() => { const s = S(); const olds = s.objs.filter(o => o.type === 'text'); const rest = s.objs.filter(o => o.type !== 'text'); const nw = layoutObjs(deck, lay); nw.forEach((o, i) => { if (olds[i] && !/^Click to add/.test(olds[i].text || '')) o.text = olds[i].text; }); const extra = olds.slice(nw.length); s.objs = nw.concat(extra, rest); s.layout = lay; selId = null; }); }
      function addObj(o) { mutate(() => { o.id = o.id || uid(); S().objs.push(o); selId = o.id; }); }
      const T0 = () => T(deck);
      const mkText = (extra) => txt(deck, Object.assign({ x: deck.w / 2 - 160, y: deck.h / 2 - 30, w: 320, h: 60, text: 'Text box', sz: 24 }, extra));
      const mkShape = k => ({ id: uid(), type: 'shape', kind: k, x: deck.w / 2 - 100, y: deck.h / 2 - 70, w: k === 'line' ? 240 : 200, h: k === 'line' ? 4 : 140, fill: T0().accent, line: null, lw: 2, text: '', font: T0().font, sz: 20, color: '#ffffff', ha: 'center', va: 'middle' });
      async function insertImage(file) {
        if (!file) return; if (!/^image\//.test(file.type)) return toast('Sirf image files'); const url = await OA.readDataURL(file);
        const im = new Image(); im.onload = () => { let w = im.naturalWidth || 400, h = im.naturalHeight || 300; const m = Math.min(1, (deck.w * 0.7) / w, (deck.h * 0.7) / h); w *= m; h *= m; addObj({ type: 'image', src: url, x: (deck.w - w) / 2, y: (deck.h - h) / 2, w, h, alt: file.name }); }; im.onerror = () => toast('Image load nahi hui'); im.src = url;
      }
      async function tableDialog() { const v = await OA.form('Insert Table', [{ id: 'c', label: 'Columns', type: 'number', value: 4, min: 1, max: 12 }, { id: 'r', label: 'Rows', type: 'number', value: 4, min: 1, max: 20 }]); if (!v) return; const c = Math.max(1, v.c | 0), r = Math.max(1, v.r | 0); const rows = Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => i === 0 ? 'Header ' + (j + 1) : '')); addObj({ type: 'table', x: 60, y: 150, w: deck.w - 120, h: Math.min(deck.h - 200, r * 46), sz: 16, font: T0().font, rows, hc: T0().accent, band: OA.mixColor ? '#' + OA.mixColor(T0().accent.slice(1), 0.85) : '#e9edf7' }); }
      async function editTable(o) {
        const wrapx = el('div'); const ta = el('textarea', { class: 'inp', rows: 10, style: 'width:100%;font-family:monospace' }); ta.value = o.rows.map(r => r.join('\t')).join('\n'); wrapx.append(el('p', { class: 'bs-p', text: 'Har row naye line mein, columns Tab se alag. (Excel se paste bhi kar sakte hain.)' }), ta);
        OA.dialog({ title: 'Edit Table', width: 520, content: wrapx, buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { mutate(() => { const rows = ta.value.replace(/\r/g, '').split('\n').map(l => l.split('\t')); const nc = Math.max(...rows.map(r => r.length)); o.rows = rows.map(r => Array.from({ length: nc }, (_, i) => r[i] || '')); }); } }] });
      }
      function fmtObj(patch) {
        const o = SEL(); if (!o) return toast('Pehle koi object select karein'); const live = editing === o.id; if (!live) { finishEdit(); snap(); } Object.assign(o, patch); Object.keys(patch).forEach(k => { if (patch[k] === null) delete o[k]; });
        if (live) { const n = nodeOf(o.id); const tx = n && n.querySelector('.tx'); if (tx) tx.style.cssText = textStyle(Object.assign({}, o, { ha: o.ha || (o.type === 'shape' ? 'center' : 'left') })); } else renderAll(); syncRibbon(); thumbsSoon();
      }
      const delSel = () => { const o = SEL(); if (!o) return; mutate(() => { S().objs = S().objs.filter(x => x.id !== o.id); selId = null; }); };
      const copySel = (cut) => { const o = SEL(); if (!o) return; clip = JSON.stringify(o); if (cut) delSel(); };
      function pasteClip(txtv) { if (clip) { const o = JSON.parse(clip); o.id = uid(); o.x += 20; o.y += 20; addObj(o); clip = JSON.stringify(o); } else if (txtv) addObj(mkText({ text: txtv, h: Math.max(60, 30 * txtv.split('\n').length) })); }
      const dupSel = () => { const o = SEL(); if (!o) return; const c = cloneObj(o); c.x += 20; c.y += 20; addObj(c); };
      function arrange(kind) { const o = SEL(); if (!o) return; mutate(() => { const a = S().objs, i = a.indexOf(o); a.splice(i, 1); if (kind === 'front') a.push(o); else if (kind === 'back') a.unshift(o); else if (kind === 'fwd') a.splice(Math.min(a.length, i + 1), 0, o); else a.splice(Math.max(0, i - 1), 0, o); }); }
      function align(k) { const o = SEL(); if (!o) return; mutate(() => { if (k === 'l') o.x = 0; if (k === 'c') o.x = (deck.w - o.w) / 2; if (k === 'r') o.x = deck.w - o.w; if (k === 't') o.y = 0; if (k === 'm') o.y = (deck.h - o.h) / 2; if (k === 'b') o.y = deck.h - o.h; }); }
      function applyTheme(t) { mutate(() => { const old = deck.theme; const mapc = c => { if (!c) return c; const l = String(c).toLowerCase(); if (l === old.fg.toLowerCase()) return t.fg; if (l === old.accent.toLowerCase()) return t.accent; if (l === old.accent2.toLowerCase()) return t.accent2; return c; }; deck.slides.forEach(s => s.objs.forEach(o => { o.color = mapc(o.color); o.fill = mapc(o.fill); o.line = mapc(o.line); if (o.hc) o.hc = mapc(o.hc); if (o.font === old.font) o.font = t.font; })); deck.theme = Object.assign({}, t); }); }
      function setSize(w, h) { mutate(() => { const kx = w / deck.w, ky = h / deck.h; deck.slides.forEach(s => s.objs.forEach(o => { o.x *= kx; o.w *= kx; o.y *= ky; o.h *= ky; })); deck.w = w; deck.h = h; }); }

      /* ---------- context menus */
      function thumbMenu(e) { OA.menu({ left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY }, [{ label: 'New Slide', onClick: () => addSlide('content') }, { label: 'Duplicate Slide', onClick: dupSlide }, { label: 'Delete Slide', onClick: delSlide }, { sep: true }, { label: 'Move Up', disabled: cur === 0, onClick: () => moveSlide(cur, cur - 1) }, { label: 'Move Down', disabled: cur === deck.slides.length - 1, onClick: () => moveSlide(cur, cur + 1) }]); }
      wrap.addEventListener('contextmenu', e => { e.preventDefault(); const n = e.target.closest('.p-obj'); const at = { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY }; if (n) { select(n.dataset.id); OA.menu(at, [{ label: 'Cut', onClick: () => copySel(true) }, { label: 'Copy', onClick: () => copySel(false) }, { label: 'Duplicate', onClick: dupSel }, { sep: true }, { label: 'Bring to Front', onClick: () => arrange('front') }, { label: 'Send to Back', onClick: () => arrange('back') }, { sep: true }, { label: 'Edit Text', onClick: () => startEdit(n.dataset.id) }, { label: 'Delete', icon: 'trash', onClick: delSel }]); } else OA.menu(at, [{ label: 'Paste', disabled: !clip, onClick: () => pasteClip() }, { label: 'New Slide', onClick: () => addSlide('content') }, { label: 'Add Text Box', onClick: () => addObj(mkText()) }]); });

      /* ---------- keyboard & clipboard */
      root.addEventListener('keydown', e => {
        const t = e.target; if (t.matches && t.matches('input,textarea,select') || (t.isContentEditable)) { if (e.key === 'z' && (e.ctrlKey || e.metaKey) && t === notes) return; return; }
        const ctrl = e.ctrlKey || e.metaKey, k = e.key;
        if (ctrl && (k === 'z' || k === 'Z')) { e.preventDefault(); e.shiftKey ? redo() : undo(); return; } if (ctrl && (k === 'y' || k === 'Y')) { e.preventDefault(); redo(); return; }
        if (k === 'F5') { e.preventDefault(); present(e.shiftKey ? cur : 0); return; }
        if (!box.contains(t) && t !== box) return; const o = SEL();
        if (ctrl && (k === 'd' || k === 'D')) { e.preventDefault(); dupSel(); return; } if (ctrl && (k === 'b' || k === 'B') && o) { e.preventDefault(); fmtObj({ b: !o.b }); return; } if (ctrl && (k === 'i' || k === 'I') && o) { e.preventDefault(); fmtObj({ i: !o.i }); return; } if (ctrl && (k === 'u' || k === 'U') && o) { e.preventDefault(); fmtObj({ u: !o.u }); return; }
        if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (o) delSel(); else delSlide(); return; }
        if (k === 'Escape') { select(null); return; } if (k === 'Enter' && o && (o.type === 'text' || o.type === 'shape')) { e.preventDefault(); startEdit(o.id); return; }
        if (k === 'PageDown' || (!o && k === 'ArrowRight') || (!o && k === 'ArrowDown')) { e.preventDefault(); gotoSlide(cur + 1); return; } if (k === 'PageUp' || (!o && k === 'ArrowLeft') || (!o && k === 'ArrowUp')) { e.preventDefault(); gotoSlide(cur - 1); return; }
        if (o && /^Arrow/.test(k)) { e.preventDefault(); const st = e.shiftKey ? 10 : 1; snap(); o.x += k === 'ArrowRight' ? st : k === 'ArrowLeft' ? -st : 0; o.y += k === 'ArrowDown' ? st : k === 'ArrowUp' ? -st : 0; moveNode(o); thumbsSoon(); syncRibbon(); return; }
        if (o && (o.type === 'text' || o.type === 'shape') && k.length === 1 && !ctrl && !e.altKey) { e.preventDefault(); snap(); o.text = (o.text && !/^Click to add/.test(o.text) ? o.text : '') + k; editing = o.id; renderStage(); syncRibbon(); }
      });
      const isTyping = t => t && (t.matches && t.matches('input,textarea,select') || t.isContentEditable);
      root.addEventListener('copy', e => { if (isTyping(e.target)) return; const o = SEL(); if (!o) return; copySel(false); e.clipboardData.setData('text/plain', o.text || ''); e.preventDefault(); });
      root.addEventListener('cut', e => { if (isTyping(e.target)) return; const o = SEL(); if (!o) return; e.clipboardData.setData('text/plain', o.text || ''); copySel(true); e.preventDefault(); });
      root.addEventListener('paste', e => { if (isTyping(e.target)) return; e.preventDefault(); const f = Array.from(e.clipboardData.files || []).find(x => /^image\//.test(x.type)); if (f) { insertImage(f); return; } const t = e.clipboardData.getData('text/plain'); if (clip && (!t || JSON.parse(clip).text === t)) pasteClip(); else pasteClip(t); });
      box.addEventListener('dragover', e => e.preventDefault()); box.addEventListener('drop', e => { e.preventDefault(); const f = Array.from(e.dataTransfer.files || []).find(x => /^image\//.test(x.type)); if (f) insertImage(f); });
      // swipe between slides on touch when nothing is selected
      let sw = null; box.addEventListener('touchstart', e => { sw = e.touches.length === 1 && !e.target.closest('.p-obj') ? { x: e.touches[0].clientX, y: e.touches[0].clientY } : null; }, { passive: true });
      box.addEventListener('touchend', e => { if (!sw) return; const t = e.changedTouches[0]; const dx = t.clientX - sw.x, dy = t.clientY - sw.y; sw = null; if (Math.abs(dx) > 70 && Math.abs(dy) < 50 && zoomMul <= 1.01) gotoSlide(cur + (dx < 0 ? 1 : -1)); }, { passive: true });

      /* ---------- slide show */
      function present(from) {
        finishEdit(); let i = Math.max(0, from || 0); const ov = el('div', { class: 'p-show' }); let sl = null, steps = [], si = 0; const ctr = el('div', { class: 'ctr' }); const ex = el('button', { class: 'exit', type: 'button', text: 'Exit (Esc)', onclick: () => close() }); ov.append(ctr, ex); document.body.appendChild(ov);
        const fsEl = ov.requestFullscreen ? ov : null; try { fsEl && fsEl.requestFullscreen().catch(() => { }); } catch (e) { }
        const sc = () => Math.min(innerWidth / deck.w, innerHeight / deck.h);
        function show(idx, dirn) {
          i = idx; const s = deck.slides[i]; const k = sc(); if (sl) sl.remove(); sl = renderSlide(s, deck, { anim: true }); sl.style.left = (innerWidth - deck.w * k) / 2 + 'px'; sl.style.top = (innerHeight - deck.h * k) / 2 + 'px'; sl.style.transform = `scale(${k})`; sl.style.setProperty('--sc', k); sl.style.setProperty('--s0', dirn < 0 ? '-100%' : '100%');
          const tr = s.trans && s.trans.type !== 'none' ? s.trans : null; if (tr && dirn !== 0) { sl.style.setProperty('--d', (tr.dur || 1) + 's'); sl.classList.add('tr-' + (tr.type === 'random' ? ['fade', 'push', 'wipe', 'zoom', 'split'][Math.floor(Math.random() * 5)] : tr.type)); }
          steps = Array.from(sl.querySelectorAll('.p-obj[data-anim]')); si = dirn < 0 ? steps.length : 0; steps.forEach((n, j) => { if (j >= si) n.style.visibility = 'hidden'; }); ov.insertBefore(sl, ctr); ctr.textContent = (i + 1) + ' / ' + deck.slides.length;
        }
        function next() { if (si < steps.length) { const n = steps[si++]; n.style.visibility = 'visible'; n.classList.add('an-' + n.dataset.anim); return; } if (i < deck.slides.length - 1) show(i + 1, 1); else close(); }
        function prev() { if (i > 0) show(i - 1, -1); }
        function close() { document.removeEventListener('keydown', onKey, true); window.removeEventListener('resize', onResize); ov.remove(); if (document.fullscreenElement) document.exitFullscreen().catch(() => { }); gotoSlide(Math.min(i, deck.slides.length - 1)); }
        const onKey = e => { const k = e.key; if (k === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); } else if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter', 'n', 'N'].includes(k)) { e.preventDefault(); e.stopPropagation(); next(); } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P'].includes(k)) { e.preventDefault(); e.stopPropagation(); prev(); } else if (k === 'Home') { show(0, 0); } else if (k === 'End') { show(deck.slides.length - 1, 0); } else if (k === 'b' || k === 'B') { ov.classList.toggle('black'); sl.style.visibility = ov.classList.contains('black') ? 'hidden' : 'visible'; } };
        const onResize = () => { const k = sc(); sl.style.left = (innerWidth - deck.w * k) / 2 + 'px'; sl.style.top = (innerHeight - deck.h * k) / 2 + 'px'; sl.style.transform = `scale(${k})`; sl.style.setProperty('--sc', k); };
        document.addEventListener('keydown', onKey, true); window.addEventListener('resize', onResize); ov.addEventListener('click', e => { if (e.target === ex) return; if (e.clientX < innerWidth * 0.25) prev(); else next(); });
        let tx0 = null; ov.addEventListener('touchstart', e => { tx0 = e.touches[0].clientX; }, { passive: true }); ov.addEventListener('touchend', e => { if (tx0 === null) return; const dx = e.changedTouches[0].clientX - tx0; tx0 = null; if (Math.abs(dx) > 60) { e.preventDefault(); dx < 0 ? next() : prev(); } });
        ov.addEventListener('fullscreenchange', () => { }); show(i, 0); ov.tabIndex = 0; ov.focus();
      }

      /* ---------- print / export / import */
      function doPrint() {
        const boxp = el('div', { class: 'p-print' }); const k = 4 / 3; deck.slides.forEach(s => { const pg = el('div', { class: 'pg' }); const sl = renderSlide(s, deck, {}); sl.style.transform = `scale(${k})`; sl.style.transformOrigin = '0 0'; pg.style.width = deck.w * k + 'px'; pg.style.height = deck.h * k + 'px'; pg.appendChild(sl); boxp.appendChild(pg); });
        const stl = el('style', { text: `@page{size:${deck.w * k}px ${deck.h * k}px;margin:0}` }); document.head.appendChild(stl); document.body.appendChild(boxp); document.body.classList.add('printing');
        const done = () => { document.body.classList.remove('printing'); boxp.remove(); stl.remove(); window.removeEventListener('afterprint', done); }; window.addEventListener('afterprint', done); setTimeout(() => window.print(), 80);
      }
      async function doExport(kind) {
        finishEdit(); const n = baseName(); deck.title = n;
        try {
          if (kind === 'pptx') { OA.download(await IO.toPptx(deck), n + '.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'); toast('Exported ' + n + '.pptx'); }
          else if (kind === 'pdf') { toast('Print dialog mein "Save as PDF" chuniye'); doPrint(); }
          else if (kind === 'json') { OA.download(JSON.stringify(deck), n + '.json', 'application/json'); toast('Exported ' + n + '.json'); }
          else if (kind === 'png') { await exportPng(); }
          dirty = false;
        } catch (e) { console.error(e); OA.alert('Export fail hua: ' + OA.esc(e.message)); }
      }
      async function exportPng() {
        const s = S(); const host = el('div'); host.appendChild(renderSlide(s, deck, {})); const css = Array.from(document.styleSheets).map(ss => { try { return Array.from(ss.cssRules).map(r => r.cssText).join('') } catch (e) { return '' } }).join('');
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${deck.w * 2}" height="${deck.h * 2}"><foreignObject width="${deck.w}" height="${deck.h}" transform="scale(2)"><div xmlns="http://www.w3.org/1999/xhtml"><style>${css.replace(/</g, '&lt;')}</style>${host.innerHTML}</div></foreignObject></svg>`;
        const img = new Image(); await new Promise((res, rej) => { img.onload = res; img.onerror = () => rej(new Error('Image render nahi hui')); img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg); });
        const c = document.createElement('canvas'); c.width = deck.w * 2; c.height = deck.h * 2; c.getContext('2d').drawImage(img, 0, 0); c.toBlob(b => { OA.download(b, baseName() + '-slide' + (cur + 1) + '.png', 'image/png'); toast('Slide image export ho gayi'); });
      }
      async function openFile(file) {
        try { const ext = OA.ext(file.name); let d; if (ext === 'json') d = JSON.parse(await file.text()); else d = await IO.fromPptx(new Uint8Array(await file.arrayBuffer())); deck = d; cur = 0; selId = null; editing = null; undoS = []; redoS = []; sh.name.value = OA.baseName(file.name); renderAll(); dirty = false; toast('Opened ' + file.name); }
        catch (e) { console.error(e); OA.alert('Yeh file open nahi ho payi: ' + OA.esc(e.message)); }
      }
      function newDoc(tpl) { deck = newDeck(tpl); cur = 0; selId = null; editing = null; undoS = []; redoS = []; sh.name.value = 'Presentation1'; renderAll(); dirty = false; }

      /* ---------- ribbon */
      const B = (id, title, icon, onClick, o) => Object.assign({ id, icon, title, onClick, size: 'sm' }, o || {});
      const L = (id, label, icon, onClick, o) => Object.assign({ id, label, icon, title: label, onClick, size: 'lg' }, o || {});
      const G = (id, gl, title, onClick, style) => ({ id, glyph: gl, title, onClick, size: 'sm', gstyle: style || '' });
      const layoutMenu = fn => LAYOUTS.map(([k, l]) => ({ label: l, onClick: () => fn(k) }));
      const shapeMenu = (fn) => SHAPES.map(([k, l]) => ({ label: l, onClick: () => fn(k) }));
      const fontSel = { type: 'combo', id: 'font', options: FONTS, width: 120, value: 'Calibri', title: 'Font', inputmode: 'text', onChange: v => v && fmtObj({ font: v }) };
      const sizeSel = { type: 'combo', id: 'size', options: ['10', '12', '14', '16', '18', '20', '24', '28', '32', '36', '44', '54', '72', '96'], width: 46, value: '18', title: 'Font size', onChange: v => { const n = parseFloat(v); if (n > 0 && n < 500) fmtObj({ sz: n }); } };
      const num = (id, label, prop, min) => ({ type: 'number', id, label, value: 0, min, width: 66, onChange: v => { const o = SEL(); if (!o || isNaN(v)) return; mutate(() => { o[prop] = prop === 'rot' ? ((v % 360) + 360) % 360 : v; }); } });
      const transTypes = [['none', 'None'], ['fade', 'Fade'], ['push', 'Push'], ['wipe', 'Wipe'], ['zoom', 'Zoom'], ['split', 'Split'], ['cover', 'Cover'], ['dissolve', 'Dissolve'], ['random', 'Random']];
      const animTypes = [['none', 'None'], ['appear', 'Appear'], ['fade', 'Fade'], ['fly', 'Fly In'], ['zoom', 'Zoom']];
      const spec = {
        file: { onClick: () => bs.open('info') },
        tabs: [
          { id: 'home', label: 'Home', groups: [
            { label: 'Clipboard', items: [L('paste', 'Paste', 'paste', async () => { try { const t = await navigator.clipboard.readText(); pasteClip(t); } catch (e) { pasteClip(); } }), { type: 'row', items: [B('cut', 'Cut', 'cut', () => copySel(true), { label: 'Cut' }), B('copy', 'Copy', 'copy', () => copySel(false), { label: 'Copy' }), B('dup', 'Duplicate', 'layers', dupSel, { label: 'Duplicate' })] }] },
            { label: 'Slides', items: [L('ns', 'New Slide', 'plus', () => addSlide('content'), { menu: layoutMenu(k => addSlide(k)) }), { type: 'row', items: [B('lay', 'Layout', 'columns', null, { label: 'Layout', menu: layoutMenu(setLayout) }), B('ds', 'Delete Slide', 'trash', delSlide, { label: 'Delete' }), B('dsl', 'Duplicate Slide', 'copy', dupSlide, { label: 'Dup slide' })] }] },
            { label: 'Font', items: [{ type: 'row', items: [fontSel, sizeSel, B('grow', 'Increase Font Size', null, () => { const o = SEL(); o && fmtObj({ sz: Math.round((o.sz || 18) * 1.1 + 1) }); }, { glyph: 'A▲', gstyle: 'font-size:11px' }), B('shrink', 'Decrease Font Size', null, () => { const o = SEL(); o && fmtObj({ sz: Math.max(6, Math.round((o.sz || 18) / 1.1 - 1)) }); }, { glyph: 'A▼', gstyle: 'font-size:10px' })] },
              { type: 'row', items: [G('b', 'B', 'Bold (Ctrl+B)', () => { const o = SEL(); o && fmtObj({ b: !o.b }); }, 'font-weight:700'), G('i', 'I', 'Italic (Ctrl+I)', () => { const o = SEL(); o && fmtObj({ i: !o.i }); }, 'font-style:italic;font-family:serif'), G('u', 'U', 'Underline (Ctrl+U)', () => { const o = SEL(); o && fmtObj({ u: !o.u }); }, 'text-decoration:underline'), { type: 'color', id: 'fc', glyph: 'A', title: 'Font Color', default: 'FF0000', noColor: 'Automatic', onPick: c => fmtObj({ color: c ? '#' + c : T0().fg }) }] }] },
            { label: 'Paragraph', items: [{ type: 'row', items: [B('bul', 'Bullets', 'list', () => { const o = SEL(); o && fmtObj({ bullet: !o.bullet }); }), B('al', 'Align Left', 'alignLeft', () => fmtObj({ ha: 'left' })), B('ac', 'Center', 'alignCenter', () => fmtObj({ ha: 'center' })), B('ar', 'Align Right', 'alignRight', () => fmtObj({ ha: 'right' })), B('aj', 'Justify', 'justify', () => fmtObj({ ha: 'justify' }))] },
              { type: 'row', items: [B('vt', 'Top', null, () => fmtObj({ va: 'top' }), { glyph: '⤒' }), B('vm', 'Middle', null, () => fmtObj({ va: 'middle' }), { glyph: '↕' }), B('vb', 'Bottom', null, () => fmtObj({ va: 'bottom' }), { glyph: '⤓' })] }] },
            { label: 'Drawing', items: [L('shp', 'Shapes', 'triangle', null, { menu: shapeMenu(k => addObj(mkShape(k))) }), { type: 'color', id: 'fill', icon: 'bucket', title: 'Shape Fill', default: '4472C4', noColor: 'No Fill', onPick: c => fmtObj({ fill: c ? '#' + c : null }) }, { type: 'color', id: 'ol', glyph: '▭', title: 'Shape Outline', default: '000000', noColor: 'No Outline', onPick: c => fmtObj({ line: c ? '#' + c : null }) }, B('arr', 'Arrange', 'layers', null, { label: 'Arrange', menu: [{ label: 'Bring to Front', onClick: () => arrange('front') }, { label: 'Bring Forward', onClick: () => arrange('fwd') }, { label: 'Send Backward', onClick: () => arrange('bwd') }, { label: 'Send to Back', onClick: () => arrange('back') }, { sep: true }, { label: 'Align Left', onClick: () => align('l') }, { label: 'Align Center', onClick: () => align('c') }, { label: 'Align Right', onClick: () => align('r') }, { label: 'Align Top', onClick: () => align('t') }, { label: 'Align Middle', onClick: () => align('m') }, { label: 'Align Bottom', onClick: () => align('b') }, { sep: true }, { label: 'Rotate Right 90°', onClick: () => { const o = SEL(); o && mutate(() => { o.rot = ((o.rot || 0) + 90) % 360; }); } }, { label: 'Rotate Left 90°', onClick: () => { const o = SEL(); o && mutate(() => { o.rot = ((o.rot || 0) + 270) % 360; }); } }] })] },
            { label: 'Editing', items: [B('del', 'Delete object', 'trash', delSel, { label: 'Delete' })] }] },
          { id: 'insert', label: 'Insert', groups: [
            { label: 'Slides', items: [L('ns2', 'New Slide', 'plus', () => addSlide('content'), { menu: layoutMenu(k => addSlide(k)) })] },
            { label: 'Tables', items: [L('tbl', 'Table', 'table', tableDialog)] },
            { label: 'Images', items: [L('pic', 'Pictures', 'image', async () => { const f = await OA.pickFile('image/*'); insertImage(f); })] },
            { label: 'Illustrations', items: [L('shp2', 'Shapes', 'triangle', null, { menu: shapeMenu(k => addObj(mkShape(k))) })] },
            { label: 'Text', items: [L('tb', 'Text Box', 'textbox', () => addObj(mkText())), B('dt', 'Date', null, () => addObj(mkText({ text: new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }), w: 260, h: 44, sz: 18 })), { glyph: '📅', label: 'Date' }), B('sn', 'Slide Number', null, () => addObj(mkText({ text: String(cur + 1), x: deck.w - 100, y: deck.h - 50, w: 70, h: 34, sz: 14, ha: 'right' })), { glyph: '#', label: 'Number' })] }] },
          { id: 'design', label: 'Design', groups: [
            { label: 'Themes', cls: 'wide', items: [{ type: 'gallery', id: 'themes', items: THEMES.map(t => ({ label: t.name, key: t.name, sample: 'Aa', style: `background:${t.bg};color:${t.accent};font-weight:700;padding:2px 8px;border:1px solid #bbb;font-family:${t.font}`, onClick: () => applyTheme(t) })) }] },
            { label: 'Background', items: [{ type: 'color', id: 'bgc', icon: 'bucket', title: 'Slide Background (current slide)', noColor: 'Theme default', onPick: c => mutate(() => { S().bg = c ? '#' + c : null; }) }, B('bga', 'Apply background to all slides', null, () => mutate(() => { const b = S().bg; deck.slides.forEach(s => s.bg = b); }), { glyph: '▣', label: 'Apply to all' })] },
            { label: 'Customize', items: [L('sz', 'Slide Size', 'expand', null, { menu: [{ label: 'Widescreen (16:9)', onClick: () => setSize(960, 540) }, { label: 'Standard (4:3)', onClick: () => setSize(720, 540) }, { label: 'Widescreen (16:10)', onClick: () => setSize(960, 600) }, { label: 'A4 Landscape', onClick: () => setSize(842, 595) }] })] }] },
          { id: 'trans', label: 'Transitions', groups: [
            { label: 'Transition to This Slide', items: [{ type: 'select', id: 'tt', options: transTypes, width: 110, title: 'Transition', onChange: v => mutate(() => { S().trans = Object.assign({ dur: 1 }, S().trans, { type: v }); }) }, B('tpv', 'Preview', 'play', () => present(cur), { label: 'Preview' })] },
            { label: 'Timing', items: [{ type: 'number', id: 'td', label: 'Duration (s)', value: 1, min: 0.2, step: 0.2, width: 60, onChange: v => mutate(() => { S().trans = Object.assign({ type: 'none' }, S().trans, { dur: v || 1 }); }) }, B('tall', 'Apply To All', null, () => mutate(() => { const t = Object.assign({}, S().trans); deck.slides.forEach(s => s.trans = Object.assign({}, t)); }), { glyph: '⧉', label: 'Apply To All' })] }] },
          { id: 'anim', label: 'Animations', groups: [
            { label: 'Animation (selected object)', items: [{ type: 'select', id: 'an', options: animTypes, width: 110, title: 'Entrance effect', onChange: v => { const o = SEL(); if (!o) return toast('Pehle object select karein'); mutate(() => { if (v === 'none') delete o.anim; else o.anim = v; }); } }, B('apv', 'Preview', 'play', () => present(cur), { label: 'Preview' })] }] },
          { id: 'show', label: 'Slide Show', groups: [{ label: 'Start Slide Show', items: [L('fb', 'From Beginning', 'play', () => present(0)), L('fc2', 'From Current Slide', 'play', () => present(cur))] }] },
          { id: 'view', label: 'View', groups: [
            { label: 'Show', items: [{ type: 'check', id: 'nv', label: 'Notes', value: true, onChange: v => { showNotes = v; notes.hidden = !v; fit(); } }, { type: 'check', id: 'gv', label: 'Gridlines', value: false, onChange: v => { showGrid = v; renderStage(); } }] },
            { label: 'Zoom', items: [L('zf', 'Fit to Window', 'expand', () => { zoomMul = 1; fit(); }), B('zi', 'Zoom In', 'zoomIn', () => { zoomMul = Math.min(2.5, zoomMul + 0.1); fit(); }, { label: 'In' }), B('zo', 'Zoom Out', 'zoomOut', () => { zoomMul = Math.max(0.3, zoomMul - 0.1); fit(); }, { label: 'Out' })] }] },
          { id: 'shape', label: 'Shape Format', contextual: true, hidden: true, groups: [
            { label: 'Size & Position', items: [{ type: 'row', items: [num('px', 'X', 'x'), num('py', 'Y', 'y')] }, { type: 'row', items: [num('pw', 'W', 'w', 5), num('ph', 'H', 'h', 5)] }, num('pr', 'Rotate', 'rot')] },
            { label: 'Style', items: [{ type: 'color', id: 'fill2', icon: 'bucket', title: 'Fill', default: '4472C4', noColor: 'No Fill', onPick: c => fmtObj({ fill: c ? '#' + c : null }) }, { type: 'color', id: 'ol2', glyph: '▭', title: 'Outline', default: '000000', noColor: 'No Outline', onPick: c => fmtObj({ line: c ? '#' + c : null }) }, { type: 'number', id: 'lw', label: 'Weight', value: 1, min: 0.5, step: 0.5, width: 60, onChange: v => fmtObj({ lw: v || 1 }) }] },
            { label: 'Arrange', items: [B('f1', 'Bring to Front', 'layers', () => arrange('front'), { label: 'Front' }), B('f2', 'Send to Back', 'layers', () => arrange('back'), { label: 'Back' })] },
            { label: 'Accessibility', items: [B('alt', 'Alt text', 'info', async () => { const o = SEL(); if (!o) return; const t = await OA.prompt('Alt text:', o.alt || '', { title: 'Alt Text' }); if (t !== null) mutate(() => { o.alt = t; }); }, { label: 'Alt Text' })] },
            { label: 'Table', items: [B('et', 'Edit table', 'table', () => { const o = SEL(); o && o.type === 'table' && editTable(o); }, { label: 'Edit table' })] }] }
        ]
      };
      spec.right = el('div', { class: 'rtab-right' }); rb = OA.buildRibbon(sh.ribbon, spec); const ctl = rb.ctl;
      function syncRibbon() {
        if (!rb) return; const o = SEL(); rb.setTabVisible('shape', !!o); if (o && !rb._o) { rb.showTab('shape'); } rb._o = !!o; if (o) { ['b', 'i', 'u'].forEach(k => rb.setOn(k, o[k])); rb.setOn('bul', o.bullet); rb.setOn('al', o.ha === 'left'); rb.setOn('ac', o.ha === 'center'); rb.setOn('ar', o.ha === 'right'); rb.setOn('aj', o.ha === 'justify');
          if (document.activeElement !== ctl.font) ctl.font.value = o.font || 'Calibri'; if (document.activeElement !== ctl.size) ctl.size.value = String(Math.round(o.sz || 18)); [['px', 'x'], ['py', 'y'], ['pw', 'w'], ['ph', 'h'], ['pr', 'rot']].forEach(([id, p]) => { if (document.activeElement !== ctl[id]) ctl[id].value = Math.round((o[p] || 0) * 10) / 10; }); if (ctl.lw && document.activeElement !== ctl.lw) ctl.lw.value = o.lw || 1; if (ctl.an) ctl.an.value = o.anim || 'none'; } else if (ctl.an) ctl.an.value = 'none';
        if (ctl.tt) { ctl.tt.value = (S().trans && S().trans.type) || 'none'; ctl.td.value = (S().trans && S().trans.dur) || 1; }
      }
      /* ---------- backstage */
      bs = OA.backstage(root, { items: [
        { id: 'info', label: 'Info', icon: 'info', render(p) { p.appendChild(el('p', { class: 'bs-p', text: 'Aapki presentation sirf is browser tab mein hai. Bachane ke liye export karein.' })); let objs = 0; deck.slides.forEach(s => objs += s.objs.length); p.appendChild(el('div', { class: 'kv', html: `<span>Name</span><span>${OA.esc(baseName())}</span><span>Slides</span><span>${deck.slides.length}</span><span>Objects</span><span>${objs}</span><span>Slide size</span><span>${Math.round(deck.w)} × ${Math.round(deck.h)}</span>` })); p.appendChild(el('div', { style: { marginTop: '18px' } }, el('button', { class: 'btn primary', type: 'button', text: 'Export as .pptx', onclick: () => doExport('pptx') }))); } },
        { id: 'new', label: 'New', icon: 'plus', render(p, api) { const go = async k => { if (!dirty || await OA.confirm('Current presentation discard karke nayi banayein?', { ok: 'Continue' })) { newDoc(k); api.close(); } }; p.appendChild(OA.bigActions([{ label: 'Blank presentation', desc: 'Ek title slide se shuru', icon: 'file', onClick: () => go() }, { label: 'Pitch Deck', desc: 'Template (dark theme)', icon: 'file', onClick: () => go('pitch') }, { label: 'Status Update', desc: 'Template with table', icon: 'file', onClick: () => go('update') }, { label: 'Lesson', desc: 'Teaching template', icon: 'file', onClick: () => go('lesson') }])); } },
        { id: 'open', label: 'Open', icon: 'folder', render(p, api) { p.appendChild(el('p', { class: 'bs-p', text: '.pptx file apne device se kholein (text, shapes, images, tables aur notes import hote hain).' })); p.appendChild(el('button', { class: 'btn primary', type: 'button', text: 'Browse…', onclick: async () => { const f = await OA.pickFile('.pptx,.json'); if (f) { api.close(); openFile(f); } } })); } },
        { id: 'export', label: 'Export / Save As', icon: 'download', render(p) { p.appendChild(el('p', { class: 'bs-p', text: 'Format chuniye. Data is website par save nahi hota.' })); p.appendChild(OA.bigActions([{ label: 'PowerPoint Presentation (.pptx)', desc: 'PowerPoint, LibreOffice, Google Slides mein khulta hai', icon: 'download', onClick: () => doExport('pptx') }, { label: 'PDF', desc: 'Print dialog mein "Save as PDF" (har slide ek page)', icon: 'print', onClick: () => doExport('pdf') }, { label: 'Current slide as PNG', desc: 'Image', icon: 'image', onClick: () => doExport('png') }])); } },
        { id: 'print', label: 'Print', icon: 'print', action: doPrint }, { sep: true }, { id: 'home', label: 'Close (Office Home)', icon: 'home', action: () => OA.go('home') }
      ] });
      sh.title.querySelector('.tb-right').append(el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('undo', 16), title: 'Undo (Ctrl+Z)', onclick: undo }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('redo', 16), title: 'Redo (Ctrl+Y)', onclick: redo }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('play', 16), title: 'Slide Show (F5)', onclick: () => present(0) }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('download', 16), title: 'Export .pptx', onclick: () => doExport('pptx') }));
      if (window.ResizeObserver) new ResizeObserver(() => fit()).observe(box); window.addEventListener('resize', fit);
      renderAll();
      return { onShow() { setTimeout(() => { fit(); box.focus({ preventScroll: true }); }, 60); }, openFile, exportDefault: () => doExport('pptx'), loadTemplate: k => newDoc(k) };
    }
  });
})();
