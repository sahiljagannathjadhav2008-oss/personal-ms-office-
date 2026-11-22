/* Excel extras 1: shapes/text boxes/WordArt/equations/pictures, sparklines, chart elements, page setup, view, comments — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, XL = window.XL, el = OA.el, key = XL.key;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  XL.gridHooks = XL.gridHooks || []; XL.extras = XL.extras || [];

  /* ---------- floating objects (drawn through the chart layer) ---------- */
  const SHAPES = [['rect', 'Rectangle'], ['round', 'Rounded Rectangle'], ['ellipse', 'Oval'], ['triangle', 'Triangle'], ['diamond', 'Diamond'], ['pentagon', 'Pentagon'], ['hexagon', 'Hexagon'], ['star', 'Star'], ['rarrow', 'Right Arrow'], ['larrow', 'Left Arrow'], ['uarrow', 'Up Arrow'], ['darrow', 'Down Arrow'], ['line', 'Line'], ['callout', 'Callout'], ['heart', 'Heart'], ['cloud', 'Cloud'], ['plus', 'Cross'], ['chevron', 'Chevron']];
  const ICONS = ['★', '☀', '☁', '☂', '♥', '✔', '✘', '⚑', '✉', '☎', '⌂', '⚙', '✎', '⚠', '☺', '♪', '⚡', '✈', '☕', '🔒', '🔍', '💡', '📌', '📈', '📊', '🏆', '🎯', '🔔'];
  function shapePath(k, W, H) {
    const p = (x, y) => `${(x * W).toFixed(1)},${(y * H).toFixed(1)}`;
    const poly = pts => `<polygon points="${pts.map(a => p(a[0], a[1])).join(' ')}"`;
    switch (k) {
      case 'triangle': return poly([[.5, 0], [1, 1], [0, 1]]); case 'diamond': return poly([[.5, 0], [1, .5], [.5, 1], [0, .5]]);
      case 'pentagon': return poly([[.5, 0], [1, .38], [.81, 1], [.19, 1], [0, .38]]); case 'hexagon': return poly([[.25, 0], [.75, 0], [1, .5], [.75, 1], [.25, 1], [0, .5]]);
      case 'rarrow': return poly([[0, .3], [.6, .3], [.6, 0], [1, .5], [.6, 1], [.6, .7], [0, .7]]); case 'larrow': return poly([[1, .3], [.4, .3], [.4, 0], [0, .5], [.4, 1], [.4, .7], [1, .7]]);
      case 'uarrow': return poly([[.3, 1], [.3, .4], [0, .4], [.5, 0], [1, .4], [.7, .4], [.7, 1]]); case 'darrow': return poly([[.3, 0], [.3, .6], [0, .6], [.5, 1], [1, .6], [.7, .6], [.7, 0]]);
      case 'plus': return poly([[.35, 0], [.65, 0], [.65, .35], [1, .35], [1, .65], [.65, .65], [.65, 1], [.35, 1], [.35, .65], [0, .65], [0, .35], [.35, .35]]);
      case 'chevron': return poly([[0, 0], [.7, 0], [1, .5], [.7, 1], [0, 1], [.3, .5]]);
      case 'star': { const pts = []; for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, r = i % 2 ? .21 : .5; pts.push([.5 + r * Math.cos(a), .5 + r * Math.sin(a)]); } return poly(pts); }
      case 'ellipse': return `<ellipse cx="${W / 2}" cy="${H / 2}" rx="${W / 2}" ry="${H / 2}"`;
      case 'round': return `<rect x="0" y="0" width="${W}" height="${H}" rx="${Math.min(W, H) * .18}"`;
      case 'callout': return `<path d="M0,0 H${W} V${H * .75} H${W * .4} L${W * .2},${H} L${W * .25},${H * .75} H0 Z"`;
      case 'heart': return `<path d="M${W / 2},${H} C${-W * .2},${H * .4} ${W * .25},${-H * .1} ${W / 2},${H * .3} C${W * .75},${-H * .1} ${W * 1.2},${H * .4} ${W / 2},${H} Z"`;
      case 'cloud': return `<path d="M${W * .25},${H * .85} A${W * .2},${H * .25} 0 0 1 ${W * .2},${H * .4} A${W * .25},${H * .3} 0 0 1 ${W * .6},${H * .2} A${W * .22},${H * .25} 0 0 1 ${W * .85},${H * .5} A${W * .15},${H * .25} 0 0 1 ${W * .75},${H * .85} Z"`;
      default: return `<rect x="0" y="0" width="${W}" height="${H}"`;
    }
  }
  XL.objSVG = function (o, W, H) {
    const sw = o.sw == null ? 1.5 : o.sw, fill = o.fill === undefined ? '#4472c4' : o.fill, stroke = o.stroke || '#2f528f', fs = (o.fs || 14) * (W / Math.max(1, o.w || W));
    let inner = '';
    const txt = (t, color, extra) => { const lines = String(t || '').split('\n'); const lh = fs * 1.2; const y0 = H / 2 - (lines.length - 1) * lh / 2; const ax = o.align === 'left' ? 6 : o.align === 'right' ? W - 6 : W / 2, an = o.align === 'left' ? 'start' : o.align === 'right' ? 'end' : 'middle'; return lines.map((l, i) => `<text x="${ax}" y="${y0 + i * lh}" text-anchor="${an}" dominant-baseline="central" fill="${color}" font-size="${fs}" font-weight="${o.bold ? 700 : 400}" font-style="${o.italic ? 'italic' : 'normal'}" ${extra || ''}>${esc(l)}</text>`).join(''); };
    const defs = o.grad ? `<defs><linearGradient id="g${o.id}" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="${fill || '#4472c4'}"/><stop offset="1" stop-color="${o.grad}"/></linearGradient></defs>` : '';
    const fl = o.grad ? `url(#g${o.id})` : (fill || 'none');
    const sh = o.shadow ? `<filter id="s${o.id}" x="-20%" y="-20%" width="150%" height="150%"><feDropShadow dx="3" dy="3" stdDeviation="2.5" flood-opacity=".45"/></filter>` : '';
    const filt = o.shadow ? ` filter="url(#s${o.id})"` : '';
    if (o.kind === 'line') inner = `<line x1="2" y1="${H / 2}" x2="${W - 2}" y2="${H / 2}" stroke="${stroke}" stroke-width="${Math.max(1, sw * 1.5)}"${o.arrow ? ` marker-end="url(#ar${o.id})"` : ''}${o.dash ? ' stroke-dasharray="6 4"' : ''}/>` + (o.arrow ? `<defs><marker id="ar${o.id}" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="${stroke}"/></marker></defs>` : '');
    else if (o.kind === 'textbox') inner = `<rect x="0" y="0" width="${W}" height="${H}" fill="${o.fill === undefined ? 'none' : (fill || 'none')}" stroke="${o.stroke || '#999'}" stroke-width="${sw}"/>` + txt(o.text, o.color || '#000');
    else if (o.kind === 'wordart') { const c1 = o.fill || '#ed7d31', c2 = o.grad || '#ffc000'; inner = `<defs><linearGradient id="w${o.id}" x1="0" y1="0" x2="1" y2="0"><stop offset="0" stop-color="${c1}"/><stop offset="1" stop-color="${c2}"/></linearGradient></defs>` + txt(o.text, `url(#w${o.id})`, `stroke="${o.stroke || '#7f3f00'}" stroke-width="1" style="font-family:Impact,Arial Black,sans-serif;font-size:${fs * 1.8}px"`).replace(/font-size="[^"]+"/, ''); }
    else if (o.kind === 'icon') inner = `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" dominant-baseline="central" font-size="${Math.min(W, H) * .8}" fill="${o.fill || '#217346'}">${esc(o.text)}</text>`;
    else if (o.kind === 'equation') { let mm = ''; try { const t = OA.texToMath ? OA.texToMath(o.text || '') : ''; mm = t && t.ml ? t.ml : (typeof t === 'string' ? t : ''); } catch (e) { } inner = `<foreignObject x="0" y="0" width="${W}" height="${H}"><div xmlns="http://www.w3.org/1999/xhtml" style="display:flex;align-items:center;justify-content:center;width:100%;height:100%;font-size:${fs * 1.3}px;font-family:Cambria Math,serif;color:${o.color || '#000'}"><math xmlns="http://www.w3.org/1998/Math/MathML">${mm && /^<math/.test(mm) ? mm.replace(/^<math[^>]*>|<\/math>$/g, '') : esc(o.text)}</math></div></foreignObject>`; }
    else if (o.kind === 'picture') inner = o.src ? `<image href="${o.src}" x="0" y="0" width="${W}" height="${H}" preserveAspectRatio="none"/>` + (o.stroke ? `<rect width="${W}" height="${H}" fill="none" stroke="${o.stroke}" stroke-width="${sw}"/>` : '') : '';
    else inner = shapePath(o.kind, W, H) + ` fill="${fl}" stroke="${o.stroke === null ? 'none' : stroke}" stroke-width="${sw}"${o.dash ? ' stroke-dasharray="6 4"' : ''}${filt}/>` + txt(o.text, o.color || '#fff');
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="overflow:visible;font-family:Calibri,Arial,sans-serif;${o.rot ? `transform:rotate(${o.rot}deg);` : ''}${o.opacity != null ? `opacity:${o.opacity};` : ''}">${defs}${sh}${inner}</svg>`;
  };
  const EDITABLE = k => k !== 'picture' && k !== 'line';
  XL.objDialog = async function (ctx, o) {
    const f = [];
    if (EDITABLE(o.kind)) f.push({ id: 'text', label: o.kind === 'equation' ? 'Equation (LaTeX, e.g. \\frac{a}{b}+x^2)' : 'Text', type: 'textarea', value: o.text || '', rows: 3 }, { id: 'fs', label: 'Font size', type: 'number', value: o.fs || 14, min: 6, max: 200 }, { id: 'bold', label: 'Bold', type: 'checkbox', value: !!o.bold }, { id: 'color', label: 'Text colour', type: 'color', value: o.color || (o.kind === 'textbox' || o.kind === 'equation' ? '#000000' : '#ffffff') });
    if (o.kind !== 'picture' && o.kind !== 'textbox') f.push({ id: 'fill', label: 'Fill / main colour', type: 'color', value: o.fill || '#4472c4' });
    f.push({ id: 'stroke', label: 'Outline colour', type: 'color', value: o.stroke || '#2f528f' }, { id: 'sw', label: 'Outline width', type: 'number', value: o.sw == null ? 1.5 : o.sw, min: 0, max: 20, step: .5 }, { id: 'rot', label: 'Rotation (deg)', type: 'number', value: o.rot || 0, min: -360, max: 360 }, { id: 'shadow', label: 'Shadow', type: 'checkbox', value: !!o.shadow });
    const v = await OA.form('Format Object', f); if (!v) return; ctx.g.updateChart(o, v);
  };
  function addObj(ctx, props) {
    const g = ctx.g, sc = g.scroller || document.querySelector('.x-scroller'); const sx = sc ? sc.scrollLeft / g.zoom : 0, sy = sc ? sc.scrollTop / g.zoom : 0;
    const o = Object.assign({ id: XL.uid(), obj: true, type: 'col', ref: '', x: sx + 120, y: sy + 80, w: 160, h: 100 }, props); g.addChart(o); return o;
  }
  const readImg = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  async function pickPicture(ctx) {
    const f = await OA.pickFile('image/*'); if (!f) return; const src = await readImg(f);
    const im = new Image(); await new Promise(r => { im.onload = r; im.onerror = r; im.src = src; }); const k = Math.min(1, 360 / (im.width || 360)); addObj(ctx, { kind: 'picture', src, w: Math.round((im.width || 300) * k), h: Math.round((im.height || 200) * k) });
  }

  /* ---------- sparklines ---------- */
  XL.drawSpark = function (c, wb, sh, r, col, X, Y, w, h) {
    for (const sp of sh.sparks) {
      if (sp.r !== r || sp.c !== col) continue; const g = XL.parseRange(sp.ref); if (!g) continue; const vals = []; for (let rr = g.r1; rr <= g.r2; rr++) for (let cc = g.c1; cc <= g.c2; cc++) { const v = wb.value(sh.name, rr, cc); vals.push(typeof v === 'number' ? v : null); }
      const nums = vals.filter(v => v !== null); if (!nums.length) continue; let mn = Math.min(...nums), mx = Math.max(...nums); if (mn === mx) { mn -= 1; mx += 1; }
      const px = 4, py = 4, iw = w - 2 * px, ih = h - 2 * py, n = vals.length, col1 = sp.color || '#217346', hi = '#c00000';
      c.save(); c.beginPath(); c.rect(X, Y, w, h); c.clip(); c.strokeStyle = col1; c.fillStyle = col1; c.lineWidth = 1.6;
      const xs = i => X + px + (n === 1 ? iw / 2 : i * iw / (n - 1)), ys = v => Y + py + ih - (v - mn) / (mx - mn) * ih;
      if (sp.type === 'line') { c.beginPath(); let st = false; vals.forEach((v, i) => { if (v === null) { st = false; return; } if (!st) { c.moveTo(xs(i), ys(v)); st = true; } else c.lineTo(xs(i), ys(v)); }); c.stroke(); if (sp.markers !== false) { vals.forEach((v, i) => { if (v === null) return; const isH = sp.high && v === Math.max(...nums), isL = sp.low && v === mn; if (isH || isL || sp.markers === true) { c.fillStyle = isH ? hi : isL ? '#ed7d31' : col1; c.beginPath(); c.arc(xs(i), ys(v), 2.5, 0, 6.3); c.fill(); } }); } }
      else { const bw = Math.max(1, iw / n * .7), zero = (mn < 0 && mx > 0) ? ys(0) : Y + py + ih; vals.forEach((v, i) => { if (v === null) return; const x = X + px + (i + .5) * iw / n - bw / 2; if (sp.type === 'winloss') { const hh = ih / 2 - 1; c.fillStyle = v >= 0 ? col1 : hi; if (v > 0) c.fillRect(x, Y + h / 2 - hh, bw, hh); else if (v < 0) c.fillRect(x, Y + h / 2, bw, hh); else c.fillRect(x, Y + h / 2 - 0.5, bw, 1); } else { c.fillStyle = (sp.high && v === Math.max(...nums)) ? hi : (sp.low && v === mn) ? '#ed7d31' : (v < 0 && sp.negColor !== false ? hi : col1); const y = ys(v); c.fillRect(x, Math.min(y, zero), bw, Math.max(1, Math.abs(zero - y))); } }); }
      c.restore();
    }
  };
  async function addSpark(ctx, type) {
    const { wb, g } = ctx; const sh = ctx.sh(); const q = g.range(); const v = await OA.form('Create Sparklines', [{ id: 'data', label: 'Data range', value: XL.rangeLabel(q) }, { id: 'loc', label: 'Location range (e.g. G2:G6)', value: '' }, { id: 'high', label: 'High point', type: 'checkbox' }, { id: 'low', label: 'Low point', type: 'checkbox' }, { id: 'color', label: 'Colour', type: 'color', value: '#217346' }], { intro: 'Har location cell ke liye ek data row (ya column) use hoti hai.' }); if (!v) return;
    const d = XL.parseRange(v.data), l = XL.parseRange(v.loc); if (!d || !l) return OA.alert('Data range aur Location range sahi likhiye.');
    const byRow = (l.c2 - l.c1 + 1) === 1 && (l.r2 - l.r1 + 1) > 1 ? true : (l.r2 - l.r1 + 1) === 1 ? false : true; const out = [];
    if (l.r1 === l.r2 && l.c1 === l.c2) out.push({ r: l.r1, c: l.c1, ref: XL.rangeLabel(d) });
    else if (byRow) for (let i = 0; i <= l.r2 - l.r1 && d.r1 + i <= d.r2; i++) out.push({ r: l.r1 + i, c: l.c1, ref: XL.rangeLabel({ r1: d.r1 + i, c1: d.c1, r2: d.r1 + i, c2: d.c2 }) });
    else for (let i = 0; i <= l.c2 - l.c1 && d.c1 + i <= d.c2; i++) out.push({ r: l.r1, c: l.c1 + i, ref: XL.rangeLabel({ r1: d.r1, c1: d.c1 + i, r2: d.r2, c2: d.c2 + i }) });
    const keep = (sh.sparks || []).filter(s => !out.some(o => o.r === s.r && o.c === s.c));
    wb.tx(() => wb.setProp(sh, 'sparks', keep.concat(out.map(o => Object.assign(o, { type, color: v.color, high: v.high, low: v.low })))));
    ctx.after();
  }
  function sparkOpts(ctx, patch) { const { wb, g } = ctx, sh = ctx.sh(), p = g.active(); const list = sh.sparks || []; if (!list.some(s => s.r === p.r && s.c === p.c)) return OA.toast('Sparkline wali cell select kijiye'); wb.tx(() => wb.setProp(sh, 'sparks', list.map(s => (s.r === p.r && s.c === p.c) ? Object.assign({}, s, typeof patch === 'function' ? patch(s) : patch) : s))); ctx.after(); }

  /* ---------- Page setup ---------- */
  XL.dlg.pageSetup = async function (ctx, tab) {
    const p = ctx.page, s = ctx.sh(); p.hf = p.hf || {}; const H = ctx.H;
    const lbl = (t, node) => el('label', { class: 'form-row' }, [el('span', { text: t }), node]);
    const sel = (opts, v) => { const x = el('select', { class: 'inp' }); opts.forEach(o => { const a = Array.isArray(o) ? o : [o, o]; const e = el('option', { value: a[0], text: a[1] }); if (String(a[0]) === String(v)) e.selected = true; x.appendChild(e); }); return x; };
    const inp = (v, t) => el('input', { class: 'inp', type: t || 'text', value: v == null ? '' : v });
    const chk = (t, v) => { const i = el('input', { type: 'checkbox' }); i.checked = !!v; return { i, row: el('label', { class: 'form-row chk' }, [i, el('span', { text: t })]) }; };
    const orient = sel([['portrait', 'Portrait'], ['landscape', 'Landscape']], p.orient), paper = sel(['A4', 'Letter', 'Legal', 'A3', 'A5'], p.paper || 'A4'), fit = chk('Fit all columns on one page', p.fit), scale = inp(p.scale || 100, 'number');
    const mg = sel(['Normal', 'Wide', 'Narrow'], p.margin || 'Normal'), ctr = chk('Center horizontally on page', p.center);
    const hdr = inp(p.hf.h || ''), ftr = inp(p.hf.f || ''), codes = el('div', { class: 'form-intro', html: 'Codes: <b>&amp;P</b> page no, <b>&amp;N</b> total pages, <b>&amp;D</b> date, <b>&amp;T</b> time, <b>&amp;F</b> file name, <b>&amp;A</b> sheet name' });
    const hdrBtns = (target) => el('div', { style: 'display:flex;gap:4px;flex-wrap:wrap;margin:2px 0 6px' }, [['Page', '&P'], ['Pages', '&N'], ['Date', '&D'], ['Time', '&T'], ['File', '&F'], ['Sheet', '&A']].map(([t, c]) => el('button', { type: 'button', class: 'btn', text: t, onclick: () => { target.value += c; } })));
    const area = inp(s.printArea ? XL.rangeLabel(s.printArea) : ''), trow = inp(s.titleRows && s.titleRows.r1 >= 0 ? (s.titleRows.r1 + 1) + ':' + (s.titleRows.r2 + 1) : ''), grid = chk('Gridlines', p.grid), heads = chk('Row and column headings', p.headings), title = inp(p.title);
    const panes = { 'Page': [lbl('Orientation', orient), lbl('Paper size', paper), lbl('Scale %', scale), fit.row], 'Margins': [lbl('Margins', mg), ctr.row], 'Header/Footer': [lbl('Header', hdr), hdrBtns(hdr), lbl('Footer', ftr), hdrBtns(ftr), codes], 'Sheet': [lbl('Print area', area), lbl('Rows to repeat at top (e.g. 1:1)', trow), lbl('Print title (top of every page)', title), grid.row, heads.row] };
    const wrap = el('div'); const bar = el('div', { style: 'display:flex;gap:2px;border-bottom:1px solid #ccc;margin-bottom:10px' }); const body = el('div', { class: 'form-grid' }); wrap.append(bar, body);
    let cur = tab || 'Page'; const show = n => { cur = n; body.innerHTML = ''; panes[n].forEach(x => body.appendChild(x)); Array.from(bar.children).forEach(b => { b.style.fontWeight = b.textContent === n ? '700' : '400'; b.style.background = b.textContent === n ? '#e2efda' : ''; }); };
    Object.keys(panes).forEach(n => bar.appendChild(el('button', { type: 'button', class: 'btn', text: n, onclick: () => show(n) }))); show(cur);
    return new Promise(res => OA.dialog({ title: 'Page Setup', content: wrap, width: 440, buttons: [{ label: 'Cancel', value: false }, { label: 'Print…', onClick: () => { apply(); setTimeout(() => H.doPrint(), 50); } }, { label: 'OK', primary: true, onClick: () => { apply(); res(true); } }], onClose: () => res(false) }));
    function apply() {
      p.orient = orient.value; p.paper = paper.value; p.fit = fit.i.checked; p.scale = Math.max(10, Math.min(400, +scale.value || 100)); p.margin = mg.value; p.center = ctr.i.checked; p.hf = { h: hdr.value, f: ftr.value }; p.grid = grid.i.checked; p.headings = heads.i.checked; p.title = title.value;
      const a = XL.parseRange(area.value.trim()); s.printArea = a || null; const m = /^(\d+):(\d+)$/.exec(trow.value.trim()); s.titleRows = m ? { r1: +m[1] - 1, r2: +m[2] - 1 } : null;
    }
  };

  /* ---------- extras: ribbon wiring ---------- */
  XL.extras.push(function (ctx, spec, H) {
    const { B, L, G } = H; const tab = id => spec.tabs.find(t => t.id === id); const add = (id, groups) => { const t = tab(id); if (t) t.groups.push(...groups); };
    const { wb } = ctx; const toast = OA.toast;
    const shapeMenu = SHAPES.map(([k, l]) => ({ label: l, onClick: () => addObj(ctx, k === 'line' ? { kind: k, w: 180, h: 30, stroke: '#2f528f', sw: 2 } : { kind: k, text: '', w: 150, h: k === 'rarrow' || k === 'larrow' ? 80 : 110 }) }));
    const sparkMenu = [['line', 'Line'], ['col', 'Column'], ['winloss', 'Win/Loss']].map(([k, l]) => ({ label: l, onClick: () => addSpark(ctx, k) }));
    add('insert', [
      { label: 'Illustrations', items: [L('shp', 'Shapes', 'edit', null, { menu: shapeMenu }), L('pic', 'Pictures', 'image', () => pickPicture(ctx)), L('ico', 'Icons', 'star', null, { menu: ICONS.map(i => ({ label: i, onClick: () => addObj(ctx, { kind: 'icon', text: i, w: 70, h: 70, fill: '#217346' }) })) })] },
      { label: 'Sparklines', items: [L('spk', 'Sparklines', 'chart', null, { menu: sparkMenu })] },
      { label: 'Text', items: [L('tbx', 'Text Box', 'edit', () => addObj(ctx, { kind: 'textbox', text: 'Text', w: 160, h: 60, color: '#000', fs: 14 })), L('wart', 'WordArt', 'palette', async () => { const t = await OA.prompt('WordArt text:', 'WordArt', { title: 'WordArt' }); if (t) addObj(ctx, { kind: 'wordart', text: t, w: 280, h: 80, fs: 36 }); }), L('hdf', 'Header & Footer', 'list', () => XL.dlg.pageSetup(ctx, 'Header/Footer'))] },
      { label: 'Symbols', items: [L('eqn', 'Equation', 'quote', async () => { const t = await OA.prompt('Equation (LaTeX subset):', '\\frac{a+b}{2}=x^2', { title: 'Insert Equation' }); if (t) addObj(ctx, { kind: 'equation', text: t, w: 220, h: 70, fs: 18 }); }), L('sym', 'Symbol', 'star', null, { menu: ['©', '®', '™', '°', '±', '×', '÷', '≠', '≤', '≥', '∞', '√', 'π', 'Σ', 'Δ', 'µ', '€', '£', '¥', '₹', '←', '→', '↑', '↓', '✓'].map(s => ({ label: s, onClick: () => { const p = ctx.g.active(); ctx.g.startEdit(p.r, p.c, s, 'edit'); } })) })] }
    ]);
    add('layout', [{ label: 'Page Setup', items: [
      L('psz', 'Size', 'file', null, { menu: ['A4', 'Letter', 'Legal', 'A3', 'A5'].map(n => ({ label: n, onClick: () => { ctx.page.paper = n; toast('Paper: ' + n); } })) }),
      L('pmg', 'Margins', 'margin', null, { menu: ['Normal', 'Wide', 'Narrow'].map(n => ({ label: n, onClick: () => { ctx.page.margin = n; toast('Margins: ' + n); } })) }),
      L('par', 'Print Area', 'print', null, { menu: [{ label: 'Set Print Area', onClick: () => { const s = ctx.sh(); wb.tx(() => wb.setProp(s, 'printArea', ctx.g.range())); toast('Print area set'); } }, { label: 'Clear Print Area', onClick: () => { wb.tx(() => wb.setProp(ctx.sh(), 'printArea', null)); toast('Print area cleared'); } }] }),
      L('ptt', 'Print Titles', 'list', () => XL.dlg.pageSetup(ctx, 'Sheet')),
      L('psc', 'Scale to Fit', 'zoomIn', null, { menu: [{ label: 'Fit all columns on one page', onClick: () => { ctx.page.fit = true; toast('Fit to 1 page wide'); } }, { label: 'Actual size', onClick: () => { ctx.page.fit = false; ctx.page.scale = 100; toast('Actual size'); } }, { label: '75%', onClick: () => { ctx.page.fit = false; ctx.page.scale = 75; } }, { label: '50%', onClick: () => { ctx.page.fit = false; ctx.page.scale = 50; } }] })
    ] }, { label: 'Themes', items: [L('thm', 'Themes', 'palette', null, { menu: [['Office', '#217346'], ['Ocean', '#1f6fb4'], ['Sunset', '#d9531e'], ['Violet', '#7030a0'], ['Slate', '#44546a']].map(([n, c]) => ({ label: n, onClick: () => { document.documentElement.style.setProperty('--x-accent', c); const q = ctx.g.range(); wb.tx(() => wb.applyStyle(ctx.sh(), q, { bg: c, color: '#ffffff' })); ctx.after(); toast('Theme colour ' + n + ' selection par lagaya'); } })) })] }]);
    add('view', [{ label: 'Workbook Views', items: [
      { type: 'check', id: 'pbv', label: 'Page Break Preview', value: false, onChange: v => { ctx.g.pbv = v; ctx.g.draw(); ctx.g.overlayUpdate && ctx.g.overlayUpdate(); } },
      L('znz', 'Zoom to Selection', 'zoomIn', () => { const g = ctx.g, q = g.range(), r = g.cellRect(q.r1, q.c1), r2 = g.cellRect(q.r2, q.c2); const w = (r2.x + r2.w - r.x) / g.zoom, h = (r2.y + r2.h - r.y) / g.zoom; const sc = document.querySelector('.x-scroller'); g.setZoom(Math.max(.4, Math.min(4, Math.min(sc.clientWidth / (w + 60), sc.clientHeight / (h + 40))))); g.ensureVisible(q.r1, q.c1); }),
      L('csv2', 'Custom Views', 'eye', async () => { const s = ctx.sh(); ctx.views = ctx.views || {}; const names = Object.keys(ctx.views); const n = await OA.prompt((names.length ? 'Saved: ' + names.join(', ') + '\n' : '') + 'View name (existing = restore, new = save):', '', { title: 'Custom Views' }); if (!n) return; if (ctx.views[n]) { const v = ctx.views[n]; ctx.activate(v.sheet); const s2 = ctx.sh(); wb.tx(() => { wb.setProp(s2, 'hiddenR', new Set(v.hr)); wb.setProp(s2, 'hiddenC', new Set(v.hc)); wb.setProp(s2, 'showGrid', v.grid); }); ctx.g.setZoom(v.zoom); ctx.g.setActive(v.r, v.c); ctx.after(); toast('View restored: ' + n); } else { const a = ctx.g.active(); ctx.views[n] = { sheet: s.name, hr: Array.from(s.hiddenR), hc: Array.from(s.hiddenC), grid: s.showGrid, zoom: ctx.g.zoom, r: a.r, c: a.c }; toast('View saved: ' + n); } }),
      L('spl', 'Split', 'columns', () => { ctx.g.split = !ctx.g.split; const sc = document.querySelector('.x-scroller'); if (!ctx.g.split) { document.querySelectorAll('.x-splitline').forEach(e => e.remove()); return toast('Split removed'); } const a = ctx.g.active(); const r = ctx.g.cellRect(a.r, a.c); const host = H.host; const mk = (css) => { const d = el('div', { class: 'x-splitline' }); d.style.cssText = 'position:absolute;background:#217346;z-index:6;pointer-events:none;' + css; host.appendChild(d); }; mk(`left:${r.x}px;top:0;bottom:0;width:3px`); mk(`top:${r.y}px;left:0;right:0;height:3px`); toast('Split at active cell (freeze panes se scroll lock karein)'); })
    ] }]);
    const sparkTab = { id: 'spark', label: 'Sparkline', contextual: true, hidden: true, groups: [] };
    add('review', [{ label: 'Comments', items: [
      L('cmn', 'New Comment', 'comment', () => commentDlg(ctx)),
      L('cdel', 'Delete', 'trash', () => { const s = ctx.sh(), p = ctx.g.active(), k = key(p.r, p.c); if (!s.notes.has(k)) return toast('Is cell me comment nahi hai'); wb.tx(() => { const old = s.notes.get(k); s.notes.delete(k); wb._op(() => s.notes.set(k, old), () => s.notes.delete(k)); }); ctx.after(); }),
      L('cprev', 'Previous', 'chevronLeft', () => navNote(ctx, -1)), L('cnext', 'Next', 'chevronRight', () => navNote(ctx, 1)),
      L('cshow', 'Show All Notes', 'eye', () => { const s = ctx.sh(); if (!s.notes.size) return toast('Koi comment nahi'); const rows = []; s.notes.forEach((v, k) => { const u = XL.unkey(k); rows.push('<b>' + XL.addr(u.r, u.c) + '</b>: ' + esc(v).replace(/\n/g, '<br>')); }); OA.dialog({ title: 'All Comments (' + s.notes.size + ')', content: el('div', { class: 'dlg-msg', html: rows.join('<hr>') }), buttons: [{ label: 'Close', primary: true }] }); })
    ] }]);
    spec.tabs.push({ id: 'shape', label: 'Shape Format', contextual: true, hidden: true, groups: [
      { label: 'Edit', items: [L('oed', 'Format / Edit Text', 'edit', () => { const o = ctx.g.chart(); o && XL.objDialog(ctx, o); })] },
      { label: 'Style', items: [L('ofl', 'Fill', 'palette', async () => { const o = ctx.g.chart(); if (!o) return; const v = await OA.form('Fill', [{ id: 'fill', label: 'Colour', type: 'color', value: o.fill || '#4472c4' }, { id: 'grad', label: 'Gradient end colour (blank = none)', value: o.grad || '' }]); if (v) ctx.g.updateChart(o, { fill: v.fill, grad: v.grad || null }); }),
        L('oop', 'Transparency', 'eye', async () => { const o = ctx.g.chart(); if (!o) return; const v = await OA.prompt('Opacity % (0-100):', String(Math.round((o.opacity == null ? 1 : o.opacity) * 100))); if (v !== null) ctx.g.updateChart(o, { opacity: Math.max(0, Math.min(100, +v || 100)) / 100 }); }),
        L('osh', 'Shadow', 'layers', () => { const o = ctx.g.chart(); o && ctx.g.updateChart(o, { shadow: !o.shadow }); }),
        L('oar', 'Arrow / Dash', 'play', () => { const o = ctx.g.chart(); o && ctx.g.updateChart(o, o.kind === 'line' ? { arrow: !o.arrow } : { dash: !o.dash }); })] },
      { label: 'Arrange', items: [L('ofr', 'Bring to Front', 'layers', () => reorder(ctx, 1)), L('obk', 'Send to Back', 'layers', () => reorder(ctx, -1)), L('orl', 'Rotate 90°', 'refresh', () => { const o = ctx.g.chart(); o && ctx.g.updateChart(o, { rot: ((o.rot || 0) + 90) % 360 }); }), L('odu', 'Duplicate', 'copy', () => { const o = ctx.g.chart(); if (o) addObj(ctx, Object.assign({}, o, { id: XL.uid(), x: o.x + 20, y: o.y + 20 })); })] },
      { label: 'Size', items: [L('osz', 'Size…', 'zoomIn', async () => { const o = ctx.g.chart(); if (!o) return; const v = await OA.form('Size', [{ id: 'w', label: 'Width (px)', type: 'number', value: o.w }, { id: 'h', label: 'Height (px)', type: 'number', value: o.h }]); if (v) ctx.g.updateChart(o, { w: Math.max(10, v.w), h: Math.max(10, v.h) }); }), L('odl', 'Delete', 'trash', () => { const o = ctx.g.chart(); o && ctx.g.deleteChart(o.id); })] }
    ] });
    // Sparkline tab content + chart elements
    const sp = (label, patch) => B('sp_' + label, label, null, () => sparkOpts(ctx, patch), { label });
    sparkTab.groups = [{ label: 'Type', items: [L('spl1', 'Line', 'chart', () => sparkOpts(ctx, { type: 'line' })), L('spc1', 'Column', 'chart', () => sparkOpts(ctx, { type: 'col' })), L('spw1', 'Win/Loss', 'chart', () => sparkOpts(ctx, { type: 'winloss' }))] },
      { label: 'Show', items: [L('sph', 'High Point', 'star', () => sparkOpts(ctx, s => ({ high: !s.high }))), L('spo', 'Low Point', 'star', () => sparkOpts(ctx, s => ({ low: !s.low }))), L('spm', 'Markers', 'star', () => sparkOpts(ctx, s => ({ markers: s.markers === true ? false : true })))] },
      { label: 'Style', items: [L('spk2', 'Colour', 'palette', async () => { const v = await OA.form('Sparkline colour', [{ id: 'color', label: 'Colour', type: 'color', value: '#217346' }]); if (v) sparkOpts(ctx, { color: v.color }); }), L('spx', 'Clear', 'trash', () => { const s = ctx.sh(), p = ctx.g.active(); wb.tx(() => wb.setProp(s, 'sparks', (s.sparks || []).filter(x => !(x.r === p.r && x.c === p.c)))); ctx.after(); })] }];
    spec.tabs.push(sparkTab);
    const ct = tab('chart'); if (ct) {
      const up = p => { const c = ctx.g.chart(); c && ctx.g.updateChart(c, typeof p === 'function' ? p(c) : p); };
      const ask = (t, k) => async () => { const c = ctx.g.chart(); if (!c) return; const v = await OA.prompt(t, c[k] || ''); if (v !== null) ctx.g.updateChart(c, { [k]: v }); };
      const tl = t => ({ label: t[1], onClick: () => up({ trendline: t[0] }) });
      const elemMenu = [{ header: 'Titles' }, { label: 'Chart Title…', onClick: ask('Chart title:', 'title') }, { label: 'Horizontal Axis Title…', onClick: ask('Horizontal axis title:', 'xTitle') }, { label: 'Vertical Axis Title…', onClick: ask('Vertical axis title:', 'yTitle') }, { label: 'Secondary Axis Title…', onClick: ask('Secondary axis title:', 'y2Title') },
        { header: 'Elements' }, { label: 'Data Labels (toggle)', onClick: () => up(c => ({ dataLabels: !c.dataLabels })) }, { label: 'Gridlines (toggle)', onClick: () => up(c => ({ gridlines: c.gridlines === false ? true : false })) }, { label: 'Legend (toggle)', onClick: () => up(c => ({ legend: c.legend === false })) },
        { header: 'Trendline' }, tl(['linear', 'Linear']), tl(['exp', 'Exponential']), tl(['log', 'Logarithmic']), tl(['poly', 'Polynomial (2)']), tl(['power', 'Power']), tl(['avg', 'Moving Average']), { label: 'None', onClick: () => up({ trendline: null }) }, { label: 'Show Trendline Equation / R²', onClick: () => up(c => ({ trendEq: !c.trendEq })) },
        { header: 'Error Bars' }, { label: 'Fixed Value…', onClick: async () => { const v = await OA.prompt('Fixed error value:', '1'); if (v !== null) up({ errBars: 'fixed', errVal: +v || 1 }); } }, { label: 'Percentage 5%', onClick: () => up({ errBars: 'pct', errVal: 5 }) }, { label: 'Standard Deviation', onClick: () => up({ errBars: 'sd' }) }, { label: 'None', onClick: () => up({ errBars: null }) }];
      ct.groups.splice(1, 0, { label: 'Chart Elements', items: [L('ace', 'Add Chart Element', 'plus', null, { menu: elemMenu }),
        L('ql', 'Quick Layout', 'layers', null, { menu: [{ label: 'Layout 1: title + legend bottom', onClick: () => up({ legend: true, dataLabels: false, gridlines: true }) }, { label: 'Layout 2: data labels, no legend', onClick: () => up({ legend: false, dataLabels: true }) }, { label: 'Layout 3: minimal', onClick: () => up({ legend: false, dataLabels: false, gridlines: false }) }] })] });
      ct.groups.push({ label: 'Data', items: [
        L('sw', 'Switch Row/Column', 'refresh', () => up(c => ({ rowsAsSeries: !c.rowsAsSeries }))),
        L('sd', 'Select Data', 'table', async () => { const c = ctx.g.chart(); if (!c) return; const v = await OA.prompt('Chart data range:', c.ref); if (v) ctx.g.updateChart(c, { ref: v }); }),
        L('s2', 'Secondary Axis', 'chart', async () => { const c = ctx.g.chart(); if (!c) return; const v = await OA.prompt('Series numbers on secondary axis (e.g. 2,3; blank = none):', (c.y2 || []).map(i => i + 1).join(',')); if (v !== null) ctx.g.updateChart(c, { y2: v.split(',').map(x => parseInt(x) - 1).filter(x => x >= 0) }); }),
        L('sty', 'Series Chart Type', 'chart', async () => { const c = ctx.g.chart(); if (!c) return; const v = await OA.form('Combo chart', [{ id: 'n', label: 'Series number', type: 'number', value: 2, min: 1 }, { id: 't', label: 'Type', type: 'select', options: [['', 'Same as chart'], ['line', 'Line'], ['col', 'Column'], ['area', 'Area']], value: 'line' }]); if (!v) return; const st = Object.assign({}, c.seriesTypes || {}); if (v.t) st[v.n - 1] = v.t; else delete st[v.n - 1]; ctx.g.updateChart(c, { seriesTypes: st }); }),
        L('smv', 'Move Chart', 'copy', async () => { const c = ctx.g.chart(); if (!c) return; const names = wb.sheets.map(s => s.name); const v = await OA.form('Move Chart', [{ id: 's', label: 'Move to sheet', type: 'select', options: names, value: ctx.sh().name }]); if (!v || v.s === ctx.sh().name) return; const to = wb.sheet(v.s); const clone = Object.assign({}, c); wb.tx(() => { wb.setProp(ctx.sh(), 'charts', ctx.sh().charts.filter(x => x.id !== c.id)); wb.setProp(to, 'charts', to.charts.concat([clone])); }); ctx.g.selChart = null; ctx.after(); ctx.activate(v.s); })] });
    }
    // right-click: sparklines tab show on selection change
    ctx._sparkSync = () => { if (!H.rb) return; const s = ctx.sh(), p = ctx.g.active(); const on = !!(s.sparks && s.sparks.some(x => x.r === p.r && x.c === p.c)); H.rb.setTabVisible('spark', on); };
  });
  function reorder(ctx, dir) { const g = ctx.g, o = g.chart(), s = ctx.sh(); if (!o) return; const rest = s.charts.filter(x => x.id !== o.id); ctx.wb.tx(() => ctx.wb.setProp(s, 'charts', dir > 0 ? rest.concat([o]) : [o].concat(rest))); g.layoutCharts(); ctx.after(); }
  function navNote(ctx, d) { const s = ctx.sh(), keys = Array.from(s.notes.keys()).map(k => XL.unkey(k)).sort((a, b) => a.r - b.r || a.c - b.c); if (!keys.length) return OA.toast('Koi comment nahi'); const p = ctx.g.active(); let i = d > 0 ? keys.findIndex(k => k.r > p.r || (k.r === p.r && k.c > p.c)) : (() => { let j = -1; keys.forEach((k, x) => { if (k.r < p.r || (k.r === p.r && k.c < p.c)) j = x; }); return j; })(); if (i < 0) i = d > 0 ? 0 : keys.length - 1; ctx.g.setActive(keys[i].r, keys[i].c); ctx.g.ensureVisible(keys[i].r, keys[i].c); OA.toast(XL.addr(keys[i].r, keys[i].c) + ': ' + String(s.notes.get(key(keys[i].r, keys[i].c))).split('\n')[0].slice(0, 80)); }
  async function commentDlg(ctx) {
    const { wb, g } = ctx, s = ctx.sh(), a = g.active(), k = key(a.r, a.c); const cur = s.notes.get(k) || '';
    const v = await OA.form('Comment — ' + XL.addr(a.r, a.c), [{ id: 'author', label: 'Author', value: OA.author || 'Sahil Jadhav' }, { id: 'text', label: cur ? 'Add reply / edit thread' : 'Comment', type: 'textarea', rows: 4, value: '' }], { intro: cur ? '<pre style="white-space:pre-wrap;margin:0;font:inherit">' + esc(cur) + '</pre>' : '' });
    if (!v || !v.text.trim()) return; const entry = '[' + v.author + '] ' + v.text.trim(); const nv = cur ? cur + '\n↳ ' + entry : entry; const old = s.notes.get(k);
    wb.tx(() => { s.notes.set(k, nv); wb._op(() => { old === undefined ? s.notes.delete(k) : s.notes.set(k, old); }, () => s.notes.set(k, nv)); }); ctx.after();
  }

  /* ---------- grid hooks: page-break overlay + sparkline tab sync ---------- */
  XL.gridHooks.push(function (g, ctx) {
    const host = ctx.H.host, ov = el('div', { class: 'x-pbv' }); ov.style.cssText = 'position:absolute;inset:0;pointer-events:none;z-index:4;overflow:hidden'; host.appendChild(ov);
    const PAPER = { A4: [210, 297], Letter: [215.9, 279.4], Legal: [215.9, 355.6], A3: [297, 420], A5: [148, 210] }, MARG = { Normal: 12, Wide: 25, Narrow: 6 };
    g.overlayUpdate = function () {
      ov.innerHTML = ''; if (!g.pbv) return; const s = g.sh(), pg = ctx.page; const pp = PAPER[pg.paper || 'A4'], land = pg.orient === 'landscape', m = MARG[pg.margin || 'Normal'] || 12;
      const pw = ((land ? pp[1] : pp[0]) - 2 * m) / 25.4 * 96 / ((pg.scale || 100) / 100), ph = ((land ? pp[0] : pp[1]) - 2 * m) / 25.4 * 96 / ((pg.scale || 100) / 100);
      const u = s.printArea || ctx.wb.usedRange(s); if (!u || u.r2 < 0) return; const a = g.cellRect(u.r1, u.c1), b = g.cellRect(u.r2, u.c2); const z = g.zoom; const rh = g.showHeaders ? 48 * z : 0, chh = g.showHeaders ? 22 * z : 0;
      const area = el('div'); area.style.cssText = `position:absolute;left:${a.x + rh}px;top:${a.y + chh}px;width:${b.x + b.w - a.x}px;height:${b.y + b.h - a.y}px;outline:2px solid #1f4e9c;background:rgba(31,78,156,.06)`; ov.appendChild(area);
      const w0 = (b.x + b.w - a.x) / z, h0 = (b.y + b.h - a.y) / z; let n = 1; for (let y = 0; y < h0 || y === 0; y += ph) { for (let x = 0; x < w0 || x === 0; x += pw) { const lab = el('div', { text: 'Page ' + n++ }); lab.style.cssText = `position:absolute;left:${a.x + rh + x * z}px;top:${a.y + chh + y * z}px;width:${Math.min(pw, w0 - x) * z}px;height:${Math.min(ph, h0 - y) * z}px;border:1px dashed #1f4e9c;color:rgba(31,78,156,.25);font:700 ${Math.max(16, 48 * z)}px sans-serif;display:flex;align-items:center;justify-content:center;box-sizing:border-box`; ov.appendChild(lab); if (x + pw >= w0) break; } if (y + ph >= h0) break; }
    };
    const od = g.draw; g.draw = function () { od(); if (g.pbv) setTimeout(g.overlayUpdate, 0); else if (ov.firstChild) ov.innerHTML = ''; ctx._sparkSync && ctx._sparkSync(); };
    host.addEventListener('scroll', () => g.pbv && g.overlayUpdate(), true);
  });
})();
