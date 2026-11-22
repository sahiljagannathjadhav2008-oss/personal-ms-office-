/* Excel grid: canvas renderer, selection, editing, clipboard, charts layer — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL, OA = global.OA, el = OA.el, key = XL.key;
  const RHW = 48, CHH = 22; // row header width, col header height
  const FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';

  XL.createGrid = function (host, wb, hooks) {
    hooks = hooks || {};
    const canvas = el('canvas', { class: 'x-canvas' }), scroller = el('div', { class: 'x-scroller', tabindex: '0' }), spacer = el('div', { class: 'x-spacer' });
    scroller.appendChild(spacer); host.append(canvas, scroller);
    const ctx = canvas.getContext('2d');
    const ed = el('textarea', { class: 'x-ed', rows: 1, spellcheck: 'false', autocomplete: 'off', autocapitalize: 'off' }); ed.hidden = true; host.appendChild(ed);
    const g = { wb, canvas, scroller, spacer, ed, zoom: 1, showHeaders: true, showFormulas: false, clip: null, editing: null, selChart: null, selMode: false };
    let L = null, W = 0, H = 0, dpr = 1, raf = 0, antsOn = 0;

    /* ---------------- sheet / selection state */
    g.sh = () => wb.sheets[wb.activeIndex || 0];
    function ui(sh) { sh = sh || g.sh(); return sh._ui || (sh._ui = { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], ar: 0, ac: 0, sx: 0, sy: 0, vr: 2000, vc: 60 }); }
    g.ui = ui;
    g.active = () => { const u = ui(); return { r: u.r, c: u.c }; };
    g.ranges = () => ui().ranges;
    g.range = () => { const u = ui(); const q = u.ranges[u.ranges.length - 1]; return { r1: q.r1, c1: q.c1, r2: q.r2, c2: q.c2 }; };
    const norm = q => ({ r1: Math.min(q.r1, q.r2), c1: Math.min(q.c1, q.c2), r2: Math.max(q.r1, q.r2), c2: Math.max(q.c1, q.c2) });

    /* ---------------- layout */
    function layout() {
      const sh = g.sh(), u = ui(sh); const ext = wb.extent(sh.name);
      u.vr = Math.max(u.vr, ext.r + 300, u.r + 200); u.vc = Math.max(u.vc, ext.c + 30, u.c + 20); u.vr = Math.min(u.vr, XL.MAXR); u.vc = Math.min(u.vc, XL.MAXC);
      const k = wb.version + ':' + wb.sheets.indexOf(sh) + ':' + u.vr + ':' + u.vc;
      if (L && L.k === k) return L;
      const cx = new Float64Array(u.vc + 1), ry = new Float64Array(u.vr + 1);
      for (let c = 0; c < u.vc; c++) cx[c + 1] = cx[c] + wb.colWidth(sh, c);
      for (let r = 0; r < u.vr; r++) ry[r + 1] = ry[r] + wb.rowHeight(sh, r);
      L = { k, cx, ry, vr: u.vr, vc: u.vc };
      return L;
    }
    g.layout = layout;
    const idxAt = (arr, v, n) => { let lo = 0, hi = n - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (arr[m] <= v) lo = m; else hi = m - 1; } return lo; };
    const RH = () => g.showHeaders ? RHW : 0, CH = () => g.showHeaders ? CHH : 0;
    function frozen() { const sh = g.sh(); const l = layout(); return { fr: Math.min(sh.freeze.r, l.vr), fc: Math.min(sh.freeze.c, l.vc), fw: l.cx[Math.min(sh.freeze.c, l.vc)], fh: l.ry[Math.min(sh.freeze.r, l.vr)] }; }
    function resize() {
      const r = host.getBoundingClientRect(); W = Math.max(50, Math.floor(r.width)); H = Math.max(50, Math.floor(r.height)); dpr = window.devicePixelRatio || 1;
      canvas.width = Math.floor(W * dpr); canvas.height = Math.floor(H * dpr); canvas.style.width = W + 'px'; canvas.style.height = H + 'px'; updateSpacer(); g.draw();
    }
    function updateSpacer() { const l = layout(); spacer.style.width = ((RH() + l.cx[l.vc]) * g.zoom + 60) + 'px'; spacer.style.height = ((CH() + l.ry[l.vr]) * g.zoom + 60) + 'px'; }
    g.relayout = () => { L = null; updateSpacer(); layoutCharts(); g.draw(); };
    // convert content coords (unzoomed, relative to cell origin) <-> screen
    function view() {
      const l = layout(), f = frozen(), z = g.zoom, sx = scroller.scrollLeft / z, sy = scroller.scrollTop / z, rh = RH(), ch = CH();
      const cStart = Math.max(f.fc, idxAt(l.cx, l.cx[f.fc] + sx, l.vc)), rStart = Math.max(f.fr, idxAt(l.ry, l.ry[f.fr] + sy, l.vr));
      let cEnd = cStart, rEnd = rStart; const wlim = W / z - rh - f.fw + sx + l.cx[f.fc] + 5, hlim = H / z - ch - f.fh + sy + l.ry[f.fr] + 5;
      while (cEnd < l.vc - 1 && l.cx[cEnd + 1] < wlim) cEnd++; while (rEnd < l.vr - 1 && l.ry[rEnd + 1] < hlim) rEnd++;
      return { l, f, z, sx, sy, rh, ch, cStart, cEnd, rStart, rEnd };
    }
    // screen position (px, zoomed) of cell rect
    function cellRect(r, c, v) {
      v = v || view(); const { l, f, z, sx, sy, rh, ch } = v;
      const x = c < f.fc ? rh + l.cx[c] : rh + f.fw + (l.cx[c] - l.cx[f.fc] - sx), y = r < f.fr ? ch + l.ry[r] : ch + f.fh + (l.ry[r] - l.ry[f.fr] - sy);
      return { x: x * z, y: y * z, w: (l.cx[c + 1] - l.cx[c]) * z, h: (l.ry[r + 1] - l.ry[r]) * z };
    }
    function mergedRect(r, c, v) {
      const sh = g.sh(); const m = wb.mergeAt(sh, r, c); if (!m) return Object.assign(cellRect(r, c, v), { m: null });
      const a = cellRect(m.r1, m.c1, v), b = cellRect(m.r2, m.c2, v); return { x: a.x, y: a.y, w: b.x + b.w - a.x, h: b.y + b.h - a.y, m };
    }
    g.cellRect = (r, c) => mergedRect(r, c);
    function hit(px, py) {
      const v = view(), { l, f, z, sx, sy, rh, ch } = v; const x = px / z, y = py / z;
      if (x < rh && y < ch) return { t: 'corner' };
      let c, r;
      if (x >= rh) { const xx = x - rh; c = xx < f.fw ? idxAt(l.cx, xx, l.vc) : idxAt(l.cx, xx - f.fw + l.cx[f.fc] + sx, l.vc); } else c = -1;
      if (y >= ch) { const yy = y - ch; r = yy < f.fh ? idxAt(l.ry, yy, l.vr) : idxAt(l.ry, yy - f.fh + l.ry[f.fr] + sy, l.vr); } else r = -1;
      if (y < ch && c >= 0) { const cr = cellRect(0, c, v); const edgeR = (cr.x + cr.w) - px; if (edgeR <= 4 && edgeR >= -4 || (px - cr.x <= 3 && c > 0)) return { t: 'colresize', c: edgeR <= 4 ? c : c - 1 }; return { t: 'colh', c }; }
      if (x < rh && r >= 0) { const cr = cellRect(r, 0, v); const edgeB = (cr.y + cr.h) - py; if (edgeB <= 3 && edgeB >= -3 || (py - cr.y <= 2 && r > 0)) return { t: 'rowresize', r: edgeB <= 3 ? r : r - 1 }; return { t: 'rowh', r }; }
      return { t: 'cell', r, c };
    }
    g.hit = hit;

    /* ---------------- rendering */
    const colorOr = (c, d) => c || d;
    function fontOf(st) { return (st.i ? 'italic ' : '') + (st.b ? 'bold ' : '') + ((st.sz || 11) * 4 / 3) + 'px ' + (st.font ? '"' + st.font + '", ' : '') + FONT; }
    function drawIcon(c, x, y, h, ic) {
      const col = ic.n === 5 ? ['#c00000', '#f4a020', '#e6c800', '#9bc400', '#00a550'][ic.idx] : ['#d9363e', '#f0b400', '#2e9d4f'][ic.idx]; c.fillStyle = col; c.beginPath();
      if (ic.set === '3traffic') c.arc(x + h / 2, y + h / 2, h * 0.32, 0, 6.3); else if (ic.set === '3flags') { c.moveTo(x + 2, y + 1); c.lineTo(x + h * 0.9, y + h * 0.35); c.lineTo(x + 2, y + h * 0.6); } else { const up = ic.idx === ic.n - 1, mid = ic.idx > 0 && ic.idx < ic.n - 1; if (mid) { c.moveTo(x + h * 0.15, y + h * 0.3); c.lineTo(x + h * 0.85, y + h * 0.5); c.lineTo(x + h * 0.15, y + h * 0.7); } else if (up) { c.moveTo(x + h / 2, y + h * 0.1); c.lineTo(x + h * 0.9, y + h * 0.8); c.lineTo(x + h * 0.1, y + h * 0.8); } else { c.moveTo(x + h / 2, y + h * 0.9); c.lineTo(x + h * 0.9, y + h * 0.2); c.lineTo(x + h * 0.1, y + h * 0.2); } }
      c.closePath(); c.fill();
    }
    function wrapLines(c, text, w) {
      const out = []; text.split('\n').forEach(par => { let line = ''; par.split(/(\s+)/).forEach(tok => { const t = line + tok; if (c.measureText(t).width > w && line) { out.push(line.trimEnd()); line = tok.trimStart(); } else line = t; }); out.push(line); }); return out;
    }
    const BW = { thin: 1, medium: 2, thick: 3, dashed: 1, dotted: 1, double: 3, hair: 1 };
    function strokeBorder(c, b, x1, y1, x2, y2) { if (!b) return; c.strokeStyle = b.c || '#000'; c.lineWidth = BW[b.s] || 1; c.setLineDash(b.s === 'dashed' ? [4, 3] : b.s === 'dotted' ? [1.5, 1.5] : []); const o = (c.lineWidth % 2) ? 0.5 : 0; c.beginPath(); c.moveTo(x1 + (x1 === x2 ? o : 0), y1 + (y1 === y2 ? o : 0)); c.lineTo(x2 + (x1 === x2 ? o : 0), y2 + (y1 === y2 ? o : 0)); c.stroke(); c.setLineDash([]); }
    g.draw = function () { if (!raf) raf = requestAnimationFrame(() => { raf = 0; paint(); }); };
    function paint() {
      wb.ensureCalc(); const sh = g.sh(), u = ui(sh); const c = ctx; const v = view(); const { l, f, z, rh, ch } = v;
      c.setTransform(dpr, 0, 0, dpr, 0, 0); c.clearRect(0, 0, W, H); c.fillStyle = '#fff'; c.fillRect(0, 0, W, H);
      const cols = [], rows = []; for (let i = 0; i < f.fc; i++) if (l.cx[i + 1] > l.cx[i]) cols.push(i); for (let i = v.cStart; i <= v.cEnd; i++) if (l.cx[i + 1] > l.cx[i]) cols.push(i);
      for (let i = 0; i < f.fr; i++) if (l.ry[i + 1] > l.ry[i]) rows.push(i); for (let i = v.rStart; i <= v.rEnd; i++) if (l.ry[i + 1] > l.ry[i]) rows.push(i);
      c.save(); c.scale(z, z);
      const gx = i => cellRect(0, i, v).x / z, gy = i => cellRect(i, 0, v).y / z;
      const x0 = rh, y0 = ch, xEnd = W / z, yEnd = H / z;
      // gridlines
      if (sh.showGrid) { c.strokeStyle = '#e1e1e1'; c.lineWidth = 1; c.beginPath(); cols.forEach(i => { const x = Math.round(gx(i) + (l.cx[i + 1] - l.cx[i])) + 0.5; c.moveTo(x, y0); c.lineTo(x, yEnd); }); rows.forEach(i => { const y = Math.round(gy(i) + (l.ry[i + 1] - l.ry[i])) + 0.5; c.moveTo(x0, y); c.lineTo(xEnd, y); }); c.stroke(); }
      const drawn = new Set(); const cfs = sh.cf.length ? true : false; const cellsToText = [];
      rows.forEach(r => cols.forEach(cc => {
        const m = wb.mergeAt(sh, r, cc); let rr = r, c2 = cc; if (m) { if (drawn.has(m)) return; drawn.add(m); rr = m.r1; c2 = m.c1; }
        const rect = mergedRect(rr, c2, v); const X = rect.x / z, Y = rect.y / z, w = rect.w / z, h = rect.h / z; if (X + w < x0 || Y + h < y0 || X > xEnd || Y > yEnd) return;
        const d = wb.display(sh, rr, c2); const st = d.st; const cf = cfs ? wb.evalCF(sh, rr, c2) : null; const bg = (cf && cf.bg) || st.bg;
        if (bg) { c.fillStyle = bg; c.fillRect(X, Y, w, h); }
        if (cf && cf.bar) { const bw = (w - 4) * cf.bar.pct; c.fillStyle = cf.bar.color; c.globalAlpha = 0.75; c.fillRect(X + 2, Y + 2, Math.max(1, bw), h - 4); c.globalAlpha = 1; }
        if (sh.sparks && sh.sparks.length && XL.drawSpark) XL.drawSpark(c, wb, sh, rr, c2, X, Y, w, h);
        cellsToText.push({ r: rr, c: c2, X, Y, w, h, d, st, cf, m });
      }));
      // borders
      cellsToText.forEach(o => { const { X, Y, w, h, st, r, c: cc, m } = o; const R2 = m ? m.r2 : r, C2 = m ? m.c2 : cc;
        const nb = (rr, c2, side) => { if (rr < 0 || c2 < 0) return null; return wb.styleOf(sh, rr, c2)[side]; };
        strokeBorder(c, st.bT || nb(r - 1, cc, 'bB'), X, Y, X + w, Y); strokeBorder(c, st.bB || nb(R2 + 1, cc, 'bT'), X, Y + h, X + w, Y + h); strokeBorder(c, st.bL || nb(r, cc - 1, 'bR'), X, Y, X, Y + h); strokeBorder(c, st.bR || nb(r, C2 + 1, 'bL'), X + w, Y, X + w, Y + h); });
      // text
      const tx = (rr, cc) => { const k = key(rr, cc); const cl = sh.cells.get(k); return (cl && (cl.f != null || (cl.v !== undefined && cl.v !== null && cl.v !== ''))) || sh.spill.has(k); };
      cellsToText.forEach(o => {
        const { X, Y, w, h, d, st, cf, r, c: cc, m } = o; let text = d.text;
        if (g.showFormulas) { const cl = sh.cells.get(key(r, cc)); if (cl && cl.f != null) text = '=' + cl.f; }
        const editingHere = g.editing && g.editing.r === r && g.editing.c === cc;
        const ds = wb.dvAt(sh, r, cc); const pad = 3; let iconW = 0;
        if (cf && cf.icon) { iconW = Math.min(h - 4, 14); drawIcon(c, X + 2, Y + (h - iconW) / 2, iconW, cf.icon); }
        if (text === '' || editingHere) return;
        c.save(); c.beginPath();
        c.font = fontOf(st); const color = (cf && cf.color) || d.color || st.color || '#000'; c.fillStyle = d.err ? (st.color || '#000') : color; c.textBaseline = 'alphabetic';
        let ha = st.ha; if (!ha || ha === 'general') ha = d.num ? 'right' : (d.bool || d.err) ? 'center' : 'left'; if (cf && cf.b) c.font = 'bold ' + c.font;
        let avail = w - pad * 2 - iconW - (st.ind ? st.ind * 9 : 0);
        let lines = [text];
        if (st.wrap) lines = wrapLines(c, text, avail);
        else { const tw = c.measureText(text).width; if (d.num && tw > avail) { text = '#'.repeat(Math.max(1, Math.floor(avail / 6))); lines = [text]; } }
        const lh = (st.sz || 11) * 4 / 3 * 1.2; const th = lines.length * lh; const va = st.va || 'bottom'; let ty = va === 'top' ? Y + 2 + (st.sz || 11) * 4 / 3 : va === 'middle' ? Y + (h - th) / 2 + (st.sz || 11) * 4 / 3 * 0.95 : Y + h - th + (st.sz || 11) * 4 / 3 * 0.85 - 2;
        // overflow into empty neighbours
        let cx0 = X, cx1 = X + w;
        if (!st.wrap && !m && !(d.num) && lines.length === 1) { const tw = c.measureText(lines[0]).width; if (tw > avail) { if (ha === 'left') { let k = cc + 1; while (cx1 < xEnd && k < l.vc && !tx(r, k) && !wb.mergeAt(sh, r, k)) { cx1 += (l.cx[k + 1] - l.cx[k]); k++; if (cx1 - X > tw + 6) break; } } else if (ha === 'right') { let k = cc - 1; while (cx0 > x0 && k >= 0 && !tx(r, k)) { cx0 -= (l.cx[k + 1] - l.cx[k]); k--; if (X + w - cx0 > tw + 6) break; } } else { let k = cc + 1, k2 = cc - 1; while (cx1 - X < (tw - w) / 2 + w && !tx(r, k) && k < l.vc) { cx1 += l.cx[k + 1] - l.cx[k]; k++; } while (X - cx0 < (tw - w) / 2 && k2 >= 0 && !tx(r, k2)) { cx0 -= l.cx[k2 + 1] - l.cx[k2]; k2--; } } } }
        c.rect(Math.max(cx0, c === ctx ? (cc >= f.fc ? x0 + f.fw * 0 : 0) : 0), Y, cx1 - cx0, h); c.clip();
        lines.forEach((ln, i) => { const tw = c.measureText(ln).width; let xx; if (ha === 'left') xx = X + pad + iconW + (st.ind ? st.ind * 9 : 0); else if (ha === 'right') xx = X + w - pad - tw; else xx = X + (w - tw) / 2 + iconW / 2; const yy = ty + i * lh; c.fillText(ln, xx, yy);
          if (st.u) { c.fillRect(xx, yy + 2, tw, 1); } if (st.st) { c.fillRect(xx, yy - (st.sz || 11) * 4 / 3 * 0.3, tw, 1); } });
        c.restore();
      });
      g._cells = cellsToText;
      // freeze lines
      c.strokeStyle = '#7f7f7f'; c.lineWidth = 1; if (f.fr) { const y = ch + f.fh + 0.5; c.beginPath(); c.moveTo(rh, y); c.lineTo(xEnd, y); c.stroke(); } if (f.fc) { const x = rh + f.fw + 0.5; c.beginPath(); c.moveTo(x, ch); c.lineTo(x, yEnd); c.stroke(); }
      // filter / validation buttons
      const btn = (X, Y, w, h) => { const s = 15, bx = X + w - s - 1, by = Y + h - s - 1; c.fillStyle = '#f3f3f3'; c.strokeStyle = '#8a8a8a'; c.lineWidth = 1; c.fillRect(bx, by, s, s); c.strokeRect(bx + 0.5, by + 0.5, s - 1, s - 1); c.fillStyle = '#333'; c.beginPath(); c.moveTo(bx + 4, by + 6); c.lineTo(bx + 11, by + 6); c.lineTo(bx + 7.5, by + 10.5); c.closePath(); c.fill(); };
      if (sh.filter) { const fr = sh.filter.range; for (let cc = fr.c1; cc <= fr.c2; cc++) { if (!cols.includes(cc) || !rows.includes(fr.r1)) continue; const rc = cellRect(fr.r1, cc, v); btn(rc.x / z, rc.y / z, rc.w / z, rc.h / z); if (sh.filter.crit[cc]) { c.fillStyle = '#217346'; const rr = cellRect(fr.r1, cc, v); c.fillRect(rr.x / z + rr.w / z - 6, rr.y / z + rr.h / z - 16, 4, 4); } } }
      const ar = mergedRect(u.r, u.c, v); const adv = wb.dvAt(sh, u.r, u.c); if (adv && adv.type === 'list' && !g.editing) btn(ar.x / z, ar.y / z, ar.w / z, ar.h / z);
      // headers
      if (g.showHeaders) {
        c.fillStyle = '#f0f0f0'; c.fillRect(0, 0, xEnd, ch); c.fillRect(0, 0, rh, yEnd); c.font = '12px ' + FONT; c.textBaseline = 'middle'; c.textAlign = 'center';
        const selC = new Set(), selR = new Set(), fullC = new Set(), fullR = new Set(); u.ranges.forEach(q => { const n = norm(q); for (let i = n.c1; i <= Math.min(n.c2, l.vc); i++) { if (n.c2 - n.c1 < 700) selC.add(i); if (n.r1 === 0 && n.r2 >= l.vr - 1) fullC.add(i); } for (let i = n.r1; i <= Math.min(n.r2, l.vr); i++) { if (n.r2 - n.r1 < 5000) selR.add(i); if (n.c1 === 0 && n.c2 >= l.vc - 1) fullR.add(i); } });
        cols.forEach(i => { const rc = cellRect(0, i, v), X = rc.x / z, w = rc.w / z; if (selC.has(i)) { c.fillStyle = fullC.has(i) ? '#217346' : '#d3e5db'; c.fillRect(X, 0, w, ch); } c.fillStyle = fullC.has(i) ? '#fff' : (selC.has(i) ? '#107c41' : '#444'); c.fillText(XL.colName(i), X + w / 2, ch / 2 + 1); c.strokeStyle = '#c4c4c4'; c.beginPath(); c.moveTo(Math.round(X + w) + 0.5, 0); c.lineTo(Math.round(X + w) + 0.5, ch); c.stroke(); });
        rows.forEach(i => { const rc = cellRect(i, 0, v), Y = rc.y / z, h = rc.h / z; if (selR.has(i)) { c.fillStyle = fullR.has(i) ? '#217346' : '#d3e5db'; c.fillRect(0, Y, rh, h); } c.fillStyle = fullR.has(i) ? '#fff' : (selR.has(i) ? '#107c41' : '#444'); c.fillText(String(i + 1), rh / 2, Y + h / 2 + 1); c.strokeStyle = '#c4c4c4'; c.beginPath(); c.moveTo(0, Math.round(Y + h) + 0.5); c.lineTo(rh, Math.round(Y + h) + 0.5); c.stroke(); });
        c.fillStyle = '#e8e8e8'; c.fillRect(0, 0, rh, ch); c.fillStyle = '#b5b5b5'; c.beginPath(); c.moveTo(rh - 4, ch - 4); c.lineTo(rh - 4, ch - 14); c.lineTo(rh - 14, ch - 4); c.closePath(); c.fill(); c.strokeStyle = '#c4c4c4'; c.beginPath(); c.moveTo(rh + 0.5, 0); c.lineTo(rh + 0.5, ch); c.moveTo(0, ch + 0.5); c.lineTo(rh, ch + 0.5); c.stroke();
        c.textAlign = 'left';
      }
      // selection
      const clipRect = (q) => { const n = norm(q); const a = mergedRect(n.r1, n.c1, v), b = mergedRect(n.r2, n.c2, v); return { x: a.x / z, y: a.y / z, w: (b.x + b.w - a.x) / z, h: (b.y + b.h - a.y) / z, n }; };
      c.save(); c.beginPath(); c.rect(rh, ch, xEnd - rh, yEnd - ch); c.clip();
      const pri = u.ranges[u.ranges.length - 1];
      u.ranges.forEach(q => { const n = norm(q); const rrR = clipRectSafe(n, v, l); if (!rrR) return; const single = n.r1 === n.r2 && n.c1 === n.c2 || (wb.mergeAt(sh, n.r1, n.c1) && wb.mergeAt(sh, n.r1, n.c1).r2 === n.r2 && wb.mergeAt(sh, n.r1, n.c1).c2 === n.c2);
        if (!single) { c.fillStyle = 'rgba(33,115,70,0.12)'; c.fillRect(rrR.x, rrR.y, rrR.w, rrR.h); const ac = mergedRect(u.r, u.c, v); if (q === pri && u.r >= n.r1 && u.r <= n.r2 && u.c >= n.c1 && u.c <= n.c2) { c.fillStyle = '#fff'; c.fillRect(ac.x / z, ac.y / z, ac.w / z, ac.h / z); const dd = wb.display(sh, u.r, u.c); } }
        c.strokeStyle = '#217346'; c.lineWidth = 2; c.strokeRect(rrR.x + 1, rrR.y + 1, rrR.w - 2, rrR.h - 2); });
      // redraw active cell text over white fill when multi-select
      const act = mergedRect(u.r, u.c, v); const apos = { x: act.x / z, y: act.y / z, w: act.w / z, h: act.h / z };
      const prn = norm(pri); if (!(prn.r1 === prn.r2 && prn.c1 === prn.c2)) { const o = cellsToText.find(q => q.r === (act.m ? act.m.r1 : u.r) && q.c === (act.m ? act.m.c1 : u.c)); if (o) { c.save(); c.beginPath(); c.rect(apos.x, apos.y, apos.w, apos.h); c.clip(); c.fillStyle = '#fff'; c.fillRect(apos.x, apos.y, apos.w, apos.h); c.font = fontOf(o.st); c.fillStyle = o.st.color || '#000'; const ha = o.st.ha || (o.d.num ? 'right' : 'left'); const tw = c.measureText(o.d.text).width; c.textBaseline = 'alphabetic'; c.fillText(o.d.text, ha === 'right' ? apos.x + apos.w - 3 - tw : ha === 'center' ? apos.x + (apos.w - tw) / 2 : apos.x + 3, apos.y + apos.h - 4); c.restore(); } }
      // fill handle
      const pr = clipRectSafe(prn, v, l); if (pr && !g.editing && !g.selChart) { c.fillStyle = '#217346'; c.fillRect(pr.x + pr.w - 4, pr.y + pr.h - 4, 7, 7); c.strokeStyle = '#fff'; c.lineWidth = 1; c.strokeRect(pr.x + pr.w - 4, pr.y + pr.h - 4, 7, 7); g._fh = { x: (pr.x + pr.w - 4) * z, y: (pr.y + pr.h - 4) * z, s: 7 * z }; } else g._fh = null;
      // copy marching border
      if (g.clip && g.clip.sheet === sh.name) { const gg = g.clip.g; const rc = clipRectSafe(gg, v, l); if (rc) { c.strokeStyle = '#217346'; c.lineWidth = 2; c.setLineDash([5, 4]); c.lineDashOffset = -(antsOn % 18); c.strokeRect(rc.x + 1, rc.y + 1, rc.w - 2, rc.h - 2); c.setLineDash([]); } }
      // fill drag preview
      if (g.fillDrag) { const rc = clipRectSafe(g.fillDrag, v, l); if (rc) { c.strokeStyle = '#555'; c.lineWidth = 1; c.setLineDash([3, 3]); c.strokeRect(rc.x + 0.5, rc.y + 0.5, rc.w, rc.h); c.setLineDash([]); } }
      // formula reference highlights
      if (g.refBoxes) g.refBoxes.forEach((b, i) => { const rc = clipRectSafe(b.g, v, l); if (rc) { c.strokeStyle = b.color; c.lineWidth = 2; c.strokeRect(rc.x + 1, rc.y + 1, rc.w - 2, rc.h - 2); } });
      c.restore(); c.restore();
      if (g.clip && g.clip.sheet === sh.name && !g.antT) { g.antT = setInterval(() => { antsOn++; if (!g.clip) { clearInterval(g.antT); g.antT = null; } else g.draw(); }, 250); } else if ((!g.clip) && g.antT) { clearInterval(g.antT); g.antT = null; }
      hooks.onPaint && hooks.onPaint();
    }
    function clipRectSafe(n, v, l) {
      const { z } = v; const a = mergedRect(Math.max(0, n.r1), Math.max(0, n.c1), v), b = mergedRect(Math.min(l.vr - 1, n.r2), Math.min(l.vc - 1, n.c2), v);
      const full = n.r2 >= l.vr - 1 || n.c2 >= l.vc - 1; if (!isFinite(a.x) || !isFinite(b.x)) return null;
      const rc = { x: a.x / z, y: a.y / z, w: (b.x + b.w - a.x) / z, h: (b.y + b.h - a.y) / z }; if (rc.w < 0 || rc.h < 0) { if (rc.w < 0) { const t = a.x; } } return rc;
    }

    /* ---------------- selection & navigation */
    const MAXR = XL.MAXR, MAXC = XL.MAXC;
    const hidR = r => wb.rowHeight(g.sh(), r) === 0, hidC = c => wb.colWidth(g.sh(), c) === 0;
    function notify() { hooks.onSel && hooks.onSel(); g.draw(); }
    g.ensureVisible = function (r, c) {
      const v = view(), { l, f, z } = v, sc = scroller; const visW = (sc.clientWidth / z) - v.rh - f.fw, visH = (sc.clientHeight / z) - v.ch - f.fh;
      if (c >= f.fc) { const a = l.cx[c] - l.cx[f.fc], b = l.cx[c + 1] - l.cx[f.fc]; if (a < v.sx) sc.scrollLeft = a * z; else if (b > v.sx + visW) sc.scrollLeft = Math.max(0, (b - visW + 4) * z); }
      if (r >= f.fr) { const a = l.ry[r] - l.ry[f.fr], b = l.ry[r + 1] - l.ry[f.fr]; if (a < v.sy) sc.scrollTop = a * z; else if (b > v.sy + visH) sc.scrollTop = Math.max(0, (b - visH + 4) * z); }
    };
    g.setActive = function (r, c, keep) {
      const u = ui(); r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c)); const m = wb.mergeAt(g.sh(), r, c); if (m) { r = m.r1; c = m.c1; }
      u.r = r; u.c = c; u.ar = r; u.ac = c; if (!keep) u.ranges = [{ r1: r, c1: c, r2: r, c2: c }]; if (m) u.ranges[u.ranges.length - 1] = { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 };
      if (u.r + 100 > u.vr) u.vr = Math.min(MAXR, u.r + 200); if (u.c + 10 > u.vc) u.vc = Math.min(MAXC, u.c + 20); updateSpacer(); g.ensureVisible(r, c); g.selChart = null; layoutCharts(); notify();
    };
    g.select = function (q) { const u = ui(); const n = wb.expandMerges(g.sh(), norm(q)); u.ranges = [n]; u.ar = n.r1; u.ac = n.c1; u.r = n.r1; u.c = n.c1; notify(); };
    g.extendTo = function (r, c, add) {
      const u = ui(); r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c)); const n = wb.expandMerges(g.sh(), norm({ r1: u.ar, c1: u.ac, r2: r, c2: c }));
      u.ranges[u.ranges.length - 1] = n; u.fr = r; u.fc = c; notify();
    };
    g.addRange = function (q) { const u = ui(); const n = wb.expandMerges(g.sh(), norm(q)); u.ranges.push(n); u.r = n.r1; u.c = n.c1; u.ar = n.r1; u.ac = n.c1; notify(); };
    g.selectAll = function () { const u = ui(); u.ranges = [{ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }]; notify(); };
    g.selectCols = function (c1, c2) { const u = ui(); const q = { r1: 0, c1: Math.min(c1, c2), r2: MAXR - 1, c2: Math.max(c1, c2) }; u.ranges = [q]; u.r = view().rStart; u.c = q.c1; u.ar = 0; u.ac = c1; notify(); };
    g.selectRows = function (r1, r2) { const u = ui(); const q = { r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAXC - 1 }; u.ranges = [q]; u.r = q.r1; u.c = view().cStart; u.ar = r1; u.ac = 0; notify(); };
    function stepFrom(r, c, dr, dc) {
      const sh = g.sh(); let m = wb.mergeAt(sh, r, c); if (m) { if (dr > 0) r = m.r2; if (dr < 0) r = m.r1; if (dc > 0) c = m.c2; if (dc < 0) c = m.c1; }
      let nr = r + dr, nc = c + dc; while (nr >= 0 && nr < MAXR && dr && hidR(nr)) nr += dr; while (nc >= 0 && nc < MAXC && dc && hidC(nc)) nc += dc;
      nr = Math.max(0, Math.min(MAXR - 1, nr)); nc = Math.max(0, Math.min(MAXC - 1, nc)); return { r: nr, c: nc };
    }
    const hasVal = (sh, r, c) => { const k = key(r, c); const cl = sh.cells.get(k); return !!(cl && (cl.f != null || (cl.v !== undefined && cl.v !== null && cl.v !== ''))) || sh.spill.has(k); };
    function jump(r, c, dr, dc) {
      const sh = g.sh(); const ext = wb.extent(sh.name); const lr = Math.max(ext.r, 0), lc = Math.max(ext.c, 0);
      const inb = (a, b) => a >= 0 && b >= 0 && a < MAXR && b < MAXC; let a = r + dr, b = c + dc; if (!inb(a, b)) return { r, c };
      const cur = hasVal(sh, r, c), nxt = hasVal(sh, a, b);
      if (cur && nxt) { while (inb(a + dr, b + dc) && hasVal(sh, a + dr, b + dc)) { a += dr; b += dc; } return { r: a, c: b }; }
      while (inb(a, b) && !hasVal(sh, a, b)) { if ((dr > 0 && a > lr + 1) || (dc > 0 && b > lc + 1)) { a = dr > 0 ? Math.min(MAXR - 1, a) : a; break; } if ((dr < 0 && a <= 0) || (dc < 0 && b <= 0)) break; a += dr; b += dc; }
      if (!inb(a, b)) return { r: Math.max(0, Math.min(MAXR - 1, a)), c: Math.max(0, Math.min(MAXC - 1, b)) };
      if (!hasVal(sh, a, b)) { if (dr > 0) a = Math.min(MAXR - 1, Math.max(a, lr)); if (dc > 0) b = Math.min(MAXC - 1, Math.max(b, lc)); if (dr < 0) a = 0; if (dc < 0) b = 0; }
      return { r: a, c: b };
    }
    g.moveActive = function (dr, dc, extend, ctrl) {
      const u = ui(); const q = u.ranges[u.ranges.length - 1];
      if (extend) { const fr = u.fr !== undefined && (q.r1 !== q.r2 || q.c1 !== q.c2) ? u.fr : u.r, fc = u.fc !== undefined && (q.r1 !== q.r2 || q.c1 !== q.c2) ? u.fc : u.c; const t = ctrl ? jump(fr, fc, dr, dc) : stepFrom(fr, fc, dr, dc); g.extendTo(t.r, t.c); g.ensureVisible(t.r, t.c); return; }
      const t = ctrl ? jump(u.r, u.c, dr, dc) : stepFrom(u.r, u.c, dr, dc); u.fr = u.fc = undefined; g.setActive(t.r, t.c);
    };
    g.activeText = () => { const u = ui(); return wb.rawText(g.sh(), u.r, u.c); };
    g.activeDisplay = () => { const u = ui(); return wb.display(g.sh(), u.r, u.c); };
    g.setZoom = function (z) {
      z = Math.max(0.4, Math.min(4, z)); const o = g.zoom; if (z === o) return; const sx = scroller.scrollLeft / o, sy = scroller.scrollTop / o; g.zoom = z; updateSpacer(); scroller.scrollLeft = sx * z; scroller.scrollTop = sy * z; layoutCharts(); placeEditor(); g.draw(); hooks.onZoom && hooks.onZoom(z);
    };
    g.setHeaders = v => { g.showHeaders = v; updateSpacer(); layoutCharts(); g.draw(); };

    /* ---------------- editor */
    const REFCOL = ['#0070c0', '#c00000', '#7030a0', '#00b050', '#ed7d31', '#00b0f0', '#bf8f00'];
    function updateRefBoxes() {
      if (!g.editing) { g.refBoxes = null; return; } const t = ed.value; if (t[0] !== '=') { g.refBoxes = null; return; }
      const out = []; const re = /(\$?[A-Z]{1,3}\$?\d+)(?::(\$?[A-Z]{1,3}\$?\d+))?/g; let m, i = 0;
      while ((m = re.exec(t))) { if (m.index > 0 && /[A-Za-z_!.'\d]/.test(t[m.index - 1])) continue; try { const a = XL.parseAddr(m[1].replace(/\$/g, '')), b = m[2] ? XL.parseAddr(m[2].replace(/\$/g, '')) : a; if (!a || !b) continue; out.push({ g: norm({ r1: a.r, c1: a.c, r2: b.r, c2: b.c }), color: REFCOL[i++ % REFCOL.length] }); } catch (e) { } }
      g.refBoxes = out;
    }
    function placeEditor() {
      if (!g.editing) return; const e = g.editing, sh = g.sh(); const rc = mergedRect(e.r, e.c, view()); const st = wb.styleOf(sh, e.r, e.c); const z = g.zoom;
      ctx.font = fontOf(st); const lines = ed.value.split('\n'); let tw = 0; lines.forEach(s => tw = Math.max(tw, ctx.measureText(s).width * z)); const fs = (st.sz || 11) * 4 / 3 * z; const lh = fs * 1.25;
      ed.style.fontSize = fs + 'px'; ed.style.fontWeight = st.b ? '700' : '400'; ed.style.fontStyle = st.i ? 'italic' : 'normal'; ed.style.fontFamily = (st.font ? '"' + st.font + '",' : '') + FONT; ed.style.color = st.color || '#000'; ed.style.textAlign = (st.ha === 'right' || (!st.ha && typeof e.num === 'number')) ? 'left' : (st.ha === 'center' ? 'center' : 'left');
      const wrapIt = !!st.wrap; ed.classList.toggle('wrap', wrapIt);
      const w = wrapIt ? rc.w : Math.min(Math.max(rc.w, tw + 24), Math.max(rc.w, W - rc.x - 4)); const h = Math.max(rc.h, (wrapIt ? Math.ceil(tw / Math.max(10, w - 8)) + lines.length - 1 : lines.length) * lh + 6);
      ed.style.left = (rc.x - 1) + 'px'; ed.style.top = (rc.y - 1) + 'px'; ed.style.width = (w + 2) + 'px'; ed.style.height = (h + 2) + 'px';
    }
    g.startEdit = function (r, c, initial, mode, viaBar) {
      if (g.editing) return; const sh = g.sh(); if (sh.protect && wb.styleOf(sh, r, c).locked !== false && sh.protect.on) { OA.toast('Sheet protected hai — pehle Unprotect Sheet karein'); return; }
      const m = wb.mergeAt(sh, r, c); if (m) { r = m.r1; c = m.c1; } if (sh.spill.has(key(r, c)) && !sh.cells.has(key(r, c))) { /* editing spill result: start empty */ }
      const raw = initial !== undefined && initial !== null ? initial : wb.rawText(sh, r, c); const cur = sh.cells.get(key(r, c));
      g.editing = { r, c, mode: mode || (initial !== undefined && initial !== null ? 'enter' : 'edit'), num: typeof (cur && cur.v) === 'number', bar: !!viaBar, orig: wb.rawText(sh, r, c) };
      ed.hidden = false; ed.value = raw; g.refIns = null; placeEditor(); updateRefBoxes();
      if (!viaBar) { ed.focus(); const n = ed.value.length; ed.setSelectionRange(n, n); } hooks.onEditState && hooks.onEditState(true); syncBar(); g.draw();
    };
    function syncBar() { if (g.bar && g.editing && g.bar.value !== ed.value) g.bar.value = ed.value; }
    g.bindBar = function (inp) {
      g.bar = inp;
      inp.addEventListener('focus', () => { if (!g.editing) { const u = ui(); g.startEdit(u.r, u.c, undefined, 'edit', true); } else g.editing.bar = true; });
      inp.addEventListener('input', () => { if (!g.editing) return; ed.value = inp.value; g.editing.bar = true; onEdInput(true); });
      inp.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); g.commit(1, 0); scroller.focus(); } else if (e.key === 'Escape') { e.preventDefault(); g.cancel(); scroller.focus(); } else if (e.key === 'Tab') { e.preventDefault(); g.commit(0, e.shiftKey ? -1 : 1); scroller.focus(); } });
    };
    g.cancel = function () { if (!g.editing) return; g.editing = null; ed.hidden = true; hideAC(); g.refBoxes = null; hooks.onEditState && hooks.onEditState(false); hooks.onSel && hooks.onSel(); scroller.focus({ preventScroll: true }); g.draw(); };
    g.commit = function (dr, dc, keepSel) {
      if (!g.editing) return true; const e = g.editing, sh = g.sh(); let text = ed.value;
      const rule = wb.dvCheck(sh, e.r, e.c, text.trim());
      if (rule && text !== '') { const style = rule.alert || 'stop'; const msg = rule.errMsg || 'Yeh value is cell ke liye valid nahi hai (Data Validation).';
        if (style === 'stop') { OA.alert(OA.esc ? OA.esc(msg) : msg, rule.errTitle || 'Microsoft Excel'); return false; }
        if (style === 'warning' || style === 'information') { OA.toast(msg); } }
      g.editing = null; ed.hidden = true; hideAC(); g.refBoxes = null;
      if (text !== e.orig || sh.spill.has(key(e.r, e.c))) { try { if (text.startsWith('=') && text.length > 1) text = balanceParens(text); wb.setInput(sh, e.r, e.c, text); } catch (err) { OA.alert('Formula mein error hai: ' + (err.message || err), 'Microsoft Excel'); } }
      hooks.onEditState && hooks.onEditState(false); updateSpacer(); layoutCharts();
      if (!keepSel) { const u = ui(); const q = u.ranges[u.ranges.length - 1]; const multi = q.r1 !== q.r2 || q.c1 !== q.c2; if (multi && (dr || dc) && u.ranges.length === 1) { stepInRange(dr, dc); } else g.moveActive(dr, dc); }
      scroller.focus({ preventScroll: true }); notify(); return true;
    };
    function balanceParens(t) { let o = 0, q = false; for (const ch of t) { if (ch === '"') q = !q; else if (!q) { if (ch === '(') o++; else if (ch === ')') o--; } } return o > 0 && !q ? t + ')'.repeat(o) : t; }
    function stepInRange(dr, dc) {
      const u = ui(), q = u.ranges[0]; let r = u.r + dr, c = u.c + dc;
      if (dr) { if (r > q.r2 || r < q.r1) { c += dr > 0 ? 1 : -1; r = dr > 0 ? q.r1 : q.r2; if (c > q.c2) c = q.c1; if (c < q.c1) c = q.c2; } }
      else { if (c > q.c2 || c < q.c1) { r += dc > 0 ? 1 : -1; c = dc > 0 ? q.c1 : q.c2; if (r > q.r2) r = q.r1; if (r < q.r1) r = q.r2; } }
      u.r = r; u.c = c; g.ensureVisible(r, c);
    }
    /* autocomplete + hint */
    let acEl = null, hintEl = null, acItems = [], acSel = 0;
    function hideAC() { if (acEl) { acEl.remove(); acEl = null; } if (hintEl) { hintEl.remove(); hintEl = null; } acItems = []; }
    function tokenAt() { const t = ed.value, p = ed.selectionStart; const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(t.slice(0, p)); return m ? { s: p - m[1].length, e: p, w: m[1] } : null; }
    function currentFn() { const t = ed.value.slice(0, ed.selectionStart); let depth = 0, q = false; for (let i = t.length - 1; i >= 0; i--) { const ch = t[i]; if (ch === '"') q = !q; if (q) continue; if (ch === ')') depth++; else if (ch === '(') { if (depth === 0) { const m = /([A-Za-z_][A-Za-z0-9_.]*)$/.exec(t.slice(0, i)); return m ? m[1].toUpperCase() : null; } depth--; } } return null; }
    function showAC() {
      hideAC(); if (!g.editing || ed.value[0] !== '=') return; const rc = mergedRect(g.editing.r, g.editing.c, view());
      const tk = tokenAt(); const prev = tk ? ed.value[tk.s - 1] : ''; if (tk && tk.w.length >= 1 && !/\d/.test(tk.w[0]) && (!prev || /[=+\-*\/^&(,<>;:\s%!]/.test(prev))) {
        const up = tk.w.toUpperCase(); acItems = Object.keys(XL.FUNCS).filter(n => n.startsWith(up) && !/^_/.test(n)).sort().slice(0, 40);
        if (acItems.length && !(acItems.length === 1 && acItems[0] === up && ed.value[tk.e] === '(')) { acSel = 0; acEl = el('div', { class: 'x-ac' }); acItems.forEach((n, i) => { const d = el('div', { class: i ? '' : 'on' }, [n, el('small', {}, [(XL.FSIG[n] || '').slice(n.length, 40)])]); d.addEventListener('mousedown', ev => { ev.preventDefault(); acSel = i; acceptAC(); }); acEl.appendChild(d); }); acEl.style.left = rc.x + 'px'; acEl.style.top = (rc.y + Math.max(rc.h, ed.offsetHeight) + 2) + 'px'; host.appendChild(acEl); }
      }
      const fn = currentFn(); if (fn && XL.FSIG[fn] && !acEl) { hintEl = el('div', { class: 'x-hint' }, [XL.FSIG[fn]]); hintEl.style.left = rc.x + 'px'; hintEl.style.top = Math.max(0, rc.y - 26) + 'px'; host.appendChild(hintEl); }
    }
    function acceptAC() { const tk = tokenAt(); if (!tk || !acItems.length) return; const n = acItems[acSel]; const t = ed.value; const after = t[tk.e] === '(' ? '' : '('; ed.value = t.slice(0, tk.s) + n + after + t.slice(tk.e); const p = tk.s + n.length + (after ? 1 : 1); ed.setSelectionRange(p, p); hideAC(); onEdInput(); }
    function moveAC(d) { acSel = (acSel + d + acItems.length) % acItems.length; Array.from(acEl.children).forEach((x, i) => { x.classList.toggle('on', i === acSel); if (i === acSel) x.scrollIntoView({ block: 'nearest' }); }); }
    function onEdInput(fromBar) { if (!g.editing) return; g.refIns = null; placeEditor(); updateRefBoxes(); if (!fromBar) syncBar(); showAC(); g.draw(); }
    ed.addEventListener('input', () => onEdInput());
    ed.addEventListener('click', () => { g.refIns = null; showAC(); });
    ed.addEventListener('blur', () => { /* click elsewhere handled by pointer logic */ });
    // formula reference pointing
    function canPoint() { if (!g.editing || ed.value[0] !== '=') return false; if (g.refIns && ed.selectionStart === g.refIns.end) return true; const p = ed.selectionStart; const before = ed.value.slice(0, p).trimEnd(); return before.length === 0 ? false : /[=+\-*\/^&(,<>;:%]$/.test(before) && ed.selectionEnd === p; }
    function refText(q) { const n = norm(q); const a = XL.addr(n.r1, n.c1), b = XL.addr(n.r2, n.c2); return (n.r1 === n.r2 && n.c1 === n.c2) ? a : a + ':' + b; }
    function pointRef(q, sheetName) {
      let s = refText(q); if (sheetName && sheetName !== g.sh().name) s = XL.quoteSheet(sheetName) + '!' + s; const t = ed.value; let st, en;
      if (g.refIns && ed.selectionStart === g.refIns.end) { st = g.refIns.start; en = g.refIns.end; } else { st = en = ed.selectionStart; }
      ed.value = t.slice(0, st) + s + t.slice(en); const p = st + s.length; ed.setSelectionRange(p, p); g.refIns = { start: st, end: p }; syncBar(); placeEditor(); updateRefBoxes(); g.draw();
    }
    g.pointRef = pointRef; g.canPoint = canPoint;
    ed.addEventListener('keydown', e => {
      if (!g.editing) return; const k = e.key;
      if (acEl && acItems.length) { if (k === 'ArrowDown') { e.preventDefault(); moveAC(1); return; } if (k === 'ArrowUp') { e.preventDefault(); moveAC(-1); return; } if (k === 'Tab' || k === 'Enter') { e.preventDefault(); acceptAC(); return; } if (k === 'Escape') { hideAC(); e.preventDefault(); return; } }
      if (k === 'Enter') { if (e.altKey) { e.preventDefault(); const p = ed.selectionStart; ed.value = ed.value.slice(0, p) + '\n' + ed.value.slice(ed.selectionEnd); ed.setSelectionRange(p + 1, p + 1); onEdInput(); return; } e.preventDefault(); g.commit(e.shiftKey ? -1 : 1, 0); return; }
      if (k === 'Tab') { e.preventDefault(); g.commit(0, e.shiftKey ? -1 : 1); return; }
      if (k === 'Escape') { e.preventDefault(); g.cancel(); return; }
      if (k === 'F2') { e.preventDefault(); g.editing.mode = g.editing.mode === 'enter' ? 'edit' : 'enter'; return; }
      if (k === 'F4' && ed.value[0] === '=') { e.preventDefault(); toggleAbs(); return; }
      if (/^Arrow/.test(k) && g.editing.mode === 'enter' && !e.altKey) {
        const dr = k === 'ArrowDown' ? 1 : k === 'ArrowUp' ? -1 : 0, dc = k === 'ArrowRight' ? 1 : k === 'ArrowLeft' ? -1 : 0;
        if (ed.value[0] === '=' && canPoint()) { e.preventDefault(); const base = g.refCur || { r: ui().r, c: ui().c, r0: ui().r, c0: ui().c }; if (!g.refIns) { g.refCur = { r: ui().r, c: ui().c, ar: ui().r, ac: ui().c }; } const rc = g.refCur; if (e.shiftKey) { rc.fr = (rc.fr === undefined ? rc.r : rc.fr) + dr; rc.fc = (rc.fc === undefined ? rc.c : rc.fc) + dc; pointRef({ r1: rc.r, c1: rc.c, r2: Math.max(0, rc.fr), c2: Math.max(0, rc.fc) }); } else { rc.r = Math.max(0, rc.r + dr); rc.c = Math.max(0, rc.c + dc); rc.fr = rc.fc = undefined; pointRef({ r1: rc.r, c1: rc.c, r2: rc.r, c2: rc.c }); } return; }
        e.preventDefault(); g.commit(dr, dc); return;
      }
    });
    function toggleAbs() { const t = ed.value, p = ed.selectionStart; const re = /\$?[A-Z]{1,3}\$?\d+/g; let m; while ((m = re.exec(t))) { if (p >= m.index && p <= m.index + m[0].length) { const mm = /^(\$?)([A-Z]+)(\$?)(\d+)$/.exec(m[0]); const state = (mm[1] ? 2 : 0) + (mm[3] ? 1 : 0); const nx = [[1, 1], [0, 1], [1, 0], [0, 0]]; const order = { 3: 1, 1: 2, 2: 3, 0: 0 }; const ns = order[state]; const [cA, rA] = ns === 0 ? [1, 1] : ns === 1 ? [0, 1] : ns === 2 ? [1, 0] : [0, 0]; const nr = (cA ? '$' : '') + mm[2] + (rA ? '$' : '') + mm[4]; ed.value = t.slice(0, m.index) + nr + t.slice(m.index + m[0].length); const np = m.index + nr.length; ed.setSelectionRange(np, np); onEdInput(); return; } } }

    /* ---------------- go to */
    g.gotoRef = function (text) {
      text = String(text).trim(); if (!text) return false; let q = null, sheet = null; const m = /^(?:'([^']+)'|([^!']+))!(.+)$/.exec(text); if (m) { sheet = m[1] || m[2]; text = m[3]; }
      try { q = XL.parseRange(text); } catch (e) { q = null; }
      if (!q) { const rf = wb.resolveName(text); if (rf) { q = rf; sheet = rf.sheet || sheet; } }
      if (!q) return false; if (sheet && sheet !== g.sh().name && wb.hasSheet(sheet)) hooks.activateSheet && hooks.activateSheet(sheet);
      const u = ui(); const n = norm(q); u.ranges = [n]; u.r = n.r1; u.c = n.c1; u.ar = n.r1; u.ac = n.c1; u.vr = Math.max(u.vr, Math.min(MAXR, n.r2 + 100)); u.vc = Math.max(u.vc, Math.min(MAXC, n.c2 + 10)); updateSpacer(); g.ensureVisible(n.r1, n.c1); notify(); return true;
    };

    /* ---------------- pointer handling */
    let drag = null, lastTap = { t: 0, r: -1, c: -1 }, dropEl = null;
    function pos(e) { const b = scroller.getBoundingClientRect(); return { x: e.clientX - b.left, y: e.clientY - b.top }; }
    function cellAtClamped(px, py) {
      const z = g.zoom, rh = RH() * z, ch = CH() * z; const h = hit(Math.max(rh + 1, px), Math.max(ch + 1, py)); if (h.t === 'cell') return h;
      const hh = hit(Math.max(rh + 1, px), Math.max(ch + 1, py)); return hh.t === 'cell' ? hh : { t: 'cell', r: 0, c: 0 };
    }
    function closeDrop() { if (dropEl) { dropEl.remove(); dropEl = null; } }
    function openDV(r, c) {
      closeDrop(); const sh = g.sh(); const rule = wb.dvAt(sh, r, c); if (!rule) return; const items = wb.dvList(sh, rule); const rc = mergedRect(r, c, view());
      dropEl = el('div', { class: 'x-ac' }); items.forEach(it => { const d = el('div', {}, [it]); d.addEventListener('mousedown', ev => { ev.preventDefault(); closeDrop(); wb.setInput(sh, r, c, it); layoutCharts(); notify(); }); dropEl.appendChild(d); });
      if (!items.length) dropEl.appendChild(el('div', {}, ['(list khali hai)'])); dropEl.style.left = rc.x + 'px'; dropEl.style.top = (rc.y + rc.h) + 'px'; dropEl.style.minWidth = rc.w + 'px'; host.appendChild(dropEl);
    }
    function openFilter(c) {
      const sh = g.sh(); const fr = sh.filter.range; const rc = mergedRect(fr.r1, c, view()); const a = el('div', { style: 'position:absolute;left:' + rc.x + 'px;top:' + rc.y + 'px;width:' + rc.w + 'px;height:' + rc.h + 'px;pointer-events:none' }); host.appendChild(a);
      hooks.filterMenu && hooks.filterMenu(c, a); setTimeout(() => a.remove(), 20000);
    }
    document.addEventListener('pointerdown', e => { if (dropEl && !dropEl.contains(e.target)) closeDrop(); if (acEl && !acEl.contains(e.target) && e.target !== ed) hideAC(); }, true);
    function startSelectDrag(h, e, p) {
      const shift = e.shiftKey, ctrl = e.ctrlKey || e.metaKey;
      if (shift) g.extendTo(h.r, h.c); else if (ctrl) g.addRange({ r1: h.r, c1: h.c, r2: h.r, c2: h.c }); else g.setActive(h.r, h.c);
      drag = { mode: 'select', id: e.pointerId, ax: h.r, ay: h.c }; if (!shift) { const u = ui(); u.ar = h.r; u.ac = h.c; }
    }
    scroller.addEventListener('pointerdown', e => {
      if (e.button === 2) return; const p = pos(e); if (p.x > scroller.clientWidth || p.y > scroller.clientHeight) return; closeDrop();
      const touch = e.pointerType === 'touch'; const h = hit(p.x, p.y);
      if (touch && !g.selMode) { drag = { mode: 'tap', x: p.x, y: p.y, t: Date.now(), h, moved: false }; return; }
      if (g._fh && !g.editing && h.t === 'cell' && p.x >= g._fh.x - 2 && p.x <= g._fh.x + g._fh.s + 3 && p.y >= g._fh.y - 2 && p.y <= g._fh.y + g._fh.s + 3) { scroller.setPointerCapture(e.pointerId); drag = { mode: 'fill', src: Object.assign({}, norm(ui().ranges[ui().ranges.length - 1])) }; e.preventDefault(); return; }
      if (h.t === 'cell') {
        if (g.editing) { if (canPoint() || (g.refIns && ed.value[0] === '=' && ed.selectionStart === g.refIns.end)) { e.preventDefault(); const rr = { r1: h.r, c1: h.c, r2: h.r, c2: h.c }; const m = wb.mergeAt(g.sh(), h.r, h.c); pointRef(m || rr); scroller.setPointerCapture(e.pointerId); drag = { mode: 'point', ar: (m || rr).r1, ac: (m || rr).c1 }; return; } if (!g.commit(0, 0, true)) { e.preventDefault(); return; } }
        // dropdown buttons
        const sh = g.sh(); const rc = mergedRect(h.r, h.c, view()); const inBtn = p.x > rc.x + rc.w - 19 * g.zoom && p.y > rc.y + rc.h - 19 * g.zoom;
        if (inBtn && sh.filter && h.r === sh.filter.range.r1 && h.c >= sh.filter.range.c1 && h.c <= sh.filter.range.c2) { e.preventDefault(); openFilter(h.c); return; }
        const u = ui(); if (inBtn && u.r === h.r && u.c === h.c) { const dv = wb.dvAt(sh, h.r, h.c); if (dv && dv.type === 'list') { e.preventDefault(); openDV(h.r, h.c); return; } }
        scroller.setPointerCapture(e.pointerId); startSelectDrag(h, e, p); hooks.onPointerCell && hooks.onPointerCell(h); return;
      }
      if (g.editing && !g.commit(0, 0, true)) return;
      if (h.t === 'corner') { g.selectAll(); return; }
      scroller.setPointerCapture(e.pointerId);
      if (h.t === 'colh') { if (e.shiftKey) { const u = ui(); g.selectCols(u.ac, h.c); } else if (e.ctrlKey || e.metaKey) { const u = ui(); u.ranges.push({ r1: 0, c1: h.c, r2: MAXR - 1, c2: h.c }); notify(); } else g.selectCols(h.c, h.c); drag = { mode: 'colh', a: ui().ac }; }
      else if (h.t === 'rowh') { if (e.shiftKey) { const u = ui(); g.selectRows(u.ar, h.r); } else if (e.ctrlKey || e.metaKey) { const u = ui(); u.ranges.push({ r1: h.r, c1: 0, r2: h.r, c2: MAXC - 1 }); notify(); } else g.selectRows(h.r, h.r); drag = { mode: 'rowh', a: ui().ar }; }
      else if (h.t === 'colresize') { const sh = g.sh(); drag = { mode: 'colresize', c: h.c, x0: p.x, w0: wb.colWidth(sh, h.c), old: sh.colW[h.c], sel: selectedCols(h.c) }; }
      else if (h.t === 'rowresize') { const sh = g.sh(); drag = { mode: 'rowresize', r: h.r, y0: p.y, h0: wb.rowHeight(sh, h.r), old: sh.rowH[h.r], sel: selectedRows(h.r) }; }
      e.preventDefault();
    });
    function selectedCols(c) { const u = ui(); const out = new Set([c]); u.ranges.forEach(q => { const n = norm(q); if (n.r1 === 0 && n.r2 >= MAXR - 1 && c >= n.c1 && c <= n.c2) for (let i = n.c1; i <= Math.min(n.c2, 2000); i++) out.add(i); }); return Array.from(out); }
    function selectedRows(r) { const u = ui(); const out = new Set([r]); u.ranges.forEach(q => { const n = norm(q); if (n.c1 === 0 && n.c2 >= MAXC - 1 && r >= n.r1 && r <= n.r2) for (let i = n.r1; i <= Math.min(n.r2, 5000); i++) out.add(i); }); return Array.from(out); }
    let lastPt = null, autoT = null;
    function autoScroll() {
      if (!drag || !lastPt || !(drag.mode === 'select' || drag.mode === 'point' || drag.mode === 'fill')) { clearInterval(autoT); autoT = null; return; }
      const sc = scroller, p = lastPt; let dx = 0, dy = 0; const m = 24; if (p.x > sc.clientWidth - m) dx = 24; else if (p.x < RH() * g.zoom + m) dx = -24; if (p.y > sc.clientHeight - m) dy = 24; else if (p.y < CH() * g.zoom + m) dy = -24;
      if (dx || dy) { sc.scrollLeft += dx; sc.scrollTop += dy; dragMove(p); }
    }
    function dragMove(p) {
      if (!drag) return; const z = g.zoom;
      if (drag.mode === 'select') { const h = cellAtClamped(p.x, p.y); g.extendTo(h.r, h.c); g.ensureVisible(h.r, h.c); }
      else if (drag.mode === 'point') { const h = cellAtClamped(p.x, p.y); pointRef({ r1: drag.ar, c1: drag.ac, r2: h.r, c2: h.c }); }
      else if (drag.mode === 'colh') { const h = hit(Math.max(RH() * z + 1, p.x), 1); if (h.c !== undefined && h.c >= 0) g.selectCols(drag.a, h.c); }
      else if (drag.mode === 'rowh') { const h = hit(1, Math.max(CH() * z + 1, p.y)); if (h.r !== undefined && h.r >= 0) g.selectRows(drag.a, h.r); }
      else if (drag.mode === 'colresize') { const sh = g.sh(); const w = Math.max(0, Math.round(drag.w0 + (p.x - drag.x0) / z)); drag.sel.forEach(c => { sh.colW[c] = w; sh.hiddenC.delete(c); }); drag.nw = w; wb.dirty(); updateSpacer(); layoutCharts(); g.draw(); hooks.onTip && hooks.onTip('Width: ' + w + ' px'); }
      else if (drag.mode === 'rowresize') { const sh = g.sh(); const hh = Math.max(0, Math.round(drag.h0 + (p.y - drag.y0) / z)); drag.sel.forEach(r => { sh.rowH[r] = hh; sh.hiddenR.delete(r); }); drag.nh = hh; wb.dirty(); updateSpacer(); layoutCharts(); g.draw(); }
      else if (drag.mode === 'fill') {
        const h = cellAtClamped(p.x, p.y), s = drag.src; let d = Object.assign({}, s); const dv = h.r - s.r2, dh = h.c - s.c2, up = s.r1 - h.r, lf = s.c1 - h.c;
        const vert = Math.max(dv, up) >= Math.max(dh, lf);
        if (vert) { if (dv > 0) d.r2 = h.r; else if (up > 0) d.r1 = h.r; } else { if (dh > 0) d.c2 = h.c; else if (lf > 0) d.c1 = h.c; }
        g.fillDrag = d; g.draw();
      }
    }
    scroller.addEventListener('pointermove', e => {
      const p = pos(e); lastPt = p;
      if (drag && drag.mode === 'tap') { if (Math.abs(p.x - drag.x) + Math.abs(p.y - drag.y) > 8) drag.moved = true; return; }
      if (drag) { e.preventDefault(); dragMove(p); if (!autoT && (drag.mode === 'select' || drag.mode === 'point' || drag.mode === 'fill')) autoT = setInterval(autoScroll, 60); return; }
      if (e.pointerType === 'touch') return; const h = hit(p.x, p.y); let cur = 'cell';
      if (h.t === 'colresize') cur = 'col-resize'; else if (h.t === 'rowresize') cur = 'row-resize'; else if (h.t === 'colh' || h.t === 'rowh' || h.t === 'corner') cur = 'default';
      else if (g._fh && p.x >= g._fh.x - 2 && p.x <= g._fh.x + g._fh.s + 3 && p.y >= g._fh.y - 2 && p.y <= g._fh.y + g._fh.s + 3) cur = 'crosshair';
      if (scroller.style.cursor !== cur) scroller.style.cursor = cur;
    });
    function endDrag(e) {
      if (!drag) return; const d = drag; drag = null; if (autoT) { clearInterval(autoT); autoT = null; } const sh = g.sh(); hooks.onTip && hooks.onTip(null);
      if (d.mode === 'tap') {
        if (d.moved || Date.now() - d.t > 600) return; const h = d.h;
        if (h.t === 'cell') {
          if (g.editing) { if (canPoint()) { pointRef({ r1: h.r, c1: h.c, r2: h.r, c2: h.c }); return; } if (!g.commit(0, 0, true)) return; }
          const u = ui(); const rc = mergedRect(h.r, h.c, view()); const inBtn = d.x > rc.x + rc.w - 22 * g.zoom && d.y > rc.y + rc.h - 22 * g.zoom;
          if (inBtn && sh.filter && h.r === sh.filter.range.r1 && h.c >= sh.filter.range.c1 && h.c <= sh.filter.range.c2) { openFilter(h.c); return; }
          if (inBtn && u.r === h.r && u.c === h.c) { const dv = wb.dvAt(sh, h.r, h.c); if (dv && dv.type === 'list') { openDV(h.r, h.c); return; } }
          const now = Date.now(); const dbl = now - lastTap.t < 380 && lastTap.r === h.r && lastTap.c === h.c; lastTap = { t: now, r: h.r, c: h.c };
          g.setActive(h.r, h.c); if (dbl) g.startEdit(h.r, h.c); return;
        }
        if (g.editing && !g.commit(0, 0, true)) return;
        if (h.t === 'corner') g.selectAll(); else if (h.t === 'colh' || h.t === 'colresize') g.selectCols(h.c, h.c); else if (h.t === 'rowh' || h.t === 'rowresize') g.selectRows(h.r, h.r);
        return;
      }
      if (d.mode === 'colresize' && d.nw !== undefined) { const nw = d.nw; d.sel.forEach(c => { if (d.old === undefined) delete sh.colW[c]; else sh.colW[c] = d.old; }); wb.tx(() => d.sel.forEach(c => wb.setProp(sh.colW, c, nw))); updateSpacer(); layoutCharts(); }
      if (d.mode === 'rowresize' && d.nh !== undefined) { const nh = d.nh; d.sel.forEach(r => { if (d.old === undefined) delete sh.rowH[r]; else sh.rowH[r] = d.old; }); wb.tx(() => d.sel.forEach(r => wb.setProp(sh.rowH, r, nh))); updateSpacer(); layoutCharts(); }
      if (d.mode === 'fill') { const dest = g.fillDrag; g.fillDrag = null; if (dest && (dest.r1 !== d.src.r1 || dest.r2 !== d.src.r2 || dest.c1 !== d.src.c1 || dest.c2 !== d.src.c2)) { try { wb.autoFill(sh, d.src, dest, { ctrl: e && (e.ctrlKey || e.metaKey) }); } catch (err) { OA.toast('Fill nahi ho paya'); } g.select(dest); layoutCharts(); } g.draw(); hooks.onChange && hooks.onChange(); }
      if (d.mode === 'select' || d.mode === 'point') { hooks.onPointerUp && hooks.onPointerUp(); }
    }
    scroller.addEventListener('pointerup', endDrag); scroller.addEventListener('pointercancel', e => { if (drag && drag.mode === 'tap') drag = null; else endDrag(e); });
    scroller.addEventListener('dblclick', e => {
      const p = pos(e); if (p.x > scroller.clientWidth || p.y > scroller.clientHeight) return; const h = hit(p.x, p.y); const sh = g.sh();
      if (h.t === 'colresize') { autofitCol(h.c); } else if (h.t === 'rowresize') { wb.tx(() => selectedRows(h.r).forEach(r => wb.setProp(sh.rowH, r, undefined))); updateSpacer(); g.draw(); }
      else if (h.t === 'cell') {
        if (g._fh && p.x >= g._fh.x - 2 && p.x <= g._fh.x + g._fh.s + 3 && p.y >= g._fh.y - 2 && p.y <= g._fh.y + g._fh.s + 3) { fillDownAuto(); return; }
        if (!g.editing) { const rc = mergedRect(h.r, h.c, view()); if (p.x > rc.x + rc.w - 19 * g.zoom && p.y > rc.y + rc.h - 19 * g.zoom && sh.filter) return; const u = ui(); g.startEdit(u.r, u.c); }
      }
    });
    function autofitCol(c) {
      const sh = g.sh(); const cols = selectedCols(c); const ext = wb.extent(sh.name); wb.ensureCalc();
      wb.tx(() => cols.forEach(cc => { let mx = 0; for (let r = 0; r <= Math.min(ext.r, 20000); r++) { const d = wb.display(sh, r, cc); if (!d.text) continue; const st = d.st; ctx.font = fontOf(st); const lines = d.text.split('\n'); lines.forEach(s => { mx = Math.max(mx, ctx.measureText(s).width + 10 + (st.ind || 0) * 9); }); } wb.setProp(sh.colW, cc, mx ? Math.max(20, Math.min(600, Math.ceil(mx))) : undefined); }));
      updateSpacer(); layoutCharts(); g.draw();
    }
    g.autofitCols = () => { const q = g.range(); const sh = g.sh(); const ext = wb.extent(sh.name); wb.ensureCalc(); wb.tx(() => { for (let cc = q.c1; cc <= Math.min(q.c2, MAXC - 1, q.c1 + 500); cc++) { let mx = 0; for (let r = 0; r <= Math.min(ext.r, 20000); r++) { const d = wb.display(sh, r, cc); if (!d.text) continue; ctx.font = fontOf(d.st); d.text.split('\n').forEach(s => mx = Math.max(mx, ctx.measureText(s).width + 10)); } wb.setProp(sh.colW, cc, mx ? Math.max(20, Math.min(600, Math.ceil(mx))) : undefined); } }); updateSpacer(); g.draw(); };
    g.autofitRows = () => { const q = g.range(); const sh = g.sh(); wb.tx(() => { for (let r = q.r1; r <= Math.min(q.r2, q.r1 + 5000); r++) wb.setProp(sh.rowH, r, undefined); }); updateSpacer(); g.draw(); };
    function fillDownAuto() {
      const sh = g.sh(), q = norm(ui().ranges[ui().ranges.length - 1]); let ref = q.c1 - 1; if (ref < 0 || !hasVal(sh, q.r2 + 1, ref)) ref = q.c2 + 1; if (ref >= MAXC) return; let r = q.r2; while (hasVal(sh, r + 1, ref)) r++; if (r <= q.r2) return;
      const dest = { r1: q.r1, c1: q.c1, r2: r, c2: q.c2 }; wb.autoFill(sh, q, dest); g.select(dest); layoutCharts(); hooks.onChange && hooks.onChange();
    }
    scroller.addEventListener('contextmenu', e => {
      e.preventDefault(); const p = pos(e); const h = hit(p.x, p.y); const u = ui();
      if (h.t === 'cell') { const inSel = u.ranges.some(q => { const n = norm(q); return h.r >= n.r1 && h.r <= n.r2 && h.c >= n.c1 && h.c <= n.c2; }); if (!inSel) g.setActive(h.r, h.c); }
      else if (h.t === 'colh') { if (!u.ranges.some(q => { const n = norm(q); return n.r1 === 0 && n.r2 >= MAXR - 1 && h.c >= n.c1 && h.c <= n.c2; })) g.selectCols(h.c, h.c); }
      else if (h.t === 'rowh') { if (!u.ranges.some(q => { const n = norm(q); return n.c1 === 0 && n.c2 >= MAXC - 1 && h.r >= n.r1 && h.r <= n.r2; })) g.selectRows(h.r, h.r); }
      hooks.contextMenu && hooks.contextMenu(e, h);
    });
    scroller.addEventListener('wheel', e => { if (e.ctrlKey) { e.preventDefault(); g.setZoom(g.zoom * (e.deltaY < 0 ? 1.1 : 0.9)); } }, { passive: false });
    scroller.addEventListener('scroll', () => {
      const u = ui(), z = g.zoom, l = layout(); let grow = false;
      if (scroller.scrollTop + scroller.clientHeight > scroller.scrollHeight - 250 && u.vr < MAXR) { u.vr = Math.min(MAXR, u.vr + 200); grow = true; }
      if (scroller.scrollLeft + scroller.clientWidth > scroller.scrollWidth - 250 && u.vc < MAXC) { u.vc = Math.min(MAXC, u.vc + 10); grow = true; }
      if (grow) updateSpacer(); layoutCharts(); placeEditor(); if (dropEl) closeDrop(); g.draw();
    });

    /* ---------------- clipboard */
    function clampRange(q) { const ext = wb.extent(g.sh().name); q = Object.assign({}, q); if (q.r2 >= MAXR - 1) q.r2 = Math.max(q.r1, ext.r); if (q.c2 >= MAXC - 1) q.c2 = Math.max(q.c1, ext.c); return q; }
    g.copy = function (cut, e) {
      const u = ui(); if (u.ranges.length > 1) OA.toast('Sirf last select ki hui range copy hogi'); const q = clampRange(g.range()); const sh = g.sh();
      const clip = wb.copyRange(sh, q, !!cut); g.clip = clip;
      if (e && e.clipboardData) { e.clipboardData.setData('text/plain', clip.text); try { e.clipboardData.setData('text/html', wb.rangeHTML(sh, q)); } catch (x) { } e.preventDefault(); }
      g.draw(); return clip;
    };
    g.copyButton = function (cut) { scroller.focus({ preventScroll: true }); let ok = false; try { ok = document.execCommand(cut ? 'cut' : 'copy'); } catch (e) { } if (!ok || !g.clip || (cut && !g.clip.cut)) { const c = g.copy(cut); try { navigator.clipboard && navigator.clipboard.writeText(c.text); } catch (e) { } } };
    const nz = s => String(s).replace(/\r/g, '').replace(/\n$/, '');
    g.pasteText = function (text, mode, opts) {
      const sh = g.sh(), u = ui(), q = norm(u.ranges[u.ranges.length - 1]); opts = opts || {};
      if (g.clip && nz(text) === nz(g.clip.text)) {
        const clip = g.clip; const R = clip.rows.length, C = clip.rows[0].length; const tr = !!opts.transpose; const oR = tr ? C : R, oC = tr ? R : C;
        if (clip.cut) { wb.pasteClip(sh, q.r1, q.c1, clip, mode || 'all', opts); g.clip = null; g.select({ r1: q.r1, c1: q.c1, r2: q.r1 + oR - 1, c2: q.c1 + oC - 1 }); }
        else if (R === 1 && C === 1 && (q.r2 > q.r1 || q.c2 > q.c1) && q.r2 - q.r1 < 5000 && q.c2 - q.c1 < 200) { wb.tx(() => { for (let r = q.r1; r <= q.r2; r++) for (let c = q.c1; c <= q.c2; c++) wb.pasteClip(sh, r, c, clip, mode || 'all', opts); }); }
        else { wb.pasteClip(sh, q.r1, q.c1, clip, mode || 'all', opts); g.select({ r1: q.r1, c1: q.c1, r2: Math.min(MAXR - 1, q.r1 + oR - 1), c2: Math.min(MAXC - 1, q.c1 + oC - 1) }); }
      } else {
        let rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t')); if (text.indexOf('\t') < 0 && rows.length === 1 && /,/.test(text) && false) rows = [text.split(',')];
        if (rows.length > 20000) { OA.toast('Paste bahut bada hai'); return; }
        wb.tx(() => { rows.forEach((row, i) => row.forEach((v, j) => { const r = q.r1 + i, c = q.c1 + j; if (r < MAXR && c < MAXC) wb.putInput(sh, r, c, (v.length > 1 && v[0] === '"' && v[v.length - 1] === '"') ? v.slice(1, -1).replace(/""/g, '"') : v); })); });
        const wid = Math.max(...rows.map(r => r.length)); g.select({ r1: q.r1, c1: q.c1, r2: Math.min(MAXR - 1, q.r1 + rows.length - 1), c2: Math.min(MAXC - 1, q.c1 + wid - 1) });
      }
      updateSpacer(); layoutCharts(); hooks.onChange && hooks.onChange(); g.draw();
    };
    scroller.addEventListener('copy', e => g.copy(false, e));
    scroller.addEventListener('cut', e => { g.copy(true, e); });
    scroller.addEventListener('paste', e => { e.preventDefault(); const t = e.clipboardData ? e.clipboardData.getData('text/plain') : ''; if (t === '' && g.clip) g.pasteText(g.clip.text); else g.pasteText(t); });
    g.pasteButton = async function (mode, opts) {
      if (g.clip && mode && mode !== 'all') { g.pasteText(g.clip.text, mode, opts); return; }
      try { const t = await navigator.clipboard.readText(); if (t || !g.clip) { g.pasteText(t || '', mode, opts); return; } } catch (e) { }
      if (g.clip) g.pasteText(g.clip.text, mode, opts); else OA.toast('Paste ke liye Ctrl+V dabayein (browser permission)');
    };

    /* ---------------- edit helpers */
    g.clearContents = function (what) { const sh = g.sh(); wb.tx(() => ui().ranges.forEach(q => wb.clearRange(sh, clampRange(norm(q)), what || 'contents'))); layoutCharts(); hooks.onChange && hooks.onChange(); notify(); };
    g.fill = function (dir) {
      const sh = g.sh(), q = norm(ui().ranges[ui().ranges.length - 1]);
      let src; if (dir === 'down') src = { r1: q.r1, c1: q.c1, r2: q.r1, c2: q.c2 }; else if (dir === 'right') src = { r1: q.r1, c1: q.c1, r2: q.r2, c2: q.c1 }; else if (dir === 'up') src = { r1: q.r2, c1: q.c1, r2: q.r2, c2: q.c2 }; else src = { r1: q.r1, c1: q.c2, r2: q.r2, c2: q.c2 };
      if ((dir === 'down' || dir === 'up') && q.r1 === q.r2) { if (dir === 'down') { src = { r1: q.r1 - 1, c1: q.c1, r2: q.r1 - 1, c2: q.c2 }; if (src.r1 < 0) return; } else return; }
      if ((dir === 'right' || dir === 'left') && q.c1 === q.c2) { if (dir === 'right') { src = { r1: q.r1, c1: q.c1 - 1, r2: q.r2, c2: q.c1 - 1 }; if (src.c1 < 0) return; } else return; }
      const dest = { r1: Math.min(src.r1, q.r1), c1: Math.min(src.c1, q.c1), r2: Math.max(src.r2, q.r2), c2: Math.max(src.c2, q.c2) }; wb.autoFill(sh, src, dest, { copy: true }); layoutCharts(); hooks.onChange && hooks.onChange(); g.draw();
    };
    g.insertText = function (text) { const u = ui(); wb.setInput(g.sh(), u.r, u.c, text); hooks.onChange && hooks.onChange(); notify(); };
    const pageRows = () => { const v = view(); return Math.max(1, v.rEnd - v.rStart - 1); };

    /* ---------------- keyboard */
    scroller.addEventListener('keydown', e => {
      if (g.editing) return; const k = e.key, ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey; const u = ui(), sh = g.sh();
      const arrow = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
      if (g.selChart && (k === 'Delete' || k === 'Backspace')) { e.preventDefault(); g.deleteChart(g.selChart); return; }
      if (arrow && !alt) { e.preventDefault(); g.moveActive(arrow[0], arrow[1], shift, ctrl); return; }
      if (ctrl) {
        const lk = k.toLowerCase();
        if (lk === 'a') { e.preventDefault(); g.selectAll(); return; }
        if (lk === 'z') { e.preventDefault(); wb.undo(); return; } if (lk === 'y') { e.preventDefault(); wb.redo(); return; }
        if (lk === 'c' || lk === 'x' || lk === 'v') return;
        if (lk === 'd') { e.preventDefault(); g.fill('down'); return; } if (lk === 'r') { e.preventDefault(); g.fill('right'); return; }
        if (lk === 'home') { e.preventDefault(); g.setActive(0, 0); return; }
        if (lk === 'end') { e.preventDefault(); const ext = wb.extent(sh.name); if (shift) g.extendTo(ext.r, ext.c); else g.setActive(ext.r, ext.c); return; }
        if (k === ' ') { e.preventDefault(); g.selectCols(u.c, shift ? u.c : u.c); if (shift) g.selectAll(); return; }
        if (k === ';') { e.preventDefault(); const d = new Date(); g.insertText((d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear()); return; }
        if (k === ':' ) { e.preventDefault(); const d = new Date(); g.insertText(d.getHours() + ':' + String(d.getMinutes()).padStart(2, '0')); return; }
        if (k === '+' || k === '=') { e.preventDefault(); hooks.cmd && hooks.cmd('insert'); return; } if (k === '-') { e.preventDefault(); hooks.cmd && hooks.cmd('delete'); return; }
        const map = { b: 'bold', i: 'italic', u: 'underline', '1': 'format', f: 'find', h: 'replace', g: 'goto', k: 'link', l: 'filter', t: 'table', p: 'print', s: 'save', o: 'open', n: 'new', '5': 'strike', '0': 'hidecol', '9': 'hiderow' };
        if (map[lk] && hooks.cmd) { e.preventDefault(); hooks.cmd(map[lk]); return; } return;
      }
      if (k === ' ' && shift) { e.preventDefault(); g.selectRows(u.r, u.r); return; }
      if (alt && k === '=') { e.preventDefault(); hooks.cmd && hooks.cmd('autosum'); return; }
      switch (k) {
        case 'Enter': e.preventDefault(); { const q = u.ranges[0]; const multi = u.ranges.length === 1 && (q.r1 !== q.r2 || q.c1 !== q.c2); if (multi) { stepInRange(shift ? -1 : 1, 0); notify(); } else g.moveActive(shift ? -1 : 1, 0); } return;
        case 'Tab': e.preventDefault(); { const q = u.ranges[0]; const multi = u.ranges.length === 1 && (q.r1 !== q.r2 || q.c1 !== q.c2); if (multi) { stepInRange(0, shift ? -1 : 1); notify(); } else g.moveActive(0, shift ? -1 : 1); } return;
        case 'Home': e.preventDefault(); if (shift) g.extendTo(u.r, 0); else g.setActive(u.r, 0); return;
        case 'PageDown': e.preventDefault(); g.moveActive(pageRows(), 0, shift); return; case 'PageUp': e.preventDefault(); g.moveActive(-pageRows(), 0, shift); return;
        case 'Delete': e.preventDefault(); g.clearContents(); return;
        case 'Backspace': e.preventDefault(); g.startEdit(u.r, u.c, ''); return;
        case 'F2': e.preventDefault(); g.startEdit(u.r, u.c, undefined, 'edit'); return;
        case 'F5': e.preventDefault(); hooks.cmd && hooks.cmd('goto'); return;
        case 'F11': if (shift) { e.preventDefault(); hooks.cmd && hooks.cmd('newsheet'); } return;
        case 'Escape': if (g.clip) { g.clip = null; g.draw(); } if (g.selChart) { g.selChart = null; layoutCharts(); } return;
        case 'ContextMenu': e.preventDefault(); { const rc = mergedRect(u.r, u.c, view()); const b = host.getBoundingClientRect(); hooks.contextMenu && hooks.contextMenu({ clientX: b.left + rc.x + 20, clientY: b.top + rc.y + rc.h, preventDefault() { } }, { t: 'cell', r: u.r, c: u.c }); } return;
      }
      if (k.length === 1 && !alt) { e.preventDefault(); g.startEdit(u.r, u.c, k, 'enter'); }
    });
    // mobile soft keyboards / IME may deliver text via beforeinput on the focused scroller
    scroller.addEventListener('beforeinput', e => { if (!g.editing && e.data) { e.preventDefault(); const u = ui(); g.startEdit(u.r, u.c, e.data, 'enter'); } });

    /* ---------------- charts layer */
    const layer = el('div', { class: 'x-chartlayer' }); layer.style.cssText = 'position:absolute;overflow:hidden;pointer-events:none;z-index:3;right:0;bottom:0'; host.appendChild(layer);
    layer.addEventListener('wheel', e => { if (e.ctrlKey) { e.preventDefault(); g.setZoom(g.zoom * (e.deltaY < 0 ? 1.1 : 0.9)); } else { scroller.scrollTop += e.deltaY; scroller.scrollLeft += e.deltaX; } }, { passive: false });
    let chartKey = '', chartEls = new Map();
    function layoutCharts() {
      const sh = g.sh(), z = g.zoom; layer.style.left = (RH() * z) + 'px'; layer.style.top = (CH() * z) + 'px'; const sx = scroller.scrollLeft / z, sy = scroller.scrollTop / z;
      const k = wb.version + ':' + sh.name + ':' + sh.charts.length + ':' + z + ':' + (g.showHeaders ? 1 : 0);
      if (k !== chartKey) {
        chartKey = k; layer.innerHTML = ''; chartEls = new Map();
        sh.charts.forEach(ch => {
          const d = el('div', { class: 'x-chart' }); d.style.pointerEvents = 'auto'; d.dataset.id = ch.id; const rs = el('div', { class: 'rs' }); d.appendChild(rs);
          try { const holder = el('div'); holder.style.cssText = 'position:absolute;inset:0'; holder.innerHTML = XL.chartSVG(ch.obj ? null : XL.chartData(wb, ch), ch, Math.round(ch.w * z), Math.round(ch.h * z)); d.insertBefore(holder, rs); } catch (err) { d.appendChild(el('div', { style: 'padding:8px;color:#a00;font-size:12px' }, ['Chart data error'])); }
          layer.appendChild(d); chartEls.set(ch.id, d); bindChart(d, ch, rs);
        });
      }
      sh.charts.forEach(ch => { const d = chartEls.get(ch.id); if (!d) return; d.style.left = ((ch.x - sx) * z) + 'px'; d.style.top = ((ch.y - sy) * z) + 'px'; d.style.width = (ch.w * z) + 'px'; d.style.height = (ch.h * z) + 'px'; d.classList.toggle('sel', g.selChart === ch.id); });
    }
    g.layoutCharts = layoutCharts;
    function selectChart(id) { const prev = g.selChart; g.selChart = id; layoutCharts(); if (prev !== id) hooks.onChartSel && hooks.onChartSel(id ? g.sh().charts.find(c => c.id === id) : null); g.draw(); }
    g.selectChart = selectChart;
    function bindChart(d, ch, rs) {
      let st = null;
      d.addEventListener('pointerdown', e => { if (e.button === 2) { selectChart(ch.id); return; } e.preventDefault(); e.stopPropagation(); if (g.editing && !g.commit(0, 0, true)) return; selectChart(ch.id); d.setPointerCapture(e.pointerId); st = { x: e.clientX, y: e.clientY, ox: ch.x, oy: ch.y, ow: ch.w, oh: ch.h, rs: e.target === rs, moved: false }; scroller.blur && 0; });
      d.addEventListener('pointermove', e => { if (!st) return; const dx = (e.clientX - st.x) / g.zoom, dy = (e.clientY - st.y) / g.zoom; if (Math.abs(dx) + Math.abs(dy) > 2) st.moved = true; if (!st.moved) return; if (st.rs) { ch.w = Math.max(120, Math.round(st.ow + dx)); ch.h = Math.max(80, Math.round(st.oh + dy)); } else { ch.x = Math.max(0, Math.round(st.ox + dx)); ch.y = Math.max(0, Math.round(st.oy + dy)); } const z = g.zoom, sx = scroller.scrollLeft / z, sy = scroller.scrollTop / z; d.style.left = ((ch.x - sx) * z) + 'px'; d.style.top = ((ch.y - sy) * z) + 'px'; d.style.width = (ch.w * z) + 'px'; d.style.height = (ch.h * z) + 'px'; });
      const up = () => { if (!st) return; const s = st; st = null; if (!s.moved) { scroller.focus({ preventScroll: true }); return; } const nx = ch.x, ny = ch.y, nw = ch.w, nh = ch.h; ch.x = s.ox; ch.y = s.oy; ch.w = s.ow; ch.h = s.oh; wb.tx(() => { wb.setProp(ch, 'x', nx); wb.setProp(ch, 'y', ny); wb.setProp(ch, 'w', nw); wb.setProp(ch, 'h', nh); }); hooks.onChange && hooks.onChange(); layoutCharts(); scroller.focus({ preventScroll: true }); };
      d.addEventListener('pointerup', up); d.addEventListener('pointercancel', up);
      d.addEventListener('dblclick', () => hooks.chartDialog && hooks.chartDialog(ch));
      d.addEventListener('contextmenu', e => { e.preventDefault(); selectChart(ch.id); hooks.chartMenu && hooks.chartMenu(e, ch); });
    }
    g.addChart = function (ch) { const sh = g.sh(); wb.tx(() => wb.setProp(sh, 'charts', sh.charts.concat([ch]))); g.selChart = null; layoutCharts(); selectChart(ch.id); hooks.onChange && hooks.onChange(); };
    g.deleteChart = function (id) { const sh = g.sh(); wb.tx(() => wb.setProp(sh, 'charts', sh.charts.filter(c => c.id !== id))); g.selChart = null; layoutCharts(); hooks.onChartSel && hooks.onChartSel(null); hooks.onChange && hooks.onChange(); g.draw(); };
    g.chart = () => g.selChart ? g.sh().charts.find(c => c.id === g.selChart) : null;
    g.updateChart = function (ch, patch) { wb.tx(() => Object.keys(patch).forEach(k => wb.setProp(ch, k, patch[k]))); layoutCharts(); hooks.onChange && hooks.onChange(); };

    /* ---------------- misc API */
    g.setSelMode = on => { g.selMode = !!on; scroller.classList.toggle('selmode', !!on); };
    g.refresh = function () { L = null; updateSpacer(); layoutCharts(); if (g.editing) placeEditor(); g.draw(); };
    g.focus = () => scroller.focus({ preventScroll: true });
    g.sheetChanged = function () { L = null; chartKey = ''; g.selChart = null; g.clip = g.clip; const u = ui(); updateSpacer(); scroller.scrollLeft = (u.sx || 0) * g.zoom; scroller.scrollTop = (u.sy || 0) * g.zoom; layoutCharts(); g.draw(); hooks.onSel && hooks.onSel(); };
    g.saveScroll = function () { const u = ui(); u.sx = scroller.scrollLeft / g.zoom; u.sy = scroller.scrollTop / g.zoom; };
    g.measure = (st, text) => { ctx.font = fontOf(st || {}); return ctx.measureText(text).width; };
    g.destroy = function () { if (ro) ro.disconnect(); if (g.antT) clearInterval(g.antT); };
    let ro = null; if (typeof ResizeObserver !== 'undefined') { ro = new ResizeObserver(() => { if (host.clientWidth > 0 && host.clientHeight > 0) resize(); }); ro.observe(host); } window.addEventListener('resize', () => { if (host.clientWidth > 0) resize(); });
    resize(); layoutCharts();
    return g;
  };
})(window);
