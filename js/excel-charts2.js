/* Extended chart types for Excel: combo, secondary axis, radar, bubble, histogram, waterfall, funnel, box & whisker, stock,
   trendlines, error bars, data labels, axis titles — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL, orig = XL.chartSVG, PAL = XL.CHART_PAL;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = v => { const a = Math.abs(v); if (a >= 1e9) return +(v / 1e9).toFixed(2) + 'B'; if (a >= 1e6) return +(v / 1e6).toFixed(2) + 'M'; if (a >= 1e4) return +(v / 1e3).toFixed(1) + 'K'; return String(+v.toFixed(a < 1 ? 3 : 2)); };
  function ticksOf(lo, hi, n) { if (hi === lo) hi = lo + 1; const raw = (hi - lo) / n, mag = Math.pow(10, Math.floor(Math.log10(raw))), nm = raw / mag; const step = (nm < 1.5 ? 1 : nm < 3 ? 2 : nm < 7 ? 5 : 10) * mag; const s = Math.floor(lo / step) * step, e = Math.ceil(hi / step) * step; const t = []; for (let v = s; v <= e + step * 1e-6; v += step) t.push(+v.toFixed(10)); return t; }
  const NEW = ['combo', 'radar', 'bubble', 'histogram', 'waterfall', 'funnel', 'box', 'stock'];
  XL.CHART_TYPES_EXTRA = [['combo', 'Combo (Column + Line)'], ['radar', 'Radar'], ['bubble', 'Bubble'], ['histogram', 'Histogram'], ['waterfall', 'Waterfall'], ['funnel', 'Funnel'], ['box', 'Box & Whisker'], ['stock', 'Stock (Open-High-Low-Close)']];

  /* ---- trend helpers */
  function solve(A, b) { const n = b.length; for (let i = 0; i < n; i++) { let p = i; for (let k = i + 1; k < n; k++) if (Math.abs(A[k][i]) > Math.abs(A[p][i])) p = k;[A[i], A[p]] = [A[p], A[i]];[b[i], b[p]] = [b[p], b[i]]; for (let k = i + 1; k < n; k++) { const f = A[k][i] / (A[i][i] || 1e-12); for (let j = i; j < n; j++) A[k][j] -= f * A[i][j]; b[k] -= f * b[i]; } } const x = new Array(n).fill(0); for (let i = n - 1; i >= 0; i--) { let s = b[i]; for (let j = i + 1; j < n; j++) s -= A[i][j] * x[j]; x[i] = s / (A[i][i] || 1e-12); } return x; }
  function polyFit(xs, ys, deg) { const m = deg + 1; const A = Array.from({ length: m }, (_, i) => Array.from({ length: m }, (_, j) => xs.reduce((s, x) => s + Math.pow(x, i + j), 0))); const b = Array.from({ length: m }, (_, i) => xs.reduce((s, x, k) => s + ys[k] * Math.pow(x, i), 0)); return solve(A, b); }
  function trendFn(kind, xs, ys, period) {
    if (kind === 'linear') { const c = polyFit(xs, ys, 1); return { f: x => c[0] + c[1] * x, eq: `y = ${fmt(c[1])}x ${c[0] < 0 ? '−' : '+'} ${fmt(Math.abs(c[0]))}`, r2: r2(xs, ys, x => c[0] + c[1] * x) }; }
    if (kind === 'poly') { const c = polyFit(xs, ys, 2); return { f: x => c[0] + c[1] * x + c[2] * x * x, eq: `y = ${fmt(c[2])}x² ${c[1] < 0 ? '−' : '+'} ${fmt(Math.abs(c[1]))}x ${c[0] < 0 ? '−' : '+'} ${fmt(Math.abs(c[0]))}` }; }
    if (kind === 'exp') { const pos = ys.every(y => y > 0); if (!pos) return null; const c = polyFit(xs, ys.map(Math.log), 1); const a = Math.exp(c[0]), b = c[1]; return { f: x => a * Math.exp(b * x), eq: `y = ${fmt(a)}e^${fmt(b)}x` }; }
    if (kind === 'log') { if (xs.some(x => x <= 0)) return null; const c = polyFit(xs.map(Math.log), ys, 1); return { f: x => c[0] + c[1] * Math.log(x), eq: `y = ${fmt(c[1])}ln(x) ${c[0] < 0 ? '−' : '+'} ${fmt(Math.abs(c[0]))}` }; }
    if (kind === 'power') { if (xs.some(x => x <= 0) || ys.some(y => y <= 0)) return null; const c = polyFit(xs.map(Math.log), ys.map(Math.log), 1); const a = Math.exp(c[0]); return { f: x => a * Math.pow(x, c[1]), eq: `y = ${fmt(a)}x^${fmt(c[1])}` }; }
    if (kind === 'avg') { const p = Math.max(2, period || 2); return { pts: xs.map((x, i) => i < p - 1 ? null : [x, ys.slice(i - p + 1, i + 1).reduce((a, b) => a + b, 0) / p]).filter(Boolean), eq: `${p}-period moving average` }; }
    return null;
  }
  function r2(xs, ys, f) { const m = ys.reduce((a, b) => a + b, 0) / ys.length; const st = ys.reduce((s, y) => s + (y - m) ** 2, 0), sr = ys.reduce((s, y, i) => s + (y - f(xs[i])) ** 2, 0); return st ? 1 - sr / st : 1; }

  XL.chartSVG = function (d, ch, W, H) {
    if (ch.obj && XL.objSVG) return XL.objSVG(ch, W, H);
    const t = ch.type, pie = t === 'pie' || t === 'doughnut';
    const extra = ch.trendline || ch.errBars || ch.dataLabels || ch.xTitle || ch.yTitle || ch.y2Title || (ch.y2 && ch.y2.length) || (ch.seriesTypes && Object.keys(ch.seriesTypes).length) || ch.gridlines === false;
    if (pie || (!NEW.includes(t) && !extra)) return orig(d, ch, W, H);
    try { return render(d, ch, W, H); } catch (e) { return orig(d, Object.assign({}, ch, { type: 'col' }), W, H); }
  };

  function render(d, ch, W, H) {
    const t = ch.type, S = d.series.map(s => ({ name: s.name, vals: s.vals.map(v => (typeof v === 'number' && isFinite(v)) ? v : 0) })), cats = d.cats.slice();
    const font = 'font-family="Segoe UI,Calibri,Arial,sans-serif"'; let o = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" ${font} font-size="11"><rect width="${W}" height="${H}" fill="#fff"/>`;
    if (!S.length || !cats.length) return o + `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" fill="#888">No data to plot – check the data range</text></svg>`;
    const titleH = ch.title ? 34 : 10, legendH = ch.legend === false ? 8 : 28; if (ch.title) o += `<text x="${W / 2}" y="22" text-anchor="middle" font-size="15" fill="#404040">${esc(ch.title)}</text>`;
    const legend = items => { if (ch.legend === false) return ''; let s = ''; const tw = items.reduce((a, it) => a + 26 + Math.min(110, it.n.length * 6.2), 0); let x = Math.max(8, (W - tw) / 2); items.forEach(it => { s += `<rect x="${x}" y="${H - 20}" width="10" height="10" fill="${it.c}"/><text x="${x + 14}" y="${H - 11}" fill="#595959">${esc(it.n.slice(0, 18))}</text>`; x += 26 + Math.min(110, it.n.length * 6.2); }); return s; };
    const dl = !!ch.dataLabels;

    /* ===== radar */
    if (t === 'radar') {
      const cx = W / 2, cy = titleH + (H - titleH - legendH) / 2, R = Math.min(W / 2 - 50, (H - titleH - legendH) / 2 - 22), n = cats.length; let mx = 0; S.forEach(s => s.vals.forEach(v => mx = Math.max(mx, v))); const tk = ticksOf(0, mx, 4), tmax = tk[tk.length - 1] || 1;
      tk.forEach(v => { const r = R * v / tmax; o += `<polygon points="${cats.map((_, i) => { const a = i / n * Math.PI * 2 - Math.PI / 2; return (cx + r * Math.cos(a)) + ',' + (cy + r * Math.sin(a)); }).join(' ')}" fill="none" stroke="#d9d9d9"/>`; if (v) o += `<text x="${cx + 3}" y="${cy - r + 11}" fill="#808080" font-size="9">${fmt(v)}</text>`; });
      cats.forEach((c, i) => { const a = i / n * Math.PI * 2 - Math.PI / 2; o += `<line x1="${cx}" y1="${cy}" x2="${cx + R * Math.cos(a)}" y2="${cy + R * Math.sin(a)}" stroke="#d9d9d9"/><text x="${cx + (R + 12) * Math.cos(a)}" y="${cy + (R + 12) * Math.sin(a) + 4}" text-anchor="${Math.cos(a) > .3 ? 'start' : Math.cos(a) < -.3 ? 'end' : 'middle'}" fill="#595959">${esc(String(c).slice(0, 12))}</text>`; });
      S.forEach((s, si) => { const pts = s.vals.map((v, i) => { const a = i / n * Math.PI * 2 - Math.PI / 2, r = R * v / tmax; return [cx + r * Math.cos(a), cy + r * Math.sin(a)]; }); o += `<polygon points="${pts.map(p => p.join(',')).join(' ')}" fill="${PAL[si % 10]}" fill-opacity=".18" stroke="${PAL[si % 10]}" stroke-width="2"/>` + pts.map((p, i) => `<circle cx="${p[0]}" cy="${p[1]}" r="3" fill="${PAL[si % 10]}"><title>${esc(s.name)} – ${esc(cats[i])}: ${s.vals[i]}</title></circle>`).join(''); });
      return o + legend(S.map((s, i) => ({ n: s.name, c: PAL[i % 10] }))) + '</svg>';
    }

    /* ===== data prep for cartesian specials */
    let kinds = S.map((s, i) => (ch.seriesTypes && ch.seriesTypes[i]) || (t === 'combo' ? (i === 0 ? 'col' : 'line') : (t === 'bar' ? 'bar' : t === 'col' ? 'col' : t === 'line' ? 'line' : t === 'area' ? 'area' : t === 'scatter' ? 'scatter' : 'col')));
    let sCats = cats, bars = null, special = null;
    if (t === 'histogram') {
      const v = S[0].vals.slice().filter(x => isFinite(x)); const lo = Math.min(...v), hi = Math.max(...v); const nb = Math.max(1, Math.round(ch.bins || Math.ceil(Math.sqrt(v.length)))); const w = (hi - lo) / nb || 1; const cnt = new Array(nb).fill(0); v.forEach(x => { cnt[Math.min(nb - 1, Math.floor((x - lo) / w))]++; });
      sCats = cnt.map((_, i) => `${fmt(lo + i * w)}–${fmt(lo + (i + 1) * w)}`); S.length = 0; S.push({ name: 'Frequency', vals: cnt }); kinds = ['col']; special = 'hist';
    } else if (t === 'waterfall') {
      const v = S[0].vals; let run = 0; const segs = v.map((x, i) => { const isTot = (ch.totals ? ch.totals.includes(i) : (/total|net|sum/i.test(String(cats[i])))); if (isTot) { const s = { a: 0, b: x === 0 ? run : x, tot: true }; run = s.b; return s; } const s = { a: run, b: run + x, tot: false }; run += x; return s; }); bars = segs; special = 'wf';
    } else if (t === 'funnel') { special = 'funnel'; }
    else if (t === 'box') { special = 'box'; }
    else if (t === 'stock') { special = 'stock'; }
    else if (t === 'bubble') { special = 'bubble'; }

    /* ===== layout frame */
    const lbl = c => String(c);
    let lo = 0, hi = 0; const y2s = new Set(ch.y2 || []); const prim = [], sec = [];
    S.forEach((s, i) => (y2s.has(i) ? sec : prim).push(i));
    const range = idxs => { let l = 0, h = 0; idxs.forEach(i => S[i].vals.forEach(v => { l = Math.min(l, v); h = Math.max(h, v); })); return [l, h]; };
    let [l1, h1] = range(prim.length ? prim : S.map((_, i) => i)); let l2 = 0, h2 = 0; if (sec.length) [l2, h2] = range(sec);
    if (special === 'wf') { l1 = 0; h1 = 0; bars.forEach(b => { l1 = Math.min(l1, b.a, b.b); h1 = Math.max(h1, b.a, b.b); }); }
    if (special === 'box') { l1 = Infinity; h1 = -Infinity; S.forEach(s => s.vals.forEach(v => { l1 = Math.min(l1, v); h1 = Math.max(h1, v); })); const pad = (h1 - l1) * .05 || 1; l1 -= pad; h1 += pad; }
    if (special === 'stock') { l1 = Infinity; h1 = -Infinity; S.slice(0, 4).forEach(s => s.vals.forEach(v => { l1 = Math.min(l1, v); h1 = Math.max(h1, v); })); const pad = (h1 - l1) * .08 || 1; l1 -= pad; h1 += pad; }
    if (special === 'bubble') { const ys = S.filter((_, i) => i % 2 === 0); l1 = Math.min(0, ...ys.flatMap(s => s.vals)); h1 = Math.max(...ys.flatMap(s => s.vals)); }
    if (ch.errBars && !special) { const ev = errAmt(S[0], ch); S.forEach((s, si) => s.vals.forEach((v, i) => { const e = ev(v); (y2s.has(si) ? (h2 = Math.max(h2, v + e), l2 = Math.min(l2, v - e)) : (h1 = Math.max(h1, v + e), l1 = Math.min(l1, v - e))); })); }
    const tk1 = ticksOf(l1, h1, 5), tk2 = sec.length ? ticksOf(l2, h2, 5) : null; const stacked = !!ch.stacked && kinds.every(k => k === 'col' || k === 'bar' || k === 'area');
    if (stacked) { let p = 0, ng = 0; cats.forEach((_, i) => { let a = 0, b = 0; prim.forEach(si => { const v = S[si].vals[i] || 0; v >= 0 ? a += v : b += v; }); p = Math.max(p, a); ng = Math.min(ng, b); }); tk1.length = 0; ticksOf(ng, p, 5).forEach(v => tk1.push(v)); }
    const horiz = t === 'bar' || special === 'funnel' && false; const lblW = Math.max(...tk1.map(v => fmt(v).length)) * 6.3 + 12 + (ch.yTitle ? 16 : 0), lblW2 = tk2 ? Math.max(...tk2.map(v => fmt(v).length)) * 6.3 + 12 + (ch.y2Title ? 16 : 0) : 8;
    const xTitleH = ch.xTitle ? 16 : 0; let L = lblW + 6, Rr = W - lblW2 - 8, T = titleH, B = H - legendH - 22 - xTitleH; if (horiz) { L = Math.max(...cats.map(c => Math.min(16, lbl(c).length))) * 6.2 + 14; B = H - legendH - 20 - xTitleH; Rr = W - 16; }
    if (special === 'funnel') { L = 90; Rr = W - 16; B = H - legendH - 10; }
    const pw = Rr - L, ph = B - T; if (pw < 20 || ph < 20) return o + '</svg>';
    const tmin1 = tk1[0], tmax1 = tk1[tk1.length - 1]; const yv1 = v => horiz ? L + (v - tmin1) / (tmax1 - tmin1) * pw : B - (v - tmin1) / (tmax1 - tmin1) * ph; const yv2 = tk2 ? (v => B - (v - tk2[0]) / (tk2[tk2.length - 1] - tk2[0]) * ph) : yv1;
    const grid = ch.gridlines !== false;

    /* ---- funnel */
    if (special === 'funnel') {
      const v = S[0].vals; const mx = Math.max(...v) || 1; const n = v.length, bh = ph / n; const cxm = L + pw / 2;
      v.forEach((x, i) => { const w = pw * x / mx, y = T + i * bh; o += `<rect x="${cxm - w / 2}" y="${y + 2}" width="${Math.max(1, w)}" height="${bh - 4}" fill="${PAL[i % 10]}"><title>${esc(cats[i])}: ${x}</title></rect><text x="${L - 6}" y="${y + bh / 2 + 4}" text-anchor="end" fill="#595959">${esc(String(cats[i]).slice(0, 14))}</text><text x="${cxm}" y="${y + bh / 2 + 4}" text-anchor="middle" fill="#fff" font-weight="600">${fmt(x)}</text>`; });
      return o + '</svg>';
    }
    /* ---- gridlines + axes */
    tk1.forEach(v => { const p = yv1(v); if (horiz) o += (grid ? `<line x1="${p}" y1="${T}" x2="${p}" y2="${B}" stroke="#d9d9d9"/>` : '') + `<text x="${p}" y="${B + 14}" text-anchor="middle" fill="#595959">${fmt(v)}</text>`; else o += (grid ? `<line x1="${L}" y1="${p}" x2="${Rr}" y2="${p}" stroke="#d9d9d9"/>` : '') + `<text x="${L - 6}" y="${p + 4}" text-anchor="end" fill="#595959">${fmt(v)}</text>`; });
    if (tk2) tk2.forEach(v => { const p = yv2(v); o += `<text x="${Rr + 6}" y="${p + 4}" fill="#ED7D31">${fmt(v)}</text>`; });
    if (ch.xTitle) o += `<text x="${L + pw / 2}" y="${B + 22 + 12}" text-anchor="middle" fill="#595959" font-weight="600">${esc(ch.xTitle)}</text>`;
    if (ch.yTitle) o += `<text transform="translate(12 ${T + ph / 2}) rotate(-90)" text-anchor="middle" fill="#595959" font-weight="600">${esc(ch.yTitle)}</text>`;
    if (ch.y2Title && tk2) o += `<text transform="translate(${W - 10} ${T + ph / 2}) rotate(90)" text-anchor="middle" fill="#ED7D31" font-weight="600">${esc(ch.y2Title)}</text>`;
    const nCat = special === 'box' ? S.length : special === 'hist' || special === 'wf' ? sCats.length : cats.length; const gw = (horiz ? ph : pw) / (nCat || 1);
    const catX = i => L + gw * i + gw / 2;
    const catName = i => special === 'box' ? S[i].name : lbl(sCats[i]);
    for (let i = 0; i < nCat; i++) { const c = catName(i); const lab = c.length > 12 && !horiz ? c.slice(0, 11) + '…' : c; if (horiz) o += `<text x="${L - 6}" y="${T + gw * (nCat - 1 - i) + gw / 2 + 4}" text-anchor="end" fill="#595959">${esc(lab.slice(0, 16))}</text>`; else { const x = catX(i); o += gw < 46 && nCat > 8 ? `<text transform="translate(${x} ${B + 14}) rotate(-35)" text-anchor="end" fill="#595959">${esc(lab)}</text>` : `<text x="${x}" y="${B + 15}" text-anchor="middle" fill="#595959">${esc(lab)}</text>`; } }
    const zero = horiz ? yv1(Math.max(tmin1, Math.min(0, tmax1))) : yv1(Math.max(tmin1, Math.min(0, tmax1)));
    o += horiz ? `<line x1="${zero}" y1="${T}" x2="${zero}" y2="${B}" stroke="#bfbfbf"/>` : `<line x1="${L}" y1="${zero}" x2="${Rr}" y2="${zero}" stroke="#bfbfbf"/>`;

    /* ---- special renderers */
    if (special === 'wf') {
      const bw = Math.min(60, gw * .7);
      bars.forEach((b, i) => { const up = b.b >= b.a; const col = b.tot ? '#5B9BD5' : up ? '#70AD47' : '#C0504D'; const y1 = yv1(Math.max(b.a, b.b)), y2 = yv1(Math.min(b.a, b.b)); const x = L + gw * i + (gw - bw) / 2; o += `<rect x="${x}" y="${y1}" width="${bw}" height="${Math.max(1, y2 - y1)}" fill="${col}"><title>${esc(cats[i])}: ${b.tot ? b.b : (b.b - b.a)}</title></rect>`; if (i < bars.length - 1) o += `<line x1="${x + bw}" y1="${yv1(b.b)}" x2="${L + gw * (i + 1) + (gw - bw) / 2}" y2="${yv1(b.b)}" stroke="#7f7f7f" stroke-dasharray="3 2"/>`; if (dl) o += `<text x="${x + bw / 2}" y="${y1 - 4}" text-anchor="middle" fill="#404040">${b.tot ? fmt(b.b) : (b.b - b.a >= 0 ? '+' : '') + fmt(b.b - b.a)}</text>`; });
      return o + legend([{ n: 'Increase', c: '#70AD47' }, { n: 'Decrease', c: '#C0504D' }, { n: 'Total', c: '#5B9BD5' }]) + '</svg>';
    }
    if (special === 'box') {
      const q = (a, p) => { const s = a.slice().sort((x, y) => x - y); const r = (s.length - 1) * p, i = Math.floor(r); return s[i] + (s[Math.min(i + 1, s.length - 1)] - s[i]) * (r - i); };
      S.forEach((s, i) => { const v = s.vals; const q1 = q(v, .25), md = q(v, .5), q3 = q(v, .75), iqr = q3 - q1; const lw = Math.min(...v.filter(x => x >= q1 - 1.5 * iqr)), hw = Math.max(...v.filter(x => x <= q3 + 1.5 * iqr)); const x = catX(i), bw = Math.min(70, gw * .55), col = PAL[i % 10]; const mean = v.reduce((a, b) => a + b, 0) / v.length;
        o += `<line x1="${x}" y1="${yv1(hw)}" x2="${x}" y2="${yv1(lw)}" stroke="#404040"/><line x1="${x - bw / 4}" y1="${yv1(hw)}" x2="${x + bw / 4}" y2="${yv1(hw)}" stroke="#404040"/><line x1="${x - bw / 4}" y1="${yv1(lw)}" x2="${x + bw / 4}" y2="${yv1(lw)}" stroke="#404040"/><rect x="${x - bw / 2}" y="${yv1(q3)}" width="${bw}" height="${Math.max(1, yv1(q1) - yv1(q3))}" fill="${col}" fill-opacity=".55" stroke="#404040"><title>${esc(s.name)}: Q1 ${fmt(q1)}, median ${fmt(md)}, Q3 ${fmt(q3)}, min ${fmt(lw)}, max ${fmt(hw)}</title></rect><line x1="${x - bw / 2}" y1="${yv1(md)}" x2="${x + bw / 2}" y2="${yv1(md)}" stroke="#404040" stroke-width="2"/><text x="${x}" y="${yv1(mean) + 4}" text-anchor="middle" fill="#222" font-size="12">×</text>`;
        v.filter(x2 => x2 < lw || x2 > hw).forEach(out => o += `<circle cx="${x}" cy="${yv1(out)}" r="3" fill="none" stroke="#404040"/>`); });
      return o + legend(S.map((s, i) => ({ n: s.name, c: PAL[i % 10] }))) + '</svg>';
    }
    if (special === 'stock') {
      const [Op, Hi, Lo, Cl] = S; const n = cats.length, bw = Math.min(24, gw * .6); for (let i = 0; i < n; i++) { const up = Cl.vals[i] >= Op.vals[i], col = up ? '#2E8B57' : '#C0392B', x = catX(i); o += `<line x1="${x}" y1="${yv1(Hi.vals[i])}" x2="${x}" y2="${yv1(Lo.vals[i])}" stroke="${col}" stroke-width="1.5"/><rect x="${x - bw / 2}" y="${yv1(Math.max(Op.vals[i], Cl.vals[i]))}" width="${bw}" height="${Math.max(1, Math.abs(yv1(Op.vals[i]) - yv1(Cl.vals[i])))}" fill="${up ? '#fff' : col}" stroke="${col}"><title>${esc(cats[i])} O ${Op.vals[i]} H ${Hi.vals[i]} L ${Lo.vals[i]} C ${Cl.vals[i]}</title></rect>`; }
      return o + legend([{ n: 'Up', c: '#2E8B57' }, { n: 'Down', c: '#C0392B' }]) + '</svg>';
    }
    if (special === 'bubble') {
      const sizes = S.filter((_, i) => i % 2 === 1).flatMap(s => s.vals); const smax = Math.max(...sizes, 1); const xv = cats.map((c, i) => isNaN(+c) ? i + 1 : +c); const xt = ticksOf(Math.min(...xv), Math.max(...xv), 6); const xmin = xt[0], xmax = xt[xt.length - 1];
      xt.forEach(v => { const x = L + (v - xmin) / (xmax - xmin) * pw; o += `<text x="${x}" y="${B + 16}" text-anchor="middle" fill="#595959">${fmt(v)}</text>`; });
      for (let si = 0; si + 1 < S.length || si < S.length; si += 2) { const Y = S[si], Z = S[si + 1] || { vals: Y.vals.map(() => 1) }; Y.vals.forEach((y, i) => { const r = 6 + 26 * Math.sqrt(Math.abs(Z.vals[i]) / smax); o += `<circle cx="${L + (xv[i] - xmin) / (xmax - xmin) * pw}" cy="${yv1(y)}" r="${r}" fill="${PAL[(si / 2) % 10]}" fill-opacity=".6" stroke="${PAL[(si / 2) % 10]}"><title>${esc(Y.name)}: x=${xv[i]}, y=${y}, size=${Z.vals[i]}</title></circle>`; }); }
      return o + legend(S.filter((_, i) => i % 2 === 0).map((s, i) => ({ n: s.name, c: PAL[i % 10] }))) + '</svg>';
    }

    /* ---- generic series drawing (col, bar, line, area, scatter, hist, combo) */
    const colSeries = S.map((_, i) => i).filter(i => kinds[i] === 'col' || kinds[i] === 'bar'); const nCol = stacked ? 1 : Math.max(1, colSeries.length); const bw = Math.min(60, gw * .7 / nCol); const pad = (gw - bw * nCol) / 2;
    const xnum = t === 'scatter' ? cats.map((c, i) => isNaN(+c) ? i + 1 : +c) : null; let xmin = 0, xmax = 1; if (xnum) { const xt = ticksOf(Math.min(...xnum), Math.max(...xnum), 6); xmin = xt[0]; xmax = xt[xt.length - 1]; xt.forEach(v => { const x = L + (v - xmin) / (xmax - xmin) * pw; o += `<text x="${x}" y="${B + 16}" text-anchor="middle" fill="#595959">${fmt(v)}</text>`; }); }
    const px = i => xnum ? L + (xnum[i] - xmin) / (xmax - xmin) * pw : L + gw * i + gw / 2;
    const pos = new Array(nCat).fill(0), neg = new Array(nCat).fill(0); const baseArea = new Array(nCat).fill(0);
    S.forEach((s, si) => {
      const kind = kinds[si], col = PAL[si % 10], yv = y2s.has(si) ? yv2 : yv1;
      if (kind === 'col' || kind === 'bar') { const slot = stacked ? 0 : colSeries.indexOf(si); s.vals.forEach((v, i) => { let a = 0, b = v; if (stacked) { if (v >= 0) { a = pos[i]; pos[i] += v; b = pos[i]; } else { a = neg[i]; neg[i] += v; b = neg[i]; } } const off = pad + slot * bw; if (horiz) { const x1 = yv(Math.min(a, b)), x2 = yv(Math.max(a, b)); o += `<rect x="${x1}" y="${T + gw * (nCat - 1 - i) + off}" width="${Math.max(.5, x2 - x1)}" height="${bw}" fill="${col}"><title>${esc(s.name)} – ${esc(catName(i))}: ${v}</title></rect>`; if (dl) o += `<text x="${x2 + 4}" y="${T + gw * (nCat - 1 - i) + off + bw / 2 + 4}" fill="#404040">${fmt(v)}</text>`; } else { const y1 = yv(Math.max(a, b)), y2 = yv(Math.min(a, b)); o += `<rect x="${L + gw * i + off}" y="${y1}" width="${special === 'hist' ? gw - 2 : bw}" height="${Math.max(.5, y2 - y1)}" fill="${col}"><title>${esc(s.name)} – ${esc(catName(i))}: ${v}</title></rect>`; if (dl) o += `<text x="${L + gw * i + off + bw / 2}" y="${y1 - 4}" text-anchor="middle" fill="#404040">${fmt(v)}</text>`; } }); }
      else if (kind === 'line' || kind === 'area') { let base = baseArea; const pts = s.vals.map((v, i) => [px(i), yv(v + (kind === 'area' && stacked ? base[i] : 0))]); if (kind === 'area') { o += `<path d="M${pts[0][0]},${zero} ${pts.map(p => 'L' + p[0] + ',' + p[1]).join(' ')} L${pts[pts.length - 1][0]},${zero} Z" fill="${col}" fill-opacity=".55" stroke="${col}" stroke-width="2"/>`; if (stacked) s.vals.forEach((v, i) => base[i] += v); } else { o += `<polyline points="${pts.map(p => p.join(',')).join(' ')}" fill="none" stroke="${col}" stroke-width="2.25" stroke-linejoin="round"/>` + pts.map((p, i) => `<circle cx="${p[0]}" cy="${p[1]}" r="3.5" fill="${col}"><title>${esc(s.name)} – ${esc(catName(i))}: ${s.vals[i]}</title></circle>`).join(''); if (dl) pts.forEach((p, i) => o += `<text x="${p[0]}" y="${p[1] - 8}" text-anchor="middle" fill="#404040">${fmt(s.vals[i])}</text>`); } }
      else if (kind === 'scatter') { s.vals.forEach((v, i) => { o += `<circle cx="${px(i)}" cy="${yv(v)}" r="4" fill="${col}"><title>${esc(s.name)}: (${xnum ? xnum[i] : i + 1}, ${v})</title></circle>`; if (dl) o += `<text x="${px(i)}" y="${yv(v) - 8}" text-anchor="middle" fill="#404040">${fmt(v)}</text>`; }); }
      // error bars
      if (ch.errBars) { const ev = errAmt(s, ch); s.vals.forEach((v, i) => { const e = ev(v); const x = (kind === 'col' || kind === 'bar') ? L + gw * i + pad + (stacked ? 0 : colSeries.indexOf(si)) * bw + bw / 2 : px(i); if (horiz) return; o += `<line x1="${x}" y1="${yv(v + e)}" x2="${x}" y2="${yv(v - e)}" stroke="#404040"/><line x1="${x - 4}" y1="${yv(v + e)}" x2="${x + 4}" y2="${yv(v + e)}" stroke="#404040"/><line x1="${x - 4}" y1="${yv(v - e)}" x2="${x + 4}" y2="${yv(v - e)}" stroke="#404040"/>`; }); }
    });
    // trendline
    if (ch.trendline) { const si = Math.min(S.length - 1, ch.trendSeries || 0); const s = S[si]; const xs = (xnum || s.vals.map((_, i) => i + 1)).slice(), ys = s.vals; const tr = trendFn(ch.trendline, xs, ys, ch.trendPeriod); const yv = y2s.has(si) ? yv2 : yv1;
      if (tr) { const col = PAL[si % 10]; const X = xnum ? (x => L + (x - xmin) / (xmax - xmin) * pw) : (x => L + gw * (x - 1) + gw / 2); let pts; if (tr.pts) pts = tr.pts; else { pts = []; const x0 = Math.min(...xs), x1 = Math.max(...xs); for (let k = 0; k <= 40; k++) { const x = x0 + (x1 - x0) * k / 40; const y = tr.f(x); if (isFinite(y)) pts.push([x, y]); } } o += `<polyline points="${pts.map(p => X(p[0]) + ',' + yv(p[1])).join(' ')}" fill="none" stroke="${col}" stroke-width="2" stroke-dasharray="6 4" opacity=".85"><title>Trendline (${esc(ch.trendline)}) ${esc(tr.eq || '')}</title></polyline>`; if (ch.trendEq && tr.eq) o += `<text x="${Rr - 6}" y="${T + 14}" text-anchor="end" fill="${col}" font-size="11">${esc(tr.eq)}${tr.r2 != null ? '  R² = ' + tr.r2.toFixed(4) : ''}</text>`; } }
    const items = S.map((s, i) => ({ n: s.name, c: PAL[i % 10] })); if (ch.trendline) items.push({ n: 'Trendline', c: '#7f7f7f' });
    return o + legend(items) + '</svg>';
  }
  function errAmt(s, ch) { const k = ch.errBars, v = +ch.errVal || (k === 'pct' ? 5 : k === 'sd' ? 1 : 1); if (k === 'pct') return x => Math.abs(x) * v / 100; if (k === 'sd') { const m = s.vals.reduce((a, b) => a + b, 0) / s.vals.length; const sd = Math.sqrt(s.vals.reduce((a, b) => a + (b - m) ** 2, 0) / Math.max(1, s.vals.length - 1)); return () => sd * v; } return () => v; }
})(window);
