/* Excel chart SVG renderer — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL;
  const PAL = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
  XL.CHART_PAL = PAL;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmt = v => { const a = Math.abs(v); if (a >= 1e9) return +(v / 1e9).toFixed(2) + 'B'; if (a >= 1e6) return +(v / 1e6).toFixed(2) + 'M'; if (a >= 1e4) return +(v / 1e3).toFixed(1) + 'K'; return String(+v.toFixed(a < 1 ? 3 : 2)); };
  function niceTicks(lo, hi, n) {
    if (hi === lo) { hi = lo + 1; }
    const raw = (hi - lo) / n, mag = Math.pow(10, Math.floor(Math.log10(raw))), norm = raw / mag; const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
    const s = Math.floor(lo / step) * step, e = Math.ceil(hi / step) * step; const t = []; for (let v = s; v <= e + step * 1e-6; v += step) t.push(+v.toFixed(10)); return t;
  }
  XL.chartSVG = function (d, ch, W, H) {
    const t = ch.type, S = d.series, cats = d.cats; const font = 'font-family="Segoe UI,Calibri,Arial,sans-serif"';
    let o = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" ${font} font-size="11">`;
    o += `<rect width="${W}" height="${H}" fill="#fff"/>`;
    if (!S.length || !cats.length) return o + `<text x="${W / 2}" y="${H / 2}" text-anchor="middle" fill="#888">No data to plot – check the data range</text></svg>`;
    const legendRight = (t === 'pie' || t === 'doughnut') ? false : false;
    const titleH = ch.title ? 34 : 10; const legendH = ch.legend === false ? 8 : 28;
    if (ch.title) o += `<text x="${W / 2}" y="22" text-anchor="middle" font-size="15" fill="#404040">${esc(ch.title)}</text>`;
    const items = (t === 'pie' || t === 'doughnut') ? cats.map((c, i) => ({ n: c, c: PAL[i % 10] })) : S.map((s, i) => ({ n: s.name, c: PAL[i % 10] }));
    if (ch.legend !== false) { let tw = items.reduce((a, it) => a + 26 + Math.min(110, it.n.length * 6.2), 0); let x = Math.max(8, (W - tw) / 2); items.forEach(it => { o += `<rect x="${x}" y="${H - 20}" width="10" height="10" fill="${it.c}"/><text x="${x + 14}" y="${H - 11}" fill="#595959">${esc(it.n.length > 18 ? it.n.slice(0, 17) + '…' : it.n)}</text>`; x += 26 + Math.min(110, it.n.length * 6.2); }); }
    if (t === 'pie' || t === 'doughnut') {
      const vals = S[0].vals.map(v => Math.max(0, v)); const tot = vals.reduce((a, b) => a + b, 0) || 1; const cx = W / 2, cy = titleH + (H - titleH - legendH) / 2, R = Math.max(20, Math.min(W / 2 - 20, (H - titleH - legendH) / 2 - 8)); let a0 = -Math.PI / 2;
      vals.forEach((v, i) => { const a1 = a0 + v / tot * Math.PI * 2; const big = a1 - a0 > Math.PI ? 1 : 0; const p = (a, r) => [cx + r * Math.cos(a), cy + r * Math.sin(a)]; const [x0, y0] = p(a0, R), [x1, y1] = p(a1, R); const inner = t === 'doughnut' ? R * 0.55 : 0;
        let path; if (v / tot > 0.9999) path = inner ? `M${cx - R},${cy} a${R},${R} 0 1,0 ${2 * R},0 a${R},${R} 0 1,0 ${-2 * R},0 M${cx - inner},${cy} a${inner},${inner} 0 1,1 ${2 * inner},0 a${inner},${inner} 0 1,1 ${-2 * inner},0` : `M${cx - R},${cy} a${R},${R} 0 1,0 ${2 * R},0 a${R},${R} 0 1,0 ${-2 * R},0`; else if (inner) { const [x2, y2] = p(a1, inner), [x3, y3] = p(a0, inner); path = `M${x0},${y0} A${R},${R} 0 ${big} 1 ${x1},${y1} L${x2},${y2} A${inner},${inner} 0 ${big} 0 ${x3},${y3} Z`; } else path = `M${cx},${cy} L${x0},${y0} A${R},${R} 0 ${big} 1 ${x1},${y1} Z`;
        o += `<path d="${path}" fill="${PAL[i % 10]}" stroke="#fff" stroke-width="1.5" fill-rule="evenodd"><title>${esc(cats[i])}: ${v} (${(v / tot * 100).toFixed(1)}%)</title></path>`;
        if (v / tot > 0.04) { const am = (a0 + a1) / 2, rr = inner ? (R + inner) / 2 : R * 0.65; o += `<text x="${cx + rr * Math.cos(am)}" y="${cy + rr * Math.sin(am) + 4}" text-anchor="middle" fill="#fff" font-size="11" font-weight="600">${(v / tot * 100).toFixed(0)}%</text>`; } a0 = a1; });
      return o + '</svg>';
    }
    // cartesian
    const stacked = !!ch.stacked && (t === 'col' || t === 'bar' || t === 'area'); const horiz = t === 'bar';
    let lo = 0, hi = 0;
    if (t === 'scatter') { lo = Infinity; hi = -Infinity; S.forEach(s => s.vals.forEach(v => { lo = Math.min(lo, v); hi = Math.max(hi, v); })); if (!isFinite(lo)) { lo = 0; hi = 1; } }
    else if (stacked) cats.forEach((_, i) => { let p = 0, n = 0; S.forEach(s => { const v = s.vals[i] || 0; if (v >= 0) p += v; else n += v; }); hi = Math.max(hi, p); lo = Math.min(lo, n); });
    else S.forEach(s => s.vals.forEach(v => { hi = Math.max(hi, v); lo = Math.min(lo, v); }));
    const ticks = niceTicks(lo, hi, 5); const tmin = ticks[0], tmax = ticks[ticks.length - 1];
    const lblW = Math.max(...ticks.map(v => fmt(v).length)) * 6.3 + 12; const catLblH = 22;
    let L, Rr, T = titleH, B = H - legendH - catLblH;
    if (horiz) { const maxC = Math.max(...cats.map(c => Math.min(16, c.length))); L = maxC * 6.2 + 14; Rr = W - 16; B = H - legendH - 20; } else { L = lblW + 6; Rr = W - 16; }
    const pw = Rr - L, ph = B - T; if (pw < 20 || ph < 20) return o + '</svg>';
    // scatter x scale
    let xs = null; if (t === 'scatter') { const xv = cats.map((c, i) => isNaN(+c) ? i + 1 : +c); const xt = niceTicks(Math.min(...xv), Math.max(...xv), 6); xs = { v: xv, t: xt, min: xt[0], max: xt[xt.length - 1] }; }
    const yv = v => horiz ? L + (v - tmin) / (tmax - tmin) * pw : B - (v - tmin) / (tmax - tmin) * ph;
    ticks.forEach(v => { const p = yv(v); if (horiz) o += `<line x1="${p}" y1="${T}" x2="${p}" y2="${B}" stroke="#d9d9d9"/><text x="${p}" y="${B + 14}" text-anchor="middle" fill="#595959">${fmt(v)}</text>`; else o += `<line x1="${L}" y1="${p}" x2="${Rr}" y2="${p}" stroke="#d9d9d9"/><text x="${L - 6}" y="${p + 4}" text-anchor="end" fill="#595959">${fmt(v)}</text>`; });
    const n = cats.length; const gw = (horiz ? ph : pw) / n;
    if (t === 'scatter') {
      xs.t.forEach(v => { const x = L + (v - xs.min) / (xs.max - xs.min) * pw; o += `<line x1="${x}" y1="${B}" x2="${x}" y2="${B + 4}" stroke="#bfbfbf"/><text x="${x}" y="${B + 16}" text-anchor="middle" fill="#595959">${fmt(v)}</text>`; });
      o += `<line x1="${L}" y1="${B}" x2="${Rr}" y2="${B}" stroke="#bfbfbf"/>`;
      S.forEach((s, si) => s.vals.forEach((v, i) => { const x = L + (xs.v[i] - xs.min) / (xs.max - xs.min) * pw; o += `<circle cx="${x}" cy="${yv(v)}" r="4" fill="${PAL[si % 10]}"><title>${esc(s.name)}: (${xs.v[i]}, ${v})</title></circle>`; }));
      return o + '</svg>';
    }
    cats.forEach((c, i) => { const lab = c.length > 12 && !horiz ? c.slice(0, 11) + '…' : c; if (horiz) o += `<text x="${L - 6}" y="${T + gw * (n - 1 - i) + gw / 2 + 4}" text-anchor="end" fill="#595959">${esc(lab.slice(0, 16))}</text>`; else { const x = L + gw * i + gw / 2; o += `<text x="${x}" y="${B + 16}" text-anchor="middle" fill="#595959">${esc(lab)}</text>`; } });
    const zero = yv(Math.max(tmin, Math.min(0, tmax)));
    o += horiz ? `<line x1="${zero}" y1="${T}" x2="${zero}" y2="${B}" stroke="#bfbfbf"/>` : `<line x1="${L}" y1="${zero}" x2="${Rr}" y2="${zero}" stroke="#bfbfbf"/>`;
    if (t === 'col' || t === 'bar') {
      const ns = stacked ? 1 : S.length; const bw = Math.min(60, gw * 0.7 / ns); const pad = (gw - bw * ns) / 2;
      cats.forEach((c, i) => { let pos = 0, neg = 0; S.forEach((s, si) => { const v = s.vals[i] || 0; let a, b; if (stacked) { if (v >= 0) { a = pos; pos += v; b = pos; } else { a = neg; neg += v; b = neg; } } else { a = 0; b = v; } const col = PAL[si % 10]; const off = stacked ? pad : pad + si * bw;
        if (horiz) { const x1 = yv(Math.min(a, b)), x2 = yv(Math.max(a, b)); o += `<rect x="${x1}" y="${T + gw * (n - 1 - i) + off}" width="${Math.max(0.5, x2 - x1)}" height="${bw}" fill="${col}"><title>${esc(s.name)} – ${esc(c)}: ${v}</title></rect>`; }
        else { const y1 = yv(Math.max(a, b)), y2 = yv(Math.min(a, b)); o += `<rect x="${L + gw * i + off}" y="${y1}" width="${bw}" height="${Math.max(0.5, y2 - y1)}" fill="${col}"><title>${esc(s.name)} – ${esc(c)}: ${v}</title></rect>`; } }); });
    } else if (t === 'line' || t === 'area') {
      S.forEach((s, si) => { const col = PAL[si % 10]; let base = new Array(n).fill(0); if (t === 'area' && stacked) for (let k = 0; k < si; k++) base = base.map((b, i) => b + (S[k].vals[i] || 0)); const pts = s.vals.map((v, i) => [L + gw * i + gw / 2, yv(v + (t === 'area' && stacked ? base[i] : 0))]);
        if (t === 'area') o += `<path d="M${pts[0][0]},${zero} ${pts.map(p => 'L' + p[0] + ',' + p[1]).join(' ')} L${pts[pts.length - 1][0]},${zero} Z" fill="${col}" fill-opacity="0.55" stroke="${col}" stroke-width="2"/>`;
        else { o += `<polyline points="${pts.map(p => p.join(',')).join(' ')}" fill="none" stroke="${col}" stroke-width="2.25" stroke-linejoin="round"/>`; pts.forEach((p, i) => { o += `<circle cx="${p[0]}" cy="${p[1]}" r="3.5" fill="${col}"><title>${esc(s.name)} – ${esc(cats[i])}: ${s.vals[i]}</title></circle>`; }); } });
    }
    return o + '</svg>';
  };
})(window);
