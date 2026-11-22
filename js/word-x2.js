/* Word extras 2: Insert, Layout, Design, Home extras — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, el = OA.el;
  OA.wordExtras = OA.wordExtras || [];

  /* ---- LaTeX-subset -> MathML + linear text */
  const SYM = { alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ε', theta: 'θ', lambda: 'λ', mu: 'μ', pi: 'π', rho: 'ρ', sigma: 'σ', tau: 'τ', phi: 'φ', omega: 'ω', Delta: 'Δ', Sigma: 'Σ', Omega: 'Ω', Pi: 'Π', pm: '±', times: '×', cdot: '·', div: '÷', leq: '≤', geq: '≥', neq: '≠', approx: '≈', infty: '∞', to: '→', rightarrow: '→', leftarrow: '←', in: '∈', partial: '∂', nabla: '∇', ldots: '…', degree: '°', sum: '∑', prod: '∏', int: '∫', forall: '∀', exists: '∃', cup: '∪', cap: '∩', subset: '⊂', angle: '∠', Rightarrow: '⇒', Leftrightarrow: '⇔' };
  const BIGOP = { sum: 1, prod: 1, int: 1 };
  function texToMath(src) {
    let i = 0; const s = src;
    const esc = OA.esc;
    function group() { // {...} or single token
      skip(); if (s[i] === '{') { i++; const r = seq('}'); i++; return r; } return atom();
    }
    function skip() { while (s[i] === ' ') i++; }
    function atom() {
      skip(); const c = s[i]; if (c === undefined) return ['', '']; 
      if (c === '\\') { i++; let n = ''; while (/[a-zA-Z]/.test(s[i] || '')) n += s[i++]; if (!n) { const ch = s[i++]; return [`<mo>${esc(ch)}</mo>`, ch]; }
        if (n === 'frac') { const a = group(), b = group(); return [`<mfrac><mrow>${a[0]}</mrow><mrow>${b[0]}</mrow></mfrac>`, `(${a[1]})/(${b[1]})`]; }
        if (n === 'sqrt') { skip(); let idx = null; if (s[i] === '[') { i++; let t = ''; while (s[i] !== ']' && i < s.length) t += s[i++]; i++; idx = t; } const a = group(); return idx ? [`<mroot><mrow>${a[0]}</mrow><mn>${esc(idx)}</mn></mroot>`, idx + '√(' + a[1] + ')'] : [`<msqrt>${a[0]}</msqrt>`, '√(' + a[1] + ')']; }
        if (n === 'text') { const a = group(); return [`<mtext>${a[1]}</mtext>`, a[1]]; }
        if (n === 'left' || n === 'right') { const ch = s[i++] || ''; return ch === '.' ? ['', ''] : [`<mo>${esc(ch)}</mo>`, ch]; }
        const g = SYM[n]; if (g) return [BIGOP[n] ? `<mo largeop="true">${g}</mo>` : (/^[a-zA-Z]/.test(g) || /[α-ωΑ-Ω∂∇]/.test(g) ? `<mi>${g}</mi>` : `<mo>${g}</mo>`), g];
        return [`<mi>${esc(n)}</mi>`, n]; }
      if (c === '{') { i++; const r = seq('}'); i++; return r; }
      i++;
      if (/[0-9.]/.test(c)) { let n = c; while (/[0-9.]/.test(s[i] || '')) n += s[i++]; return [`<mn>${n}</mn>`, n]; }
      if (/[a-zA-Z]/.test(c)) return [`<mi>${c}</mi>`, c];
      return [`<mo>${esc(c)}</mo>`, c];
    }
    function seq(end) {
      let m = '', t = '';
      while (i < s.length && s[i] !== end) {
        const a = atom(); let base = a;
        for (;;) {
          skip(); if (s[i] !== '^' && s[i] !== '_') break;
          let sub = null, sup = null;
          while (s[i] === '^' || s[i] === '_') { const k = s[i++]; const g = group(); if (k === '^') sup = g; else sub = g; skip(); }
          const big = /largeop/.test(base[0]);
          if (sub && sup) base = [`<m${big ? 'under' : 'sub'}sup>${base[0]}<mrow>${sub[0]}</mrow><mrow>${sup[0]}</mrow></m${big ? 'under' : 'sub'}sup>`, base[1] + '_' + sub[1] + '^' + sup[1]];
          else if (sup) base = [`<msup>${base[0]}<mrow>${sup[0]}</mrow></msup>`, base[1] + '^' + sup[1]];
          else if (sub) base = [`<msub>${base[0]}<mrow>${sub[0]}</mrow></msub>`, base[1] + '_' + sub[1]];
        }
        m += base[0]; t += base[1];
      }
      return [m, t];
    }
    const r = seq('\0'); return { ml: `<math xmlns="http://www.w3.org/1998/Math/MathML" display="inline"><mrow>${r[0]}</mrow></math>`, lin: r[1] };
  }
  OA.texToMath = texToMath;
  const EQS = [['Quadratic formula', 'x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}'], ['Pythagorean theorem', 'a^2 + b^2 = c^2'], ['Area of a circle', 'A = \\pi r^2'], ['Binomial theorem', '(a+b)^n = \\sum_{k=0}^{n} \\frac{n!}{k!(n-k)!} a^{n-k} b^k'], ['Euler identity', 'e^{i\\pi} + 1 = 0'], ['Einstein', 'E = mc^2'], ['Definite integral', '\\int_{a}^{b} f(x) dx'], ['Slope', 'm = \\frac{y_2 - y_1}{x_2 - x_1}'], ['Compound interest', 'A = P \\left(1 + \\frac{r}{n}\\right)^{nt}']];

  function svgToPng(svg, w, h, cb) {
    const img = new Image(); img.onload = () => { const c = document.createElement('canvas'); c.width = w * 2; c.height = h * 2; const x = c.getContext('2d'); x.fillStyle = 'rgba(0,0,0,0)'; x.scale(2, 2); x.drawImage(img, 0, 0, w, h); cb(c.toDataURL('image/png')); }; img.onerror = () => cb('data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)))); img.src = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svg)));
  }
  const COLS = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

  OA.wordExtras.push(function (ed, H, spec) {
    const P = ed.page, { B, L } = H; const X = ed.x = ed.x || {};
    const tabAdd = (id, groups) => { const t = spec.tabs.find(x => x.id === id); if (t) t.groups.push(...groups); };
    const sel = () => { const s = getSelection(); return s.rangeCount && P.contains(s.getRangeAt(0).commonAncestorContainer) ? s.getRangeAt(0) : null; };
    const insImg = (svg, w, h, alt) => svgToPng(svg, w, h, url => ed.insertImageURL(url, alt));

    /* ============ Equation */
    X.equation = async () => {
      const ta = el('textarea', { class: 'inp', rows: 3, style: { width: '100%' }, placeholder: 'e.g. x = \\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}' }); const pv = el('div', { style: { minHeight: '60px', margin: '8px 0', padding: '10px', border: '1px solid #ddd', fontSize: '20px', textAlign: 'center' } });
      const upd = () => { try { pv.innerHTML = texToMath(ta.value).ml; } catch (e) { pv.textContent = '…'; } }; ta.addEventListener('input', upd); ta.value = EQS[0][1]; upd();
      const gal = el('div', { style: { display: 'flex', flexWrap: 'wrap', gap: '4px', marginBottom: '6px' } }, EQS.map(([n, t]) => el('button', { class: 'btn', type: 'button', text: n, onclick: () => { ta.value = t; upd(); } })));
      const sy = el('div', { style: { fontSize: '11px', color: '#666' }, text: 'Use LaTeX: \\frac{a}{b}  \\sqrt{x}  x^{2}  x_{i}  \\sum_{i=1}^{n}  \\int  \\alpha \\pi \\theta  \\pm \\times \\leq \\geq \\neq \\infty' });
      OA.dialog({ title: 'Insert Equation', width: 520, content: el('div', null, gal, ta, pv, sy), buttons: [{ label: 'Cancel', value: false }, { label: 'Insert', primary: true, onClick: () => { const r = texToMath(ta.value); X.insInline(`<span class="eq" contenteditable="false" data-tex="${OA.esc(ta.value)}" data-lin="${OA.esc(r.lin)}">${r.ml}</span>&nbsp;`); } }] });
    };
    P.addEventListener('dblclick', e => { const q = e.target.closest && e.target.closest('.eq'); if (!q) return; (async () => { const v = await OA.prompt('Edit equation (LaTeX):', q.dataset.tex, { title: 'Equation' }); if (v) ed.mutate(() => { const r = texToMath(v); q.dataset.tex = v; q.dataset.lin = r.lin; q.innerHTML = r.ml; }); })(); });

    /* ============ Chart */
    X.chart = async () => {
      const v = await OA.form('Insert Chart', [{ id: 't', label: 'Chart type', type: 'select', value: 'bar', options: [['bar', 'Column'], ['hbar', 'Bar'], ['line', 'Line'], ['pie', 'Pie'], ['donut', 'Doughnut'], ['area', 'Area']] }, { id: 'title', label: 'Title', value: 'Chart Title' }, { id: 'data', label: 'Data — first row = series names; first column = categories (comma or tab separated)', type: 'textarea', rows: 7, value: ',Series 1,Series 2\nQ1,4.3,2.4\nQ2,2.5,4.4\nQ3,3.5,1.8\nQ4,4.5,2.8' }], { width: 480, ok: 'Insert' });
      if (!v) return; const rows = v.data.trim().split(/\r?\n/).map(l => l.split(/\t|,/).map(x => x.trim())); const names = rows[0].slice(1); const cats = rows.slice(1).map(r => r[0]); const ser = names.map((n, i) => rows.slice(1).map(r => parseFloat(r[i + 1]) || 0));
      if (!cats.length) return OA.alert('Need at least one data row.'); insImg(chartSvg(v.t, v.title, names, cats, ser), 480, 300, 'Chart: ' + v.title);
    };
    function chartSvg(type, title, names, cats, ser) {
      const W = 480, Hh = 300, ml = 44, mr = 14, mt = 38, mb = 46, pw = W - ml - mr, ph = Hh - mt - mb; const e = OA.esc;
      let g = `<rect width="${W}" height="${Hh}" fill="#fff"/><text x="${W / 2}" y="22" text-anchor="middle" font-family="Calibri,Arial" font-size="16" fill="#404040">${e(title)}</text>`;
      if (type === 'pie' || type === 'donut') {
        const vals = ser[0], tot = vals.reduce((a, b) => a + Math.abs(b), 0) || 1; let a0 = -Math.PI / 2; const cx = 190, cy = 160, R = 105;
        vals.forEach((val, i) => { const a1 = a0 + Math.abs(val) / tot * Math.PI * 2, lg = a1 - a0 > Math.PI ? 1 : 0; g += `<path d="M${cx},${cy} L${cx + R * Math.cos(a0)},${cy + R * Math.sin(a0)} A${R},${R} 0 ${lg} 1 ${cx + R * Math.cos(a1)},${cy + R * Math.sin(a1)} Z" fill="${COLS[i % 8]}" stroke="#fff"/>`; const am = (a0 + a1) / 2; g += `<text x="${cx + R * .65 * Math.cos(am)}" y="${cy + R * .65 * Math.sin(am)}" text-anchor="middle" font-size="11" fill="#fff" font-family="Calibri,Arial">${Math.round(Math.abs(val) / tot * 100)}%</text>`; a0 = a1; });
        if (type === 'donut') g += `<circle cx="${cx}" cy="${cy}" r="${R * .5}" fill="#fff"/>`;
        cats.forEach((c, i) => { g += `<rect x="340" y="${70 + i * 20}" width="11" height="11" fill="${COLS[i % 8]}"/><text x="357" y="${80 + i * 20}" font-size="12" font-family="Calibri,Arial" fill="#404040">${e(c)}</text>`; });
        return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}">${g}</svg>`;
      }
      const all = ser.flat(); let mx = Math.max(...all, 0), mn = Math.min(...all, 0); const step = niceStep((mx - mn) || 1); mx = Math.ceil(mx / step) * step; mn = Math.floor(mn / step) * step; const horiz = type === 'hbar';
      const yv = v => mt + ph - (v - mn) / (mx - mn || 1) * ph;
      for (let t = mn; t <= mx + 1e-9; t += step) { const y = yv(t); g += `<line x1="${ml}" x2="${ml + pw}" y1="${y}" y2="${y}" stroke="#d9d9d9"/><text x="${ml - 6}" y="${y + 4}" text-anchor="end" font-size="11" font-family="Calibri,Arial" fill="#595959">${+t.toFixed(2)}</text>`; }
      const n = cats.length, bw = pw / n;
      if (horiz) { g = g.replace(/<line[^>]*\/>|<text x="\d+" y="[\d.]+" text-anchor="end"[^<]*<\/text>/g, ''); const xv = v => ml + (v - mn) / (mx - mn || 1) * pw, bh = ph / n; for (let t = mn; t <= mx + 1e-9; t += step) { const x = xv(t); g += `<line x1="${x}" x2="${x}" y1="${mt}" y2="${mt + ph}" stroke="#d9d9d9"/><text x="${x}" y="${mt + ph + 14}" text-anchor="middle" font-size="11" font-family="Calibri,Arial" fill="#595959">${+t.toFixed(2)}</text>`; }
        cats.forEach((c, ci) => { g += `<text x="${ml - 6}" y="${mt + bh * ci + bh / 2 + 4}" text-anchor="end" font-size="11" font-family="Calibri,Arial" fill="#595959">${e(c)}</text>`; ser.forEach((s, si) => { const bb = bh * .7 / ser.length; g += `<rect x="${Math.min(xv(0), xv(s[ci]))}" y="${mt + bh * ci + bh * .15 + si * bb}" width="${Math.abs(xv(s[ci]) - xv(0))}" height="${bb - 1}" fill="${COLS[si % 8]}"/>`; }); }); }
      else {
        cats.forEach((c, ci) => { g += `<text x="${ml + bw * ci + bw / 2}" y="${mt + ph + 16}" text-anchor="middle" font-size="11" font-family="Calibri,Arial" fill="#595959">${e(c)}</text>`; });
        if (type === 'bar') ser.forEach((s, si) => s.forEach((val, ci) => { const bb = bw * .7 / ser.length; g += `<rect x="${ml + bw * ci + bw * .15 + si * bb}" y="${Math.min(yv(val), yv(0))}" width="${bb - 1}" height="${Math.abs(yv(val) - yv(0))}" fill="${COLS[si % 8]}"/>`; }));
        else ser.forEach((s, si) => { const pts = s.map((val, ci) => [ml + bw * ci + bw / 2, yv(val)]); if (type === 'area') g += `<path d="M${pts[0][0]},${yv(0)} ${pts.map(p => 'L' + p.join(',')).join(' ')} L${pts[pts.length - 1][0]},${yv(0)} Z" fill="${COLS[si % 8]}" opacity=".55"/>`; g += `<polyline fill="none" stroke="${COLS[si % 8]}" stroke-width="2.5" points="${pts.map(p => p.join(',')).join(' ')}"/>` + pts.map(p => `<circle cx="${p[0]}" cy="${p[1]}" r="3.5" fill="${COLS[si % 8]}"/>`).join(''); });
      }
      g += `<line x1="${ml}" x2="${ml + pw}" y1="${yv(0)}" y2="${yv(0)}" stroke="#808080"/>`;
      if (names.length > 1) names.forEach((nm, i) => { g += `<rect x="${ml + i * 90}" y="${Hh - 18}" width="10" height="10" fill="${COLS[i % 8]}"/><text x="${ml + i * 90 + 14}" y="${Hh - 9}" font-size="11" font-family="Calibri,Arial" fill="#595959">${e(nm)}</text>`; });
      return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${Hh}" viewBox="0 0 ${W} ${Hh}">${g}</svg>`;
    }
    function niceStep(range) { const raw = range / 5, p = Math.pow(10, Math.floor(Math.log10(raw))), f = raw / p; return (f < 1.5 ? 1 : f < 3 ? 2 : f < 7 ? 5 : 10) * p; }

    /* ============ SmartArt */
    X.smartart = async () => {
      const v = await OA.form('Choose a SmartArt Graphic', [{ id: 'k', label: 'Layout', type: 'select', value: 'process', options: [['process', 'Basic Process'], ['list', 'Vertical List'], ['cycle', 'Basic Cycle'], ['pyramid', 'Basic Pyramid'], ['hier', 'Organization Chart'], ['venn', 'Basic Venn']] }, { id: 'items', label: 'Text (one item per line; for the org chart, first line = top)', type: 'textarea', rows: 6, value: 'Plan\nBuild\nTest\nRelease' }], { ok: 'Insert', width: 420 });
      if (!v) return; const it = v.items.split('\n').map(s => s.trim()).filter(Boolean).slice(0, 8); if (!it.length) return; const e = OA.esc; let W = 480, h = 200, g = '';
      const box = (x, y, w, hh, t, i, shape) => `<${shape || 'rect'} ${shape === 'ellipse' ? `cx="${x + w / 2}" cy="${y + hh / 2}" rx="${w / 2}" ry="${hh / 2}"` : `x="${x}" y="${y}" width="${w}" height="${hh}" rx="6"`} fill="${COLS[i % 8]}"/><text x="${x + w / 2}" y="${y + hh / 2 + 5}" text-anchor="middle" font-size="14" font-family="Calibri,Arial" fill="#fff">${e(t.slice(0, 18))}</text>`;
      if (v.k === 'process') { const n = it.length, w = Math.min(120, (W - 20) / n - 22); it.forEach((t, i) => { const x = 10 + i * (w + 22); g += box(x, 70, w, 60, t, i); if (i < n - 1) g += `<path d="M${x + w + 4},100 l12,0 m-5,-5 l5,5 l-5,5" stroke="#555" stroke-width="2" fill="none"/>`; }); }
      else if (v.k === 'list') { h = it.length * 48 + 16; it.forEach((t, i) => g += box(60, 8 + i * 48, 360, 40, t, i)); }
      else if (v.k === 'cycle') { h = 280; const cx = 240, cy = 140, R = 90; it.forEach((t, i) => { const a = i / it.length * Math.PI * 2 - Math.PI / 2; g += box(cx + R * 1.2 * Math.cos(a) - 50, cy + R * Math.sin(a) - 22, 100, 44, t, i, 'ellipse'); }); }
      else if (v.k === 'pyramid') { h = it.length * 52 + 10; it.forEach((t, i) => { const w = 90 + i * (360 / it.length), x = 240 - w / 2; g += `<rect x="${x}" y="${6 + i * 52}" width="${w}" height="46" fill="${COLS[i % 8]}"/><text x="240" y="${34 + i * 52}" text-anchor="middle" font-size="14" font-family="Calibri,Arial" fill="#fff">${e(t)}</text>`; }); }
      else if (v.k === 'hier') { h = 220; g += box(190, 10, 100, 44, it[0], 0); const k = it.slice(1), w = Math.min(100, (W - 20) / Math.max(1, k.length) - 10); k.forEach((t, i) => { const x = 10 + i * (w + 10) + ((W - 20) - k.length * (w + 10) + 10) / 2; g += `<path d="M240,54 V82 H${x + w / 2} V120" stroke="#777" fill="none"/>` + box(x, 120, w, 44, t, i + 1); }); }
      else if (v.k === 'venn') { h = 260; const n = Math.min(it.length, 4), cx = 240, cy = 130; it.slice(0, n).forEach((t, i) => { const a = i / n * Math.PI * 2 - Math.PI / 2, x = cx + (n > 1 ? 50 : 0) * Math.cos(a), y = cy + (n > 1 ? 50 : 0) * Math.sin(a); g += `<circle cx="${x}" cy="${y}" r="80" fill="${COLS[i]}" opacity=".6"/><text x="${cx + 95 * Math.cos(a)}" y="${cy + 95 * Math.sin(a) + 4}" text-anchor="middle" font-size="13" font-family="Calibri,Arial" fill="#222">${e(t.slice(0, 16))}</text>`; }); }
      insImg(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${h}" viewBox="0 0 ${W} ${h}"><rect width="${W}" height="${h}" fill="#fff"/>${g}</svg>`, W, h, 'SmartArt: ' + v.k);
    };

    /* ============ WordArt */
    const WA = [['Fill – Blue', 'color:#2f5496;font-weight:700;text-shadow:2px 2px 3px rgba(0,0,0,.35)'], ['Gradient – Sunset', 'background:linear-gradient(90deg,#ff512f,#f09819);-webkit-background-clip:text;background-clip:text;color:transparent;font-weight:800'], ['Outline', '-webkit-text-stroke:1.5px #2b579a;color:transparent;font-weight:800'], ['Gold Shadow', 'color:#c9a227;font-weight:800;text-shadow:3px 3px 0 #6b5b14'], ['Neon', 'color:#fff;font-weight:700;text-shadow:0 0 6px #0ff,0 0 14px #0ff,0 0 24px #08f;background:#111;padding:2px 10px'], ['Rainbow', 'background:linear-gradient(90deg,red,orange,gold,green,blue,violet);-webkit-background-clip:text;background-clip:text;color:transparent;font-weight:800']];
    X.wordart = async (anchor) => {
      const r = sel(); const t0 = r && !r.collapsed ? r.toString() : '';
      OA.menu(anchor, WA.map(([l, st]) => ({ label: l, style: st.includes('transparent') && !st.includes('gradient') ? '' : '', onClick: async () => { const t = await OA.prompt('WordArt text:', t0 || 'Your text here', { title: 'WordArt' }); if (t) X.insInline(`<span class="wordart" style="font-size:32pt;${st}">${OA.esc(t)}</span>&nbsp;`); } })));
    };

    /* ============ Drop cap / Text box / Quick parts / Signature */
    X.dropcap = (mode) => ed.mutate(() => { const s = sel(); const n = s && s.startContainer; const b = n && (n.nodeType === 3 ? n.parentNode : n).closest('p'); if (!b) return; b.classList.remove('dc'); b.querySelectorAll('.dcap').forEach(d => d.replaceWith(d.textContent)); if (mode === 'none') return; b.normalize(); const w = document.createTreeWalker(b, NodeFilter.SHOW_TEXT); let tn; while ((tn = w.nextNode())) if (tn.nodeValue.trim()) break; if (!tn) return; const i = tn.nodeValue.search(/\S/); const ch = tn.nodeValue[i]; const sp = document.createElement('span'); sp.className = 'dcap' + (mode === 'margin' ? ' dcm' : ''); sp.textContent = ch; tn.nodeValue = tn.nodeValue.slice(0, i) + tn.nodeValue.slice(i + 1); tn.parentNode.insertBefore(sp, tn); if (i) tn.parentNode.insertBefore(document.createTextNode(tn.nodeValue.slice(0, i)), sp), tn.nodeValue = tn.nodeValue.slice(i); });
    const TB = { simple: 'border:1px solid #000;padding:8px;', side: 'border-left:4px solid #2f5496;background:#f2f2f2;padding:8px 12px;', quote: 'border-top:3px solid #2f5496;border-bottom:3px solid #2f5496;padding:10px;text-align:center;font-style:italic;font-size:14pt;' };
    X.textbox = anchor => OA.menu(anchor, [['simple', 'Simple Text Box'], ['side', 'Sidebar'], ['quote', 'Pull Quote']].map(([k, l]) => ({ label: l, onClick: () => X.insBlock(`<div class="tbx" style="${TB[k]}width:60%;margin:6pt auto;"><p>Type here…</p></div><p><br></p>`) })));
    const qp = X.qp = [{ n: 'Confidential notice', h: '<p style="text-align:center;color:#c00"><b>CONFIDENTIAL — Do not distribute</b></p>' }, { n: 'Thank-you closing', h: '<p>Thank you for your time and consideration.</p><p>Sincerely,<br><br>[Your Name]</p>' }];
    X.quickParts = anchor => OA.menu(anchor, [{ header: 'Building Blocks' }, ...qp.map(q => ({ label: q.n, onClick: () => /^<p|^<div/.test(q.h) ? X.insBlock(q.h) : X.insInline(q.h) })), { sep: true }, { header: 'Document Property' }, ...[['Title', () => (document.querySelector('.app-word .tb-name input,.app-word input.docname') || {}).value || 'Document'], ['Author', () => 'Sahil Jadhav'], ['Date (today)', () => new Date().toLocaleDateString()], ['Page count', () => String(ed.pageCount || 1)], ['Word count', () => String(ed.stats().words)]].map(([l, f]) => ({ label: l, onClick: () => ed.insertText(f()) })), { sep: true }, { label: 'Save selection to Quick Part Gallery…', onClick: async () => { const r = sel(); if (!r || r.collapsed) return OA.toast('Select content first'); const d = document.createElement('div'); d.appendChild(r.cloneContents()); const n = await OA.prompt('Name:', d.textContent.slice(0, 24), { title: 'New Building Block' }); if (n) { qp.push({ n, h: d.innerHTML }); OA.toast('Saved “' + n + '” (kept until you close the site)'); } } }]);
    X.signature = async () => { const v = await OA.form('Signature Setup', [{ id: 'n', label: 'Suggested signer', value: '' }, { id: 't', label: 'Signer\'s title', value: '' }, { id: 'e', label: 'Signer\'s e-mail', value: '' }], { ok: 'OK' }); if (!v) return; X.insBlock(`<p class="sigline" contenteditable="false"><span class="sigx">✕</span><span class="sigl"></span><br><span style="font-size:10pt">${OA.esc(v.n)}${v.t ? '<br>' + OA.esc(v.t) : ''}${v.e ? '<br>' + OA.esc(v.e) : ''}</span></p><p><br></p>`); };

    /* ============ Layout: breaks, columns, line numbers, hyphenation, watermark, borders */
    X.sectionBreak = (kind) => X.insBlock(kind === 'page' ? '<div class="pb" contenteditable="false"></div>' : kind === 'next' ? '<div class="pb sect" contenteditable="false" data-k="next"></div>' : kind === 'cont' ? '<div class="sect cont" contenteditable="false" data-k="cont"></div>' : '<div class="colbrk" contenteditable="false"></div>');
    X.columns = n => ed.mutate(() => {
      const bl = ed.selBlocks(); if (!bl.length) return; const tops = Array.from(new Set(bl.map(b => { while (b.parentNode !== P && b.parentNode) b = b.parentNode; return b; })));
      const cs = tops.find(t => t.classList && t.classList.contains('colsec'));
      if (n === 1) { tops.forEach(t => { if (t.classList.contains('colsec')) t.replaceWith(...t.childNodes); }); return; }
      if (cs) { cs.style.columnCount = n; return; }
      const d = document.createElement('div'); d.className = 'colsec'; d.style.columnCount = n; d.style.columnGap = '36px'; tops[0].before(d); tops.forEach(t => d.appendChild(t));
    });
    X.lineNumbers = v => { P.classList.toggle('lnum', v); };
    X.hyphen = v => { P.style.hyphens = v ? 'auto' : 'manual'; P.setAttribute('lang', 'en'); };
    X.watermark = async () => {
      const cur = P.querySelector('.wmark'); const v = await OA.form('Printed Watermark', [{ id: 't', label: 'Text (empty to remove)', value: cur ? cur.textContent : 'DRAFT' }, { id: 'c', label: 'Colour', type: 'color', value: '#c0c0c0' }, { id: 'l', label: 'Layout', type: 'select', value: 'diag', options: [['diag', 'Diagonal'], ['horiz', 'Horizontal']] }]); if (!v) return;
      ed.mutate(() => { P.querySelectorAll('.wmark').forEach(w => w.remove()); if (v.t.trim()) { const w = document.createElement('div'); w.className = 'wmark'; w.contentEditable = 'false'; w.setAttribute('data-noexport', '1'); w.textContent = v.t; w.style.color = v.c; w.style.transform = v.l === 'diag' ? 'translate(-50%,-50%) rotate(-40deg)' : 'translate(-50%,-50%)'; P.insertBefore(w, P.firstChild); } });
    };
    X.pageBorder = async () => { const v = await OA.form('Borders and Shading — Page Border', [{ id: 's', label: 'Style', type: 'select', value: 'solid', options: ['none', 'solid', 'double', 'dashed', 'dotted', 'groove'] }, { id: 'w', label: 'Width (px)', type: 'number', value: 3, min: 1, max: 12 }, { id: 'c', label: 'Colour', type: 'color', value: '#2f5496' }]); if (!v) return; P.style.outline = v.s === 'none' ? '' : `${v.w}px ${v.s} ${v.c}`; P.style.outlineOffset = v.s === 'none' ? '' : '-24px'; ed.changed(true); };

    /* ============ Design: themes, colours, fonts */
    const THEMES = { Office: ['#2f5496', '#1f3763', 'Calibri'], Ion: ['#c0392b', '#7b241c', 'Georgia'], Facet: ['#2e7d32', '#1b5e20', 'Cambria'], Retrospect: ['#d35400', '#873600', 'Garamond'], Slate: ['#455a64', '#263238', 'Arial'], Violet: ['#6a1b9a', '#4a148c', 'Candara'] };
    function applyTheme(h1, h3, font, bodyFont) { const id = 'wthemecss'; let st = document.getElementById(id); if (!st) { st = el('style', { id }); document.head.appendChild(st); } st.textContent = `.w-page h1,.w-page h2,.w-page h4,.w-page h5{color:${h1}!important}.w-page h3,.w-page h6{color:${h3}!important}` + (font ? `.w-page h1,.w-page h2,.w-page h3,.w-page h4{font-family:'${font}',Calibri,sans-serif!important}` : '') + (bodyFont ? `.w-page{font-family:'${bodyFont}',Calibri,sans-serif}` : ''); ed.changed(true); }
    X.themes = anchor => OA.menu(anchor, Object.entries(THEMES).map(([n, t]) => ({ label: n, onClick: () => applyTheme(t[0], t[1], t[2], null) })));
    X.themeColors = anchor => OA.menu(anchor, [['Blue', '#2f5496', '#1f3763'], ['Red', '#c0392b', '#7b241c'], ['Green', '#2e7d32', '#1b5e20'], ['Orange', '#d35400', '#873600'], ['Purple', '#6a1b9a', '#4a148c'], ['Gray', '#455a64', '#263238'], ['Black', '#000000', '#000000']].map(([n, a, b]) => ({ label: n, style: 'color:' + a, onClick: () => applyTheme(a, b) })));
    X.themeFonts = anchor => OA.menu(anchor, [['Calibri Light / Calibri', 'Calibri Light', 'Calibri'], ['Arial / Arial', 'Arial', 'Arial'], ['Cambria / Calibri', 'Cambria', 'Calibri'], ['Georgia / Garamond', 'Georgia', 'Garamond'], ['Consolas / Courier New', 'Consolas', 'Courier New']].map(([n, h, b]) => ({ label: n, onClick: () => { const st = document.getElementById('wthemecss'); const c = st ? /color:(#\w+)!important}\.w-page h3,\.w-page h6\{color:(#\w+)/.exec(st.textContent) : null; applyTheme(c ? c[1] : '#2f5496', c ? c[2] : '#1f3763', h, b); } })));

    /* ============ Styles: create/modify, multilevel list, borders&shading, tabs, sort, outline, content controls, macros */
    const userStyles = X.styles = {};
    X.createStyle = async () => {
      const r = sel(); const bl = ed.selBlocks && ed.selBlocks()[0]; const cs = bl ? getComputedStyle(bl) : null;
      const v = await OA.form('Create New Style from Formatting', [{ id: 'n', label: 'Name', value: 'Style1' }, { id: 'f', label: 'Font', type: 'select', value: cs ? cs.fontFamily.split(',')[0].replace(/["']/g, '') : 'Calibri', options: Array.from(new Set([cs ? cs.fontFamily.split(',')[0].replace(/["']/g, '') : 'Calibri', ...OA.wordFonts])) }, { id: 's', label: 'Size (pt)', type: 'number', value: cs ? Math.round(parseFloat(cs.fontSize) * .75 * 2) / 2 : 11 }, { id: 'c', label: 'Colour', type: 'color', value: '#000000' }, { id: 'b', label: 'Bold', type: 'checkbox', value: !!cs && +cs.fontWeight >= 600 }, { id: 'i', label: 'Italic', type: 'checkbox', value: !!cs && cs.fontStyle === 'italic' }, { id: 'a', label: 'Alignment', type: 'select', value: cs ? (cs.textAlign === 'start' ? 'left' : cs.textAlign) : 'left', options: ['left', 'center', 'right', 'justify'] }, { id: 'sp', label: 'Space after (pt)', type: 'number', value: 8 }], { ok: 'Create' });
      if (!v) return; userStyles[v.n] = v; X.applyStyle(v.n); OA.toast('Style “' + v.n + '” created — find it under Styles → Custom'); };
    X.applyStyle = n => { const v = userStyles[n]; if (!v) return; ed.mutate(() => { ed.selBlocks().forEach(b => { b.style.fontFamily = v.f; b.style.fontSize = v.s + 'pt'; b.style.color = v.c; b.style.fontWeight = v.b ? '700' : ''; b.style.fontStyle = v.i ? 'italic' : ''; b.style.textAlign = v.a; b.style.marginBottom = v.sp + 'pt'; b.dataset.st = n; }); }); };
    X.stylesMenu = anchor => OA.menu(anchor, [{ label: 'Create a Style…', onClick: X.createStyle }, { label: 'Modify Style (update from selection)…', onClick: async () => { const k = Object.keys(userStyles); if (!k.length) return OA.toast('Create a style first'); const n = await OA.form('Modify Style', [{ id: 'n', label: 'Style', type: 'select', value: k[0], options: k }]); if (n) { const old = userStyles[n.n]; const v = await OA.form('Modify Style: ' + n.n, [{ id: 'f', label: 'Font', type: 'select', value: old.f, options: OA.wordFonts }, { id: 's', label: 'Size (pt)', type: 'number', value: old.s }, { id: 'c', label: 'Colour', type: 'color', value: old.c }, { id: 'b', label: 'Bold', type: 'checkbox', value: old.b }, { id: 'i', label: 'Italic', type: 'checkbox', value: old.i }, { id: 'a', label: 'Alignment', type: 'select', value: old.a, options: ['left', 'center', 'right', 'justify'] }, { id: 'sp', label: 'Space after (pt)', type: 'number', value: old.sp }]); if (v) { Object.assign(old, v); ed.mutate(() => { P.querySelectorAll('[data-st="' + n.n + '"]').forEach(b => { b.style.fontFamily = old.f; b.style.fontSize = old.s + 'pt'; b.style.color = old.c; b.style.fontWeight = old.b ? '700' : ''; b.style.fontStyle = old.i ? 'italic' : ''; b.style.textAlign = old.a; b.style.marginBottom = old.sp + 'pt'; }); }); } } } }, { sep: true }, { header: 'Custom styles' }, ...Object.keys(userStyles).map(n => ({ label: n, onClick: () => X.applyStyle(n) })), { label: 'Clear Formatting', onClick: () => ed.clearFormat() }]);
    X.multilevel = anchor => OA.menu(anchor, [['1) a) i)', ['decimal', 'lower-alpha', 'lower-roman']], ['1. 1.1 1.1.1', 'legal'], ['I. A. 1.', ['upper-roman', 'upper-alpha', 'decimal']], ['● ○ ■', null]].map(([l, t]) => ({ label: l, onClick: () => ed.mutate(() => {
      const bl = ed.selBlocks(); if (!bl.length) return; if (!t) { ed.list('ul'); return; }
      const ol = document.createElement('ol'); bl[0].before(ol); const mk = (list, text, lvl) => { const li = document.createElement('li'); li.textContent = text; list.appendChild(li); return li; };
      bl.forEach(b => { const lvl = Math.min(2, Math.round((parseFloat(b.style.marginLeft) || 0) / 36)); let host = ol; for (let k = 0; k < lvl; k++) { let last = host.lastElementChild; if (!last) last = mk(host, '', 0); let sub = last.querySelector(':scope > ol'); if (!sub) { sub = document.createElement('ol'); last.appendChild(sub); } host = sub; } mk(host, b.textContent, lvl); b.remove(); });
      if (t === 'legal') { ol.classList.add('legal'); } else { ol.style.listStyleType = t[0]; ol.querySelectorAll('ol').forEach(o => { const d = (() => { let n = 0, p = o; while ((p = p.parentNode.closest('ol'))) n++; return n; })(); o.style.listStyleType = t[Math.min(d, 2)]; }); }
    }) })));
    X.bordersShading = async () => {
      const v = await OA.form('Borders and Shading', [{ id: 'b', label: 'Border', type: 'select', value: 'box', options: [['none', 'None'], ['box', 'Box'], ['bottom', 'Bottom only'], ['top', 'Top only'], ['left', 'Left only']] }, { id: 's', label: 'Style', type: 'select', value: 'solid', options: ['solid', 'dashed', 'dotted', 'double'] }, { id: 'w', label: 'Width (px)', type: 'number', value: 1, min: 1, max: 10 }, { id: 'c', label: 'Border colour', type: 'color', value: '#000000' }, { id: 'f', label: 'Shading (fill) colour', type: 'color', value: '#ffffff' }, { id: 'nf', label: 'No shading', type: 'checkbox', value: true }]); if (!v) return;
      ed.mutate(() => ed.selBlocks().forEach(b => { b.style.border = ''; ['Top', 'Bottom', 'Left'].forEach(k => b.style['border' + k] = ''); const css = `${v.w}px ${v.s} ${v.c}`; if (v.b === 'box') { b.style.border = css; b.style.padding = '3px 5px'; } else if (v.b !== 'none') b.style['border' + v.b[0].toUpperCase() + v.b.slice(1)] = css; b.style.background = v.nf ? '' : v.f; }));
    };
    X.tabStops = async () => {
      const v = await OA.form('Tab Stops', [{ id: 'p', label: 'Tab stop position (cm) — adds a tab leader', type: 'number', value: 8, step: .5 }, { id: 'a', label: 'Alignment', type: 'select', value: 'left', options: ['left', 'center', 'right'] }, { id: 'l', label: 'Leader', type: 'select', value: 'none', options: [['none', 'None'], ['dots', '.......'], ['dash', '-------'], ['line', '_______']] }, { id: 'ins', label: 'Insert a tab at cursor', type: 'checkbox', value: true }], { ok: 'Set' }); if (!v) return;
      ed.mutate(() => { ed.selBlocks().forEach(b => { b.style.tabSize = (v.p * 37.8) + 'px'; b.dataset.tab = v.p; }); if (v.ins) document.execCommand('insertText', false, '\t'); });
      if (v.l !== 'none') OA.toast('Tab leaders are shown with a fill character in this version'); if (v.l !== 'none') ed.insertText({ dots: '.'.repeat(30), dash: '-'.repeat(30), line: '_'.repeat(30) }[v.l]);
    };
    X.outline = v => { P.classList.toggle('outline-view', v); if (v) { ed.showNav('headings'); } };
    X.contentControl = anchor => OA.menu(anchor, [{ label: 'Rich Text Content Control', onClick: () => X.insInline('<span class="cc" data-t="rich" title="Rich text">[Click to enter text]</span>&nbsp;') }, { label: 'Date Picker Content Control', onClick: () => X.insInline('<span class="cc" data-t="date" title="Click to pick a date">[Click to pick a date]</span>&nbsp;') }, { label: 'Drop-Down List Content Control…', onClick: async () => { const o = await OA.prompt('Choices (comma-separated):', 'Option 1, Option 2, Option 3', { title: 'Drop-Down List' }); if (o) X.insInline(`<span class="cc" data-t="list" data-o="${OA.esc(o)}" title="Click to choose">[Choose an item]</span>&nbsp;`); } }, { label: 'Check Box Content Control', onClick: () => X.insInline('<span class="cc" data-t="check" contenteditable="false">☐</span>&nbsp;') }]);
    P.addEventListener('click', e => {
      const c = e.target.closest && e.target.closest('.cc'); if (!c) return; const t = c.dataset.t;
      if (t === 'check') ed.mutate(() => { c.textContent = c.textContent === '☐' ? '☒' : '☐'; });
      else if (t === 'date') { const i = el('input', { type: 'date' }); OA.dialog({ title: 'Pick a date', content: i, buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { if (i.value) ed.mutate(() => { c.textContent = new Date(i.value + 'T00:00').toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }); c.classList.add('filled'); }); } }] }); }
      else if (t === 'list') OA.menu(c, c.dataset.o.split(',').map(s => s.trim()).map(s => ({ label: s, onClick: () => ed.mutate(() => { c.textContent = s; c.classList.add('filled'); }) })));
    });

    /* ---- Macros (JavaScript-based) */
    const macros = X.macros = JSON.parse((() => { try { return sessionStorage.getItem('wmacros') || '{}'; } catch (e) { return '{}'; } })());
    X.macroDialog = () => {
      const names = () => Object.keys(macros); const ta = el('textarea', { class: 'inp', rows: 9, style: { width: '100%', fontFamily: 'Consolas,monospace', fontSize: '12px' }, placeholder: '// Macro code (JavaScript). The object "ed" is the editor, e.g.\ned.insertText("Hello");\ned.cmd("bold");' });
      const sel2 = el('select', { class: 'inp', style: { width: '100%', marginBottom: '6px' } }); const fill = () => { sel2.innerHTML = '<option value="">— new macro —</option>' + names().map(n => `<option>${OA.esc(n)}</option>`).join(''); }; fill();
      sel2.addEventListener('change', () => { ta.value = macros[sel2.value] || ''; });
      const save = n => { macros[n] = ta.value; try { sessionStorage.setItem('wmacros', JSON.stringify(macros)); } catch (e) { } fill(); sel2.value = n; };
      OA.dialog({ title: 'Macros', width: 520, content: el('div', null, sel2, ta, el('div', { style: { fontSize: '11px', color: '#666', marginTop: '4px' }, text: 'Macros are JavaScript and live only until you close the tab. Only run code you trust.' })), buttons: [{ label: 'Save…', onClick: async () => { const n = sel2.value || await OA.prompt('Macro name:', 'Macro1', { title: 'Save Macro' }); if (n) save(n); return false; } }, { label: 'Delete', onClick: () => { if (sel2.value) { delete macros[sel2.value]; fill(); ta.value = ''; } return false; } }, { label: 'Close', value: false }, { label: 'Run', primary: true, onClick: () => { try { new Function('ed', 'OA', ta.value)(ed, OA); } catch (e) { OA.alert('Macro error: ' + OA.esc(e.message)); } } }] });
    };
    X.macroRecord = () => { const rec = X.recording; if (!rec) { X.recording = { start: P.innerHTML, steps: [] }; OA.toast('Recording: typed text and formatting are captured. Click again to stop.'); const h = e => { if (e.inputType === 'insertText' && e.data) X.recording.steps.push('ed.insertText(' + JSON.stringify(e.data) + ');'); }; X.recording.h = h; P.addEventListener('beforeinput', h); const oc = ed.cmd; X.recording.oc = oc; ed.cmd = function (n, v) { X.recording.steps.push('ed.cmd(' + JSON.stringify(n) + (v !== undefined ? ',' + JSON.stringify(v) : '') + ');'); return oc.apply(this, arguments); }; return; } P.removeEventListener('beforeinput', rec.h); ed.cmd = rec.oc; X.recording = null; (async () => { const n = await OA.prompt('Macro name:', 'Macro1', { title: 'Stop Recording' }); if (n) { macros[n] = rec.steps.join('\n').replace(/(ed\.insertText\("[^"]*"\);\n?)+/g, m => 'ed.insertText(' + JSON.stringify(m.match(/ed\.insertText\(("(?:[^"\\]|\\.)*")\)/g).map(s => JSON.parse(s.slice(14, -1))).join('')) + ');\n'); try { sessionStorage.setItem('wmacros', JSON.stringify(macros)); } catch (e) { } OA.toast('Macro “' + n + '” saved'); } })(); };
    X.macroMenu = anchor => OA.menu(anchor, [{ label: 'View Macros…', onClick: X.macroDialog }, { label: X.recording ? 'Stop Recording' : 'Record Macro', onClick: X.macroRecord }, { sep: true }, ...Object.keys(macros).map(n => ({ label: 'Run: ' + n, onClick: () => { try { new Function('ed', 'OA', macros[n])(ed, OA); } catch (e) { OA.alert(OA.esc(e.message)); } } }))]);

    /* ============ Ribbon wiring */
    spec.tabs.forEach(t => t.groups.forEach(g => g.items.forEach(it => { if (it.id === 'col') it.menu = [1, 2, 3, 4].map(n => ({ label: ['One', 'Two', 'Three', 'Four'][n - 1], onClick: () => { const r = sel(); if (r && !r.collapsed) X.columns(n); else ed.setLayout({ cols: n }); } })).concat([{ sep: true }, { label: 'Remove columns from selection', onClick: () => X.columns(1) }]); })));
    tabAdd('insert', [
      { label: 'Illustrations+', items: [L('chart', 'Chart', 'chart', X.chart), L('smart', 'SmartArt', 'columns', X.smartart)] },
      { label: 'Text+', items: [L('tbx', 'Text Box', 'file', (e, b) => X.textbox(b)), L('wart', 'WordArt', 'edit', (e, b) => X.wordart(b)), { type: 'row', items: [B('dcap', 'Drop Cap', null, null, { glyph: 'A', label: 'Drop Cap', menu: [{ label: 'None', onClick: () => X.dropcap('none') }, { label: 'Dropped', onClick: () => X.dropcap('drop') }, { label: 'In margin', onClick: () => X.dropcap('margin') }] }), B('qp', 'Quick Parts', null, (e, b) => X.quickParts(b), { glyph: '⚡', label: 'Quick Parts' })] }, { type: 'row', items: [B('sig', 'Signature Line', null, X.signature, { glyph: '✍', label: 'Signature Line' }), B('ccx', 'Content Controls', null, (e, b) => X.contentControl(b), { glyph: '▣', label: 'Content Control' })] }] },
      { label: 'Equations', items: [L('eq', 'Equation', 'sigma', X.equation)] }]);
    tabAdd('layout', [
      { label: 'Breaks & Columns', items: [L('brk', 'Breaks', 'pagebreak', null, { menu: [{ header: 'Page Breaks' }, { label: 'Page', onClick: () => X.sectionBreak('page') }, { label: 'Column', onClick: () => X.sectionBreak('col') }, { header: 'Section Breaks' }, { label: 'Next Page', onClick: () => X.sectionBreak('next') }, { label: 'Continuous', onClick: () => X.sectionBreak('cont') }] }), { type: 'check', id: 'lnx', label: 'Line Numbers', value: false, onChange: v => X.lineNumbers(v) }, { type: 'check', id: 'hyp', label: 'Hyphenation', value: false, onChange: v => X.hyphen(v) }] },
      { label: 'Page Background+', items: [L('wmk', 'Watermark', 'stamp', X.watermark), L('pbd', 'Page Borders', 'square', X.pageBorder)] },
      { label: 'Paragraph+', items: [L('bds', 'Borders & Shading', 'square', X.bordersShading), L('tbs', 'Tab Stops', 'indent', X.tabStops)] }]);
    spec.tabs.splice(spec.tabs.findIndex(t => t.id === 'layout') + 1, 0, { id: 'design', label: 'Design', groups: [
      { label: 'Document Formatting', items: [L('thm', 'Themes', 'palette', (e, b) => X.themes(b)), L('thc', 'Colors', 'palette', (e, b) => X.themeColors(b)), L('thf', 'Fonts', 'edit', (e, b) => X.themeFonts(b))] },
      { label: 'Page Background', items: [L('wmk2', 'Watermark', 'stamp', X.watermark), { type: 'color', id: 'pgc2', icon: 'bucket', title: 'Page Color', noColor: 'No Color', onPick: c => ed.setLayout({ pageColor: c }) }, L('pbd2', 'Page Borders', 'square', X.pageBorder)] }] });
    tabAdd('home', [{ label: 'Styles+', items: [L('stm', 'Styles', 'edit', (e, b) => X.stylesMenu(b))] }, { label: 'Lists+', items: [L('mlt', 'Multilevel List', 'list', (e, b) => X.multilevel(b)), B('srt', 'Sort', 'sort', () => ed.sortSel && ed.sortSel(), { label: 'Sort' })] }]);
    tabAdd('view', [{ label: 'Views+', items: [{ type: 'check', id: 'olv', label: 'Outline view', value: false, onChange: v => X.outline(v) }] }, { label: 'Macros', items: [L('mac', 'Macros', 'play', (e, b) => X.macroMenu(b))] }]);
  });
})();
