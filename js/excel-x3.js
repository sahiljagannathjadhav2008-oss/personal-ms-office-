/* Excel extras 3: Group/Subtotal, Flash Fill, Consolidate, Advanced Filter, Circle Invalid, JSON import, formula auditing, names, calc options, protection, Paste Special — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, XL = window.XL, el = OA.el, key = XL.key;
  XL.extras = XL.extras || []; XL.gridHooks = XL.gridHooks || [];
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const num = v => typeof v === 'number' && isFinite(v);
  const text = (ctx, sh, r, c) => { const v = ctx.wb.value(sh.name, r, c); return v === null || v === undefined ? '' : String(v); };
  const inputText = (sh, r, c) => { const x = sh.cells.get(key(r, c)); return x ? (x.f != null ? '=' + x.f : String(x.v == null ? '' : x.v)) : ''; };
  const colL = c => XL.colName(c);

  /* ---------- Flash Fill ---------- */
  const proper = s => s.toLowerCase().replace(/(^|[\s\-'])([a-z])/g, (m, a, b) => a + b.toUpperCase());
  const TOK = /([\s,;._@\/\-]+)/;
  const split = s => s.split(TOK);
  function learnFlash(pairs) {
    const cands = [];
    const words = s => s.split(/[\s,;]+/).filter(Boolean);
    cands.push(s => s.toUpperCase(), s => s.toLowerCase(), proper, s => s.replace(/\s+/g, ''), s => s.replace(/\D/g, ''), s => s.replace(/[^A-Za-z]/g, ''), s => s.trim(), s => words(s).map(w => w[0]).join('').toUpperCase(), s => words(s).map(w => w[0]).join(''), s => s.split('').reverse().join(''));
    for (let n = 1; n <= 8; n++) cands.push(s => s.slice(0, n), s => s.slice(-n), s => s.slice(n), s => s.slice(0, -n));
    for (let i = 0; i < 6; i++) { cands.push(s => words(s)[i] || '', s => words(s).slice(-1 - i)[0] || ''); cands.push(s => proper(words(s)[i] || ''), s => (words(s)[i] || '').toUpperCase()); }
    ['@', ' ', ',', '-', '_', '/', '.', ':'].forEach(d => { cands.push(s => { const i = s.indexOf(d); return i < 0 ? '' : s.slice(0, i); }, s => { const i = s.indexOf(d); return i < 0 ? '' : s.slice(i + d.length); }, s => { const i = s.lastIndexOf(d); return i < 0 ? '' : s.slice(i + d.length); }); });
    for (const f of cands) if (pairs.every(([a, b]) => { try { return f(a) === b && b !== ''; } catch (e) { return false; } })) return f;
    // token template learner: output made of source tokens (with case transforms) and literals
    const tr = [x => x, x => x.toUpperCase(), x => x.toLowerCase(), proper, x => x[0] || '', x => (x[0] || '').toUpperCase(), x => (x[0] || '').toUpperCase() + '.'];
    const src0 = split(pairs[0][0]).filter((t, i) => i % 2 === 0), out0 = split(pairs[0][1]); const plan = [];
    for (let i = 0; i < out0.length; i++) {
      const o = out0[i]; if (i % 2 === 1) { plan.push({ lit: o }); continue; } let found = null;
      for (let si = 0; si < src0.length && !found; si++) for (let ti = 0; ti < tr.length && !found; ti++) if (tr[ti](src0[si]) === o && pairs.every(([a, b]) => { const sp = split(a).filter((t, k) => k % 2 === 0), op = split(b); return sp[si] !== undefined && op[i] !== undefined && tr[ti](sp[si]) === op[i]; })) found = { si, ti };
      plan.push(found || { lit: o });
    }
    const f = s => { const sp = split(s).filter((t, k) => k % 2 === 0); return plan.map(p => p.lit !== undefined ? p.lit : (sp[p.si] === undefined ? '' : tr[p.ti](sp[p.si]))).join(''); };
    if (pairs.every(([a, b]) => f(a) === b) && plan.some(p => p.lit === undefined)) return f;
    return null;
  }
  function flashFill(ctx) {
    const { wb, g } = ctx, s = ctx.sh(), a = g.active(); const col = a.c; let src = col - 1; if (src < 0) return OA.toast('Flash Fill ke liye source column chahiye (left side).');
    let top = a.r; while (top > 0 && inputText(s, top - 1, col) !== '' && text(ctx, s, top - 1, src) !== '') top--;
    const pairs = []; for (let r = top; r <= a.r; r++) { const o = inputText(s, r, col), i = text(ctx, s, r, src); if (o !== '' && i !== '') pairs.push([i, o]); }
    if (!pairs.length) return OA.toast('Pehle ek example output type kijiye (source ke bagal wali cell me).');
    const f = learnFlash(pairs); if (!f) return OA.alert('Pattern samajh nahi aaya. 2–3 examples type karke dobara try kijiye.');
    let r = a.r + (inputText(s, a.r, col) === '' ? 0 : 1), n = 0; wb.tx(() => { for (; ; r++) { const i = text(ctx, s, r, src); if (i === '') break; if (inputText(s, r, col) !== '' && r <= a.r) continue; wb.putInput(s, r, col, f(i)); n++; } }); ctx.after(); OA.toast('Flash Fill: ' + n + ' cells bhari');
  }

  /* ---------- Group / outline / subtotal ---------- */
  function groupSel(ctx, kind) { const s = ctx.sh(), q = ctx.g.range(), wb = ctx.wb; const key2 = kind === 'col' ? 'cgroups' : 'groups'; const g1 = kind === 'col' ? q.c1 : q.r1, g2 = kind === 'col' ? q.c2 : q.r2; wb.tx(() => wb.setProp(s, key2, (s[key2] || []).concat([{ a: g1, b: g2 }]))); OA.toast((kind === 'col' ? 'Columns ' : 'Rows ') + (g1 + 1) + '–' + (g2 + 1) + ' grouped'); }
  function ungroupSel(ctx, kind) { const s = ctx.sh(), q = ctx.g.range(), key2 = kind === 'col' ? 'cgroups' : 'groups'; const a = kind === 'col' ? q.c1 : q.r1, b = kind === 'col' ? q.c2 : q.r2; ctx.wb.tx(() => ctx.wb.setProp(s, key2, (s[key2] || []).filter(x => x.b < a || x.a > b))); }
  function detail(ctx, hide, all) { const s = ctx.sh(), q = ctx.g.range(), wb = ctx.wb; const hr = new Set(s.hiddenR), hc = new Set(s.hiddenC); (s.groups || []).forEach(x => { if (all || (q.r1 >= x.a - 1 && q.r2 <= x.b + 1)) for (let r = x.a; r <= x.b; r++) hide ? hr.add(r) : hr.delete(r); }); (s.cgroups || []).forEach(x => { if (all || (q.c1 >= x.a - 1 && q.c2 <= x.b + 1)) for (let c = x.a; c <= x.b; c++) hide ? hc.add(c) : hc.delete(c); }); wb.tx(() => { wb.setProp(s, 'hiddenR', hr); wb.setProp(s, 'hiddenC', hc); }); ctx.after(); }
  async function subtotal(ctx) {
    const s = ctx.sh(), wb = ctx.wb; let q = ctx.g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, s, q.r1, q.c1); wb.ensureCalc();
    const heads = []; for (let c = q.c1; c <= q.c2; c++) heads.push([c, text(ctx, s, q.r1, c) || colL(c)]);
    const FN = { Sum: 9, Count: 2, Average: 1, Max: 4, Min: 5 };
    const v = await OA.form('Subtotal', [{ id: 'by', label: 'At each change in', type: 'select', options: heads, value: heads[0][0] }, { id: 'fn', label: 'Use function', type: 'select', options: Object.keys(FN), value: 'Sum' }, { id: 'cols', label: 'Add subtotal to (column letters, comma)', value: colL(q.c2) }, { id: 'grand', label: 'Grand total', type: 'checkbox', value: true }], { width: 400 }); if (!v) return;
    const by = +v.by, tc = v.cols.split(/[,\s]+/).filter(Boolean).map(x => XL.colIndex ? XL.colIndex(x.toUpperCase()) : (XL.parseRange(x.toUpperCase() + '1') || {}).c1).filter(x => x != null && x >= q.c1 && x <= q.c2); if (!tc.length) return OA.alert('Subtotal column sahi dein.');
    const groups = []; let st = q.r1 + 1; for (let r = q.r1 + 1; r <= q.r2; r++) { if (r === q.r2 || text(ctx, s, r + 1, by) !== text(ctx, s, r, by)) { groups.push({ a: st, b: r, name: text(ctx, s, r, by) }); st = r + 1; } }
    wb.tx(() => { let shift = 0; const detailGroups = []; groups.forEach(gr => { const at = gr.b + 1 + shift; wb.insertDelete(s, 'row', at, 1); wb.putInput(s, at, by, gr.name + ' ' + (v.fn === 'Average' ? 'Average' : v.fn === 'Count' ? 'Count' : v.fn === 'Sum' ? 'Total' : v.fn)); tc.forEach(c => wb.putInput(s, at, c, '=SUBTOTAL(' + FN[v.fn] + ',' + XL.addr(gr.a + shift, c) + ':' + XL.addr(gr.b + shift, c) + ')')); wb.applyStyle(s, { r1: at, c1: q.c1, r2: at, c2: q.c2 }, { b: true, bg: '#e2efda' }); detailGroups.push({ a: gr.a + shift, b: gr.b + shift }); shift++; });
      if (v.grand) { const at = q.r2 + 1 + shift; wb.insertDelete(s, 'row', at, 1); wb.putInput(s, at, by, 'Grand ' + (v.fn === 'Sum' ? 'Total' : v.fn)); tc.forEach(c => wb.putInput(s, at, c, '=SUBTOTAL(' + FN[v.fn] + ',' + XL.addr(q.r1 + 1, c) + ':' + XL.addr(q.r2 + shift, c) + ')')); wb.applyStyle(s, { r1: at, c1: q.c1, r2: at, c2: q.c2 }, { b: true, bg: '#c6e0b4' }); }
      wb.setProp(s, 'groups', (s.groups || []).concat(detailGroups)); }); ctx.after(); OA.toast('Subtotals insert ho gaye');
  }
  async function consolidate(ctx) {
    const wb = ctx.wb; const v = await OA.form('Consolidate', [{ id: 'fn', label: 'Function', type: 'select', options: ['Sum', 'Average', 'Count', 'Max', 'Min'], value: 'Sum' }, { id: 'refs', label: 'References (one per line, e.g. Sheet1!A1:B5)', type: 'textarea', rows: 4, value: ctx.sh().name + '!' + XL.rangeLabel(ctx.g.range()) }, { id: 'top', label: 'Top row labels', type: 'checkbox', value: true }, { id: 'left', label: 'Left column labels', type: 'checkbox', value: true }], { width: 440 }); if (!v) return;
    const refs = v.refs.split('\n').map(x => x.trim()).filter(Boolean).map(t => { const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(t); return { sh: m[1] ? wb.sheet(m[1]) : ctx.sh(), g: XL.parseRange(m[2]) }; }).filter(x => x.sh && x.g); if (!refs.length) return OA.alert('Valid reference dein.');
    wb.ensureCalc(); const tops = [], lefts = [], map = new Map(); const A = { Sum: a => a.reduce((x, y) => x + y, 0), Average: a => a.reduce((x, y) => x + y, 0) / a.length, Count: a => a.length, Max: a => Math.max(...a), Min: a => Math.min(...a) }[v.fn];
    refs.forEach(rf => { const { sh, g } = rf; const r0 = v.top ? g.r1 + 1 : g.r1, c0 = v.left ? g.c1 + 1 : g.c1; for (let r = r0; r <= g.r2; r++) for (let c = c0; c <= g.c2; c++) { const x = wb.value(sh.name, r, c); if (!num(x)) continue; const rk = v.left ? text(ctx, sh, r, g.c1) : String(r - r0), ck = v.top ? text(ctx, sh, g.r1, c) : String(c - c0); if (!lefts.includes(rk)) lefts.push(rk); if (!tops.includes(ck)) tops.push(ck); const k = rk + '\u0001' + ck; (map.get(k) || map.set(k, []).get(k)).push(x); } });
    const out = []; out.push([''].concat(v.top ? tops : tops.map((_, i) => 'Col' + (i + 1)))); lefts.forEach(l => out.push([v.left ? l : ''].concat(tops.map(t => { const a = map.get(l + '\u0001' + t); return a ? String(+A(a).toFixed(6)) : ''; }))));
    let ns; wb.tx(() => { ns = wb.addSheet(wb.uniqueName('Consolidate'), wb.sheets.length); wb.loadRows(ns, out); wb.applyStyle(ns, { r1: 0, c1: 0, r2: 0, c2: out[0].length - 1 }, { b: true, color: '#ffffff', bg: '#217346' }); }); ctx.activate(ns.name);
  }
  function critMatch(cv, crit) { crit = String(crit); if (crit === '') return true; const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(crit); const op = m[1] || '=', rhs = m[2]; const n = parseFloat(rhs), isN = rhs !== '' && !isNaN(n) && isFinite(rhs); const lv = typeof cv === 'number' ? cv : parseFloat(cv); if (isN && !isNaN(lv) && (op !== '=' || typeof cv === 'number')) return { '<': lv < n, '>': lv > n, '<=': lv <= n, '>=': lv >= n, '=': lv === n, '<>': lv !== n }[op]; const re = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i'); const t = re.test(String(cv)); return op === '<>' ? !t : t; }
  async function advFilter(ctx) {
    const s = ctx.sh(), wb = ctx.wb; let q = ctx.g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, s, q.r1, q.c1);
    const v = await OA.form('Advanced Filter', [{ id: 'list', label: 'List range', value: XL.rangeLabel(q) }, { id: 'crit', label: 'Criteria range (header + rows)', value: '' }, { id: 'act', label: 'Action', type: 'select', options: [['in', 'Filter the list, in-place'], ['copy', 'Copy to another location']], value: 'in' }, { id: 'to', label: 'Copy to (top-left cell)', value: '' }, { id: 'uniq', label: 'Unique records only', type: 'checkbox' }], { width: 420 }); if (!v) return;
    const L = XL.parseRange(v.list), C = XL.parseRange(v.crit); if (!L || !C) return OA.alert('List aur Criteria range dein.'); wb.ensureCalc();
    const heads = []; for (let c = L.c1; c <= L.c2; c++) heads.push(text(ctx, s, L.r1, c)); const crits = []; for (let r = C.r1 + 1; r <= C.r2; r++) { const row = []; for (let c = C.c1; c <= C.c2; c++) { const t = text(ctx, s, r, c); if (t !== '') row.push([heads.indexOf(text(ctx, s, C.r1, c)), inputText(s, r, c)]); } crits.push(row); }
    const match = r => !crits.length || crits.some(row => row.every(([ci, cr]) => ci >= 0 && critMatch(wb.value(s.name, r, L.c1 + ci), cr))); const seen = new Set(), hits = [];
    for (let r = L.r1 + 1; r <= L.r2; r++) { if (!match(r)) continue; if (v.uniq) { const k = heads.map((_, i) => text(ctx, s, r, L.c1 + i)).join('\u0001'); if (seen.has(k)) continue; seen.add(k); } hits.push(r); }
    if (v.act === 'in') { const hid = new Set(s.hiddenR); for (let r = L.r1 + 1; r <= L.r2; r++) hits.includes(r) ? hid.delete(r) : hid.add(r); wb.tx(() => wb.setProp(s, 'hiddenR', hid)); ctx.after(); OA.toast(hits.length + ' records match'); }
    else { const T = XL.parseRange(v.to); if (!T) return OA.alert('Copy-to cell dein.'); wb.tx(() => { for (let c = 0; c < heads.length; c++) wb.putInput(s, T.r1, T.c1 + c, heads[c]); hits.forEach((r, i) => { for (let c = 0; c < heads.length; c++) wb.putInput(s, T.r1 + 1 + i, T.c1 + c, inputText(s, r, L.c1 + c)); }); }); ctx.after(); OA.toast(hits.length + ' records copy hue'); }
  }
  async function importJSON(ctx, urlMode) {
    let t; if (urlMode) { const u = await OA.prompt('URL (CSV ya JSON; CORS allow hona chahiye):', 'https://', { title: 'From Web' }); if (!u) return; try { t = await (await fetch(u)).text(); } catch (e) { return OA.alert('Download nahi hua (network ya CORS): ' + esc(e.message)); } }
    else { const f = await OA.pickFile('.json,application/json'); if (!f) return; t = await f.text(); }
    let rows; try { const j = JSON.parse(t); const arr = Array.isArray(j) ? j : (Object.values(j).find(Array.isArray) || [j]); const keys = []; arr.forEach(o => { if (o && typeof o === 'object') Object.keys(o).forEach(k => { if (!keys.includes(k)) keys.push(k); }); }); rows = keys.length ? [keys].concat(arr.map(o => keys.map(k => { const x = o[k]; return x == null ? '' : typeof x === 'object' ? JSON.stringify(x) : String(x); }))) : arr.map(x => [String(x)]); }
    catch (e) { if (urlMode) rows = t.split(/\r?\n/).filter(Boolean).map(l => l.split(',')); else return OA.alert('JSON parse nahi hua: ' + esc(e.message)); }
    const p = ctx.g.active(); ctx.wb.loadRows(ctx.sh(), rows, p.r, p.c); ctx.after(); OA.toast(rows.length + ' rows import hui');
  }

  /* ---------- auditing overlay (trace arrows + invalid-data circles) ---------- */
  const refsIn = (f, curSheet) => { const out = []; const re = /(?:(?:'([^']+)'|([A-Za-z_][\w.]*))!)?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/g; let m; while ((m = re.exec(f))) { if (m[1] || m[2]) { const nm = m[1] || m[2]; if (nm !== curSheet) continue; } const g = XL.parseRange(m[3] + m[4] + (m[5] ? ':' + m[5] + m[6] : '')); if (g) out.push(g); } return out; };
  function trace(ctx, dir) {
    const s = ctx.sh(), a = ctx.g.active(); ctx.audit = ctx.audit || { arrows: [], circles: [] }; const here = { r: a.r, c: a.c }; let found = 0;
    if (dir === 'pre') { const cl = s.cells.get(key(a.r, a.c)); if (!cl || cl.f == null) return OA.toast('Is cell me formula nahi hai'); refsIn(cl.f, s.name).forEach(g => { ctx.audit.arrows.push({ from: g, to: here }); found++; }); }
    else s.cells.forEach((cl, k) => { if (cl.f == null) return; const u = XL.unkey(k); if (refsIn(cl.f, s.name).some(g => a.r >= g.r1 && a.r <= g.r2 && a.c >= g.c1 && a.c <= g.c2)) { ctx.audit.arrows.push({ from: { r1: a.r, c1: a.c, r2: a.r, c2: a.c }, to: u }); found++; } });
    if (!found) OA.toast(dir === 'pre' ? 'Koi precedent nahi mila' : 'Koi dependent nahi mila'); ctx.g.overlayAudit(); }
  async function evaluateFormula(ctx) {
    const s = ctx.sh(), a = ctx.g.active(), cl = s.cells.get(key(a.r, a.c)); if (!cl || cl.f == null) return OA.alert('Formula wali cell select kijiye.'); ctx.wb.ensureCalc();
    const f0 = '=' + cl.f; let stage = 0; const stages = [f0]; stages.push(f0.replace(/(?:'[^']+'!|[A-Za-z_][\w.]*!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/g, m => { const g = XL.parseRange(m.replace(/^.*!/, '').replace(/\$/g, '')); if (!g) return m; if (g.r1 === g.r2 && g.c1 === g.c2) { const v = ctx.wb.value(s.name, g.r1, g.c1); return typeof v === 'string' ? '"' + v + '"' : String(v); } const vs = []; for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 7); r++) for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 7); c++) vs.push(ctx.wb.value(s.name, r, c)); return '{' + vs.join(',') + (((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1)) > 64 ? ',…' : '') + '}'; }));
    stages.push('= ' + ctx.wb.display(s, a.r, a.c).text);
    const box = el('div', { style: 'font-family:Consolas,monospace;font-size:14px;padding:10px;background:#f6f6f6;border:1px solid #ccc;min-height:44px;word-break:break-all' }); const lab = el('div', { class: 'form-intro' });
    const show = () => { box.textContent = stages[stage]; lab.textContent = 'Step ' + (stage + 1) + ' / ' + stages.length + (stage === 0 ? ' — original formula' : stage === 1 ? ' — references replaced by values' : ' — final result'); };
    show(); OA.dialog({ title: 'Evaluate Formula — ' + XL.addr(a.r, a.c), content: el('div', {}, [lab, box]), width: 480, buttons: [{ label: 'Restart', onClick: () => { stage = 0; show(); return false; } }, { label: 'Evaluate ▶', primary: true, onClick: () => { if (stage < stages.length - 1) { stage++; show(); } return false; } }, { label: 'Close' }] });
  }
  function errorCheck(ctx) {
    const s = ctx.sh(), wb = ctx.wb; wb.ensureCalc(); const errs = []; s.cells.forEach((cl, k) => { const u = XL.unkey(k), v = wb.value(s.name, u.r, u.c); if (v && typeof v === 'object' && v.err || (typeof v === 'string' && /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/.test(v))) errs.push({ r: u.r, c: u.c, t: typeof v === 'string' ? v : (v.err || v.toString()) }); else if (cl.f != null && typeof cl.f === 'string' && /^\s*\d+\s*[+\-*\/]\s*\d+\s*$/.test(cl.f) && false) { } });
    errs.sort((x, y) => x.r - y.r || x.c - y.c); if (!errs.length) return OA.alert('Error Checking: koi formula error nahi mila. ✔');
    const list = el('div', { style: 'max-height:260px;overflow:auto' }); errs.forEach(e => list.appendChild(el('div', { style: 'padding:4px 6px;border-bottom:1px solid #eee;cursor:pointer', html: '<b>' + XL.addr(e.r, e.c) + '</b> — ' + esc(e.t) + ' <span style="color:#888">=' + esc((s.cells.get(key(e.r, e.c)) || {}).f || '') + '</span>', onclick: () => { ctx.g.setActive(e.r, e.c); ctx.g.ensureVisible(e.r, e.c); } })));
    OA.dialog({ title: 'Error Checking (' + errs.length + ')', content: list, width: 480, buttons: [{ label: 'Close', primary: true }] });
  }
  function watchWindow(ctx) {
    ctx.watch = ctx.watch || []; const body = el('div', { style: 'max-height:240px;overflow:auto' });
    const draw = () => { ctx.wb.ensureCalc(); body.innerHTML = ''; const t = el('table', { style: 'width:100%;border-collapse:collapse' }); t.innerHTML = '<tr style="background:#eee"><th align=left>Sheet</th><th align=left>Cell</th><th align=left>Value</th><th align=left>Formula</th><th></th></tr>'; ctx.watch.forEach((w, i) => { const sh = ctx.wb.sheet(w.s); const tr = el('tr'); tr.innerHTML = `<td>${esc(w.s)}</td><td>${XL.addr(w.r, w.c)}</td><td>${esc(sh ? text(ctx, sh, w.r, w.c) : '')}</td><td>${esc(sh ? inputText(sh, w.r, w.c) : '')}</td>`; const d = el('td'); d.appendChild(el('button', { class: 'btn', type: 'button', text: '✕', onclick: () => { ctx.watch.splice(i, 1); draw(); } })); tr.appendChild(d); t.appendChild(tr); }); body.appendChild(t); if (!ctx.watch.length) body.appendChild(el('div', { class: 'form-intro', text: 'Koi watch nahi. "Add Watch" se active cell add karein.' })); };
    draw(); OA.dialog({ title: 'Watch Window', content: body, width: 520, buttons: [{ label: 'Add Watch', onClick: () => { const a = ctx.g.active(); ctx.watch.push({ s: ctx.sh().name, r: a.r, c: a.c }); draw(); return false; } }, { label: 'Refresh', onClick: () => { draw(); return false; } }, { label: 'Close', primary: true }] });
  }
  async function createNames(ctx) {
    const q = ctx.g.range(), s = ctx.sh(), wb = ctx.wb; const v = await OA.form('Create Names from Selection', [{ id: 'top', label: 'Top row', type: 'checkbox', value: true }, { id: 'left', label: 'Left column', type: 'checkbox' }, { id: 'bottom', label: 'Bottom row', type: 'checkbox' }, { id: 'right', label: 'Right column', type: 'checkbox' }]); if (!v) return; let n = 0; const san = t => String(t).trim().replace(/[^\w.]/g, '_').replace(/^(\d)/, '_$1'); const ref = (r1, c1, r2, c2) => "'" + s.name + "'!" + XL.rangeLabel({ r1, c1, r2, c2 }).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
    wb.tx(() => { if (v.top) for (let c = q.c1; c <= q.c2; c++) { const t = text(ctx, s, q.r1, c); if (t && q.r2 > q.r1) { wb.defineName(san(t), ref(q.r1 + 1, c, q.r2, c)); n++; } } if (v.left) for (let r = q.r1; r <= q.r2; r++) { const t = text(ctx, s, r, q.c1); if (t && q.c2 > q.c1) { wb.defineName(san(t), ref(r, q.c1 + 1, r, q.c2)); n++; } } if (v.bottom) for (let c = q.c1; c <= q.c2; c++) { const t = text(ctx, s, q.r2, c); if (t && q.r2 > q.r1) { wb.defineName(san(t), ref(q.r1, c, q.r2 - 1, c)); n++; } } if (v.right) for (let r = q.r1; r <= q.r2; r++) { const t = text(ctx, s, r, q.c2); if (t && q.c2 > q.c1) { wb.defineName(san(t), ref(r, q.c1, r, q.c2 - 1)); n++; } } }); OA.toast(n + ' names bane');
  }
  async function pasteSpecial(ctx) {
    const g = ctx.g; if (!g.clip) return OA.alert('Pehle koi range copy (Ctrl+C) kijiye.');
    const v = await OA.form('Paste Special', [{ id: 'what', label: 'Paste', type: 'select', options: [['all', 'All'], ['values', 'Values'], ['formulas', 'Formulas'], ['formats', 'Formats']], value: 'all' }, { id: 'op', label: 'Operation', type: 'select', options: [['', 'None'], ['+', 'Add'], ['-', 'Subtract'], ['*', 'Multiply'], ['/', 'Divide']], value: '' }, { id: 'blank', label: 'Skip blanks', type: 'checkbox' }, { id: 'tr', label: 'Transpose', type: 'checkbox' }]); if (!v) return;
    const wb = ctx.wb, s = ctx.sh(), q = g.range(), clip = g.clip;
    if (v.op) { const src = wb.sheet(clip.sheet) || s; wb.ensureCalc(); wb.tx(() => { for (let i = 0; i <= clip.g.r2 - clip.g.r1; i++) for (let j = 0; j <= clip.g.c2 - clip.g.c1; j++) { const x = wb.value(src.name, clip.g.r1 + i, clip.g.c1 + j); if (!num(x)) continue; const r = q.r1 + i, c = q.c1 + j, cur = wb.value(s.name, r, c), y = num(cur) ? cur : 0; wb.putInput(s, r, c, String(v.op === '+' ? y + x : v.op === '-' ? y - x : v.op === '*' ? y * x : (x === 0 ? NaN : y / x))); } }); ctx.after(); return; }
    g.pasteButton(v.what, { transpose: v.tr, skipBlanks: v.blank });
  }
  function goToSpecial(ctx) {
    OA.form('Go To Special', [{ id: 'k', label: 'Select', type: 'select', options: [['blanks', 'Blanks'], ['formulas', 'Formulas'], ['constants', 'Constants'], ['errors', 'Errors'], ['notes', 'Comments'], ['numbers', 'Numbers'], ['text', 'Text']], value: 'blanks' }]).then(v => {
      if (!v) return; const s = ctx.sh(), wb = ctx.wb; wb.ensureCalc(); const u = ctx.g.range().r1 === ctx.g.range().r2 && ctx.g.range().c1 === ctx.g.range().c2 ? wb.usedRange(s) : ctx.g.range(); if (!u || u.r2 < 0) return; const hits = [];
      for (let r = u.r1; r <= u.r2; r++) for (let c = u.c1; c <= u.c2; c++) { const cl = s.cells.get(key(r, c)), x = wb.value(s.name, r, c); let ok = false; if (v.k === 'blanks') ok = !cl || cl.v === '' || cl.v == null && cl.f == null; else if (v.k === 'formulas') ok = !!cl && cl.f != null; else if (v.k === 'constants') ok = !!cl && cl.f == null && cl.v !== '' && cl.v != null; else if (v.k === 'errors') ok = typeof x === 'string' && /^#[A-Z\/0!?]+/.test(x); else if (v.k === 'notes') ok = s.notes.has(key(r, c)); else if (v.k === 'numbers') ok = num(x); else ok = typeof x === 'string' && x !== '' && !/^#/.test(x); if (ok) hits.push({ r1: r, c1: c, r2: r, c2: c }); }
      if (!hits.length) return OA.toast('Koi cell nahi mili'); ctx.g.select ? (ctx.g.select(hits[0]), hits.length > 1 && ctx.g.addRange && 0) : 0; if (ctx.g.setRanges) ctx.g.setRanges(hits.slice(0, 2000)); OA.toast(hits.length + ' cells mili (pehli select ki)');
    });
  }
  function accessibility(ctx) {
    const issues = []; ctx.wb.sheets.forEach(s => { s.charts.forEach(c => { if (!c.alt && (c.kind === 'picture' || !c.obj)) issues.push(s.name + ': ' + (c.obj ? 'Picture' : 'Chart') + ' me alt text nahi hai'); }); const m = s.merges ? s.merges.length : 0; if (m) issues.push(s.name + ': ' + m + ' merged cells hain (screen reader ke liye mushkil)'); if (/^Sheet\d+$/.test(s.name)) issues.push(s.name + ': generic sheet name — descriptive name dein'); const u = ctx.wb.usedRange(s); if (u && u.r2 >= 0) { let blankHead = 0; for (let c = u.c1; c <= u.c2; c++) if (text(ctx, s, u.r1, c) === '') blankHead++; if (blankHead) issues.push(s.name + ': header row me ' + blankHead + ' blank cells'); } });
    OA.dialog({ title: 'Accessibility Checker', content: el('div', { class: 'dlg-msg', html: issues.length ? '<ul style="padding-left:18px">' + issues.map(i => '<li>' + esc(i) + '</li>').join('') + '</ul>' : 'Koi accessibility issue nahi mila. ✔' }), buttons: [{ label: 'Close', primary: true }] });
  }
  function workbookStats(ctx) { let cells = 0, f = 0, ch = 0, nt = 0; ctx.wb.sheets.forEach(s => { s.cells.forEach(c => { cells++; if (c.f != null) f++; }); ch += s.charts.length; nt += s.notes.size; }); OA.alert('<b>Sheets:</b> ' + ctx.wb.sheets.length + '<br><b>Cells with data:</b> ' + cells + '<br><b>Formulas:</b> ' + f + '<br><b>Charts/objects:</b> ' + ch + '<br><b>Comments:</b> ' + nt + '<br><b>Named ranges:</b> ' + Object.keys(ctx.wb.names).length, 'Workbook Statistics'); }
  function circleInvalid(ctx) {
    const s = ctx.sh(), wb = ctx.wb; wb.ensureCalc(); ctx.audit = ctx.audit || { arrows: [], circles: [] }; let n = 0;
    (s.dv || []).forEach(rule => { const q = rule.range; for (let r = q.r1; r <= q.r2; r++) for (let c = q.c1; c <= q.c2; c++) { const t = text(ctx, s, r, c); if (t === '') continue; if (wb.dvCheck(s, r, c, t)) { ctx.audit.circles.push({ r, c }); n++; } } });
    ctx.g.overlayAudit(); OA.toast(n ? n + ' invalid cells circle ki gayi' : 'Koi invalid data nahi mila');
  }

  XL.x3 = { learnFlash, critMatch, refsIn };
  XL.gridHooks.push(function (g, ctx) {
    const host = ctx.H.host, ov = el('div'); ov.style.cssText = 'position:absolute;inset:0;pointer-events:none;z-index:5;overflow:hidden'; host.appendChild(ov);
    g.overlayAudit = function () {
      ov.innerHTML = ''; const a = ctx.audit; if (!a || (!a.arrows.length && !a.circles.length)) return; const z = g.zoom, ox = g.showHeaders ? 48 * z : 0, oy = g.showHeaders ? 22 * z : 0; const mid = (r, c) => { const x = g.cellRect(r, c); return [x.x + ox + x.w / 2, x.y + oy + x.h / 2, x]; };
      let svg = '<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" style="position:absolute;inset:0"><defs><marker id="ta" markerWidth="8" markerHeight="8" refX="7" refY="4" orient="auto"><path d="M0,0 L8,4 L0,8 Z" fill="#1f4e9c"/></marker></defs>';
      a.arrows.forEach(ar => { const f = g.cellRect(ar.from.r1, ar.from.c1), f2 = g.cellRect(ar.from.r2, ar.from.c2); const fx = f.x + ox, fy = f.y + oy, fw = f2.x + f2.w - f.x, fh = f2.y + f2.h - f.y; const [tx, ty] = mid(ar.to.r, ar.to.c); if (ar.from.r1 !== ar.from.r2 || ar.from.c1 !== ar.from.c2) svg += `<rect x="${fx}" y="${fy}" width="${fw}" height="${fh}" fill="none" stroke="#1f4e9c" stroke-width="1.5" stroke-dasharray="4 3"/>`; svg += `<circle cx="${fx + 4}" cy="${fy + fh / 2}" r="3" fill="#1f4e9c"/><line x1="${fx + 4}" y1="${fy + fh / 2}" x2="${tx}" y2="${ty}" stroke="#1f4e9c" stroke-width="1.6" marker-end="url(#ta)"/>`; });
      a.circles.forEach(ci => { const [, , x] = mid(ci.r, ci.c); svg += `<ellipse cx="${x.x + ox + x.w / 2}" cy="${x.y + oy + x.h / 2}" rx="${x.w / 2 + 4}" ry="${x.h / 2 + 3}" fill="none" stroke="#e00" stroke-width="2"/>`; });
      ov.innerHTML = svg + '</svg>';
    };
    const od = g.draw; g.draw = function () { od(); if (ctx.audit && (ctx.audit.arrows.length || ctx.audit.circles.length)) setTimeout(g.overlayAudit, 0); };
    // manual calculation support
    const wbk = ctx.wb; if (!wbk._ec) { wbk._ec = wbk.ensureCalc; wbk.ensureCalc = function () { if (ctx.calcManual && !this._force) return; this._ec(); }; }
  });
  XL.extras.push(function (ctx, spec, H) {
    const { B, L } = H; const tab = id => spec.tabs.find(t => t.id === id); const add = (id, groups) => { const t = tab(id); if (t) t.groups.push(...groups); }; const wb = () => ctx.wb;
    add('data', [
      { label: 'Data Tools', items: [L('ffl', 'Flash Fill', 'play', () => flashFill(ctx)), L('cns', 'Consolidate', 'layers', () => consolidate(ctx)), L('afl', 'Advanced Filter', 'filter', () => advFilter(ctx)), L('civ', 'Circle Invalid Data', 'eye', null, { menu: [{ label: 'Circle Invalid Data', onClick: () => circleInvalid(ctx) }, { label: 'Clear Validation Circles', onClick: () => { ctx.audit && (ctx.audit.circles = []); ctx.g.overlayAudit(); } }] })] },
      { label: 'Get Data', items: [L('gjs', 'From JSON', 'file', () => importJSON(ctx, false)), L('gwb', 'From Web', 'globe', () => importJSON(ctx, true))] },
      { label: 'Outline', items: [L('grp', 'Group', 'layers', null, { menu: [{ label: 'Group Rows', onClick: () => groupSel(ctx, 'row') }, { label: 'Group Columns', onClick: () => groupSel(ctx, 'col') }] }), L('ugr', 'Ungroup', 'layers', null, { menu: [{ label: 'Ungroup Rows', onClick: () => ungroupSel(ctx, 'row') }, { label: 'Ungroup Columns', onClick: () => ungroupSel(ctx, 'col') }] }), L('sub', 'Subtotal', 'sum', () => subtotal(ctx)), B('hdt', 'Hide Detail', 'eye', () => detail(ctx, true, false), { label: 'Hide Detail' }), B('sdt', 'Show Detail', 'eye', () => detail(ctx, false, false), { label: 'Show Detail' }), B('hda', 'Collapse All', null, () => detail(ctx, true, true), { glyph: '1', label: 'Level 1' }), B('sda', 'Expand All', null, () => detail(ctx, false, true), { glyph: '2', label: 'Level 2' })] }
    ]);
    const ft = tab('formulas'); if (ft) {
      const calc = ft.groups.find(g => g.label === 'Calculation'); if (calc) { const i = calc.items.findIndex(x => x.id === 'calc'); calc.items[i] = L('calc', 'Calculate Now', 'play', () => { wb()._force = true; wb().dirty(); wb().ensureCalc(); wb()._force = false; ctx.g.refresh(); OA.toast('Recalculated'); }); calc.items.push(L('copt', 'Calculation Options', 'settings', null, { menu: [{ label: 'Automatic', onClick: () => { ctx.calcManual = false; wb().dirty(); ctx.g.refresh(); OA.toast('Calculation: Automatic'); } }, { label: 'Manual (F9 / Calculate Now)', onClick: () => { ctx.calcManual = true; OA.toast('Calculation: Manual'); } }] }), L('csh', 'Calculate Sheet', 'play', () => { wb()._force = true; wb().dirty(); wb().ensureCalc(); wb()._force = false; ctx.g.refresh(); })); }
      ft.groups.push({ label: 'Formula Auditing', items: [L('tpr', 'Trace Precedents', 'chevronLeft', () => trace(ctx, 'pre')), L('tde', 'Trace Dependents', 'chevronRight', () => trace(ctx, 'dep')), L('rar', 'Remove Arrows', 'eraser', () => { ctx.audit = { arrows: [], circles: [] }; ctx.g.overlayAudit(); }), L('evf', 'Evaluate Formula', 'play', () => evaluateFormula(ctx)), L('erc', 'Error Checking', 'warning', () => errorCheck(ctx)), L('wwn', 'Watch Window', 'eye', () => watchWindow(ctx))] });
      const dn = ft.groups.find(g => g.label === 'Defined Names'); if (dn) dn.items.push(L('cfs', 'Create from Selection', 'list', () => createNames(ctx)), L('ufm', 'Use in Formula', 'edit', async () => { const ns = Object.values(ctx.wb.names).map(n => n.name); if (!ns.length) return OA.toast('Koi name defined nahi'); const v = await OA.form('Paste Name', [{ id: 'n', label: 'Name', type: 'select', options: ns, value: ns[0] }]); if (!v) return; const p = ctx.g.active(); if (ctx.g.editing) ctx.g.ed.value += v.n; else ctx.g.startEdit(p.r, p.c, '=' + v.n, 'edit'); }));
    }
    add('review', [{ label: 'Protect', items: [{ type: 'check', id: 'pwb', label: 'Protect Workbook', value: false, onChange: v => { ctx.calcManual = ctx.calcManual; ctx.wbProtect = v; OA.toast(v ? 'Workbook structure protected' : 'Workbook structure unprotected'); } }, L('aer', 'Allow Edit Ranges', 'check', async () => { const q = ctx.g.range(); const v = await OA.form('Allow Edit Ranges', [{ id: 'on', label: 'Selected range ko editable rakhein (unlocked) jab sheet protected ho', type: 'checkbox', value: true }]); if (!v) return; ctx.wb.tx(() => ctx.wb.applyStyle(ctx.sh(), q, { locked: v.on ? false : null })); ctx.after(); })] },
      { label: 'Proofing', items: [L('acc', 'Check Accessibility', 'check2', () => accessibility(ctx)), L('wst', 'Workbook Statistics', 'info', () => workbookStats(ctx))] }]);
    const hm = tab('home'); if (hm) { const ed = hm.groups.find(g => g.label === 'Editing'); if (ed) ed.items.push(L('pss', 'Paste Special', 'paste', () => pasteSpecial(ctx)), L('gts', 'Go To Special', 'search', () => goToSpecial(ctx)), L('ffh', 'Flash Fill', 'play', () => flashFill(ctx))); }
    document.addEventListener('keydown', e => { if (e.ctrlKey && e.key.toLowerCase() === 'e' && document.querySelector('.app-excel') && !e.shiftKey && document.body.contains(H.root)) { if (H.root.offsetParent) { e.preventDefault(); flashFill(ctx); } } if (e.key === 'F9' && H.root.offsetParent) { e.preventDefault(); ctx.wb._force = true; ctx.wb.dirty(); ctx.wb.ensureCalc(); ctx.wb._force = false; ctx.g.refresh(); } });
  });
})();
