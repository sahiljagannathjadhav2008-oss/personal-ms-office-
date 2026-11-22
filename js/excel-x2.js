/* Excel extras 2: Tables (structured refs, Table Design, Slicer), PivotTable/PivotChart, What-If (Goal Seek, Scenarios, Data Table, Forecast, Solver), Analysis ToolPak — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, XL = window.XL, el = OA.el, key = XL.key;
  XL.extras = XL.extras || []; XL.gridHooks = XL.gridHooks || [];
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const num = v => typeof v === 'number' && isFinite(v);

  /* ---------- helpers ---------- */
  const val = (ctx, sh, r, c) => ctx.wb.value(sh.name, r, c);
  function readTable(ctx, sh, q) { ctx.wb.ensureCalc(); const rows = []; for (let r = q.r1; r <= q.r2; r++) { const row = []; for (let c = q.c1; c <= q.c2; c++) { const v = val(ctx, sh, r, c); row.push(v === undefined || v === null ? '' : v); } rows.push(row); } return rows; }
  function outSheet(ctx, name, rows, styleFn) {
    const wb = ctx.wb; let s; wb.tx(() => { s = wb.addSheet(wb.uniqueName(name), wb.sheets.length); wb.loadRows(s, rows.map(r => r.map(v => v === null || v === undefined ? '' : (typeof v === 'number' ? String(v) : String(v))))); styleFn && styleFn(s); });
    ctx.activate(s.name); return s;
  }
  const hdrStyle = { b: true, color: '#ffffff', bg: '#217346' };
  const agg = {
    Sum: a => a.reduce((x, y) => x + y, 0), Count: a => a.length, Average: a => a.reduce((x, y) => x + y, 0) / (a.length || 1), Max: a => Math.max(...a), Min: a => Math.min(...a),
    Product: a => a.reduce((x, y) => x * y, 1), 'Std Dev': a => { const m = agg.Average(a); return Math.sqrt(a.reduce((s, v) => s + (v - m) ** 2, 0) / Math.max(1, a.length - 1)); }
  };
  const parseSel = (ctx, txt) => { const m = /^(?:'?([^'!]+)'?!)?(.+)$/.exec(String(txt).trim()); if (!m) return null; const g = XL.parseRange(m[2]); if (!g) return null; return { sh: m[1] ? ctx.wb.sheet(m[1]) : ctx.sh(), g }; };

  /* ---------- structured references (Table1[Col], [@Col]) ---------- */
  const tablesOf = wb => { const out = []; wb.sheets.forEach(s => (s.tables || []).forEach(t => out.push({ sh: s, t }))); return out; };
  function rewriteStructured(text, sh, r, wb) {
    if (typeof text !== 'string' || text[0] !== '=' || text.indexOf('[') < 0) return text;
    const tabs = tablesOf(wb); if (!tabs.length) return text;
    return text.replace(/(?:([A-Za-z_][\w.]*))?\[(?:\[?#(Headers|Data|All|Totals)\]?,?)?\[?@?([^\[\]#]*)\]?\]?/g, (m, tn, sect, col) => {
      let ent;
      if (tn) ent = tabs.find(x => x.t.name.toLowerCase() === tn.toLowerCase());
      else ent = tabs.find(x => x.sh === sh && r >= x.t.r1 && r <= x.t.r2);
      if (!ent) return m; const t = ent.t; const hdrRow = t.header === false ? t.r1 - 1 : t.r1; const d1 = t.header === false ? t.r1 : t.r1 + 1, d2 = t.total ? t.r2 - 1 : t.r2;
      const colIx = name => { for (let c = t.c1; c <= t.c2; c++) if (String(wb.value(ent.sh.name, t.r1, c)).toLowerCase() === name.toLowerCase()) return c; return -1; };
      const pre = ent.sh === sh ? '' : "'" + ent.sh.name + "'!"; const A = XL.addr;
      const atMarker = m.indexOf('@') >= 0;
      if (atMarker) { const c = colIx(col.trim()); if (c < 0) return m; return pre + A(r, c); }
      if (col && col.trim()) { const c = colIx(col.trim()); if (c < 0) return m; let a = d1, b = d2; if (sect === 'Headers') a = b = t.r1; else if (sect === 'All') { a = t.r1; b = t.r2; } else if (sect === 'Totals') a = b = t.r2; return pre + A(a, c) + ':' + A(b, c); }
      let a = d1, b = d2; if (sect === 'Headers') a = b = t.r1; else if (sect === 'All') { a = t.r1; b = t.r2; } else if (sect === 'Totals') a = b = t.r2; return pre + A(a, t.c1) + ':' + A(b, t.c2);
    });
  }
  const origPut = XL.Workbook.prototype.putInput;
  XL.Workbook.prototype.putInput = function (sh, r, c, text, opts) { try { text = rewriteStructured(text, sh, r, this); } catch (e) { } return origPut.call(this, sh, r, c, text, opts); };

  const TSTYLES = { Green: ['#217346', '#e2efda'], Blue: ['#4472c4', '#d9e1f2'], Orange: ['#ed7d31', '#fce4d6'], Gray: ['#595959', '#ededed'], Purple: ['#7030a0', '#e4d6f2'] };
  function tableAt(ctx) { const s = ctx.sh(), p = ctx.g.active(); return (s.tables || []).find(t => p.r >= t.r1 && p.r <= t.r2 && p.c >= t.c1 && p.c <= t.c2); }
  function paintTable(ctx, s, t) {
    const wb = ctx.wb, [hc, band] = TSTYLES[t.style || 'Blue'] || TSTYLES.Blue; const w = { r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 };
    wb.applyStyle(s, w, { bg: null, b: false, color: null });
    if (t.banded !== false) for (let r = t.r1 + (t.header === false ? 0 : 1); r <= t.r2; r++) if ((r - t.r1) % 2 === (t.header === false ? 1 : 0)) wb.applyStyle(s, { r1: r, c1: t.c1, r2: r, c2: t.c2 }, { bg: band });
    if (t.header !== false) wb.applyStyle(s, { r1: t.r1, c1: t.c1, r2: t.r1, c2: t.c2 }, { b: true, color: '#ffffff', bg: hc });
    if (t.total) wb.applyStyle(s, { r1: t.r2, c1: t.c1, r2: t.r2, c2: t.c2 }, { b: true, bg: band });
  }
  function makeTable(ctx, style) {
    const s = ctx.sh(), wb = ctx.wb; let q = ctx.g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, s, q.r1, q.c1);
    const names = (s.tables || []).map(t => t.name); let n = 1; while (tablesOf(wb).some(x => x.t.name === 'Table' + n)) n++;
    const t = { name: 'Table' + n, r1: q.r1, c1: q.c1, r2: q.r2, c2: q.c2, style: style || 'Blue', header: true, banded: true, total: false };
    wb.tx(() => { wb.setProp(s, 'tables', (s.tables || []).concat([t])); paintTable(ctx, s, t); wb.setFilter(s, { r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 }); }); ctx.after(); return t;
  }
  function tableUpd(ctx, patch) {
    const t = tableAt(ctx); if (!t) return OA.toast('Table ke andar cell select kijiye'); const s = ctx.sh(), wb = ctx.wb;
    wb.tx(() => { const nt = Object.assign({}, t, typeof patch === 'function' ? patch(t) : patch);
      if (patch.total === true || (typeof patch === 'function' && nt.total && !t.total)) { nt.r2 = t.r2 + 1; for (let c = t.c1; c <= t.c2; c++) wb.putInput(s, nt.r2, c, c === t.c1 ? 'Total' : '=SUBTOTAL(109,' + XL.addr(t.r1 + 1, c) + ':' + XL.addr(t.r2, c) + ')'); }
      if (nt.total === false && t.total) { for (let c = t.c1; c <= t.c2; c++) wb.putInput(s, t.r2, c, ''); nt.r2 = t.r2 - 1; }
      wb.setProp(s, 'tables', s.tables.map(x => x === t ? nt : x)); paintTable(ctx, s, nt); });
    ctx.after();
  }
  async function slicerDlg(ctx) {
    const t = tableAt(ctx) || (ctx.sh().filter ? Object.assign({ name: 'Range' }, ctx.sh().filter) : null); if (!t) return OA.alert('Pehle Table banaiye (Insert → Table) ya filter lagaiye.');
    const s = ctx.sh(), wb = ctx.wb; wb.ensureCalc(); const heads = []; for (let c = t.c1; c <= t.c2; c++) heads.push([c, String(val(ctx, s, t.r1, c))]);
    const v = await OA.form('Insert Slicer', [{ id: 'c', label: 'Field', type: 'select', options: heads.map(h => [h[0], h[1]]), value: heads[0][0] }]); if (!v) return; const col = +v.c;
    const items = new Map(); for (let r = t.r1 + 1; r <= t.r2; r++) { const x = String(val(ctx, s, r, col)); items.set(x, (items.get(x) || 0) + 1); }
    const sel = new Set(items.keys()); const box = el('div', { style: 'display:flex;flex-wrap:wrap;gap:6px;max-height:260px;overflow:auto' });
    const apply = () => { const hid = new Set(s.hiddenR); for (let r = t.r1 + 1; r <= t.r2; r++) { if (sel.has(String(val(ctx, s, r, col)))) hid.delete(r); else hid.add(r); } wb.tx(() => wb.setProp(s, 'hiddenR', hid)); ctx.after(); };
    const draw = () => { box.innerHTML = ''; items.forEach((n, k) => box.appendChild(el('button', { type: 'button', class: 'btn' + (sel.has(k) ? ' primary' : ''), text: (k || '(blank)') + ' (' + n + ')', onclick: () => { sel.has(k) ? sel.delete(k) : sel.add(k); draw(); apply(); } }))); };
    draw(); OA.dialog({ title: 'Slicer — ' + heads.find(h => h[0] === col)[1], content: el('div', {}, [el('div', { class: 'form-intro', text: 'Click karke values filter karein; sab select = clear filter.' }), box, el('div', { style: 'margin-top:8px' }, [el('button', { class: 'btn', type: 'button', text: 'Clear Filter', onclick: () => { items.forEach((n, k) => sel.add(k)); draw(); apply(); } })])]), width: 460, buttons: [{ label: 'Close', primary: true }] });
  }

  /* ---------- PivotTable ---------- */
  function buildPivot(ctx, def) {
    const src = parseSel(ctx, def.src); if (!src) throw new Error('Source range galat hai'); const rows = readTable(ctx, src.sh, src.g); const head = rows[0].map(String), data = rows.slice(1);
    const ix = n => head.indexOf(n); const ri = ix(def.row), ci = def.col ? ix(def.col) : -1, vi = ix(def.val), fi = def.filt ? ix(def.filt) : -1; if (ri < 0 || vi < 0) throw new Error('Field names match nahi hue');
    const keep = data.filter(r => fi < 0 || !def.filtVal || String(r[fi]) === String(def.filtVal));
    const rk = [...new Set(keep.map(r => String(r[ri])))].sort((a, b) => isNaN(a) || isNaN(b) ? a.localeCompare(b) : a - b), ck = ci >= 0 ? [...new Set(keep.map(r => String(r[ci])))].sort((a, b) => isNaN(a) || isNaN(b) ? a.localeCompare(b) : a - b) : [null];
    const f = agg[def.agg] || agg.Sum; const cell = (rf, cf) => { const a = keep.filter(r => (rf === null || String(r[ri]) === rf) && (cf === null || String(r[ci]) === cf)).map(r => r[vi]).filter(num); if (def.agg === 'Count') return keep.filter(r => (rf === null || String(r[ri]) === rf) && (cf === null || String(r[ci]) === cf)).length; return a.length ? +f(a).toFixed(6) : ''; };
    const lab = def.agg + ' of ' + def.val; const out = []; const top = [lab].concat(ci >= 0 ? ck.map(String) : []); if (def.grand !== false && ci >= 0) top.push('Grand Total'); if (ci < 0) top.push('');
    if (ci >= 0) out.push([def.row + ' \\ ' + def.col].concat(ck, def.grand !== false ? ['Grand Total'] : [])); else out.push([def.row, lab]);
    rk.forEach(rv => { if (ci >= 0) out.push([rv].concat(ck.map(cv => cell(rv, cv)), def.grand !== false ? [cell(rv, null)] : [])); else out.push([rv, cell(rv, null)]); });
    if (def.grand !== false) { if (ci >= 0) out.push(['Grand Total'].concat(ck.map(cv => cell(null, cv)), [cell(null, null)])); else out.push(['Grand Total', cell(null, null)]); }
    return out;
  }
  async function pivotDlg(ctx) {
    const s = ctx.sh(); let q = ctx.g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(ctx.wb, s, q.r1, q.c1); const heads = readTable(ctx, s, { r1: q.r1, c1: q.c1, r2: q.r1, c2: q.c2 })[0].map(String); const opt = (blank) => (blank ? [['', '(none)']] : []).concat(heads.map(h => [h, h]));
    const v = await OA.form('Create PivotTable', [{ id: 'src', label: 'Source range', value: s.name + '!' + XL.rangeLabel(q) }, { id: 'row', label: 'Rows field', type: 'select', options: opt(), value: heads[0] }, { id: 'col', label: 'Columns field', type: 'select', options: opt(true), value: '' }, { id: 'val', label: 'Values field', type: 'select', options: opt(), value: heads[heads.length - 1] }, { id: 'agg', label: 'Summarize by', type: 'select', options: Object.keys(agg), value: 'Sum' }, { id: 'filt', label: 'Filter field (optional)', type: 'select', options: opt(true), value: '' }, { id: 'filtVal', label: 'Filter value', value: '' }, { id: 'grand', label: 'Grand totals', type: 'checkbox', value: true }, { id: 'chart', label: 'Also create PivotChart', type: 'checkbox', value: false }], { width: 420 }); if (!v) return;
    try { const out = buildPivot(ctx, v); const ns = outSheet(ctx, 'PivotTable', out, sh => { ctx.wb.applyStyle(sh, { r1: 0, c1: 0, r2: 0, c2: out[0].length - 1 }, hdrStyle); ctx.wb.applyStyle(sh, { r1: out.length - 1, c1: 0, r2: out.length - 1, c2: out[0].length - 1 }, { b: true, bg: '#e2efda' }); ctx.wb.setProp(sh.colW, 0, 150); }); ctx.wb.tx(() => ctx.wb.setProp(ns, 'pivot', Object.assign({}, v))); if (v.chart) addPivotChart(ctx, ns, out); }
    catch (e) { OA.alert(esc(e.message)); }
  }
  function addPivotChart(ctx, ns, out) { const ci = out[0].length - ((out[0].length > 2) ? 1 : 0); const last = (ctx.g.sh().pivot && ctx.g.sh().pivot.grand !== false) ? out.length - 1 : out.length; ctx.g.addChart({ id: XL.uid(), type: 'col', title: 'PivotChart', ref: ns.name + '!A1:' + XL.addr(last - 1, ci - 1), rowsAsSeries: false, x: 20, y: (out.length + 3) * 20, w: 480, h: 280, legend: true }); }
  function refreshPivots(ctx) { let n = 0; ctx.wb.sheets.forEach(s => { if (!s.pivot) return; try { const out = buildPivot(ctx, s.pivot); ctx.wb.tx(() => { for (let r = 0; r < out.length + 3; r++) for (let c = 0; c < out[0].length + 2; c++) ctx.wb.putInput(s, r, c, r < out.length && c < out[0].length ? String(out[r][c] == null ? '' : out[r][c]) : ''); }); n++; } catch (e) { } }); ctx.after(); OA.toast(n ? n + ' PivotTable refresh hui' : 'Koi PivotTable nahi mili'); }

  /* ---------- What-If ---------- */
  async function goalSeek(ctx) {
    const a = ctx.g.active(); const v = await OA.form('Goal Seek', [{ id: 'set', label: 'Set cell (formula)', value: XL.addr(a.r, a.c) }, { id: 'to', label: 'To value', type: 'number', value: 0 }, { id: 'by', label: 'By changing cell', value: '' }]); if (!v) return;
    const S = parseSel(ctx, v.set), B = parseSel(ctx, v.by); if (!S || !B) return OA.alert('Cell reference galat hai.'); const s = S.sh, b = B.sh, wb = ctx.wb; const orig = wb.value(b.name, B.g.r1, B.g.c1);
    const f = x => { wb.setInput(b, B.g.r1, B.g.c1, String(x)); wb.ensureCalc(); const y = wb.value(s.name, S.g.r1, S.g.c1); return (typeof y === 'number' ? y : NaN) - v.to; };
    let x0 = typeof orig === 'number' ? orig : 1, x1 = x0 + 1, f0 = f(x0), f1 = f(x1), ok = false; wb.hist.undo.length = Math.max(0, wb.hist.undo.length - 2);
    for (let i = 0; i < 100; i++) { if (Math.abs(f1) < 1e-9) { ok = true; break; } const d = f1 - f0; if (!d || !isFinite(d)) { x1 += 1; f1 = f(x1); continue; } const x2 = x1 - f1 * (x1 - x0) / d; x0 = x1; f0 = f1; x1 = x2; f1 = f(x1); }
    ok = ok || Math.abs(f1) < 1e-6; wb.hist.undo.length = Math.max(0, wb.hist.undo.length - 100);
    if (ok) { wb.setInput(b, B.g.r1, B.g.c1, String(+x1.toPrecision(12))); ctx.after(); OA.alert('Goal Seeking found a solution.<br>Cell ' + v.by + ' = <b>' + +x1.toPrecision(8) + '</b>'); }
    else { wb.setInput(b, B.g.r1, B.g.c1, typeof orig === 'number' ? String(orig) : ''); ctx.after(); OA.alert('Goal Seek solution nahi dhoondh paya.'); }
  }
  async function scenarioMgr(ctx) {
    const wb = ctx.wb, s = ctx.sh(); s.scenarios = s.scenarios || []; const list = s.scenarios; const q = ctx.g.range();
    const v = await OA.form('Scenario Manager', [{ id: 'act', label: 'Action', type: 'select', options: [['add', 'Add scenario from changing cells (selection)'], ['show', 'Show scenario'], ['sum', 'Summary report'], ['del', 'Delete scenario']], value: 'add' }, { id: 'name', label: 'Scenario name', value: list.length ? list[0].name : '' }, { id: 'res', label: 'Result cells for summary (e.g. D10)', value: '' }], { intro: list.length ? 'Saved: <b>' + list.map(x => esc(x.name)).join(', ') + '</b>' : 'Koi scenario saved nahi.', width: 420 }); if (!v) return;
    const cellsOf = g => { const o = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) o.push([r, c]); return o; };
    if (v.act === 'add') { const vals = cellsOf(q).map(([r, c]) => { const cl = s.cells.get(key(r, c)); return [r, c, cl ? (cl.f != null ? '=' + cl.f : String(cl.v == null ? '' : cl.v)) : '']; }); s.scenarios = list.filter(x => x.name !== v.name).concat([{ name: v.name, vals }]); OA.toast('Scenario saved: ' + v.name); }
    else if (v.act === 'show') { const sc = list.find(x => x.name === v.name); if (!sc) return OA.alert('Scenario nahi mila'); wb.tx(() => sc.vals.forEach(([r, c, t]) => wb.putInput(s, r, c, t))); ctx.after(); }
    else if (v.act === 'del') { s.scenarios = list.filter(x => x.name !== v.name); }
    else { const R = parseSel(ctx, v.res); if (!R || !list.length) return OA.alert('Result cell aur scenarios chahiye'); const cur = list[0].vals.map(([r, c]) => [r, c]); const snap = cur.map(([r, c]) => { const cl = s.cells.get(key(r, c)); return [r, c, cl ? (cl.f != null ? '=' + cl.f : String(cl.v == null ? '' : cl.v)) : '']; }); const rows = [['Scenario Summary'].concat(list.map(x => x.name))];
      cur.forEach(([r, c]) => rows.push(['Changing: ' + XL.addr(r, c)].concat(list.map(x => x.vals.find(y => y[0] === r && y[1] === c)[2])))); const resRow = ['Result: ' + v.res];
      list.forEach(x => { wb.tx(() => x.vals.forEach(([r, c, t]) => wb.putInput(s, r, c, t))); wb.ensureCalc(); resRow.push(String(wb.value(s.name, R.g.r1, R.g.c1))); }); wb.tx(() => snap.forEach(([r, c, t]) => wb.putInput(s, r, c, t))); rows.push(resRow); outSheet(ctx, 'Scenario Summary', rows, sh => { wb.applyStyle(sh, { r1: 0, c1: 0, r2: 0, c2: list.length }, hdrStyle); wb.setProp(sh.colW, 0, 180); }); }
  }
  async function dataTable(ctx) {
    const q = ctx.g.range(), s = ctx.sh(), wb = ctx.wb; const v = await OA.form('Data Table', [{ id: 'row', label: 'Row input cell (optional)', value: '' }, { id: 'col', label: 'Column input cell (optional)', value: '' }], { intro: 'Pehle select karein: top-left me formula, upar row values, left me column values. (Single variable: sirf ek input cell dein.)', width: 420 }); if (!v) return;
    const RI = v.row ? parseSel(ctx, v.row) : null, CI = v.col ? parseSel(ctx, v.col) : null; if (!RI && !CI) return OA.alert('Kam se kam ek input cell dein.');
    const save = (c) => { const x = c.sh.cells.get(key(c.g.r1, c.g.c1)); return x ? (x.f != null ? '=' + x.f : String(x.v == null ? '' : x.v)) : ''; }; const sr = RI && save(RI), sc = CI && save(CI);
    wb.tx(() => {
      if (RI && CI) { for (let r = q.r1 + 1; r <= q.r2; r++) for (let c = q.c1 + 1; c <= q.c2; c++) { wb.putInput(RI.sh, RI.g.r1, RI.g.c1, String(val(ctx, s, q.r1, c))); wb.putInput(CI.sh, CI.g.r1, CI.g.c1, String(val(ctx, s, r, q.c1))); wb.ensureCalc(); const y = wb.value(s.name, q.r1, q.c1); wb.putInput(s, r, c, String(y)); } }
      else if (RI) { for (let c = q.c1 + 1; c <= q.c2; c++) { wb.putInput(RI.sh, RI.g.r1, RI.g.c1, String(val(ctx, s, q.r1, c))); for (let r = q.r1 + 1; r <= q.r2; r++) { wb.ensureCalc(); wb.putInput(s, r, c, String(wb.value(s.name, r, q.c1))); } } }
      else { for (let r = q.r1 + 1; r <= q.r2; r++) { wb.putInput(CI.sh, CI.g.r1, CI.g.c1, String(val(ctx, s, r, q.c1))); wb.ensureCalc(); for (let c = q.c1 + 1; c <= q.c2; c++) wb.putInput(s, r, c, String(wb.value(s.name, q.r1, c))); } }
      RI && wb.putInput(RI.sh, RI.g.r1, RI.g.c1, sr); CI && wb.putInput(CI.sh, CI.g.r1, CI.g.c1, sc);
    }); ctx.after(); OA.toast('Data Table values fill ho gayi');
  }
  async function forecastSheet(ctx) {
    const s = ctx.sh(); let q = ctx.g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(ctx.wb, s, q.r1, q.c1); const rows = readTable(ctx, s, q);
    const v = await OA.form('Create Forecast Worksheet', [{ id: 'n', label: 'Forecast periods', type: 'number', value: 6, min: 1, max: 200 }], { intro: 'Selection: 2 columns — timeline/period (ya numbers) aur values (header row ke saath).' }); if (!v) return;
    const body = rows.slice(1).filter(r => num(r[1])); if (body.length < 3) return OA.alert('Kam se kam 3 numeric values chahiye.'); const isDate = false; const ys = body.map(r => r[1]), xs = body.map((r, i) => num(r[0]) ? r[0] : i + 1);
    const n = ys.length, mx = xs.reduce((a, b) => a + b, 0) / n, my = ys.reduce((a, b) => a + b, 0) / n; let sxy = 0, sxx = 0; xs.forEach((x, i) => { sxy += (x - mx) * (ys[i] - my); sxx += (x - mx) ** 2; }); const b = sxy / sxx, a = my - b * mx; const se = Math.sqrt(ys.reduce((t, y, i) => t + (y - (a + b * xs[i])) ** 2, 0) / Math.max(1, n - 2)); const step = (xs[n - 1] - xs[0]) / (n - 1) || 1;
    const out = [[String(rows[0][0]), String(rows[0][1]), 'Forecast', 'Lower Bound', 'Upper Bound']]; body.forEach((r, i) => out.push([xs[i], ys[i], '', '', ''])); out[out.length - 1][2] = ys[n - 1]; out[out.length - 1][3] = ys[n - 1]; out[out.length - 1][4] = ys[n - 1];
    for (let i = 1; i <= v.n; i++) { const x = xs[n - 1] + step * i, f = a + b * x, ci = 1.96 * se * Math.sqrt(1 + 1 / n + (x - mx) ** 2 / sxx); out.push([+x.toFixed(6), '', +f.toFixed(4), +(f - ci).toFixed(4), +(f + ci).toFixed(4)]); }
    const ns = outSheet(ctx, 'Forecast', out, sh => ctx.wb.applyStyle(sh, { r1: 0, c1: 0, r2: 0, c2: 4 }, hdrStyle)); ctx.g.addChart({ id: XL.uid(), type: 'line', title: 'Forecast', ref: ns.name + '!A1:E' + out.length, rowsAsSeries: false, x: 460, y: 20, w: 520, h: 300, legend: true });
  }
  async function solver(ctx) {
    const wb = ctx.wb; const a = ctx.g.active(); const v = await OA.form('Solver', [{ id: 'obj', label: 'Set objective cell', value: XL.addr(a.r, a.c) }, { id: 'goal', label: 'To', type: 'select', options: [['max', 'Max'], ['min', 'Min'], ['val', 'Value of']], value: 'max' }, { id: 'tv', label: 'Value (if "Value of")', type: 'number', value: 0 }, { id: 'vars', label: 'By changing cells (e.g. B1:B3)', value: '' }, { id: 'lo', label: 'Lower bound for variables', type: 'number', value: 0 }, { id: 'hi', label: 'Upper bound for variables', type: 'number', value: 1000 }, { id: 'cons', label: 'Constraints (e.g. D5<=100; D6>=10)', value: '' }], { width: 440, intro: 'Simple Solver: coordinate-search + hill-climb (non-linear OK, no GRG/Simplex).' }); if (!v) return;
    const O = parseSel(ctx, v.obj), V = parseSel(ctx, v.vars); if (!O || !V) return OA.alert('Cells sahi likhiye.'); const cells = []; for (let r = V.g.r1; r <= V.g.r2; r++) for (let c = V.g.c1; c <= V.g.c2; c++) cells.push([r, c]);
    const cons = v.cons.split(/[;\n]/).map(t => /^\s*([A-Za-z]+\d+)\s*(<=|>=|=)\s*(-?[\d.]+|[A-Za-z]+\d+)\s*$/.exec(t)).filter(Boolean).map(m => ({ ref: parseSel(ctx, m[1]), op: m[2], rhs: m[3] }));
    const set = xs => { cells.forEach(([r, c], i) => wb.putInput(V.sh, r, c, String(xs[i]))); wb.ensureCalc(); };
    const getv = (ref) => { const y = wb.value(ref.sh.name, ref.g.r1, ref.g.c1); return num(y) ? y : NaN; };
    const pen = () => { let p = 0; cons.forEach(k => { const lhs = getv(k.ref), rhs = /^-?[\d.]+$/.test(k.rhs) ? +k.rhs : getv(parseSel(ctx, k.rhs)); if (k.op === '<=' && lhs > rhs) p += lhs - rhs; else if (k.op === '>=' && lhs < rhs) p += rhs - lhs; else if (k.op === '=' && Math.abs(lhs - rhs) > 1e-6) p += Math.abs(lhs - rhs); }); return p; };
    const score = xs => { set(xs); const y = getv(O), p = pen(); if (isNaN(y)) return -1e18; const base = v.goal === 'max' ? y : v.goal === 'min' ? -y : -Math.abs(y - v.tv); return base - 1e6 * p; };
    const before = cells.map(([r, c]) => { const x = val(ctx, V.sh, r, c); return num(x) ? x : 0; }); let x = before.map(b => Math.min(v.hi, Math.max(v.lo, b))), best = score(x), step = (v.hi - v.lo) / 4 || 1;
    for (let it = 0; it < 400 && step > 1e-7; it++) { let improved = false; for (let i = 0; i < x.length; i++) for (const d of [step, -step]) { const y = x.slice(); y[i] = Math.min(v.hi, Math.max(v.lo, y[i] + d)); const sc = score(y); if (sc > best + 1e-12) { best = sc; x = y; improved = true; } } if (!improved) step /= 2; }
    wb.tx(() => set(x.map(n => +n.toPrecision(10)))); ctx.after(); OA.alert('Solver finished.<br>' + cells.map((c, i) => XL.addr(c[0], c[1]) + ' = <b>' + +x[i].toPrecision(6) + '</b>').join('<br>') + '<br>Objective = <b>' + getv(O) + '</b>' + (pen() > 1e-6 ? '<br><i>Constraints poori tarah satisfy nahi hui.</i>' : ''));
  }

  /* ---------- Analysis ToolPak ---------- */
  const mean = a => a.reduce((x, y) => x + y, 0) / a.length, vari = a => { const m = mean(a); return a.reduce((s, v) => s + (v - m) ** 2, 0) / Math.max(1, a.length - 1); };
  const colsOf = (ctx, q) => { const rows = readTable(ctx, ctx.sh(), q); const out = []; for (let c = 0; c <= q.c2 - q.c1; c++) out.push({ name: String(rows[0][c]) || 'Col' + (c + 1), v: rows.slice(1).map(r => r[c]).filter(num) }); if (!out.some(o => o.v.length)) { const o2 = []; for (let c = 0; c <= q.c2 - q.c1; c++) o2.push({ name: 'Col' + (c + 1), v: rows.map(r => r[c]).filter(num) }); return o2; } return out; };
  async function toolpak(ctx) {
    const q0 = ctx.g.range(); const v = await OA.form('Data Analysis', [{ id: 't', label: 'Tool', type: 'select', options: ['Descriptive Statistics', 'Correlation', 'Covariance', 'Histogram', 'Regression', 'Moving Average', 'Exponential Smoothing', 'Rank and Percentile', 'Random Number Generation', 't-Test: Two-Sample (equal variance)', 'z-Test / One-sample mean', 'ANOVA: Single Factor'], value: 'Descriptive Statistics' }, { id: 'rng', label: 'Input range (with header)', value: XL.rangeLabel(q0) }, { id: 'p', label: 'Parameter (bins count / period / alpha / count)', type: 'number', value: 5 }], { width: 440 }); if (!v) return;
    const R = parseSel(ctx, v.rng); if (!R) return OA.alert('Range galat'); const q = R.g, cols = colsOf(ctx, q); let out = [], title = v.t;
    if (!cols.length || !cols[0].v.length) return OA.alert('Numeric data nahi mila.');
    const stat = a => { const s = a.slice().sort((x, y) => x - y), n = s.length, m = mean(a), sd = Math.sqrt(vari(a)); const med = n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; const cnt = {}; a.forEach(x => cnt[x] = (cnt[x] || 0) + 1); const mode = Object.entries(cnt).sort((x, y) => y[1] - x[1])[0]; const k = n > 3 ? n * (n + 1) / ((n - 1) * (n - 2) * (n - 3)) * a.reduce((t, x) => t + ((x - m) / sd) ** 4, 0) - 3 * (n - 1) ** 2 / ((n - 2) * (n - 3)) : ''; const sk = n > 2 ? n / ((n - 1) * (n - 2)) * a.reduce((t, x) => t + ((x - m) / sd) ** 3, 0) : ''; return [['Mean', m], ['Standard Error', sd / Math.sqrt(n)], ['Median', med], ['Mode', +mode[0]], ['Standard Deviation', sd], ['Sample Variance', vari(a)], ['Kurtosis', k], ['Skewness', sk], ['Range', s[n - 1] - s[0]], ['Minimum', s[0]], ['Maximum', s[n - 1]], ['Sum', a.reduce((x, y) => x + y, 0)], ['Count', n]]; };
    const corr = (x, y, cov) => { const n = Math.min(x.length, y.length), mx = mean(x.slice(0, n)), my = mean(y.slice(0, n)); let sxy = 0, sxx = 0, syy = 0; for (let i = 0; i < n; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; } return cov ? sxy / n : sxy / Math.sqrt(sxx * syy); };
    switch (v.t) {
      case 'Descriptive Statistics': out = [['Statistic'].concat(cols.map(c => c.name))]; stat(cols[0].v).forEach((r, i) => out.push([r[0]].concat(cols.map(c => { const x = stat(c.v)[i][1]; return typeof x === 'number' ? +x.toFixed(6) : x; })))); break;
      case 'Correlation': case 'Covariance': out = [[v.t].concat(cols.map(c => c.name))]; cols.forEach(a => out.push([a.name].concat(cols.map(b => +corr(a.v, b.v, v.t === 'Covariance').toFixed(6))))); break;
      case 'Histogram': { const a = cols[0].v, mn = Math.min(...a), mx = Math.max(...a), k = Math.max(1, Math.round(v.p)), w = (mx - mn) / k || 1; out = [['Bin (upper)', 'Frequency']]; const f = new Array(k).fill(0); a.forEach(x => { f[Math.min(k - 1, Math.floor((x - mn) / w))]++; }); f.forEach((n, i) => out.push([+(mn + w * (i + 1)).toFixed(4), n])); break; }
      case 'Regression': { if (cols.length < 2) return OA.alert('Regression ke liye 2 columns (Y pehle, X baad me) chahiye.'); const y = cols[0].v, x = cols[1].v, n = Math.min(x.length, y.length), mx = mean(x), my = mean(y); let sxy = 0, sxx = 0, syy = 0; for (let i = 0; i < n; i++) { sxy += (x[i] - mx) * (y[i] - my); sxx += (x[i] - mx) ** 2; syy += (y[i] - my) ** 2; } const b = sxy / sxx, a = my - b * mx, r2 = sxy * sxy / (sxx * syy), sse = syy * (1 - r2), se = Math.sqrt(sse / Math.max(1, n - 2)); out = [['SUMMARY OUTPUT', ''], ['Multiple R', +Math.sqrt(r2).toFixed(6)], ['R Square', +r2.toFixed(6)], ['Adjusted R Square', +(1 - (1 - r2) * (n - 1) / Math.max(1, n - 2)).toFixed(6)], ['Standard Error', +se.toFixed(6)], ['Observations', n], ['', ''], ['Intercept', +a.toFixed(6)], ['Slope (' + cols[1].name + ')', +b.toFixed(6)], ['Slope Std Error', +(se / Math.sqrt(sxx)).toFixed(6)], ['t Stat (slope)', +(b / (se / Math.sqrt(sxx))).toFixed(4)]]; break; }
      case 'Moving Average': { const a = cols[0].v, k = Math.max(2, Math.round(v.p)); out = [['Period', cols[0].name, 'Moving Avg(' + k + ')']]; a.forEach((x, i) => out.push([i + 1, x, i < k - 1 ? '#N/A' : +mean(a.slice(i - k + 1, i + 1)).toFixed(6)])); break; }
      case 'Exponential Smoothing': { const a = cols[0].v, al = v.p > 0 && v.p < 1 ? v.p : 0.3; out = [['Period', cols[0].name, 'Smoothed (α=' + al + ')']]; let s = a[0]; a.forEach((x, i) => { if (i) s = al * x + (1 - al) * s; out.push([i + 1, x, +s.toFixed(6)]); }); break; }
      case 'Rank and Percentile': { const a = cols[0].v, s = a.slice().sort((x, y) => y - x); out = [['Point', cols[0].name, 'Rank', 'Percent']]; a.forEach((x, i) => out.push([i + 1, x, s.indexOf(x) + 1, +(((a.length - (s.indexOf(x) + 1)) / Math.max(1, a.length - 1)) * 100).toFixed(2) + '%'])); break; }
      case 'Random Number Generation': { out = [['Random']]; for (let i = 0; i < Math.max(1, Math.round(v.p)); i++) out.push([+Math.random().toFixed(6)]); break; }
      case 't-Test: Two-Sample (equal variance)': { if (cols.length < 2) return OA.alert('2 columns chahiye'); const a = cols[0].v, b = cols[1].v, n1 = a.length, n2 = b.length, sp = ((n1 - 1) * vari(a) + (n2 - 1) * vari(b)) / (n1 + n2 - 2), t = (mean(a) - mean(b)) / Math.sqrt(sp * (1 / n1 + 1 / n2)); out = [['t-Test', cols[0].name, cols[1].name], ['Mean', +mean(a).toFixed(6), +mean(b).toFixed(6)], ['Variance', +vari(a).toFixed(6), +vari(b).toFixed(6)], ['Observations', n1, n2], ['Pooled Variance', +sp.toFixed(6), ''], ['df', n1 + n2 - 2, ''], ['t Stat', +t.toFixed(6), '']]; break; }
      case 'z-Test / One-sample mean': { const a = cols[0].v, n = a.length, z = (mean(a) - v.p) / (Math.sqrt(vari(a)) / Math.sqrt(n)); out = [['One-sample test', 'Value'], ['Hypothesised mean', v.p], ['Mean', +mean(a).toFixed(6)], ['Std dev', +Math.sqrt(vari(a)).toFixed(6)], ['n', n], ['z / t statistic', +z.toFixed(6)]]; break; }
      case 'ANOVA: Single Factor': { const all = [].concat(...cols.map(c => c.v)), gm = mean(all), k = cols.length, N = all.length; const ssb = cols.reduce((t, c) => t + c.v.length * (mean(c.v) - gm) ** 2, 0), ssw = cols.reduce((t, c) => t + c.v.reduce((s, x) => s + (x - mean(c.v)) ** 2, 0), 0); const msb = ssb / (k - 1), msw = ssw / (N - k); out = [['Source', 'SS', 'df', 'MS', 'F'], ['Between Groups', +ssb.toFixed(6), k - 1, +msb.toFixed(6), +(msb / msw).toFixed(6)], ['Within Groups', +ssw.toFixed(6), N - k, +msw.toFixed(6), ''], ['Total', +(ssb + ssw).toFixed(6), N - 1, '', '']]; break; }
    }
    outSheet(ctx, title.split(':')[0].replace(/[^\w ]/g, ''), out, sh => { ctx.wb.applyStyle(sh, { r1: 0, c1: 0, r2: 0, c2: out[0].length - 1 }, hdrStyle); ctx.wb.setProp(sh.colW, 0, 170); });
  }

  XL.x2 = { buildPivot, rewriteStructured, agg };
  XL.gridHooks.push((g, ctx) => { XL.curCtx = ctx; });

  /* ---------- ribbon ---------- */
  XL.extras.push(function (ctx, spec, H) {
    const { B, L } = H; const tab = id => spec.tabs.find(t => t.id === id); const add = (id, groups) => { const t = tab(id); if (t) t.groups.push(...groups); };
    const insTab = tab('insert'); if (insTab) { const g0 = insTab.groups.find(g => g.label === 'Tables'); if (g0) { const i = g0.items.findIndex(x => x.id === 'tbl'); g0.items[i] = L('tbl', 'Table', 'table', () => makeTable(ctx)); g0.items.splice(1, 0, L('pvt2', 'PivotTable', 'sort', () => pivotDlg(ctx)), L('pch', 'PivotChart', 'chart', async () => { const s = ctx.sh(); if (!s.pivot) return OA.alert('PivotTable sheet par jaakar use kijiye, ya pehle "PivotTable" with PivotChart option banayein.'); const out = buildPivot(ctx, s.pivot); addPivotChart(ctx, s, out); }), L('slc', 'Slicer', 'filter', () => slicerDlg(ctx))); } }
    add('data', [
      { label: 'Forecast', items: [L('wif', 'What-If Analysis', 'refresh', null, { menu: [{ label: 'Scenario Manager…', onClick: () => scenarioMgr(ctx) }, { label: 'Goal Seek…', onClick: () => goalSeek(ctx) }, { label: 'Data Table…', onClick: () => dataTable(ctx) }] }), L('fcs', 'Forecast Sheet', 'chart', () => forecastSheet(ctx))] },
      { label: 'Analyze', items: [L('slv', 'Solver', 'play', () => solver(ctx)), L('dan', 'Data Analysis', 'chart', () => toolpak(ctx)), L('rfa', 'Refresh All', 'refresh', () => refreshPivots(ctx))] }
    ]);
    spec.tabs.push({ id: 'tdesign', label: 'Table Design', contextual: true, hidden: true, groups: [
      { label: 'Properties', items: [L('tnm', 'Table Name', 'edit', async () => { const t = tableAt(ctx); if (!t) return; const n = await OA.prompt('Table name:', t.name, { title: 'Table Name' }); if (n && /^[A-Za-z_][\w.]*$/.test(n)) tableUpd(ctx, { name: n }); }), L('trs', 'Resize Table', 'columns', async () => { const t = tableAt(ctx); if (!t) return; const n = await OA.prompt('New range:', XL.rangeLabel(t)); const g = n && XL.parseRange(n); if (g) tableUpd(ctx, { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }); })] },
      { label: 'Style Options', items: [L('thr', 'Header Row', 'check', () => tableUpd(ctx, t => ({ header: t.header === false }))), L('ttr', 'Total Row', 'sum', () => tableUpd(ctx, t => ({ total: !t.total }))), L('tbr', 'Banded Rows', 'layers', () => tableUpd(ctx, t => ({ banded: t.banded === false })))] },
      { label: 'Table Styles', items: [L('tst', 'Style', 'palette', null, { menu: Object.keys(TSTYLES).map(k => ({ label: k, onClick: () => tableUpd(ctx, { style: k }) })) })] },
      { label: 'Tools', items: [L('tpv', 'Summarize with PivotTable', 'sort', () => pivotDlg(ctx)), L('trd', 'Remove Duplicates', 'eraser', () => XL.dlg.removeDuplicates(ctx)), L('tsl', 'Insert Slicer', 'filter', () => slicerDlg(ctx)), L('tcr', 'Convert to Range', 'undo', () => { const t = tableAt(ctx); if (!t) return; const s = ctx.sh(); ctx.wb.tx(() => ctx.wb.setProp(s, 'tables', s.tables.filter(x => x !== t))); ctx.after(); OA.toast('Table ko range me convert kiya'); })] }
    ] });
    ctx._tableSync = () => { if (!H.rb) return; H.rb.setTabVisible('tdesign', !!tableAt(ctx)); };
  });
  XL.gridHooks.push((g, ctx) => { const od = g.draw; g.draw = function () { od(); ctx._tableSync && ctx._tableSync(); }; });
})();
