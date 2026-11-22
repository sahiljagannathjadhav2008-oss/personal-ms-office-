/* Excel app UI — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, XL = window.XL, el = OA.el, key = XL.key;
  const MAXR = XL.MAXR, MAXC = XL.MAXC;
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

  /* ---------- templates (built into a fresh workbook) */
  const hdr = { b: true, color: '#ffffff', bg: '#217346' };
  function fill(wb, sh, rows, r0, c0) { wb.loadRows(sh, rows, r0 || 0, c0 || 0); }
  function st(wb, sh, r1, c1, r2, c2, patch) { wb.applyStyle(sh, { r1, c1, r2, c2 }, patch); }
  const TEMPLATES = {
    budget: { name: 'Monthly Budget', build(wb) {
      const sh = wb.sheets[0]; sh.name = 'Budget';
      fill(wb, sh, [['Monthly Budget', '', ''], ['', '', ''], ['Income', 'Planned', 'Actual'], ['Salary', '50000', '50000'], ['Freelance', '10000', '8500'], ['Other', '2000', '2500'], ['Total Income', '=SUM(B4:B6)', '=SUM(C4:C6)'], ['', '', ''], ['Expenses', 'Planned', 'Actual'], ['Rent', '15000', '15000'], ['Groceries', '8000', '8900'], ['Transport', '3000', '2700'], ['Utilities', '2500', '2650'], ['Entertainment', '2000', '3200'], ['Savings', '10000', '10000'], ['Total Expenses', '=SUM(B10:B15)', '=SUM(C10:C15)'], ['', '', ''], ['Balance', '=B7-B16', '=C7-C16']]);
      st(wb, sh, 0, 0, 0, 0, { b: true, sz: 18, color: '#217346' }); [[2, 2], [8, 8]].forEach(([r]) => st(wb, sh, r, 0, r, 2, hdr)); [6, 15, 17].forEach(r => st(wb, sh, r, 0, r, 2, { b: true, bg: '#e2efda' })); st(wb, sh, 3, 1, 17, 2, { nf: '#,##0' });
      wb.setProp(sh.colW, 0, 160); wb.setProp(sh.colW, 1, 100); wb.setProp(sh.colW, 2, 100);
      wb.addCF(sh, { type: 'cell', op: '<', v1: '0', range: { r1: 17, c1: 1, r2: 17, c2: 2 }, style: { color: '#c00000', b: true } });
      sh.charts.push({ id: XL.uid(), type: 'col', title: 'Expenses: Planned vs Actual', ref: 'Budget!A9:C15', rowsAsSeries: false, x: 380, y: 40, w: 460, h: 280, legend: true });
    } },
    invoice: { name: 'Invoice', build(wb) {
      const sh = wb.sheets[0]; sh.name = 'Invoice'; sh.showGrid = false;
      fill(wb, sh, [['INVOICE', '', '', '', ''], ['Your Company Name', '', '', 'Invoice #', 'INV-001'], ['Address line, City', '', '', 'Date', '=TODAY()'], ['', '', '', '', ''], ['Bill To:', '', '', '', ''], ['Customer Name', '', '', '', ''], ['', '', '', '', ''], ['Description', 'Qty', 'Rate', 'GST %', 'Amount'], ['Product / Service 1', '2', '1500', '18', '=B9*C9*(1+D9/100)'], ['Product / Service 2', '1', '3200', '18', '=B10*C10*(1+D10/100)'], ['Product / Service 3', '5', '400', '12', '=B11*C11*(1+D11/100)'], ['', '', '', '', ''], ['', '', '', 'Total', '=SUM(E9:E11)']]);
      st(wb, sh, 0, 0, 0, 0, { b: true, sz: 24, color: '#217346' }); st(wb, sh, 7, 0, 7, 4, hdr); st(wb, sh, 4, 0, 4, 0, { b: true }); st(wb, sh, 12, 3, 12, 4, { b: true, sz: 13, bg: '#e2efda' }); st(wb, sh, 8, 2, 12, 4, { nf: '#,##0.00' }); st(wb, sh, 2, 4, 2, 4, { nf: 'dd-mmm-yyyy', ha: 'right' });
      wb.setProp(sh.colW, 0, 220); [1, 2, 3].forEach(c => wb.setProp(sh.colW, c, 80)); wb.setProp(sh.colW, 4, 110);
    } },
    todo: { name: 'To-do Tracker', build(wb) {
      const sh = wb.sheets[0]; sh.name = 'Tasks';
      fill(wb, sh, [['Task', 'Owner', 'Due Date', 'Priority', 'Status'], ['Plan project scope', 'Sahil', '12/1/2026', 'High', 'In Progress'], ['Design screens', 'Team', '12/5/2026', 'Medium', 'Not Started'], ['Write documentation', 'Sahil', '12/10/2026', 'Low', 'Not Started'], ['Release v1', 'Team', '12/20/2026', 'High', 'Done']]);
      st(wb, sh, 0, 0, 0, 4, hdr); st(wb, sh, 1, 2, 30, 2, { nf: 'dd-mmm-yyyy' }); wb.setProp(sh.colW, 0, 200); [1, 2, 3, 4].forEach(c => wb.setProp(sh.colW, c, 110));
      wb.addDV(sh, { range: { r1: 1, c1: 4, r2: 30, c2: 4 }, type: 'list', list: 'Not Started,In Progress,Done', allowBlank: true, alert: 'stop' }); wb.addDV(sh, { range: { r1: 1, c1: 3, r2: 30, c2: 3 }, type: 'list', list: 'High,Medium,Low', allowBlank: true, alert: 'stop' });
      wb.addCF(sh, { type: 'text', op: 'contains', v1: 'Done', range: { r1: 1, c1: 4, r2: 30, c2: 4 }, style: { bg: '#c6efce', color: '#006100' } }); wb.addCF(sh, { type: 'text', op: 'contains', v1: 'High', range: { r1: 1, c1: 3, r2: 30, c2: 3 }, style: { bg: '#ffc7ce', color: '#9c0006' } });
      wb.setFilter(sh, { r1: 0, c1: 0, r2: 4, c2: 4 });
    } }
  };
  function newBook(tpl) { const wb = new XL.Workbook(); wb.activeIndex = 0; if (tpl && TEMPLATES[tpl]) { wb.tx(() => TEMPLATES[tpl].build(wb)); wb.hist.undo.length = 0; wb.hist.redo.length = 0; } return wb; }

  const CELLSTYLES = [['Normal', {}], ['Good', { bg: '#c6efce', color: '#006100' }], ['Bad', { bg: '#ffc7ce', color: '#9c0006' }], ['Neutral', { bg: '#ffeb9c', color: '#9c5700' }], ['Input', { bg: '#ffcc99', color: '#3f3f76' }], ['Calculation', { bg: '#f2f2f2', color: '#fa7d00', b: true }], ['Note', { bg: '#ffffcc' }], ['Title', { b: true, sz: 18, color: '#44546a' }], ['Heading 1', { b: true, sz: 15, color: '#44546a' }], ['Heading 2', { b: true, sz: 13, color: '#44546a' }], ['Total', { b: true, bg: '#e2efda' }], ['Accent (Green)', { bg: '#217346', color: '#ffffff', b: true }]];
  const TABLESTYLES = [['Green', '#217346', '#e2efda'], ['Blue', '#4472c4', '#d9e1f2'], ['Orange', '#ed7d31', '#fce4d6'], ['Gray', '#595959', '#ededed'], ['Purple', '#7030a0', '#e4d6f2']];

  OA.registerApp({
    id: 'excel', name: 'Excel', letter: 'X', color: '#217346', accept: ['xlsx', 'csv', 'tsv', 'txt'], defaultName: 'Book1',
    desc: 'Spreadsheets with 400+ formulas, charts, formatting, filters, sorting, validation, conditional formatting, and .xlsx / .csv export.',
    templates: Object.keys(TEMPLATES).map(k => ({ id: k, label: TEMPLATES[k].name })),
    create(root) {
      const a = OA.apps.excel; const sh = OA.appShell(root, a); root.classList.add('app-excel');
      let wb, g, rb, bs, dirty = false, editing = false; const page = { orient: 'portrait', grid: false, title: '' };
      const ctx = { get wb() { return wb; }, get g() { return g; }, sh: () => g.sh(), page, after() { dirty = true; g.refresh(); updateBar(); syncRibbon(); }, activate: n => activate(wb.sheets.findIndex(s => s.name === n)) };
      const baseName = () => (sh.name.value.trim() || 'Book1').replace(/\.(xlsx|csv|tsv|txt)$/i, '');
      const toast = OA.toast;

      /* ---------- formula bar + grid host + tabs */
      const nameBox = el('input', { class: 'x-name', type: 'text', spellcheck: 'false', 'aria-label': 'Name Box' });
      const fx = el('button', { class: 'x-fx', type: 'button', title: 'Insert Function', text: 'fx' });
      const fin = el('input', { class: 'x-fin', type: 'text', spellcheck: 'false', autocomplete: 'off', 'aria-label': 'Formula Bar' });
      const fbar = el('div', { class: 'x-fbar' }, [nameBox, fx, fin]);
      const host = el('div', { class: 'x-grid' }); const tabs = el('div', { class: 'x-tabs' }); sh.content.append(fbar, host, tabs);
      const sbMode = el('span', { text: 'Ready' }), sbStats = el('span', { class: 'sb-stats' }), zoomLbl = el('button', { type: 'button', text: '100%', onclick: () => g.setZoom(1) }), zoomRange = el('input', { type: 'range', min: 40, max: 300, value: 100, 'aria-label': 'Zoom' });
      const selBtn = el('button', { type: 'button', title: 'Touch: drag se cells select karne ka mode', text: '☝ Select', onclick: () => { const on = !g.selMode; g.setSelMode(on); selBtn.classList.toggle('on', on); toast(on ? 'Select mode ON: ab drag karke cells select hongi' : 'Scroll mode'); } });
      zoomRange.addEventListener('input', () => g.setZoom(zoomRange.value / 100));
      sh.status.append(sbMode, sbStats, el('span', { class: 'sb-fill' }), el('span', { class: 'sb-dev sb-hide-s', text: 'Developer: ' + OA.DEV }), selBtn, el('button', { type: 'button', html: OA.icon('zoomOut', 15), title: 'Zoom out', onclick: () => g.setZoom(g.zoom - 0.1) }), zoomRange, el('button', { type: 'button', html: OA.icon('zoomIn', 15), title: 'Zoom in', onclick: () => g.setZoom(g.zoom + 0.1) }), zoomLbl);

      const rangeName = q => { if (q.r1 === 0 && q.r2 >= MAXR - 1) return XL.colName(q.c1) + ':' + XL.colName(q.c2); if (q.c1 === 0 && q.c2 >= MAXC - 1) return (q.r1 + 1) + ':' + (q.r2 + 1); return XL.rangeLabel(q); };
      function updateBar() { if (!g) return; const q = g.range(); const m = wb.mergeAt(g.sh(), q.r1, q.c1); if (document.activeElement !== nameBox) nameBox.value = (q.r1 === q.r2 && q.c1 === q.c2) || m ? XL.addr(g.active().r, g.active().c) : rangeName(q); if (!editing && document.activeElement !== fin) fin.value = g.activeText(); stats(); }
      function stats() {
        const s = g.sh(), ranges = g.ranges(); let n = 0, cnt = 0, sum = 0, mn = Infinity, mx = -Infinity; wb.ensureCalc();
        const q = ranges[ranges.length - 1]; const area = (Math.abs(q.r2 - q.r1) + 1) * (Math.abs(q.c2 - q.c1) + 1);
        if (area > 1) { const iter = (r, c) => { const v = wb.value(s.name, r, c); if (v === null || v === '' || v === undefined) return; cnt++; if (typeof v === 'number') { n++; sum += v; if (v < mn) mn = v; if (v > mx) mx = v; } };
          ranges.forEach(rg => { const r1 = Math.min(rg.r1, rg.r2), r2 = Math.max(rg.r1, rg.r2), c1 = Math.min(rg.c1, rg.c2), c2 = Math.max(rg.c1, rg.c2); if ((r2 - r1 + 1) * (c2 - c1 + 1) > 60000) { s.cells.forEach((cl, k) => { const u = XL.unkey(k); if (u.r >= r1 && u.r <= r2 && u.c >= c1 && u.c <= c2) iter(u.r, u.c); }); } else for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) iter(r, c); }); }
        const f = v => XL.formatValue(v, 'General').text; sbStats.innerHTML = ''; if (cnt) { sbStats.append(el('span', { text: 'Count: ' + cnt })); if (n) sbStats.append(el('span', { text: 'Average: ' + f(sum / n) }), el('span', { text: 'Sum: ' + f(sum) }), el('span', { class: 'sb-hide-s', text: 'Min: ' + f(mn) }), el('span', { class: 'sb-hide-s', text: 'Max: ' + f(mx) })); }
      }
      nameBox.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); const t = nameBox.value; if (!g.gotoRef(t)) { const q = g.range(); if (/^[A-Za-z_][\w.]*$/.test(t) && !/^[A-Za-z]{1,3}\d+$/.test(t)) { wb.tx(() => wb.defineName(t, XL.sheetRef(g.sh(), q), g.sh().name)); toast('Name "' + t + '" define ho gaya'); } else OA.alert('Reference valid nahi hai.'); } g.focus(); } else if (e.key === 'Escape') { updateBar(); g.focus(); } });
      nameBox.addEventListener('focus', () => nameBox.select());
      fx.addEventListener('click', () => XL.dlg.insertFunction(ctx, n => { const p = g.active(); if (!g.editing) g.startEdit(p.r, p.c, '=' + n + '(', 'edit'); else { g.ed.value += n + '('; } }));

      /* ---------- hooks */
      const hooks = {
        onSel() { updateBar(); syncRibbon(); }, onEditState(on) { editing = on; sbMode.textContent = on ? (g.editing && g.editing.mode === 'edit' ? 'Edit' : 'Enter') : 'Ready'; if (!on) { fin.value = g.activeText(); } },
        onZoom(z) { zoomLbl.textContent = Math.round(z * 100) + '%'; zoomRange.value = Math.round(z * 100); }, onChange() { dirty = true; updateBar(); },
        activateSheet: n => activate(wb.sheets.findIndex(s => s.name === n)), cmd: c => command(c), filterMenu: (c, anchor) => XL.dlg.filterMenu(ctx, c, anchor),
        contextMenu: (e, h) => ctxMenu(e, h), chartDialog: ch => ch.obj ? XL.objDialog(ctx, ch) : XL.dlg.chart(ctx, null, ch), chartMenu: (e, ch) => OA.menu({ left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY }, [{ label: 'Change Chart…', onClick: () => XL.dlg.chart(ctx, null, ch) }, { label: 'Delete Chart', icon: 'trash', onClick: () => g.deleteChart(ch.id) }]),
        onChartSel: ch => { if (rb) { rb.setTabVisible('chart', !!ch && !ch.obj); rb.setTabVisible('shape', !!ch && !!ch.obj); if (ch) rb.showTab(ch.obj ? 'shape' : 'chart'); } }, onTip: t => { sbMode.textContent = t || (editing ? 'Edit' : 'Ready'); }
      };
      function setBook(nb) {
        if (g) { g.destroy(); host.innerHTML = ''; }
        wb = nb; wb.activeIndex = Math.min(wb.activeIndex || 0, wb.sheets.length - 1); wb.onChange = () => { dirty = true; g.refresh(); updateBar(); syncRibbon(); };
        g = XL.createGrid(host, wb, hooks); (XL.gridHooks || []).forEach(f => f(g, ctx)); g.bindBar(fin); g.sheetChanged(); renderTabs(); updateBar(); syncRibbon(); if (rb) rb.showTab('home'); hooks.onChartSel(null); g.focus(); dirty = false;
      }
      function activate(i) { if (i < 0 || i >= wb.sheets.length) return; if (g.editing) g.commit(0, 0, true); g.saveScroll(); wb.activeIndex = i; g.sheetChanged(); renderTabs(); updateBar(); syncRibbon(); }
      function renderTabs() {
        tabs.innerHTML = ''; const list = el('div', { class: 'tl' });
        wb.sheets.forEach((s, i) => { const t = el('button', { class: 'x-tab' + (i === wb.activeIndex ? ' on' : ''), type: 'button', text: s.name }); if (s.tabColor) t.appendChild(el('span', { class: 'tc', style: { background: s.tabColor } }));
          t.onclick = () => activate(i); t.ondblclick = () => renameSheet(i); t.oncontextmenu = e => { e.preventDefault(); activate(i); OA.menu({ left: e.clientX, top: e.clientY - 150, right: e.clientX, bottom: e.clientY }, [{ label: 'Insert Sheet', onClick: addSheet }, { label: 'Delete', onClick: () => delSheet(i) }, { label: 'Rename', onClick: () => renameSheet(i) }, { label: 'Move Left', disabled: i === 0, onClick: () => moveSheet(i, i - 1) }, { label: 'Move Right', disabled: i === wb.sheets.length - 1, onClick: () => moveSheet(i, i + 1) }, { label: 'Duplicate', onClick: () => { wb.tx(() => wb.duplicateSheet(s)); activate(wb.sheets.length - 1); } }, { sep: true }, { label: 'Tab Color…', onClick: () => OA.colorPicker(t, { onPick: c => { wb.setProp(s, 'tabColor', c ? '#' + c : null); renderTabs(); } }) }]); }; list.appendChild(t); });
        tabs.append(el('button', { class: 'tbtn', type: 'button', text: '‹', title: 'Scroll tabs left', onclick: () => list.scrollBy(-120, 0) }), el('button', { class: 'tbtn', type: 'button', text: '›', title: 'Scroll tabs right', onclick: () => list.scrollBy(120, 0) }), list, el('button', { class: 'tbtn', type: 'button', text: '+', title: 'New sheet (Shift+F11)', onclick: addSheet }));
        const on = list.querySelector('.on'); on && on.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
      function addSheet() { if (ctx.wbProtect) return toast('Workbook structure protected hai'); wb.tx(() => wb.addSheet(wb.uniqueName('Sheet'), wb.activeIndex + 1)); activate(wb.activeIndex + 1); dirty = true; }
      async function renameSheet(i) { if (ctx.wbProtect) return toast('Workbook structure protected hai'); const s = wb.sheets[i]; const n = await OA.prompt('Sheet name:', s.name, { title: 'Rename Sheet' }); if (n && n !== s.name) { if (/[\\/?*\[\]:]/.test(n) || n.length > 31) return OA.alert('Sheet name mein \\ / ? * [ ] : nahi ho sakte (max 31 chars).'); if (wb.hasSheet(n)) return OA.alert('Is naam ka sheet pehle se hai.'); wb.tx(() => wb.renameSheet(s, n)); renderTabs(); g.refresh(); } }
      async function delSheet(i) { if (ctx.wbProtect) return toast('Workbook structure protected hai'); if (wb.sheets.length < 2) return OA.alert('Workbook mein kam se kam ek sheet rehni chahiye.'); if (!await OA.confirm('Sheet "' + wb.sheets[i].name + '" delete karein? Ise Undo se wapas la sakte hain.', { ok: 'Delete' })) return; wb.tx(() => wb.deleteSheet(wb.sheets[i])); wb.activeIndex = Math.min(wb.activeIndex, wb.sheets.length - 1); g.sheetChanged(); renderTabs(); updateBar(); }
      function moveSheet(i, to) { if (ctx.wbProtect) return toast('Workbook structure protected hai'); const act = wb.sheets[wb.activeIndex]; wb.tx(() => wb.moveSheet(wb.sheets[i], to)); wb.activeIndex = wb.sheets.indexOf(act); renderTabs(); g.refresh(); }

      /* ---------- formatting helpers */
      const eachRange = fn => g.ranges().forEach(q => { const r = { r1: Math.min(q.r1, q.r2), c1: Math.min(q.c1, q.c2), r2: Math.max(q.r1, q.r2), c2: Math.max(q.c1, q.c2) }; if (r.r1 === 0 && r.r2 >= MAXR - 1) r.fullCol = true; else if (r.c1 === 0 && r.c2 >= MAXC - 1) r.fullRow = true; fn(r); });
      const fmt = patch => { wb.tx(() => eachRange(r => wb.applyStyle(g.sh(), r, patch))); g.focus(); };
      const curSt = () => { const p = g.active(); return wb.styleOf(g.sh(), p.r, p.c); };
      const toggle = k => fmt({ [k]: !curSt()[k] });
      const decimals = d => { const s = curSt(); let nf = s.nf || 'General'; if (nf === 'General') nf = d > 0 ? '0.0' : '0'; else if (!/0/.test(nf)) return; const m = /0\.(0+)/.exec(nf); const cur = m ? m[1].length : 0; const nd = Math.max(0, cur + d); nf = m ? nf.replace(/0\.0+/, nd ? '0.' + '0'.repeat(nd) : '0') : (nd ? nf.replace(/0(?!.*0)/, '0.' + '0'.repeat(nd)) : nf); fmt({ nf }); };
      const hideSet = (kind, on) => { const s = g.sh(); const q = g.range(); const prop = kind === 'row' ? 'hiddenR' : 'hiddenC'; const ns = new Set(s[prop]); const a = kind === 'row' ? q.r1 : q.c1, b = Math.min(kind === 'row' ? q.r2 : q.c2, a + 3000); for (let i = a; i <= b; i++) on ? ns.add(i) : ns.delete(i); wb.tx(() => wb.setProp(s, prop, ns)); };
      const unhideAll = kind => { const s = g.sh(); const prop = kind === 'row' ? 'hiddenR' : 'hiddenC'; wb.tx(() => wb.setProp(s, prop, new Set())); };
      async function sizeDlg(kind) { const s = g.sh(), q = g.range(); const cur = kind === 'row' ? wb.rowHeight(s, q.r1) : wb.colWidth(s, q.c1); const v = await OA.prompt((kind === 'row' ? 'Row height' : 'Column width') + ' (pixels):', String(cur), { title: kind === 'row' ? 'Row Height' : 'Column Width' }); const n = parseFloat(v); if (!(n >= 0)) return; wb.tx(() => { if (kind === 'row') for (let r = q.r1; r <= Math.min(q.r2, q.r1 + 5000); r++) wb.setProp(s.rowH, r, n); else for (let c = q.c1; c <= Math.min(q.c2, q.c1 + 500); c++) wb.setProp(s.colW, c, n); }); }
      function insertDel(kind, del) {
        const s = g.sh(), q = g.range(); const rowsSel = q.c1 === 0 && q.c2 >= MAXC - 1, colsSel = q.r1 === 0 && q.r2 >= MAXR - 1;
        const doIt = k => { const at = k === 'row' ? q.r1 : q.c1, n = k === 'row' ? Math.min(q.r2 - q.r1 + 1, 1000) : Math.min(q.c2 - q.c1 + 1, 200); wb.insertDelete(s, k, at, del ? -n : n); ctx.after(); };
        if (kind) return doIt(kind); if (rowsSel) return doIt('row'); if (colsSel) return doIt('col');
        const sel = el('select', { class: 'inp' }, [el('option', { value: 'row', text: 'Entire row' }), el('option', { value: 'col', text: 'Entire column' })]);
        OA.dialog({ title: del ? 'Delete' : 'Insert', content: el('div', { class: 'form-grid' }, [el('label', { class: 'form-row' }, [el('span', { text: del ? 'Delete' : 'Insert' }), sel])]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => doIt(sel.value) }] });
      }
      function autosum(fn) {
        fn = fn || 'SUM'; const s = g.sh(), q = g.range(); wb.ensureCalc(); const num = (r, c) => typeof wb.value(s.name, r, c) === 'number';
        if (q.r1 !== q.r2 || q.c1 !== q.c2) { // fill totals below (columns) or right if one row
          if (q.r2 > q.r1) { wb.tx(() => { for (let c = q.c1; c <= q.c2; c++) wb.putInput(s, q.r2 + 1, c, '=' + fn + '(' + XL.addr(q.r1, c) + ':' + XL.addr(q.r2, c) + ')'); }); g.select({ r1: q.r2 + 1, c1: q.c1, r2: q.r2 + 1, c2: q.c2 }); }
          else { wb.setInput(s, q.r1, q.c2 + 1, '=' + fn + '(' + XL.addr(q.r1, q.c1) + ':' + XL.addr(q.r1, q.c2) + ')'); g.setActive(q.r1, q.c2 + 1); } return; }
        const p = g.active(); let r = p.r - 1; while (r >= 0 && num(r, p.c)) r--; r++;
        if (r < p.r) { g.startEdit(p.r, p.c, '=' + fn + '(' + XL.addr(r, p.c) + ':' + XL.addr(p.r - 1, p.c) + ')', 'enter'); return; }
        let c = p.c - 1; while (c >= 0 && num(p.r, c)) c--; c++; if (c < p.c) { g.startEdit(p.r, p.c, '=' + fn + '(' + XL.addr(p.r, c) + ':' + XL.addr(p.r, p.c - 1) + ')', 'enter'); return; }
        g.startEdit(p.r, p.c, '=' + fn + '(', 'edit');
      }
      function formatAsTable(ts) {
        const s = g.sh(); let q = g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, s, q.r1, q.c1); const [, hc, band] = ts || TABLESTYLES[0];
        wb.tx(() => { wb.applyStyle(s, q, { bg: null }); for (let r = q.r1 + 1; r <= q.r2; r++) if ((r - q.r1) % 2 === 0) wb.applyStyle(s, { r1: r, c1: q.c1, r2: r, c2: q.c2 }, { bg: band }); wb.applyStyle(s, { r1: q.r1, c1: q.c1, r2: q.r1, c2: q.c2 }, { b: true, color: '#ffffff', bg: hc }); wb.applyBorders(s, q, 'outer', 'thin', hc); wb.setFilter(s, q); }); g.select(q);
      }
      async function importCSV() { const f = await OA.pickFile('.csv,.tsv,.txt'); if (!f) return; const t = await f.text(); const d = /\t/.test(t.split('\n')[0]) ? '\t' : (t.split('\n')[0].split(';').length > t.split('\n')[0].split(',').length ? ';' : ','); const rows = XL.parseCSV(t, d); const p = g.active(); wb.loadRows(g.sh(), rows, p.r, p.c); ctx.after(); toast('Data import ho gaya (' + rows.length + ' rows)'); }
      const funcMenu = cat => (XL.FDOC[cat] || []).map(n => ({ label: n, hint: '', onClick: () => { const p = g.active(); g.startEdit(p.r, p.c, '=' + n + '(', 'edit'); } }));
      function freeze(mode) { const s = g.sh(), p = g.active(); const fz = mode === 'row' ? { r: 1, c: 0 } : mode === 'col' ? { r: 0, c: 1 } : mode === 'pane' ? { r: p.r, c: p.c } : { r: 0, c: 0 }; wb.tx(() => wb.setProp(s, 'freeze', fz)); g.refresh(); }
      function sortFilterToggle() { const s = g.sh(); if (s.filter) wb.setFilter(s, null); else { let q = g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, s, q.r1, q.c1); wb.setFilter(s, q); } wb.applyFilter(s); g.refresh(); }
      function protect() { const s = g.sh(); const on = !(s.protect && s.protect.on); s.protect = { on }; toast(on ? 'Sheet protected: locked cells edit nahi honge' : 'Sheet unprotected'); }
      function command(c) {
        switch (c) {
          case 'bold': toggle('b'); break; case 'italic': toggle('i'); break; case 'underline': toggle('u'); break; case 'strike': toggle('st'); break;
          case 'format': XL.dlg.formatCells(ctx); break; case 'find': XL.dlg.findReplace(ctx, false); break; case 'replace': XL.dlg.findReplace(ctx, true); break; case 'goto': XL.dlg.goTo(ctx); break;
          case 'link': XL.dlg.hyperlink(ctx); break; case 'filter': sortFilterToggle(); break; case 'table': formatAsTable(); break; case 'print': doPrint(); break; case 'save': doExport('xlsx'); break;
          case 'open': openPicker(); break; case 'new': bs.open('new'); break; case 'hidecol': hideSet('col', true); break; case 'hiderow': hideSet('row', true); break;
          case 'insert': insertDel(null, false); break; case 'delete': insertDel(null, true); break; case 'autosum': autosum(); break; case 'newsheet': addSheet(); break;
        }
      }
      function ctxMenu(e, h) {
        const at = { left: e.clientX, top: e.clientY, right: e.clientX, bottom: e.clientY }; let items;
        if (h.t === 'colh') items = [{ label: 'Cut', onClick: () => g.copyButton(true) }, { label: 'Copy', onClick: () => g.copyButton(false) }, { label: 'Paste', onClick: () => g.pasteButton() }, { sep: true }, { label: 'Insert Column', onClick: () => insertDel('col') }, { label: 'Delete Column', onClick: () => insertDel('col', true) }, { label: 'Clear Contents', onClick: () => g.clearContents() }, { sep: true }, { label: 'Column Width…', onClick: () => sizeDlg('col') }, { label: 'AutoFit Column Width', onClick: () => g.autofitCols() }, { label: 'Hide', onClick: () => hideSet('col', true) }, { label: 'Unhide All', onClick: () => unhideAll('col') }];
        else if (h.t === 'rowh') items = [{ label: 'Cut', onClick: () => g.copyButton(true) }, { label: 'Copy', onClick: () => g.copyButton(false) }, { label: 'Paste', onClick: () => g.pasteButton() }, { sep: true }, { label: 'Insert Row', onClick: () => insertDel('row') }, { label: 'Delete Row', onClick: () => insertDel('row', true) }, { label: 'Clear Contents', onClick: () => g.clearContents() }, { sep: true }, { label: 'Row Height…', onClick: () => sizeDlg('row') }, { label: 'AutoFit Row Height', onClick: () => g.autofitRows() }, { label: 'Hide', onClick: () => hideSet('row', true) }, { label: 'Unhide All', onClick: () => unhideAll('row') }];
        else items = [{ label: 'Cut', hint: 'Ctrl+X', onClick: () => g.copyButton(true) }, { label: 'Copy', hint: 'Ctrl+C', onClick: () => g.copyButton(false) }, { label: 'Paste', hint: 'Ctrl+V', onClick: () => g.pasteButton() }, { label: 'Paste Special…', onClick: () => XL.dlg.pasteSpecial(ctx) }, { sep: true }, { label: 'Insert…', onClick: () => insertDel(null, false) }, { label: 'Delete…', onClick: () => insertDel(null, true) }, { label: 'Clear Contents', onClick: () => g.clearContents() }, { sep: true }, { label: 'Sort A to Z', onClick: () => XL.dlg.sort(ctx, 'asc') }, { label: 'Sort Z to A', onClick: () => XL.dlg.sort(ctx, 'desc') }, { label: 'Filter', onClick: sortFilterToggle }, { sep: true }, { label: 'Insert Note…', onClick: () => XL.dlg.note(ctx) }, { label: 'Link…', onClick: () => XL.dlg.hyperlink(ctx) }, { label: 'Format Cells…', hint: 'Ctrl+1', onClick: () => XL.dlg.formatCells(ctx) }];
        OA.menu(at, items);
      }

      /* ---------- file I/O */
      const used = s => { const u = wb.usedRange(s); return u && u.r2 >= 0 ? u : null; };
      function printBox() {
        wb.ensureCalc(); const box = el('div', { class: 'x-print' + (page.grid ? ' grid' : '') + (page.headings ? ' heads' : '') }); if (page.title) box.appendChild(el('div', { style: 'text-align:center;font-weight:bold;margin-bottom:6px', text: page.title }));
        (page.activeOnly ? [g.sh()] : wb.sheets).forEach(s => {
          const u = s.printArea || used(s); if (!u) return; box.appendChild(el('h3', { text: s.name })); const d = el('div'); d.innerHTML = wb.rangeHTML(s, u);
          const tbl = d.querySelector('table'); const t = s.titleRows;
          if (tbl && t && t.r1 >= 0) { const thead = document.createElement('thead'); if (t.r1 >= u.r1 && t.r1 <= u.r2) { const n = Math.min(t.r2, u.r2) - t.r1 + 1; for (let i = 0; i < n && tbl.rows.length; i++) thead.appendChild(tbl.rows[t.r1 - u.r1]); } else { const th = el('div'); th.innerHTML = wb.rangeHTML(s, { r1: t.r1, c1: u.c1, r2: t.r2, c2: u.c2 }); const tt = th.querySelector('table'); if (tt) Array.from(tt.rows).forEach(r => thead.appendChild(r)); } if (thead.rows.length) tbl.insertBefore(thead, tbl.firstChild); }
          if (page.center) d.style.textAlign = 'center'; box.appendChild(d);
        });
        return box;
      }
      const PAPER = { A4: [210, 297], Letter: [215.9, 279.4], Legal: [215.9, 355.6], A3: [297, 420], A5: [148, 210] }, MARG = { Normal: 12, Wide: 25, Narrow: 6 };
      function hfContent(t) { if (!t) return '""'; const parts = []; String(t).split(/(&P|&N|&D|&T|&F|&A)/).forEach(x => { if (!x) return; if (x === '&P') parts.push('counter(page)'); else if (x === '&N') parts.push('counter(pages)'); else if (x === '&D') parts.push(JSON.stringify(new Date().toLocaleDateString())); else if (x === '&T') parts.push(JSON.stringify(new Date().toLocaleTimeString())); else if (x === '&F') parts.push(JSON.stringify(baseName())); else if (x === '&A') parts.push(JSON.stringify(g.sh().name)); else parts.push(JSON.stringify(x)); }); return parts.join(' ') || '""'; }
      function doPrint() {
        const box = printBox(); if (!box.children.length) return toast('Print karne ke liye kuch data nahi hai');
        const pp = PAPER[page.paper || 'A4'], land = page.orient === 'landscape', m = MARG[page.margin || 'Normal'] || 12; const hf = page.hf || {};
        const css = '@page{size:' + (page.paper || 'A4') + ' ' + (land ? 'landscape' : 'portrait') + ';margin:' + m + 'mm' + (hf.h || hf.f ? ' ' + Math.max(m, 15) + 'mm' : '') + ';' + [['top', 'h'], ['bottom', 'f']].map(([pos, k]) => ['left', 'center', 'right'].map(al => '@' + pos + '-' + al + '{content:' + hfContent((hf[k] || {})[al]) + ';font:9pt Calibri,Arial,sans-serif;color:#333}').join('')).join('') + '}';
        const stl = el('style', { text: css }); document.head.appendChild(stl); document.body.appendChild(box); document.body.classList.add('printing');
        let sc = 1; if (page.fit) { const availPx = ((land ? pp[1] : pp[0]) - 2 * m) / 25.4 * 96; const w = Math.max(...Array.from(box.querySelectorAll('table')).map(t => t.scrollWidth), 1); sc = Math.min(1, availPx / w); } else if (page.scale && page.scale !== 100) sc = page.scale / 100; if (sc !== 1) box.style.zoom = sc;
        const done = () => { document.body.classList.remove('printing'); box.remove(); stl.remove(); window.removeEventListener('afterprint', done); }; window.addEventListener('afterprint', done); setTimeout(() => window.print(), 60);
      }
      async function doExport(kind) {
        const n = baseName();
        try {
          if (g.editing) g.commit(0, 0, true); wb.ensureCalc();
          if (kind === 'xlsx') { OA.download(await XL.toXlsx(wb), n + '.xlsx', MIME); toast('Exported ' + n + '.xlsx'); }
          else if (kind === 'csv' || kind === 'tsv') { const s = g.sh(); const u = used(s) || { r1: 0, c1: 0, r2: 0, c2: 0 }; OA.download('﻿' + XL.toCSV(wb, s, u, kind === 'tsv' ? '\t' : ','), n + (kind === 'tsv' ? '.tsv' : '.csv'), 'text/csv;charset=utf-8'); toast('Exported ' + n + '.' + kind); }
          else if (kind === 'html') { const b = printBox(); OA.download('<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + OA.esc(n) + '</title><style>body{font-family:Calibri,Arial,sans-serif;margin:20px}table{border-collapse:collapse;margin-bottom:24px}td{padding:2px 6px;border:1px solid #bbb}</style></head><body>' + b.innerHTML + '</body></html>', n + '.html', 'text/html'); toast('Exported ' + n + '.html'); }
          else if (kind === 'pdf') { toast('Print dialog mein "Save as PDF" chuniye'); doPrint(); }
          dirty = false;
        } catch (e) { console.error(e); OA.alert('Export fail hua: ' + OA.esc(e.message)); }
      }
      async function openFile(file) {
        try {
          const ext = OA.ext(file.name); let nb;
          if (ext === 'xlsx') nb = await XL.fromXlsx(new Uint8Array(await file.arrayBuffer()));
          else { const t = await file.text(); const first = t.split('\n')[0]; const d = ext === 'tsv' || /\t/.test(first) ? '\t' : (first.split(';').length > first.split(',').length ? ';' : ','); nb = newBook(); wb && 0; nb.loadRows(nb.sheets[0], XL.parseCSV(t, d)); nb.hist.undo.length = 0; }
          sh.name.value = OA.baseName(file.name); setBook(nb); toast('Opened ' + file.name);
        } catch (e) { console.error(e); OA.alert('Yeh file open nahi ho payi: ' + OA.esc(e.message)); }
      }
      async function openPicker() { const f = await OA.pickFile('.xlsx,.csv,.tsv,.txt'); if (f) openFile(f); }
      function newDoc(tpl) { setBook(newBook(tpl)); sh.name.value = 'Book1'; }

      /* ---------- ribbon */
      const B = (id, title, icon, onClick, o) => Object.assign({ id, icon, title, onClick, size: 'sm' }, o || {});
      const L = (id, label, icon, onClick, o) => Object.assign({ id, label, icon, title: label, onClick, size: 'lg' }, o || {});
      const G = (id, gl, title, onClick, style) => ({ id, glyph: gl, title, onClick, size: 'sm', gstyle: style || '' });
      const NF = { General: 'General', Number: '0.00', Currency: '"₹"#,##0.00', Accounting: '"₹" #,##0.00;("₹" #,##0.00)', 'Short Date': 'm/d/yyyy', 'Long Date': 'dddd, mmmm d, yyyy', Time: 'h:mm:ss AM/PM', Percentage: '0.00%', Fraction: '# ?/?', Scientific: '0.00E+00', Text: '@' };
      const borderMenu = [['All Borders', 'all'], ['Outside Borders', 'outer'], ['Inside Borders', 'inner'], ['Top Border', 'top'], ['Bottom Border', 'bottom'], ['Left Border', 'left'], ['Right Border', 'right'], ['Top and Bottom Border', 'topbottom'], ['No Border', 'none']].map(([l, m]) => ({ label: l, onClick: () => { wb.tx(() => eachRange(r => wb.applyBorders(g.sh(), r, m, 'thin', '#000000'))); } })).concat([{ sep: true }, { label: 'Thick Outside Border', onClick: () => wb.tx(() => eachRange(r => wb.applyBorders(g.sh(), r, 'outer', 'thick', '#000000'))) }, { label: 'More Borders…', onClick: () => XL.dlg.formatCells(ctx, 'Border') }]);
      const mergeMenu = [{ label: 'Merge & Center', onClick: () => { wb.tx(() => eachRange(r => { if (r.r2 > r.r1 || r.c2 > r.c1) { wb.merge(g.sh(), r); wb.applyStyle(g.sh(), { r1: r.r1, c1: r.c1, r2: r.r1, c2: r.c1 }, { ha: 'center' }); } })); } }, { label: 'Merge Cells', onClick: () => wb.tx(() => eachRange(r => (r.r2 > r.r1 || r.c2 > r.c1) && wb.merge(g.sh(), r))) }, { label: 'Unmerge Cells', onClick: () => wb.tx(() => eachRange(r => wb.unmerge(g.sh(), r))) }];
      const cfMenu = [{ header: 'Highlight Cells Rules' }, { label: 'Greater Than / Less Than / Between…', onClick: () => XL.dlg.cfRule(ctx, 'cell') }, { label: 'Text that Contains…', onClick: () => XL.dlg.cfRule(ctx, 'text') }, { label: 'Duplicate Values…', onClick: () => XL.dlg.cfRule(ctx, 'dup') }, { header: 'Top/Bottom Rules' }, { label: 'Top 10 Items…', onClick: () => XL.dlg.cfRule(ctx, 'top') }, { label: 'Bottom 10 Items…', onClick: () => XL.dlg.cfRule(ctx, 'bottom') }, { label: 'Above / Below Average…', onClick: () => XL.dlg.cfRule(ctx, 'avg') }, { header: 'Visuals' }, { label: 'Data Bars', onClick: () => XL.dlg.cfRule(ctx, 'databar') }, { label: 'Color Scales', onClick: () => XL.dlg.cfRule(ctx, 'scale') }, { label: 'Icon Sets', onClick: () => XL.dlg.cfRule(ctx, 'icons') }, { header: 'Other' }, { label: 'New Rule (Formula)…', onClick: () => XL.dlg.cfRule(ctx, 'formula') }, { label: 'Clear Rules from Selected Cells', onClick: () => { eachRange(r => wb.clearCF(g.sh(), r)); ctx.after(); } }, { label: 'Clear Rules from Entire Sheet', onClick: () => { wb.clearCF(g.sh(), null); ctx.after(); } }, { label: 'Manage Rules…', onClick: () => XL.dlg.manageCF(ctx) }];
      const chartMenu = [['col', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'], ['scatter', 'Scatter']].concat(XL.CHART_TYPES_EXTRA || []).map(([k, l]) => ({ label: l + ' Chart…', onClick: () => XL.dlg.chart(ctx, k) }));
      const fontSel = { type: 'combo', id: 'font', options: XL.FONTS, width: 120, value: 'Calibri', title: 'Font', inputmode: 'text', onChange: v => v && fmt({ font: v === 'Calibri' ? null : v }) };
      const sizeSel = { type: 'combo', id: 'size', options: ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48', '72'], width: 46, value: '11', title: 'Font size', onChange: v => { const n = parseFloat(v); if (n > 0 && n < 410) fmt({ sz: n === 11 ? null : n }); } };
      const nfSel = { type: 'select', id: 'nf', options: Object.keys(NF), width: 120, title: 'Number Format', onChange: v => fmt({ nf: NF[v] === 'General' ? null : NF[v] }) };
      const stepFont = d => { const n = (curSt().sz || 11) + d; if (n > 0) fmt({ sz: n }); };
      const spec = {
        file: { onClick: () => bs.open('info') },
        tabs: [
          { id: 'home', label: 'Home', groups: [
            { label: 'Clipboard', items: [L('paste', 'Paste', 'paste', () => g.pasteButton(), { menu: [{ label: 'Paste', onClick: () => g.pasteButton() }, { label: 'Values only', onClick: () => g.pasteButton('values') }, { label: 'Formulas', onClick: () => g.pasteButton('formulas') }, { label: 'Formatting only', onClick: () => g.pasteButton('formats') }, { label: 'Transpose', onClick: () => g.pasteButton('all', { transpose: true }) }, { label: 'Paste Special…', onClick: () => XL.dlg.pasteSpecial(ctx) }] }), { type: 'row', items: [B('cut', 'Cut', 'cut', () => g.copyButton(true), { label: 'Cut' }), B('copy', 'Copy', 'copy', () => g.copyButton(false), { label: 'Copy' })] }] },
            { label: 'Font', items: [{ type: 'row', items: [fontSel, sizeSel, B('grow', 'Increase Font Size', null, () => stepFont(1), { glyph: 'A▲', gstyle: 'font-size:11px' }), B('shrink', 'Decrease Font Size', null, () => stepFont(-1), { glyph: 'A▼', gstyle: 'font-size:10px' })] },
              { type: 'row', items: [G('b', 'B', 'Bold (Ctrl+B)', () => toggle('b'), 'font-weight:700'), G('i', 'I', 'Italic (Ctrl+I)', () => toggle('i'), 'font-style:italic;font-family:serif'), G('u', 'U', 'Underline (Ctrl+U)', () => toggle('u'), 'text-decoration:underline'), B('brd', 'Borders', 'border', null, { menu: borderMenu }), { type: 'color', id: 'bg', icon: 'bucket', title: 'Fill Color', default: 'FFFF00', noColor: 'No Fill', onPick: c => fmt({ bg: c ? '#' + c : null }) }, { type: 'color', id: 'fc', glyph: 'A', title: 'Font Color', default: 'FF0000', noColor: 'Automatic', onPick: c => fmt({ color: c ? '#' + c : null }) }] }] },
            { label: 'Alignment', items: [{ type: 'row', items: [B('vt', 'Top Align', null, () => fmt({ va: 'top' }), { glyph: '⤒' }), B('vm', 'Middle Align', null, () => fmt({ va: 'middle' }), { glyph: '↕' }), B('vb', 'Bottom Align', null, () => fmt({ va: null }), { glyph: '⤓' }), B('wrap', 'Wrap Text', null, () => toggle('wrap'), { glyph: '↵', label: 'Wrap' })] },
              { type: 'row', items: [B('al', 'Align Left', 'alignLeft', () => fmt({ ha: 'left' })), B('ac', 'Center', 'alignCenter', () => fmt({ ha: 'center' })), B('ar', 'Align Right', 'alignRight', () => fmt({ ha: 'right' })), B('outd', 'Decrease Indent', 'outdent', () => fmt({ ind: Math.max(0, (curSt().ind || 0) - 1) || null })), B('ind', 'Increase Indent', 'indent', () => fmt({ ind: (curSt().ind || 0) + 1 })), B('mrg', 'Merge & Center', 'merge', null, { menu: mergeMenu })] }] },
            { label: 'Number', items: [{ type: 'row', items: [nfSel] }, { type: 'row', items: [G('cur', '₹', 'Currency', () => fmt({ nf: NF.Currency })), G('pct', '%', 'Percent Style', () => fmt({ nf: '0%' })), G('com', ',', 'Comma Style', () => fmt({ nf: '#,##0.00' })), G('d0', '.00→.0', 'Decrease Decimal', () => decimals(-1), 'font-size:10px'), G('d1', '.0→.00', 'Increase Decimal', () => decimals(1), 'font-size:10px')] }] },
            { label: 'Styles', items: [L('cf', 'Conditional Formatting', 'layers', null, { menu: cfMenu }), L('fat', 'Format as Table', 'table', null, { menu: TABLESTYLES.map(t => ({ label: t[0] + ' table style', onClick: () => formatAsTable(t) })) }), L('cs', 'Cell Styles', 'pen', null, { menu: CELLSTYLES.map(([l, p]) => ({ label: l, onClick: () => { wb.tx(() => eachRange(r => { wb.applyStyle(g.sh(), r, { b: null, i: null, bg: null, color: null, sz: null }); wb.applyStyle(g.sh(), r, p); })); } })) })] },
            { label: 'Cells', items: [B('ins', 'Insert', 'plus', null, { label: 'Insert', menu: [{ label: 'Insert Sheet Rows', onClick: () => insertDel('row') }, { label: 'Insert Sheet Columns', onClick: () => insertDel('col') }, { label: 'Insert Sheet', onClick: addSheet }] }), B('del', 'Delete', 'minus', null, { label: 'Delete', menu: [{ label: 'Delete Sheet Rows', onClick: () => insertDel('row', true) }, { label: 'Delete Sheet Columns', onClick: () => insertDel('col', true) }, { label: 'Delete Sheet', onClick: () => delSheet(wb.activeIndex) }] }), B('fmtm', 'Format', 'columns', null, { label: 'Format', menu: [{ label: 'Row Height…', onClick: () => sizeDlg('row') }, { label: 'AutoFit Row Height', onClick: () => g.autofitRows() }, { label: 'Column Width…', onClick: () => sizeDlg('col') }, { label: 'AutoFit Column Width', onClick: () => g.autofitCols() }, { sep: true }, { label: 'Hide Rows', onClick: () => hideSet('row', true) }, { label: 'Hide Columns', onClick: () => hideSet('col', true) }, { label: 'Unhide Rows', onClick: () => unhideAll('row') }, { label: 'Unhide Columns', onClick: () => unhideAll('col') }, { sep: true }, { label: 'Rename Sheet', onClick: () => renameSheet(wb.activeIndex) }, { label: 'Protect Sheet', onClick: protect }, { label: 'Format Cells…', onClick: () => XL.dlg.formatCells(ctx) }] })] },
            { label: 'Editing', items: [{ type: 'row', items: [B('sum', 'AutoSum', null, () => autosum(), { glyph: 'Σ', label: 'AutoSum', menu: ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map(f => ({ label: f, onClick: () => autosum(f) })) }), B('fill', 'Fill', 'arrow', null, { label: 'Fill', menu: [['Down', 'down'], ['Right', 'right'], ['Up', 'up'], ['Left', 'left']].map(([l, d]) => ({ label: l, onClick: () => g.fill(d) })) }), B('clr', 'Clear', 'eraser', null, { label: 'Clear', menu: [{ label: 'Clear All', onClick: () => g.clearContents('all') }, { label: 'Clear Formats', onClick: () => g.clearContents('formats') }, { label: 'Clear Contents', onClick: () => g.clearContents('contents') }] })] },
              { type: 'row', items: [B('srt', 'Sort & Filter', 'sort', null, { label: 'Sort & Filter', menu: [{ label: 'Sort A to Z', onClick: () => XL.dlg.sort(ctx, 'asc') }, { label: 'Sort Z to A', onClick: () => XL.dlg.sort(ctx, 'desc') }, { label: 'Custom Sort…', onClick: () => XL.dlg.sort(ctx) }, { sep: true }, { label: 'Filter on/off', onClick: sortFilterToggle }] }), B('fnd', 'Find & Select', 'search', null, { label: 'Find', menu: [{ label: 'Find…', onClick: () => command('find') }, { label: 'Replace…', onClick: () => command('replace') }, { label: 'Go To…', onClick: () => command('goto') }, { label: 'Select All', onClick: () => g.selectAll() }] })] }] }] },
          { id: 'insert', label: 'Insert', groups: [
            { label: 'Tables', items: [L('tbl', 'Table', 'table', () => formatAsTable()), L('pvt', 'Summary', 'sort', () => { const s = g.sh(); const q = g.range(); wb.ensureCalc(); autosum(); }, { label: 'AutoSum' })] },
            { label: 'Charts', items: [L('chart', 'Chart', 'chart', null, { menu: chartMenu }), B('cc', 'Column', null, () => XL.dlg.chart(ctx, 'col'), { glyph: '▥', label: 'Column' }), B('cl', 'Line', null, () => XL.dlg.chart(ctx, 'line'), { glyph: '〰', label: 'Line' }), B('cp', 'Pie', null, () => XL.dlg.chart(ctx, 'pie'), { glyph: '◔', label: 'Pie' })] },
            { label: 'Links', items: [L('lnk', 'Link', 'link', () => XL.dlg.hyperlink(ctx))] }, { label: 'Comments', items: [L('note', 'Note', 'comment', () => XL.dlg.note(ctx))] },
            { label: 'Text', items: [B('dt', 'Date', null, () => command('insert-date'), { glyph: '📅', label: 'Today', onClick: () => { const d = new Date(); g.insertText((d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear()); } })] }] },
          { id: 'layout', label: 'Page Layout', groups: [
            { label: 'Page Setup', items: [L('ori', 'Orientation', 'orientation', null, { menu: [{ label: 'Portrait', checked: false, onClick: () => { page.orient = 'portrait'; } }, { label: 'Landscape', onClick: () => { page.orient = 'landscape'; } }] }), L('ps', 'Page Setup…', 'margin', () => XL.dlg.pageSetup(ctx))] },
            { label: 'Sheet Options', items: [{ type: 'check', id: 'gv', label: 'Gridlines: View', value: true, onChange: v => { wb.setProp(g.sh(), 'showGrid', v); g.refresh(); } }, { type: 'check', id: 'gp', label: 'Gridlines: Print', value: false, onChange: v => { page.grid = v; } }, { type: 'check', id: 'hv', label: 'Headings: View', value: true, onChange: v => g.setHeaders(v) }] },
            { label: 'Print', items: [L('prt', 'Print', 'print', doPrint)] }] },
          { id: 'formulas', label: 'Formulas', groups: [
            { label: 'Function Library', items: [L('ifn', 'Insert Function', 'plus', () => fx.click()), L('as', 'AutoSum', null, () => autosum(), { glyph: 'Σ' })].concat(['Math', 'Statistical', 'Logical', 'Text', 'Lookup & Reference', 'Date & Time', 'Financial'].map((c, i) => B('fl' + i, c, null, null, { glyph: c[0], label: c.replace(' & Reference', ''), menu: funcMenu(c) }))) },
            { label: 'Defined Names', items: [L('nm', 'Name Manager', 'list', () => XL.dlg.nameManager(ctx)), B('dn', 'Define Name', null, () => XL.dlg.defineName(ctx), { glyph: '⊞', label: 'Define Name' })] },
            { label: 'Calculation', items: [{ type: 'check', id: 'sf', label: 'Show Formulas', value: false, onChange: v => { g.showFormulas = v; g.draw(); } }, L('calc', 'Calculate Now', 'play', () => { wb.dirty(); g.refresh(); toast('Recalculated'); })] }] },
          { id: 'data', label: 'Data', groups: [
            { label: 'Get Data', items: [L('csv', 'From Text/CSV', 'folder', importCSV)] },
            { label: 'Sort & Filter', items: [B('sa', 'Sort A to Z', null, () => XL.dlg.sort(ctx, 'asc'), { glyph: 'A↓Z', label: 'A→Z' }), B('sz', 'Sort Z to A', null, () => XL.dlg.sort(ctx, 'desc'), { glyph: 'Z↓A', label: 'Z→A' }), L('sort', 'Sort', 'sort', () => XL.dlg.sort(ctx)), L('flt', 'Filter', 'filter', sortFilterToggle), B('rf', 'Reapply', null, () => { wb.applyFilter(g.sh()); g.refresh(); }, { glyph: '↻', label: 'Reapply' })] },
            { label: 'Data Tools', items: [L('ttc', 'Text to Columns', 'columns', () => XL.dlg.textToColumns(ctx)), L('rd', 'Remove Duplicates', 'eraser', () => XL.dlg.removeDuplicates(ctx)), L('dv', 'Data Validation', 'check2', () => XL.dlg.dataValidation(ctx))] }] },
          { id: 'review', label: 'Review', groups: [
            { label: 'Notes', items: [L('nn', 'New Note', 'comment', () => XL.dlg.note(ctx))] }, { label: 'Protect', items: [L('pr', 'Protect Sheet', 'check', protect)] },
            { label: 'Proofing', items: [L('wc', 'Workbook Info', 'info', () => bs.open('info'))] }] },
          { id: 'view', label: 'View', groups: [
            { label: 'Show', items: [{ type: 'check', id: 'fbv', label: 'Formula Bar', value: true, onChange: v => { fbar.hidden = !v; g.refresh(); resizeSoon(); } }, { type: 'check', id: 'gv2', label: 'Gridlines', value: true, onChange: v => { wb.setProp(g.sh(), 'showGrid', v); g.refresh(); } }, { type: 'check', id: 'hv2', label: 'Headings', value: true, onChange: v => g.setHeaders(v) }] },
            { label: 'Zoom', items: [L('z100', '100%', 'zoomIn', () => g.setZoom(1)), B('zi', 'Zoom In', 'zoomIn', () => g.setZoom(g.zoom + 0.1), { label: 'In' }), B('zo', 'Zoom Out', 'zoomOut', () => g.setZoom(g.zoom - 0.1), { label: 'Out' })] },
            { label: 'Window', items: [L('frz', 'Freeze Panes', 'layers', null, { menu: [{ label: 'Freeze Panes (at selected cell)', onClick: () => freeze('pane') }, { label: 'Freeze Top Row', onClick: () => freeze('row') }, { label: 'Freeze First Column', onClick: () => freeze('col') }, { label: 'Unfreeze Panes', onClick: () => freeze('none') }] })] }] },
          { id: 'chart', label: 'Chart Design', contextual: true, hidden: true, groups: [
            { label: 'Type', items: [L('cty', 'Change Chart', 'chart', () => g.chart() && XL.dlg.chart(ctx, null, g.chart()))] },
            { label: 'Chart Layouts', items: [B('ctt', 'Chart Title', null, async () => { const c = g.chart(); if (!c) return; const t = await OA.prompt('Chart title:', c.title || '', { title: 'Chart Title' }); if (t !== null) g.updateChart(c, { title: t }); }, { glyph: 'T', label: 'Title' }), B('clg', 'Legend', null, () => { const c = g.chart(); c && g.updateChart(c, { legend: c.legend === false }); }, { glyph: '▤', label: 'Legend' }), B('cst', 'Stacked', null, () => { const c = g.chart(); c && g.updateChart(c, { stacked: !c.stacked }); }, { glyph: '▦', label: 'Stack' }), B('csw', 'Switch Row/Column', null, () => { const c = g.chart(); c && g.updateChart(c, { rowsAsSeries: !c.rowsAsSeries }); }, { glyph: '⇄', label: 'Switch' })] },
            { label: 'Actions', items: [L('cdl', 'Delete Chart', 'trash', () => { const c = g.chart(); c && g.deleteChart(c.id); })] }] }
        ]
      };
      const resizeSoon = () => setTimeout(() => window.dispatchEvent(new Event('resize')), 30);
      ctx.H = { B, L, G, el, fmt, eachRange, curSt, toast, page, command, doPrint, autosum, importCSV, root, host, get rb() { return rb; }, get dirty() { return dirty; }, set dirty(v) { dirty = v; }, updateBar, syncRibbon, activate, renderTabs };
      (XL.extras || []).forEach(f => f(ctx, spec, ctx.H));
      spec.right = el('div', { class: 'rtab-right' }); rb = OA.buildRibbon(sh.ribbon, spec); const ctl = rb.ctl;
      function syncRibbon() {
        if (!rb || !g) return; const s = curSt(); rb.setOn('b', s.b); rb.setOn('i', s.i); rb.setOn('u', s.u); rb.setOn('wrap', s.wrap); rb.setOn('al', s.ha === 'left'); rb.setOn('ac', s.ha === 'center'); rb.setOn('ar', s.ha === 'right');
        if (document.activeElement !== ctl.font) ctl.font.value = s.font || 'Calibri'; if (document.activeElement !== ctl.size) ctl.size.value = String(s.sz || 11);
        const nf = s.nf || 'General'; const key2 = Object.keys(NF).find(k => NF[k] === nf) || (/%/.test(nf) ? 'Percentage' : XL.isDateFormat(nf) ? 'Short Date' : /[$€£¥₹]/.test(nf) ? 'Currency' : nf === '0.00' || /^#,##0/.test(nf) ? 'Number' : 'General'); if (document.activeElement !== ctl.nf) ctl.nf.value = key2;
        const gv = g.sh().showGrid; if (ctl.gv) ctl.gv.checked = gv; if (ctl.gv2) ctl.gv2.checked = gv;
      }

      /* ---------- backstage */
      bs = OA.backstage(root, { items: [
        { id: 'info', label: 'Info', icon: 'info', render(p) {
          p.appendChild(el('p', { class: 'bs-p', text: 'Aapka workbook sirf is browser tab mein hai. Isse bachane ke liye export karein.' }));
          let cells = 0; wb.sheets.forEach(s => cells += s.cells.size);
          p.appendChild(el('div', { class: 'kv', html: `<span>Name</span><span>${OA.esc(baseName())}</span><span>Sheets</span><span>${wb.sheets.length}</span><span>Cells with data</span><span>${cells}</span>` }));
          p.appendChild(el('div', { style: { marginTop: '18px' } }, el('button', { class: 'btn primary', type: 'button', text: 'Export as .xlsx', onclick: () => doExport('xlsx') })));
        } },
        { id: 'new', label: 'New', icon: 'plus', render(p, api) { const go = async k => { if (!dirty || await OA.confirm('Current workbook discard karke naya banayein?', { ok: 'Continue' })) { newDoc(k); api.close(); } };
          p.appendChild(OA.bigActions([{ label: 'Blank workbook', desc: 'Khali spreadsheet', icon: 'table', onClick: () => go() }].concat(Object.keys(TEMPLATES).map(k => ({ label: TEMPLATES[k].name, desc: 'Template', icon: 'table', onClick: () => go(k) }))))); } },
        { id: 'open', label: 'Open', icon: 'folder', render(p, api) { p.appendChild(el('p', { class: 'bs-p', text: 'Apne device se .xlsx, .csv, .tsv ya .txt file kholein.' })); p.appendChild(el('button', { class: 'btn primary', type: 'button', text: 'Browse…', onclick: async () => { const f = await OA.pickFile('.xlsx,.csv,.tsv,.txt'); if (f) { api.close(); openFile(f); } } })); } },
        { id: 'export', label: 'Export / Save As', icon: 'download', render(p) { p.appendChild(el('p', { class: 'bs-p', text: 'Format chuniye. Data is website par save nahi hota.' })); p.appendChild(OA.bigActions([
          { label: 'Excel Workbook (.xlsx)', desc: 'Excel, LibreOffice, Google Sheets mein khulta hai', icon: 'download', onClick: () => doExport('xlsx') }, { label: 'CSV (current sheet)', desc: 'Comma separated', icon: 'download', onClick: () => doExport('csv') }, { label: 'TSV (current sheet)', desc: 'Tab separated', icon: 'download', onClick: () => doExport('tsv') }, { label: 'Web Page (.html)', desc: 'Sabhi sheets tables ke roop mein', icon: 'download', onClick: () => doExport('html') }, { label: 'PDF', desc: 'Print dialog mein "Save as PDF"', icon: 'print', onClick: () => doExport('pdf') }])); } },
        { id: 'print', label: 'Print', icon: 'print', action: doPrint }, { sep: true }, { id: 'home', label: 'Close (Office Home)', icon: 'home', action: () => OA.go('home') }
      ] });
      sh.title.querySelector('.tb-right').append(el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('undo', 16), title: 'Undo (Ctrl+Z)', onclick: () => { wb.undo(); g.focus(); } }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('redo', 16), title: 'Redo (Ctrl+Y)', onclick: () => { wb.redo(); g.focus(); } }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('download', 16), title: 'Export .xlsx', onclick: () => doExport('xlsx') }));
      setBook(newBook()); if (window.innerWidth < 760) g.setZoom(0.85);
      return { onShow() { setTimeout(() => { g.refresh(); g.focus(); }, 60); }, openFile, exportDefault: () => doExport('xlsx'), loadTemplate: k => newDoc(k) };
    }
  });
})();
