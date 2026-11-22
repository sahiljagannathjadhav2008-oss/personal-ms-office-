/* Excel dialogs — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL, OA = global.OA, el = OA.el, key = XL.key;
  const D = XL.dlg = {};
  const hex = c => c ? (c[0] === '#' ? c : '#' + c) : '';
  const FONTS = ['Calibri', 'Arial', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma', 'Georgia', 'Trebuchet MS', 'Comic Sans MS', 'Impact', 'Segoe UI', 'Cambria', 'Consolas', 'Garamond', 'Palatino Linotype', 'Century Gothic'];
  XL.FONTS = FONTS;
  const lab = (t, n) => el('label', { class: 'form-row' }, [el('span', { text: t }), n]);
  const sel = (opts, val, cls) => { const s = el('select', { class: 'inp' }); opts.forEach(o => { const v = Array.isArray(o) ? o[0] : o, l = Array.isArray(o) ? o[1] : o; const e = el('option', { value: v, text: l }); if (String(v) === String(val)) e.selected = true; s.appendChild(e); }); return s; };
  const inp = (v, t, extra) => el('input', Object.assign({ class: 'inp', type: t || 'text', value: v == null ? '' : v }, extra || {}));
  const chk = (t, v) => { const i = el('input', { type: 'checkbox' }); i.checked = !!v; return { i, row: el('label', { class: 'form-row chk' }, [i, el('span', { text: t })]) }; };

  // contiguous region around a cell (like Ctrl+A / Current Region)
  XL.region = function (wb, sh, r, c) {
    wb.ensureCalc(); const has = (a, b) => { if (a < 0 || b < 0 || a >= XL.MAXR || b >= XL.MAXC) return false; const k = key(a, b); const cl = sh.cells.get(k); return !!(cl && (cl.f != null || (cl.v !== undefined && cl.v !== null && cl.v !== ''))) || sh.spill.has(k); };
    if (!has(r, c) && !has(r - 1, c) && !has(r + 1, c) && !has(r, c - 1) && !has(r, c + 1)) return { r1: r, c1: c, r2: r, c2: c };
    let g = { r1: r, c1: c, r2: r, c2: c }, ch = true, n = 0;
    while (ch && n++ < 200) { ch = false;
      const rowHas = rr => { for (let cc = g.c1 - 1; cc <= g.c2 + 1; cc++) if (has(rr, cc)) return true; return false; }, colHas = cc => { for (let rr = g.r1 - 1; rr <= g.r2 + 1; rr++) if (has(rr, cc)) return true; return false; };
      if (g.r1 > 0 && rowHas(g.r1 - 1)) { g.r1--; ch = true; } if (rowHas(g.r2 + 1) && g.r2 < wb.extent(sh.name).r) { g.r2++; ch = true; }
      if (g.c1 > 0 && colHas(g.c1 - 1)) { g.c1--; ch = true; } if (colHas(g.c2 + 1) && g.c2 < wb.extent(sh.name).c) { g.c2++; ch = true; } }
    return g;
  };
  XL.sheetRef = (sh, g) => XL.quoteSheet(sh.name) + '!' + XL.addr(g.r1, g.c1, true, true) + ':' + XL.addr(g.r2, g.c2, true, true);
  XL.rangeLabel = g => g.r1 === g.r2 && g.c1 === g.c2 ? XL.addr(g.r1, g.c1) : XL.addr(g.r1, g.c1) + ':' + XL.addr(g.r2, g.c2);

  /* ---------------- Format Cells */
  const FMT = {
    Number: null, Currency: null, Accounting: null, Date: ['m/d/yyyy', 'd-mmm-yy', 'dd/mm/yyyy', 'yyyy-mm-dd', 'dddd, mmmm d, yyyy', 'mmmm d, yyyy', 'd-mmm', 'mmm-yy', 'mmmm yyyy'], Time: ['h:mm AM/PM', 'h:mm:ss AM/PM', 'h:mm', 'h:mm:ss', 'm/d/yyyy h:mm', '[h]:mm:ss'],
    Percentage: null, Fraction: ['# ?/?', '# ??/??', '?/?', '# ?/2', '# ?/4', '# ?/8'], Scientific: null, Text: ['@'], General: ['General']
  };
  const SYMS = [['', 'None'], ['₹', '₹ Rupee'], ['$', '$ Dollar'], ['€', '€ Euro'], ['£', '£ Pound'], ['¥', '¥ Yen']];
  function catOf(nf) {
    if (!nf || nf === 'General') return 'General'; if (nf === '@') return 'Text'; if (/%/.test(nf)) return 'Percentage'; if (/E[+-]/i.test(nf)) return 'Scientific'; if (/\?\/|\?\?\//.test(nf)) return 'Fraction';
    if (XL.isDateFormat(nf)) return /[ymd]/i.test(nf.replace(/\[[^\]]*\]/g, '').replace(/AM\/PM/ig, '')) ? 'Date' : 'Time'; if (/[$€£¥₹]/.test(nf)) return 'Currency'; if (/^#,##0|^0/.test(nf)) return 'Number'; return 'Custom';
  }
  D.formatCells = function (ctx, tab0) {
    const { wb, g } = ctx; const sh = ctx.sh(); const a = g.active(); const st = Object.assign({}, wb.styleOf(sh, a.r, a.c)); const q = g.range();
    const val = (() => { wb.ensureCalc(); const v = wb.value(sh.name, a.r, a.c); return v === null || v === undefined || v === '' ? 1234.567 : v; })();
    const patch = {}; let nf = st.nf || 'General'; let bdr = { mode: null, style: 'thin', color: '#000000' };
    const tabs = ['Number', 'Alignment', 'Font', 'Border', 'Fill', 'Protection']; const bar = el('div', { class: 'x-fmt-tabs' }), body = el('div', { class: 'x-fmt-body' }); let cur = tab0 || 'Number';
    const panes = {};
    // Number
    (function () {
      const w = el('div'); const cats = ['General', 'Number', 'Currency', 'Accounting', 'Date', 'Time', 'Percentage', 'Fraction', 'Scientific', 'Text', 'Custom'];
      const list = el('div', { class: 'cl' }), right = el('div'), sample = el('div', { class: 'x-sample' }), grid = el('div', { class: 'x-cat' }); let cat = catOf(nf);
      const upd = () => { let t; try { t = XL.formatValue(val, nf).text; } catch (e) { t = '#####'; } sample.textContent = 'Sample: ' + t; };
      const dec = inp(2, 'number', { min: 0, max: 15 }), thou = chk('Use 1000 Separator (,)', true), sym = sel(SYMS, '₹'), codeIn = inp(nf), listBox = el('div', { class: 'x-listbox' });
      const gen = () => {
        const d = Math.max(0, Math.min(15, +dec.value || 0)); const dz = d ? '.' + '0'.repeat(d) : '';
        if (cat === 'Number') nf = (thou.i.checked ? '#,##0' : '0') + dz; else if (cat === 'Currency') { const s = sym.value ? '"' + sym.value + '"' : ''; nf = s + (thou.i.checked ? '#,##0' : '0') + dz + ';[Red]-' + s + (thou.i.checked ? '#,##0' : '0') + dz; if (!sym.value) nf = (thou.i.checked ? '#,##0' : '0') + dz; }
        else if (cat === 'Accounting') { const s = sym.value ? '"' + sym.value + '" ' : ''; nf = s + '#,##0' + dz + ';(' + s + '#,##0' + dz + ')'; } else if (cat === 'Percentage') nf = '0' + dz + '%'; else if (cat === 'Scientific') nf = '0' + dz + 'E+00'; codeIn.value = nf; upd();
      };
      const draw = () => {
        Array.from(list.children).forEach(x => x.classList.toggle('on', x.textContent === cat)); right.innerHTML = ''; listBox.innerHTML = '';
        if (['Number', 'Currency', 'Accounting', 'Percentage', 'Scientific'].includes(cat)) { right.appendChild(lab('Decimal places', dec)); if (cat === 'Number') right.appendChild(thou.row); if (cat === 'Currency' || cat === 'Accounting') right.appendChild(lab('Symbol', sym)); gen(); }
        else if (FMT[cat]) { FMT[cat].forEach(f => { const b = el('div', { class: nf === f ? 'on' : '', text: f === '@' || f === 'General' ? f : f + '   →   ' + XL.formatValue(val, f).text }); b.style.cssText = 'padding:4px 8px;cursor:pointer'; b.onclick = () => { nf = f; Array.from(listBox.children).forEach(x => x.classList.remove('on')); b.classList.add('on'); codeIn.value = nf; upd(); }; listBox.appendChild(b); }); right.appendChild(el('div', { text: 'Type:' })); right.appendChild(listBox); if (cat === 'General') { nf = 'General'; } else if (!FMT[cat].includes(nf)) { nf = FMT[cat][0]; } codeIn.value = nf; upd(); }
        else { right.appendChild(lab('Format code', codeIn)); const all = Object.values(XL.BUILTIN_FMT).concat(['0.00', '#,##0.00_);(#,##0.00)', '[>=100000]##\\,##\\,##0;##,##0', 'dd-mmm-yyyy', '0.0%', '000000', '@']); const lb = el('div', { class: 'x-listbox' }); Array.from(new Set(all)).forEach(f => { const b = el('div', { text: f }); b.style.cssText = 'padding:3px 8px;cursor:pointer'; b.onclick = () => { codeIn.value = f; nf = f; upd(); }; lb.appendChild(b); }); right.appendChild(lb); codeIn.oninput = () => { nf = codeIn.value; upd(); }; upd(); }
      };
      cats.forEach(c => { const d = el('div', { text: c }); d.onclick = () => { cat = c; if (c === 'General') nf = 'General'; draw(); patch.nf = nf; }; list.appendChild(d); });
      [dec, thou.i, sym].forEach(n => n.addEventListener('input', () => { gen(); })); [dec, thou.i, sym, codeIn].forEach(n => n.addEventListener('change', () => { patch.nf = nf; })); right.addEventListener('click', () => { patch.nf = nf; });
      grid.append(list, right); w.append(sample, grid); panes.Number = w; w._init = draw; draw(); patch.nf = undefined; delete patch.nf;
      w._commit = () => { if (nf !== (st.nf || 'General')) patch.nf = nf === 'General' ? null : nf; };
    })();
    // Alignment
    (function () {
      const w = el('div', { class: 'x-bgrid' }); const ha = sel([['general', 'General'], ['left', 'Left (Indent)'], ['center', 'Center'], ['right', 'Right (Indent)']], st.ha || 'general'), va = sel([['top', 'Top'], ['middle', 'Center'], ['bottom', 'Bottom']], st.va || 'bottom'), ind = inp(st.ind || 0, 'number', { min: 0, max: 15 }), wrap = chk('Wrap text', st.wrap), mg = chk('Merge cells', !!wb.mergeAt(sh, q.r1, q.c1)), rot = inp(st.rot || 0, 'number', { min: -90, max: 90 });
      w.append(lab('Horizontal', ha), lab('Vertical', va), lab('Indent', ind), lab('Orientation (degrees)', rot), wrap.row, mg.row); panes.Alignment = w;
      w._commit = () => { patch.ha = ha.value === 'general' ? null : ha.value; patch.va = va.value === 'bottom' ? null : va.value; patch.ind = +ind.value || null; patch.wrap = wrap.i.checked || false; patch.rot = +rot.value || null; patch.merge = mg.i.checked; };
    })();
    // Font
    (function () {
      const w = el('div', { class: 'x-bgrid' }); const f = sel(FONTS, st.font || 'Calibri'), sz = inp(st.sz || 11, 'number', { min: 1, max: 409, step: 0.5 }), col = inp(hex(st.color) || '#000000', 'color'), b = chk('Bold', st.b), i = chk('Italic', st.i), u = chk('Underline', st.u), s = chk('Strikethrough', st.st), auto = chk('Automatic colour', !st.color), prev = el('div', { class: 'x-sample', text: 'AaBbCcYyZz' });
      const up = () => { prev.style.cssText = 'font-family:' + f.value + ';font-size:' + sz.value + 'px;font-weight:' + (b.i.checked ? 700 : 400) + ';font-style:' + (i.i.checked ? 'italic' : 'normal') + ';text-decoration:' + [u.i.checked ? 'underline' : '', s.i.checked ? 'line-through' : ''].join(' ') + ';color:' + (auto.i.checked ? '#000' : col.value); };
      [f, sz, col, b.i, i.i, u.i, s.i, auto.i].forEach(n => n.addEventListener('input', up)); up();
      w.append(lab('Font', f), lab('Size', sz), lab('Color', col), auto.row, b.row, i.row, u.row, s.row, prev); panes.Font = w;
      w._commit = () => { patch.font = f.value === 'Calibri' ? null : f.value; patch.sz = +sz.value === 11 ? null : +sz.value; patch.color = auto.i.checked ? null : col.value; patch.b = b.i.checked; patch.i = i.i.checked; patch.u = u.i.checked; patch.st = s.i.checked; };
    })();
    // Border
    (function () {
      const w = el('div'); const style = sel([['thin', 'Thin'], ['medium', 'Medium'], ['thick', 'Thick'], ['dashed', 'Dashed'], ['dotted', 'Dotted'], ['double', 'Double'], ['hair', 'Hair']], 'thin'), color = inp('#000000', 'color'); const row = el('div', { class: 'x-sortrow' });
      const modes = [['none', 'None'], ['outer', 'Outline'], ['inner', 'Inside'], ['all', 'All borders'], ['top', 'Top'], ['bottom', 'Bottom'], ['left', 'Left'], ['right', 'Right'], ['topbottom', 'Top & Bottom']]; const sels = new Set();
      modes.forEach(([m, l]) => { const b = el('button', { class: 'btn', type: 'button', text: l }); b.onclick = () => { if (m === 'none') sels.clear(); sels.add(m); if (m === 'none') { Array.from(row.children).forEach(x => x.classList.remove('primary')); } b.classList.toggle('primary', true); }; row.appendChild(b); });
      w.append(el('p', { text: 'Pehle border ki jagah choose karein, phir style/colour:' }), row, lab('Style', style), lab('Color', color)); panes.Border = w;
      w._commit = () => { patch.borders = { modes: Array.from(sels), style: style.value, color: color.value }; };
    })();
    // Fill
    (function () {
      const w = el('div'); const col = inp(hex(st.bg) || '#ffff00', 'color'), none = chk('No colour', !st.bg); const sw = el('div', { class: 'cp-grid', style: 'margin:8px 0' });
      ['FFFFFF', 'F2F2F2', 'D9D9D9', 'FFFF00', 'FFC000', 'F4B084', 'FF9999', 'C6E0B4', '92D050', '9DC3E6', 'B4C6E7', 'D9B3FF', '00B0F0', '4472C4', '70AD47', 'C00000'].forEach(c => { const b = el('button', { type: 'button', class: 'cp-sw', style: { background: '#' + c }, title: '#' + c }); b.onclick = () => { col.value = '#' + c.toLowerCase(); none.i.checked = false; }; sw.appendChild(b); });
      w.append(none.row, sw, lab('Custom colour', col)); panes.Fill = w; w._commit = () => { patch.bg = none.i.checked ? null : col.value; };
    })();
    // Protection
    (function () { const w = el('div'); const lk = chk('Locked', st.locked !== false); w.append(lk.row, el('p', { class: 'bs-p', text: 'Locked cells tabhi protect hote hain jab Review → Protect Sheet on ho.' })); panes.Protection = w; w._commit = () => { patch.locked = lk.i.checked ? null : false; }; })();
    const show = t => { cur = t; Array.from(bar.children).forEach(b => b.classList.toggle('on', b.textContent === t)); body.innerHTML = ''; body.appendChild(panes[t]); };
    tabs.forEach(t => { const b = el('button', { type: 'button', text: t }); b.onclick = () => show(t); bar.appendChild(b); }); show(cur);
    OA.dialog({ title: 'Format Cells', width: 560, content: el('div', {}, [bar, body]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => {
      tabs.forEach(t => panes[t]._commit && panes[t]._commit()); const bp = patch.borders, mg = patch.merge; delete patch.borders; delete patch.merge;
      wb.tx(() => { g.ranges().forEach(r0 => { const rr = { r1: Math.min(r0.r1, r0.r2), c1: Math.min(r0.c1, r0.c2), r2: Math.max(r0.r1, r0.r2), c2: Math.max(r0.c1, r0.c2) }; if (rr.r2 >= XL.MAXR - 1) rr.fullCol = true; else if (rr.c2 >= XL.MAXC - 1) rr.fullRow = true;
        wb.applyStyle(sh, rr, patch); if (bp && bp.modes.length) bp.modes.forEach(m => { if (m === 'none') wb.applyBorders(sh, rr, 'none'); else wb.applyBorders(sh, rr, m, bp.style, bp.color); });
        if (mg && !wb.mergeAt(sh, rr.r1, rr.c1) && (rr.r2 > rr.r1 || rr.c2 > rr.c1)) wb.merge(sh, rr); else if (mg === false && wb.mergeAt(sh, rr.r1, rr.c1)) wb.unmerge(sh, rr); }); });
      ctx.after();
    } }] });
  };

  /* ---------------- Paste special */
  D.pasteSpecial = function (ctx) {
    const modes = [['all', 'All'], ['values', 'Values'], ['formulas', 'Formulas'], ['formats', 'Formats'], ['colwidths', 'Column widths'], ['comments', 'Notes']]; const s = sel(modes, 'all'), tr = chk('Transpose', false), skip = chk('Skip blanks', false), op = sel([['none', 'None'], ['add', 'Add'], ['subtract', 'Subtract'], ['multiply', 'Multiply'], ['divide', 'Divide']], 'none');
    if (!ctx.g.clip) { OA.toast('Pehle koi range copy karein'); return; }
    OA.dialog({ title: 'Paste Special', content: el('div', { class: 'form-grid' }, [lab('Paste', s), lab('Operation', op), tr.row, skip.row]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { ctx.g.pasteText(ctx.g.clip.text, s.value, { transpose: tr.i.checked, skipBlanks: skip.i.checked, op: op.value }); } }] });
  };

  /* ---------------- Sort */
  D.sort = function (ctx, quick) {
    const { wb, g } = ctx; const sh = ctx.sh(); let q = g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, sh, q.r1, q.c1); if (q.r2 >= XL.MAXR - 1 || q.c2 >= XL.MAXC - 1) { const e = wb.extent(sh.name); q.r2 = Math.min(q.r2, e.r); q.c2 = Math.min(q.c2, e.c); }
    if (quick) { wb.ensureCalc(); const hasH = typeof wb.value(sh.name, q.r1, q.c1) === 'string' && q.r2 > q.r1 && typeof wb.value(sh.name, q.r1 + 1, q.c1) !== 'string'; wb.sortRange(sh, q, [{ col: Math.max(q.c1, Math.min(q.c2, g.active().c)), asc: quick === 'asc' }], hasH); ctx.after(); return; }
    const hdr = chk('My data has headers', true); const keys = []; const wrap = el('div');
    const cols = () => { const out = []; for (let c = q.c1; c <= q.c2; c++) out.push([c, hdr.i.checked ? (String(wb.display(sh, q.r1, c).text) || 'Column ' + XL.colName(c)) : 'Column ' + XL.colName(c)]); return out; };
    const rows = []; for (let i = 0; i < 3; i++) { const cs = sel(cols(), Math.min(q.c2, q.c1 + i)), ord = sel([['asc', 'A to Z / Smallest to Largest'], ['desc', 'Z to A / Largest to Smallest']], 'asc'); rows.push({ cs, ord }); wrap.appendChild(el('div', { class: 'x-sortrow' }, [el('b', { text: i ? 'Then by' : 'Sort by' }), cs, ord])); }
    hdr.i.addEventListener('change', () => rows.forEach(r => { const v = r.cs.value; r.cs.innerHTML = ''; cols().forEach(([c, l]) => r.cs.appendChild(el('option', { value: c, text: l }))); r.cs.value = v; }));
    rows[1].cs.insertBefore(el('option', { value: '-1', text: '(none)' }), rows[1].cs.firstChild); rows[1].cs.value = '-1'; rows[2].cs.insertBefore(el('option', { value: '-1', text: '(none)' }), rows[2].cs.firstChild); rows[2].cs.value = '-1';
    OA.dialog({ title: 'Sort — ' + XL.rangeLabel(q), width: 520, content: el('div', {}, [wrap, hdr.row]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { const ks = rows.filter(r => +r.cs.value >= 0).map(r => ({ col: +r.cs.value, asc: r.ord.value === 'asc' })); wb.sortRange(sh, q, ks, hdr.i.checked); g.select(q); ctx.after(); } }] });
  };

  /* ---------------- Find & Replace */
  D.findReplace = function (ctx, replace) {
    const { wb, g } = ctx; const f = inp(ctx.lastFind || ''), r = inp(''), mc = chk('Match case', false), whole = chk('Match entire cell contents', false), scope = sel([['sheet', 'Sheet'], ['book', 'Workbook']], 'sheet'); let idx = -1; const info = el('div', { class: 'bs-p', text: '' });
    const opts = () => ({ sheet: ctx.sh(), matchCase: mc.i.checked, whole: whole.i.checked, allSheets: scope.value === 'book' });
    const list = () => wb.findAll(f.value, opts()).filter(x => x.sheet);
    const goto = hit => { if (hit.sheet !== ctx.sh()) ctx.activate(hit.sheet.name); g.setActive(hit.r, hit.c); };
    const next = () => { if (!f.value) return; ctx.lastFind = f.value; const L = list(); if (!L.length) { info.textContent = 'Kuch nahi mila.'; return; } const a = g.active(); let k = L.findIndex(h => h.sheet === ctx.sh() && (h.r > a.r || (h.r === a.r && h.c > a.c))); if (k < 0) k = 0; goto(L[k]); info.textContent = (k + 1) + ' of ' + L.length; };
    const rep1 = () => { const a = g.active(); const L = list(); const cur = L.find(h => h.sheet === ctx.sh() && h.r === a.r && h.c === a.c); if (cur) { wb.replaceAt(cur.sheet, cur.r, cur.c, f.value, r.value, { matchCase: mc.i.checked, whole: whole.i.checked }); ctx.after(); } next(); };
    const all = () => { const L = list(); wb.tx(() => L.forEach(h => wb.replaceAt(h.sheet, h.r, h.c, f.value, r.value, { matchCase: mc.i.checked, whole: whole.i.checked }))); info.textContent = L.length + ' replacement(s) ho gaye.'; ctx.after(); };
    const all2 = () => { const L = list(); info.textContent = L.length + ' cell(s) mile: ' + L.slice(0, 12).map(h => (scope.value === 'book' ? h.sheet.name + '!' : '') + XL.addr(h.r, h.c)).join(', ') + (L.length > 12 ? '…' : ''); };
    const c = el('div', { class: 'form-grid' }, [lab('Find what', f)]); if (replace) c.appendChild(lab('Replace with', r)); c.append(lab('Within', scope), mc.row, whole.row, info);
    const btns = [{ label: 'Close', value: false }, { label: 'Find All', onClick: () => { all2(); return false; } }]; if (replace) btns.push({ label: 'Replace', onClick: () => { rep1(); return false; } }, { label: 'Replace All', onClick: () => { all(); return false; } }); btns.push({ label: 'Find Next', primary: true, onClick: () => { next(); return false; } });
    OA.dialog({ title: replace ? 'Find and Replace' : 'Find', content: c, buttons: btns, width: 420 });
  };

  /* ---------------- Go To */
  D.goTo = async function (ctx) { const v = await OA.prompt('Reference (jaise B5, A1:C10, ya koi Name):', '', { title: 'Go To' }); if (v) { if (!ctx.g.gotoRef(v)) OA.alert('Reference valid nahi hai.'); } };

  /* ---------------- Insert Function */
  D.insertFunction = function (ctx, onPick) {
    const cats = Object.keys(XL.FDOC); const all = [].concat(...cats.map(k => XL.FDOC[k])); let cat = 'Math'; const search = inp('', 'text', { placeholder: 'Function search…' }), cl = sel(['All'].concat(cats), 'Math'), lb = el('div', { class: 'x-listbox', style: 'height:200px' }), desc = el('div', { class: 'x-sample' }); let chosen = null;
    const draw = () => { lb.innerHTML = ''; const base = cl.value === 'All' ? all : XL.FDOC[cl.value]; const qq = search.value.toUpperCase(); base.filter(n => !qq || n.includes(qq)).sort().forEach(n => { const d = el('div', { text: n }); d.style.cssText = 'padding:3px 8px;cursor:pointer'; d.onclick = () => { chosen = n; Array.from(lb.children).forEach(x => x.style.background = ''); d.style.background = '#cfe3f7'; desc.textContent = XL.FSIG[n] || n + '(…)'; }; d.ondblclick = () => { chosen = n; ok.click(); }; lb.appendChild(d); }); };
    search.oninput = draw; cl.onchange = draw; draw(); let ok;
    const api = OA.dialog({ title: 'Insert Function', width: 460, content: el('div', { class: 'form-grid' }, [search, lab('Category', cl), lb, desc]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { if (!chosen) return false; onPick(chosen); } }] }); ok = api.el.querySelector('.btn.primary');
  };

  /* ---------------- Name manager */
  D.nameManager = function (ctx, preset) {
    const { wb } = ctx; const wrap = el('div'); const list = el('div', { class: 'x-listbox', style: 'height:160px' }); const nm = inp(preset ? preset.name : ''), rf = inp(preset ? preset.ref : '');
    const draw = () => { list.innerHTML = ''; Object.values(wb.names).forEach(n => { const d = el('div', { style: 'padding:4px 8px;display:flex;gap:8px;cursor:pointer' }, [el('b', { text: n.name }), el('span', { text: '=' + n.ref, style: 'color:#666;flex:1' }), el('button', { class: 'btn', type: 'button', text: '✕', onclick: e => { e.stopPropagation(); wb.deleteName(n.name); ctx.after(); draw(); } })]); d.onclick = () => { nm.value = n.name; rf.value = '=' + n.ref; }; list.appendChild(d); }); if (!list.children.length) list.appendChild(el('div', { style: 'padding:8px;color:#777', text: 'Koi defined name nahi hai.' })); };
    draw(); const add = el('button', { class: 'btn primary', type: 'button', text: 'Add / Update', onclick: () => { const n = nm.value.trim(); if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n)) { OA.toast('Name valid nahi hai'); return; } if (!rf.value.trim()) { OA.toast('Refers to bhariye'); return; } wb.tx(() => wb.defineName(n, rf.value.trim(), ctx.sh().name)); ctx.after(); draw(); } });
    wrap.append(list, lab('Name', nm), lab('Refers to', rf), add); OA.dialog({ title: 'Name Manager', content: wrap, width: 460, buttons: [{ label: 'Close', primary: true }] });
  };
  D.defineName = async function (ctx) { const q = ctx.g.range(); D.nameManager(ctx, { name: '', ref: '=' + XL.sheetRef(ctx.sh(), q) }); };

  /* ---------------- Data validation */
  D.dataValidation = function (ctx) {
    const { wb, g } = ctx; const sh = ctx.sh(); const q = g.range(); const a = g.active(); const ex = wb.dvAt(sh, a.r, a.c) || {};
    const type = sel([['any', 'Any value'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['list', 'List'], ['date', 'Date'], ['textLength', 'Text length'], ['custom', 'Custom (formula)']], ex.type || 'any');
    const op = sel([['between', 'between'], ['notbetween', 'not between'], ['=', 'equal to'], ['<>', 'not equal to'], ['>', 'greater than'], ['<', 'less than'], ['>=', '>='], ['<=', '<=']], ex.op || 'between');
    const v1 = inp(Array.isArray(ex.list) ? ex.list.join(',') : (ex.type === 'list' ? ex.list : ex.v1) || ''), v2 = inp(ex.v2 || ''), blank = chk('Ignore blank', ex.allowBlank !== false), al = sel([['stop', 'Stop'], ['warning', 'Warning'], ['information', 'Information']], ex.alert || 'stop'), msg = inp(ex.errMsg || '');
    const hint = el('div', { class: 'bs-p' });
    const upd = () => { const t = type.value; hint.textContent = t === 'list' ? 'Source: comma se alag items (Haan,Nahi) ya =A1:A5' : t === 'custom' ? 'Formula jaise =A1>0' : t === 'any' ? '' : 'Minimum / Maximum ya Value'; };
    type.onchange = upd; upd();
    OA.dialog({ title: 'Data Validation — ' + XL.rangeLabel(q), width: 420, content: el('div', { class: 'form-grid' }, [lab('Allow', type), lab('Data', op), lab('Minimum / Value / Source', v1), lab('Maximum', v2), hint, blank.row, lab('Error alert', al), lab('Error message', msg)]), buttons: [{ label: 'Clear All', onClick: () => { wb.clearDV(sh, q); ctx.after(); } }, { label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => {
      if (type.value === 'any') { wb.clearDV(sh, q); ctx.after(); return; }
      const rule = { range: q, type: type.value, op: op.value, allowBlank: blank.i.checked, alert: al.value, errMsg: msg.value }; if (type.value === 'list') rule.list = v1.value; else { rule.v1 = v1.value; rule.v2 = v2.value; }
      wb.addDV(sh, rule); ctx.after();
    } }] });
  };

  /* ---------------- Conditional formatting */
  D.cfRule = function (ctx, kind) {
    const { wb, g } = ctx; const sh = ctx.sh(); const q = g.range(); const bgs = { red: ['#ffc7ce', '#9c0006'], yellow: ['#ffeb9c', '#9c5700'], green: ['#c6efce', '#006100'], blue: ['#bdd7ee', '#1f4e78'] };
    const sty = sel([['red', 'Light red fill with dark red text'], ['yellow', 'Yellow fill with dark yellow text'], ['green', 'Green fill with dark green text'], ['blue', 'Light blue fill'], ['redtext', 'Red text'], ['bold', 'Bold red border-less']], 'red');
    const style = () => { const v = sty.value; if (bgs[v]) return { bg: bgs[v][0], color: bgs[v][1] }; if (v === 'redtext') return { color: '#9c0006' }; return { b: true, color: '#c00000' }; };
    const done = rule => { rule.range = q; wb.addCF(sh, rule); ctx.after(); };
    if (kind === 'databar' || kind === 'scale' || kind === 'icons') {
      if (kind === 'databar') { const c = inp('#638ec6', 'color'); OA.dialog({ title: 'Data Bars', content: lab('Bar colour', c), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => done({ type: 'databar', color: c.value }) }] }); }
      else if (kind === 'scale') { const s = sel([['rwg', 'Red – Yellow – Green'], ['gwr', 'Green – Yellow – Red'], ['wg', 'White – Green'], ['wr', 'White – Red'], ['bwr', 'Blue – White – Red']], 'rwg'); const M = { rwg: ['#f8696b', '#ffeb84', '#63be7b'], gwr: ['#63be7b', '#ffeb84', '#f8696b'], wg: ['#ffffff', '#63be7b'], wr: ['#ffffff', '#f8696b'], bwr: ['#5a8ac6', '#ffffff', '#f8696b'] }; OA.dialog({ title: 'Color Scale', content: lab('Scale', s), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => done({ type: 'scale', colors: M[s.value] }) }] }); }
      else { const s = sel([['3arrows', '3 Arrows'], ['3traffic', '3 Traffic lights'], ['3flags', '3 Flags'], ['5arrows', '5 Arrows']], '3arrows'); OA.dialog({ title: 'Icon Sets', content: lab('Icon set', s), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => done({ type: 'icons', set: s.value }) }] }); }
      return;
    }
    const f = []; let title = 'Rule';
    if (kind === 'cell') { const op = sel([['>', 'Greater than'], ['<', 'Less than'], ['between', 'Between'], ['=', 'Equal to'], ['<>', 'Not equal to'], ['>=', '≥'], ['<=', '≤'], ['notbetween', 'Not between']], '>'), a = inp(''), b = inp(''); title = 'Highlight Cells Rules'; f.push(lab('Cells that are', op), lab('Value', a), lab('and (for between)', b)); return open(() => ({ type: 'cell', op: op.value, v1: a.value, v2: b.value })); }
    if (kind === 'text') { const op = sel([['contains', 'Contains'], ['notcontains', 'Does not contain'], ['begins', 'Begins with'], ['ends', 'Ends with']], 'contains'), a = inp(''); title = 'Text Rule'; f.push(lab('Cells that', op), lab('Text', a)); return open(() => ({ type: 'text', op: op.value, v1: a.value })); }
    if (kind === 'dup') { const op = sel([['dup', 'Duplicate'], ['unique', 'Unique']], 'dup'); title = 'Duplicate Values'; f.push(lab('Format', op)); return open(() => ({ type: 'dup', op: op.value })); }
    if (kind === 'top' || kind === 'bottom') { const n = inp(10, 'number', { min: 1 }), pc = chk('% of selected cells', false); title = kind === 'top' ? 'Top Items' : 'Bottom Items'; f.push(lab('Number of items', n), pc.row); return open(() => ({ type: kind, v1: +n.value || 10, percent: pc.i.checked })); }
    if (kind === 'avg') { const op = sel([['above', 'Above average'], ['below', 'Below average']], 'above'); title = 'Average'; f.push(lab('Format cells', op)); return open(() => ({ type: 'avg', op: op.value })); }
    if (kind === 'formula') { const a = inp('=A1>0'); title = 'Formula rule'; f.push(lab('Formula (top-left cell ke liye)', a)); return open(() => ({ type: 'formula', v1: a.value })); }
    function open(mk) { f.push(lab('Format with', sty)); OA.dialog({ title, width: 400, content: el('div', { class: 'form-grid' }, f), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { const r = mk(); r.style = style(); done(r); } }] }); }
  };
  D.manageCF = function (ctx) {
    const { wb } = ctx; const sh = ctx.sh(); const lb = el('div', { class: 'x-listbox', style: 'height:200px' });
    const draw = () => { lb.innerHTML = ''; sh.cf.forEach(r => { const d = el('div', { style: 'padding:4px 8px;display:flex;gap:8px;align-items:center' }, [el('span', { style: 'flex:1', text: r.type + (r.op ? ' ' + r.op : '') + (r.v1 !== undefined ? ' ' + r.v1 : '') + '  →  ' + XL.rangeLabel(r.range) }), el('button', { class: 'btn', type: 'button', text: 'Delete', onclick: () => { wb.tx(() => { const old = sh.cf.slice(); sh.cf = sh.cf.filter(x => x !== r); const nw = sh.cf.slice(); wb._op(() => { sh.cf = old; }, () => { sh.cf = nw; }); }); ctx.after(); draw(); } })]); lb.appendChild(d); }); if (!sh.cf.length) lb.appendChild(el('div', { style: 'padding:8px;color:#777', text: 'Is sheet par koi rule nahi hai.' })); };
    draw(); OA.dialog({ title: 'Conditional Formatting Rules Manager', width: 480, content: lb, buttons: [{ label: 'Close', primary: true }] });
  };

  /* ---------------- Chart */
  D.chart = function (ctx, type0, existing) {
    const { wb, g } = ctx; const sh = ctx.sh(); let q = g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, sh, q.r1, q.c1); const e = wb.extent(sh.name); q = { r1: q.r1, c1: q.c1, r2: Math.min(q.r2, e.r), c2: Math.min(q.c2, e.c) };
    const types = [['col', 'Column'], ['bar', 'Bar'], ['line', 'Line'], ['area', 'Area'], ['pie', 'Pie'], ['doughnut', 'Doughnut'], ['scatter', 'Scatter (XY)']].concat(XL.CHART_TYPES_EXTRA || []);
    const t = sel(types, existing ? existing.type : (type0 || 'col')), title = inp(existing ? existing.title : 'Chart Title'), ref = inp(existing ? existing.ref : XL.sheetRef(sh, q)), rs = chk('Series rows mein hain (Switch Row/Column)', existing ? existing.rowsAsSeries : false), lg = chk('Legend dikhayein', existing ? existing.legend !== false : true), st = chk('Stacked', existing ? !!existing.stacked : false);
    OA.dialog({ title: existing ? 'Change Chart' : 'Insert Chart', width: 420, content: el('div', { class: 'form-grid' }, [lab('Chart type', t), lab('Title', title), lab('Data range', ref), rs.row, lg.row, st.row]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => {
      const patch = { type: t.value, title: title.value, ref: ref.value.replace(/^=/, ''), rowsAsSeries: rs.i.checked, legend: lg.i.checked, stacked: st.i.checked };
      if (existing) { g.updateChart(existing, patch); return; }
      const L = g.layout(); const ch = Object.assign({ id: XL.uid(), x: Math.round(L.cx[Math.min(L.vc - 1, q.c2 + 1)] + 24), y: Math.round(L.ry[q.r1]), w: 480, h: 290 }, patch); g.addChart(ch); ctx.after();
    } }] });
  };

  /* ---------------- Text to columns, remove duplicates, link */
  D.textToColumns = function (ctx) {
    const { wb, g } = ctx; const sh = ctx.sh(); const q = g.range(); if (q.c1 !== q.c2) { OA.toast('Ek hi column select karein'); return; }
    const d = sel([[',', 'Comma'], ['\t', 'Tab'], [';', 'Semicolon'], [' ', 'Space'], ['|', 'Pipe'], ['other', 'Other…']], ','), o = inp('', 'text', { placeholder: 'Custom delimiter' });
    OA.dialog({ title: 'Convert Text to Columns', content: el('div', { class: 'form-grid' }, [lab('Delimiter', d), lab('Other', o)]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => {
      const sep = d.value === 'other' ? o.value : d.value; if (!sep) return false; const e = wb.extent(sh.name); const r2 = Math.min(q.r2, e.r);
      wb.tx(() => { for (let r = q.r1; r <= r2; r++) { const t = wb.rawText(sh, r, q.c1); if (!t) continue; t.split(sep).forEach((p, i) => wb.putInput(sh, r, q.c1 + i, p.trim())); } }); ctx.after();
    } }] });
  };
  D.removeDuplicates = function (ctx) {
    const { wb, g } = ctx; const sh = ctx.sh(); let q = g.range(); if (q.r1 === q.r2 && q.c1 === q.c2) q = XL.region(wb, sh, q.r1, q.c1); const e = wb.extent(sh.name); q = { r1: q.r1, c1: q.c1, r2: Math.min(q.r2, e.r), c2: Math.min(q.c2, e.c) };
    const hdr = chk('Meri data mein headers hain', true); const lb = el('div', { class: 'x-listbox' }); const cks = [];
    for (let c = q.c1; c <= q.c2; c++) { const k = chk('Column ' + XL.colName(c), true); cks.push([c, k.i]); lb.appendChild(k.row); }
    OA.dialog({ title: 'Remove Duplicates', content: el('div', {}, [hdr.row, lb]), buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { const n = wb.removeDuplicates(sh, q, cks.filter(x => x[1].checked).map(x => x[0]), hdr.i.checked); OA.alert(n + ' duplicate row(s) hata diye gaye.'); ctx.after(); } }] });
  };
  D.hyperlink = async function (ctx) {
    const { wb, g } = ctx; const sh = ctx.sh(); const a = g.active(); const cur = sh.links.get(key(a.r, a.c)) || '';
    const v = await OA.form('Insert Hyperlink', [{ id: 'text', label: 'Text to display', value: wb.rawText(sh, a.r, a.c) }, { id: 'url', label: 'Address', value: cur || 'https://' }]);
    if (v && v.url) { wb.tx(() => { if (v.text) wb.putInput(sh, a.r, a.c, v.text); const old = sh.links.get(key(a.r, a.c)); sh.links.set(key(a.r, a.c), v.url); wb._op(() => { old === undefined ? sh.links.delete(key(a.r, a.c)) : sh.links.set(key(a.r, a.c), old); }, () => sh.links.set(key(a.r, a.c), v.url)); wb.applyStyle(sh, { r1: a.r, c1: a.c, r2: a.r, c2: a.c }, { color: '#0563c1', u: true }); }); ctx.after(); }
  };
  D.note = async function (ctx) {
    const { wb, g } = ctx; const sh = ctx.sh(); const a = g.active(); const k = key(a.r, a.c); const v = await OA.prompt('Note / comment:', sh.notes.get(k) || '', { title: 'Note — ' + XL.addr(a.r, a.c) });
    if (v !== null && v !== undefined) { const old = sh.notes.get(k); wb.tx(() => { v === '' ? sh.notes.delete(k) : sh.notes.set(k, v); wb._op(() => { old === undefined ? sh.notes.delete(k) : sh.notes.set(k, old); }, () => { v === '' ? sh.notes.delete(k) : sh.notes.set(k, v); }); }); ctx.after(); }
  };

  /* ---------------- Auto-filter menu */
  D.filterMenu = function (ctx, col, anchor) {
    const { wb } = ctx; const sh = ctx.sh(); const f = sh.filter; if (!f) return; const vals = wb.filterValues(sh, col); const cur = f.crit[col]; const picked = new Set(cur && cur.set ? cur.set : vals.map(v => v[0]));
    const box = el('div', { style: 'padding:8px;min-width:230px' }); const mk = (t, fn) => el('button', { class: 'mi', type: 'button', onclick: () => { OA.closePopup(); fn(); } }, [el('span', { class: 'mlb', text: t })]);
    const range = f.range; const sortBy = asc => () => { wb.sortRange(sh, { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2 }, [{ col, asc }], true); ctx.after(); };
    box.append(mk('Sort A to Z', sortBy(true)), mk('Sort Z to A', sortBy(false)), mk('Clear filter from this column', () => { wb.tx(() => { delete f.crit[col]; wb.applyFilter(sh); }); ctx.after(); }));
    const txt = inp(cur && cur.text || '', 'text', { placeholder: 'Condition: >10, =abc, *text*' }); const s = inp('', 'search', { placeholder: 'Search' }); const lb = el('div', { class: 'x-listbox', style: 'max-height:190px;margin:6px 0' }); const all = chk('(Select All)', picked.size === vals.length);
    const items = vals.map(([v, n]) => { const c = chk((v === '' ? '(Blanks)' : v) + '  (' + n + ')', picked.has(v)); c.v = v; return c; });
    const draw = () => { lb.innerHTML = ''; lb.appendChild(all.row); items.forEach(c => { if (!s.value || c.row.textContent.toLowerCase().includes(s.value.toLowerCase())) lb.appendChild(c.row); }); }; s.oninput = draw; all.i.onchange = () => items.forEach(c => c.i.checked = all.i.checked); draw();
    const ok = el('button', { class: 'btn primary', type: 'button', text: 'OK', onclick: () => { const set = items.filter(c => c.i.checked).map(c => c.v); wb.tx(() => { const cr = {}; if (set.length < items.length) cr.set = set; if (txt.value.trim()) cr.text = txt.value.trim(); if (cr.set || cr.text) f.crit[col] = cr; else delete f.crit[col]; wb.applyFilter(sh); }); OA.closePopup(); ctx.after(); } });
    box.append(el('hr'), txt, s, lb, ok); box.addEventListener('mousedown', e => { if (!/INPUT/.test(e.target.tagName)) e.preventDefault(); }); OA.popup(anchor, box);
  };

  /* ---------------- Page setup / print */
  D.pageSetup = async function (ctx) {
    const p = ctx.page; const v = await OA.form('Page Setup', [{ id: 'orient', label: 'Orientation', type: 'select', options: [['portrait', 'Portrait'], ['landscape', 'Landscape']], value: p.orient }, { id: 'grid', label: 'Print gridlines', type: 'checkbox', value: p.grid }, { id: 'title', label: 'Header text', value: p.title }]);
    if (v) Object.assign(p, v);
  };
})(window);
