/* Excel workbook model: cells, styles, formulas, undo/redo, sort, fill, CF, DV — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL;
  const KC = XL.MAXC;
  const key = (r, c) => r * KC + c;
  const unkey = k => ({ r: Math.floor(k / KC), c: k % KC });
  XL.key = key; XL.unkey = unkey;
  const E = XL.ERR;
  const isErr = XL.isErr, isMat = XL.isMat;
  const clone = o => JSON.parse(JSON.stringify(o));

  XL.Sheet = function (name) {
    this.name = name; this.cells = new Map(); this.colW = {}; this.rowH = {}; this.hiddenC = new Set(); this.hiddenR = new Set(); this.filterHidden = new Set();
    this.merges = []; this.freeze = { r: 0, c: 0 }; this.cf = []; this.dv = []; this.charts = []; this.filter = null; this.showGrid = true; this.tabColor = null;
    this.spill = new Map(); this.notes = new Map(); this.links = new Map(); this.colStyle = {}; this.rowStyle = {}; this.defaultW = 64; this.defaultH = 20; this._ext = null; this.zoom = 1;
  };

  /* ---------------------------------------------------------------- style helpers */
  function mergeStyle(base, patch) {
    const o = Object.assign({}, base);
    for (const k in patch) { if (patch[k] === null || patch[k] === undefined || patch[k] === false && k !== 'wrap' && !['b', 'i', 'u', 'st'].includes(k)) delete o[k]; else if (patch[k] === false) delete o[k]; else o[k] = patch[k]; }
    return o;
  }

  XL.Workbook = function () {
    this.sheets = []; this.names = {}; this.styles = [{}]; this._styleMap = { '{}': 0 }; this.pass = 0; this.needCalc = true; this.hist = { undo: [], redo: [] }; this._tx = null; this.onChange = null; this.circular = false; this.cfCache = new Map(); this.version = 0;
    this.addSheet('Sheet1');
  };
  const W = XL.Workbook.prototype;

  /* ---------------------------------------------------------------- sheets */
  W.sheet = function (n) { if (typeof n === 'number') return this.sheets[n]; const l = String(n).toLowerCase(); return this.sheets.find(s => s.name.toLowerCase() === l) || null; };
  W.hasSheet = function (n) { return !!this.sheet(n); };
  W.uniqueName = function (base) { let n = base, i = 2; while (this.sheet(n)) n = base.replace(/\d+$/, '') + (i++); return n; };
  W.addSheet = function (name, at) {
    const sh = new XL.Sheet(name || this.uniqueName('Sheet' + (this.sheets.length + 1)));
    const idx = at == null ? this.sheets.length : at;
    const doIt = () => { this.sheets.splice(idx, 0, sh); this.dirty(); }, undo = () => { this.sheets.splice(this.sheets.indexOf(sh), 1); this.dirty(); };
    doIt(); this._op(undo, doIt); return sh;
  };
  W.deleteSheet = function (sh) {
    if (this.sheets.length < 2) return false; const idx = this.sheets.indexOf(sh);
    const doIt = () => { this.sheets.splice(this.sheets.indexOf(sh), 1); this.dirty(); }, undo = () => { this.sheets.splice(idx, 0, sh); this.dirty(); };
    doIt(); this._op(undo, doIt); return true;
  };
  W.renameSheet = function (sh, nn) {
    nn = nn.trim(); if (!nn || /[\\/?*[\]:]/.test(nn) || nn.length > 31) return 'Sheet names cannot contain \\ / ? * [ ] : and must be 1–31 characters.';
    const o = this.sheet(nn); if (o && o !== sh) return 'That sheet name is already in use.';
    const old = sh.name; if (old === nn) return null;
    const apply = (from, to) => { sh.name = to; this.sheets.forEach(s => s.cells.forEach((c, k) => { if (c.f != null && c.ast) { const na = XL.renameSheetAst(c.ast, from, to); if (na !== c.ast) { c.ast = na; c.f = XL.astToStr(na); } } })); Object.values(this.names).forEach(n => { n.ref = n.ref.replace(new RegExp("(^|=)'?" + from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'?!", 'i'), '$1' + XL.quoteSheet(to) + '!'); n.ast = undefined; }); this.sheets.forEach(s => s.charts.forEach(ch => { ch.ref = ch.ref.replace(new RegExp("^'?" + from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'?!", 'i'), XL.quoteSheet(to) + '!'); })); this.dirty(); };
    apply(old, nn); this._op(() => apply(nn, old), () => apply(old, nn)); return null;
  };
  W.moveSheet = function (sh, to) { const from = this.sheets.indexOf(sh); if (from === to || to < 0 || to >= this.sheets.length) return; const mv = (a, b) => { const s = this.sheets.splice(a, 1)[0]; this.sheets.splice(b, 0, s); this.dirty(); }; mv(from, to); this._op(() => mv(to, from), () => mv(from, to)); };
  W.duplicateSheet = function (sh) {
    const n = this.addSheet(this.uniqueName(sh.name + ' (2)'), this.sheets.indexOf(sh) + 1);
    sh.cells.forEach((c, k) => n.cells.set(k, Object.assign({}, c))); ['colW', 'rowH', 'colStyle', 'rowStyle'].forEach(p => n[p] = Object.assign({}, sh[p])); n.merges = clone(sh.merges); n.cf = clone(sh.cf); n.dv = clone(sh.dv); n.freeze = Object.assign({}, sh.freeze); n.showGrid = sh.showGrid; n.hiddenC = new Set(sh.hiddenC); n.hiddenR = new Set(sh.hiddenR);
    n.charts = clone(sh.charts).map(c => { c.id = XL.uid(); return c; }); sh.notes.forEach((v, k) => n.notes.set(k, v)); sh.links.forEach((v, k) => n.links.set(k, v)); n._ext = null; this.dirty(); return n;
  };
  XL.uid = () => Math.random().toString(36).slice(2, 9);

  /* ---------------------------------------------------------------- transactions */
  W.dirty = function () { this.needCalc = true; this.version++; this.cfCache.clear(); this.sheets.forEach(s => { s._ext = null; s._mergeMap = null; }); };
  W._op = function (undo, redo) { if (this._tx) this._tx.ops.push({ undo, redo }); else { this.hist.undo.push({ ops: [{ undo, redo }] }); this.hist.redo.length = 0; this.onChange && this.onChange(); } };
  W.tx = function (fn) {
    if (this._tx) return fn();
    const t = this._tx = { ops: [] }; let r;
    try { r = fn(); } finally {
      this._tx = null;
      if (t.ops.length) { this.hist.undo.push(t); this.hist.redo.length = 0; if (this.hist.undo.length > 200) this.hist.undo.shift(); this.dirty(); this.onChange && this.onChange(); }
    }
    return r;
  };
  W.undo = function () { const t = this.hist.undo.pop(); if (!t) return false; for (let i = t.ops.length - 1; i >= 0; i--) t.ops[i].undo(); this.hist.redo.push(t); this.dirty(); this.onChange && this.onChange(true); return true; };
  W.redo = function () { const t = this.hist.redo.pop(); if (!t) return false; t.ops.forEach(o => o.redo()); this.hist.undo.push(t); this.dirty(); this.onChange && this.onChange(true); return true; };
  W.setProp = function (obj, prop, v) { const old = obj[prop]; if (old === v) return; const ap = x => { if (x === undefined) delete obj[prop]; else obj[prop] = x; this.dirty(); }; ap(v); this._op(() => ap(old), () => ap(v)); };
  W.setCell = function (sh, k, cell) {
    const old = sh.cells.get(k); if (old === cell) return;
    const ap = x => { if (x) sh.cells.set(k, x); else sh.cells.delete(k); sh._ext = null; };
    ap(cell); this._op(() => ap(old), () => ap(cell)); this.needCalc = true;
  };
  // snapshot-based op for structural changes
  W._snapAll = function () { return { cells: this.sheets.map(s => ({ sh: s, cells: new Map(s.cells), colW: Object.assign({}, s.colW), rowH: Object.assign({}, s.rowH), colStyle: Object.assign({}, s.colStyle), rowStyle: Object.assign({}, s.rowStyle), merges: clone(s.merges), cf: clone(s.cf), dv: clone(s.dv), freeze: Object.assign({}, s.freeze), hiddenC: new Set(s.hiddenC), hiddenR: new Set(s.hiddenR), notes: new Map(s.notes), links: new Map(s.links), charts: clone(s.charts), filter: s.filter ? clone(s.filter) : null })), names: clone(this.names) }; };
  W._restoreAll = function (sn) { sn.cells.forEach(x => { const s = x.sh; s.cells = new Map(x.cells); s.colW = Object.assign({}, x.colW); s.rowH = Object.assign({}, x.rowH); s.colStyle = Object.assign({}, x.colStyle); s.rowStyle = Object.assign({}, x.rowStyle); s.merges = clone(x.merges); s.cf = clone(x.cf); s.dv = clone(x.dv); s.freeze = Object.assign({}, x.freeze); s.hiddenC = new Set(x.hiddenC); s.hiddenR = new Set(x.hiddenR); s.notes = new Map(x.notes); s.links = new Map(x.links); s.charts = clone(x.charts); s.filter = x.filter ? clone(x.filter) : null; s._ext = null; }); this.names = clone(sn.names); this.dirty(); };
  W._structural = function (fn) { const before = this._snapAll(); fn(); const after = this._snapAll(); this._op(() => this._restoreAll(before), () => this._restoreAll(after)); this.dirty(); };

  /* ---------------------------------------------------------------- styles */
  W.styleIdx = function (o) { const k = JSON.stringify(o, Object.keys(o).sort()); let i = this._styleMap[k]; if (i === undefined) { i = this.styles.length; this.styles.push(clone(o)); this._styleMap[k] = i; } return i; };
  W.styleOf = function (sh, r, c) { const cell = sh.cells.get(key(r, c)); if (cell) return this.styles[cell.s || 0]; const i = sh.rowStyle[r] !== undefined ? sh.rowStyle[r] : sh.colStyle[c]; return i ? this.styles[i] : this.styles[0]; };
  W.applyStyle = function (sh, g, patch) {
    this.tx(() => {
      const one = (r, c) => { const k = key(r, c); const old = sh.cells.get(k); const base = this.styleOf(sh, r, c); const ns = this.styleIdx(mergeStyle(base, patch)); if (old) { if ((old.s || 0) !== ns) this.setCell(sh, k, Object.assign({}, old, { s: ns })); } else if (ns) this.setCell(sh, k, { s: ns }); };
      if (g.fullCol) { for (let c = g.c1; c <= g.c2; c++) { const ns = this.styleIdx(mergeStyle(this.styles[sh.colStyle[c] || 0], patch)); this.setProp(sh.colStyle, c, ns || undefined); sh.cells.forEach((cell, k) => { if (k % KC === c) { const nn = this.styleIdx(mergeStyle(this.styles[cell.s || 0], patch)); if (nn !== (cell.s || 0)) this.setCell(sh, k, Object.assign({}, cell, { s: nn })); } }); } }
      else if (g.fullRow) { for (let r = g.r1; r <= g.r2; r++) { const ns = this.styleIdx(mergeStyle(this.styles[sh.rowStyle[r] || 0], patch)); this.setProp(sh.rowStyle, r, ns || undefined); sh.cells.forEach((cell, k) => { if (Math.floor(k / KC) === r) { const nn = this.styleIdx(mergeStyle(this.styles[cell.s || 0], patch)); if (nn !== (cell.s || 0)) this.setCell(sh, k, Object.assign({}, cell, { s: nn })); } }); } }
      else for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) one(r, c);
    });
  };
  W.applyBorders = function (sh, g, mode, style, color) {
    const b = style ? { s: style, c: color || '#000000' } : null;
    this.tx(() => {
      const set = (r, c, side, v) => { const st = mergeStyle(this.styleOf(sh, r, c), { [side]: v }); const k = key(r, c); const old = sh.cells.get(k); const ns = this.styleIdx(st); if (old) { if ((old.s || 0) !== ns) this.setCell(sh, k, Object.assign({}, old, { s: ns })); } else if (ns) this.setCell(sh, k, { s: ns }); };
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const top = r === g.r1, bot = r === g.r2, lef = c === g.c1, rig = c === g.c2;
        if (mode === 'none') { ['bT', 'bB', 'bL', 'bR'].forEach(s => set(r, c, s, null)); continue; }
        if (mode === 'all') { ['bT', 'bB', 'bL', 'bR'].forEach(s => set(r, c, s, b)); continue; }
        if (mode === 'outer' || mode === 'thickOuter') { if (top) set(r, c, 'bT', b); if (bot) set(r, c, 'bB', b); if (lef) set(r, c, 'bL', b); if (rig) set(r, c, 'bR', b); }
        if (mode === 'inner') { if (!top) set(r, c, 'bT', b); if (!bot) set(r, c, 'bB', b); if (!lef) set(r, c, 'bL', b); if (!rig) set(r, c, 'bR', b); }
        if (mode === 'top' && top) set(r, c, 'bT', b); if (mode === 'bottom' && bot) set(r, c, 'bB', b); if (mode === 'left' && lef) set(r, c, 'bL', b); if (mode === 'right' && rig) set(r, c, 'bR', b);
        if (mode === 'topbottom') { if (top) set(r, c, 'bT', b); if (bot) set(r, c, 'bB', b); }
      }
    });
  };

  /* ---------------------------------------------------------------- reading */
  W.cell = (sh, r, c) => sh.cells.get(key(r, c));
  W.extent = function (n) {
    const sh = this.sheet(n); if (!sh) return { r: 0, c: 0 }; if (sh._ext) return sh._ext;
    let mr = 0, mc = 0; sh.cells.forEach((cell, k) => { if (cell.f != null || (cell.v !== undefined && cell.v !== null && cell.v !== '')) { const r = Math.floor(k / KC), c = k % KC; if (r > mr) mr = r; if (c > mc) mc = c; } });
    sh.spill.forEach((v, k) => { const r = Math.floor(k / KC), c = k % KC; if (r > mr) mr = r; if (c > mc) mc = c; });
    return sh._ext = { r: mr, c: mc };
  };
  W.usedRange = function (sh) { let r1 = 1e9, c1 = 1e9, r2 = -1, c2 = -1; const f = (k, cell) => { const r = Math.floor(k / KC), c = k % KC; if (r < r1) r1 = r; if (c < c1) c1 = c; if (r > r2) r2 = r; if (c > c2) c2 = c; }; sh.cells.forEach((cell, k) => { if (cell.f != null || (cell.v !== undefined && cell.v !== null && cell.v !== '') || cell.s) f(k, cell); }); sh.spill.forEach((v, k) => f(k)); return r2 < 0 ? null : { r1, c1, r2, c2 }; };
  W.value = function (shn, r, c, ctx) {
    const sh = typeof shn === 'string' ? this.sheet(shn) : shn; if (!sh) return E('#REF!');
    const k = key(r, c); const cell = sh.cells.get(k);
    if (cell) { if (cell.f != null) return this._calc(sh, cell, r, c); return cell.v === undefined ? null : cell.v; }
    const sp = sh.spill.get(k); return sp !== undefined ? sp : null;
  };
  W._calc = function (sh, cell, r, c) {
    if (cell.pass === this.pass) return cell.val;
    if (cell.busy) { this.circular = true; return 0; }
    cell.busy = true; let v;
    try { v = XL.evalFormula(cell.ast, { wb: this, sheet: sh.name, r, c }); } catch (e) { v = E('#VALUE!'); console.warn(e); }
    cell.busy = false; cell.pass = this.pass;
    // clear previous spill
    if (cell.spillKeys) { cell.spillKeys.forEach(k => sh.spill.delete(k)); cell.spillKeys = null; }
    if (isMat(v)) {
      const m = v.mat; const R = m.length, C = m[0] ? m[0].length : 0;
      if (R * C > 1) {
        let ok = true; for (let i = 0; i < R && ok; i++) for (let j = 0; j < C && ok; j++) { if (!i && !j) continue; const kk = key(r + i, c + j); const ex = sh.cells.get(kk); if (r + i >= XL.MAXR || c + j >= XL.MAXC || (ex && (ex.f != null || (ex.v !== undefined && ex.v !== null && ex.v !== ''))) || (sh.spill.has(kk))) ok = false; }
        if (!ok) v = E('#SPILL!'); else { cell.spillKeys = []; for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { if (!i && !j) continue; const kk = key(r + i, c + j); sh.spill.set(kk, m[i][j]); cell.spillKeys.push(kk); } cell.spillRect = { r1: r, c1: c, r2: r + R - 1, c2: c + C - 1 }; this._spilled = true; v = m[0][0]; }
      } else v = m[0][0];
      if (v === undefined) v = null;
    }
    if (!cell.spillKeys) cell.spillRect = null;
    cell.val = v; return v;
  };
  W.recalc = function () {
    this.circular = false;
    for (let round = 0; round < 3; round++) {
      this.pass++; this._spilled = false;
      this.sheets.forEach(sh => sh._ext = null);
      this.sheets.forEach(sh => sh.cells.forEach((cell, k) => { if (cell.f != null) this._calc(sh, cell, Math.floor(k / KC), k % KC); }));
      if (!this._spilled) break; this.sheets.forEach(sh => sh._ext = null);
    }
    this.needCalc = false;
  };
  W.ensureCalc = function () { if (this.needCalc) this.recalc(); };
  W.getValue = function (sh, r, c) { this.ensureCalc(); return this.value(sh, r, c); };
  W.rawText = function (sh, r, c) {
    const cell = sh.cells.get(key(r, c)); if (!cell) return '';
    if (cell.f != null) return '=' + cell.f;
    const v = cell.v; if (v === undefined || v === null) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (isErr(v)) return v.err;
    if (typeof v === 'number') { const nf = this.styles[cell.s || 0].nf; if (nf && XL.isDateFormat(nf)) { const t = XL.formatValue(v, /h|s/.test(nf.replace(/mmm/g, '')) && v % 1 ? 'm/d/yyyy h:mm:ss' : 'm/d/yyyy').text; if (t && !t.includes('#')) return t; } if (nf && /%/.test(nf)) return XL.general(v * 100) + '%'; return XL.general(+v.toPrecision(15)); }
    return String(v);
  };
  W.display = function (sh, r, c) {
    this.ensureCalc(); const k = key(r, c); const cell = sh.cells.get(k);
    let v = cell ? (cell.f != null ? this._calc(sh, cell, r, c) : cell.v) : sh.spill.get(k);
    if (v === undefined) v = null;
    const st = this.styles[cell ? (cell.s || 0) : ((sh.rowStyle[r] !== undefined ? sh.rowStyle[r] : sh.colStyle[c]) || 0)];
    const nf = st.nf || 'General';
    if (typeof v === 'string' && cell && cell.f == null && st.nf && !/@/.test(st.nf) && !/;/.test(st.nf)) return { text: v, v, num: false, st };
    const f = XL.formatValue(v, nf);
    return { text: f.text, color: f.color, v, num: typeof v === 'number', err: !!f.err, bool: typeof v === 'boolean', st };
  };

  /* ---------------------------------------------------------------- input */
  XL.parseInput = function (text) {
    if (text === null || text === undefined) return { kind: 'empty' };
    if (typeof text !== 'string') return { kind: 'v', v: text };
    if (text === '') return { kind: 'empty' };
    if (text[0] === '=' && text.length > 1) return { kind: 'f', src: text.slice(1) };
    if (text[0] === "'") return { kind: 'v', v: text.slice(1), forceText: true };
    const t = text.trim(); const u = t.toUpperCase();
    if (u === 'TRUE') return { kind: 'v', v: true }; if (u === 'FALSE') return { kind: 'v', v: false };
    const p = XL.parseNumber(t); if (p) return { kind: 'v', v: p.v, nf: p.nf };
    const d = XL.parseDate(t); if (d) return { kind: 'v', v: d.v, nf: d.nf };
    if (/^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/.test(u)) return { kind: 'v', v: E(u) };
    return { kind: 'v', v: text };
  };
  const AUTO_FMT = { DATE: 'm/d/yyyy', TODAY: 'm/d/yyyy', EDATE: 'm/d/yyyy', EOMONTH: 'm/d/yyyy', WORKDAY: 'm/d/yyyy', NOW: 'm/d/yyyy h:mm', TIME: 'h:mm AM/PM' };
  W.putInput = function (sh, r, c, text, opts) {
    opts = opts || {};
    const pi = XL.parseInput(text); const k = key(r, c); const old = sh.cells.get(k); const s = old ? (old.s || 0) : ((sh.rowStyle[r] !== undefined ? sh.rowStyle[r] : sh.colStyle[c]) || 0);
    let cell = null, ns = s;
    if (pi.kind === 'f') {
      let ast; try { ast = XL.parse(pi.src); } catch (e) { return { error: e.message }; }
      cell = { f: XL.astToStr(ast), ast, s: s };
      if (AUTO_FMT[ast.t === 'fn' ? ast.name : ''] && !this.styles[s].nf) ns = this.styleIdx(mergeStyle(this.styles[s], { nf: AUTO_FMT[ast.name] }));
    } else if (pi.kind === 'v') {
      cell = { v: pi.v, s: s };
      if (pi.nf && !this.styles[s].nf && !pi.forceText) ns = this.styleIdx(mergeStyle(this.styles[s], { nf: pi.nf }));
      if (pi.forceText && this.styles[s].nf && !/@/.test(this.styles[s].nf)) { }
    }
    if (cell) { cell.s = ns; if (!cell.s) delete cell.s; }
    if (!cell && old && old.s) cell = { s: old.s };
    if (!cell && !old && !s) return {};
    this.setCell(sh, k, cell || (s ? { s } : null));
    if (!cell && !s) this.setCell(sh, k, null);
    return {};
  };
  W.setInput = function (sh, r, c, text) { let res; this.tx(() => { res = this.putInput(sh, r, c, text); }); return res; };
  W.clearRange = function (sh, g, what) { // what: 'contents' | 'formats' | 'all'
    this.tx(() => {
      const ks = []; sh.cells.forEach((cell, k) => { const r = Math.floor(k / KC), c = k % KC; if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) ks.push(k); });
      ks.forEach(k => { const cell = sh.cells.get(k); let nc;
        if (what === 'all') nc = null; else if (what === 'formats') { nc = Object.assign({}, cell); delete nc.s; if (nc.f == null && nc.v === undefined) nc = null; } else { nc = cell.s ? { s: cell.s } : null; }
        this.setCell(sh, k, nc); });
      if (what !== 'formats') { for (const [kk] of Array.from(sh.notes)) { const u = unkey(kk); if (what === 'all' && u.r >= g.r1 && u.r <= g.r2 && u.c >= g.c1 && u.c <= g.c2) { const t = sh.notes.get(kk); sh.notes.delete(kk); this._op(() => sh.notes.set(kk, t), () => sh.notes.delete(kk)); } } }
    });
  };

  /* ---------------------------------------------------------------- merges */
  W.mergeMap = function (sh) { if (sh._mergeMap) return sh._mergeMap; const m = new Map(); sh.merges.forEach(g => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) m.set(key(r, c), g); }); return sh._mergeMap = m; };
  W.mergeAt = function (sh, r, c) { if (!sh.merges.length) return null; return this.mergeMap(sh).get(key(r, c)) || null; };
  W.merge = function (sh, g) {
    if (g.r1 === g.r2 && g.c1 === g.c2) return;
    this.tx(() => {
      const keep = sh.cells.get(key(g.r1, g.c1)); const merges = sh.merges.filter(m => !(m.r1 >= g.r1 && m.r2 <= g.r2 && m.c1 >= g.c1 && m.c2 <= g.c2));
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== g.r1 || c !== g.c1) { const k = key(r, c); const x = sh.cells.get(k); if (x) this.setCell(sh, k, x.s ? { s: x.s } : null); }
      const old = sh.merges.slice(); sh.merges = merges.concat([{ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }]); const nw = sh.merges.slice(); sh._mergeMap = null;
      this._op(() => { sh.merges = old; sh._mergeMap = null; }, () => { sh.merges = nw; sh._mergeMap = null; });
    });
  };
  W.unmerge = function (sh, g) {
    this.tx(() => { const old = sh.merges.slice(); sh.merges = sh.merges.filter(m => !(m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1)); const nw = sh.merges.slice(); sh._mergeMap = null; this._op(() => { sh.merges = old; sh._mergeMap = null; }, () => { sh.merges = nw; sh._mergeMap = null; }); });
  };
  W.expandMerges = function (sh, g) { // grow rect to include whole merges
    let ch = true; g = Object.assign({}, g); while (ch) { ch = false; sh.merges.forEach(m => { if (m.r1 <= g.r2 && m.r2 >= g.r1 && m.c1 <= g.c2 && m.c2 >= g.c1) { if (m.r1 < g.r1) { g.r1 = m.r1; ch = true; } if (m.c1 < g.c1) { g.c1 = m.c1; ch = true; } if (m.r2 > g.r2) { g.r2 = m.r2; ch = true; } if (m.c2 > g.c2) { g.c2 = m.c2; ch = true; } } }); } return g;
  };

  /* ---------------------------------------------------------------- insert / delete rows & columns */
  const shiftRange = (g, isR, at, cnt) => {
    const a = isR ? 'r1' : 'c1', b = isR ? 'r2' : 'c2'; g = Object.assign({}, g);
    if (cnt > 0) { if (g[a] >= at) g[a] += cnt; if (g[b] >= at) g[b] += cnt; return g; }
    const k = -cnt; if (g[a] >= at && g[b] < at + k) return null;
    g[a] = g[a] < at ? g[a] : (g[a] >= at + k ? g[a] - k : at); g[b] = g[b] < at ? g[b] : (g[b] >= at + k ? g[b] - k : at - 1); return g;
  };
  W._adjustAll = function (target, kind, at, count) {
    this.sheets.forEach(sh => sh.cells.forEach((cell, k) => { if (cell.f != null) { const na = XL.adjustAst(cell.ast, target.name, sh.name, kind, at, count); if (na !== cell.ast) { const nc = Object.assign({}, cell, { ast: na, f: XL.astToStr(na) }); sh.cells.set(k, nc); } } }));
    Object.values(this.names).forEach(n => { try { const a = XL.adjustAst(XL.parse(n.ref.replace(/^=/, '')), target.name, target.name, kind, at, count); n.ref = XL.astToStr(a); n.ast = undefined; } catch (e) { } });
  };
  W.insertDelete = function (sh, kind, at, count) { // kind 'row'|'col'; count>0 insert, <0 delete
    const isR = kind === 'row';
    this._structural(() => {
      this._adjustAll(sh, kind, at, count);
      const nm = new Map(); const k0 = Math.abs(count);
      sh.cells.forEach((cell, k) => { const r = Math.floor(k / KC), c = k % KC; const idx = isR ? r : c; let ni = idx;
        if (count > 0) { if (idx >= at) ni = idx + count; } else { if (idx >= at && idx < at + k0) return; if (idx >= at + k0) ni = idx - k0; }
        nm.set(isR ? key(ni, c) : key(r, ni), cell); });
      sh.cells = nm;
      const mapSparse = (obj) => { const o = {}; Object.keys(obj).forEach(i => { i = +i; let ni = i; if (count > 0) { if (i >= at) ni = i + count; } else { if (i >= at && i < at + k0) return; if (i >= at + k0) ni = i - k0; } o[ni] = obj[i]; }); return o; };
      const mapSet = s => { const o = new Set(); s.forEach(i => { let ni = i; if (count > 0) { if (i >= at) ni = i + count; } else { if (i >= at && i < at + k0) return; if (i >= at + k0) ni = i - k0; } o.add(ni); }); return o; };
      if (isR) { sh.rowH = mapSparse(sh.rowH); sh.rowStyle = mapSparse(sh.rowStyle); sh.hiddenR = mapSet(sh.hiddenR); } else { sh.colW = mapSparse(sh.colW); sh.colStyle = mapSparse(sh.colStyle); sh.hiddenC = mapSet(sh.hiddenC); }
      sh.merges = sh.merges.map(g => shiftRange(g, isR, at, count)).filter(Boolean);
      sh.cf.forEach(rule => { rule.range = shiftRange(rule.range, isR, at, count); }); sh.cf = sh.cf.filter(r => r.range);
      sh.dv.forEach(rule => { rule.range = shiftRange(rule.range, isR, at, count); }); sh.dv = sh.dv.filter(r => r.range);
      const mk = new Map(); const mv = m => { const o = new Map(); m.forEach((v, k) => { const r = Math.floor(k / KC), c = k % KC; const idx = isR ? r : c; let ni = idx; if (count > 0) { if (idx >= at) ni = idx + count; } else { if (idx >= at && idx < at + k0) return; if (idx >= at + k0) ni = idx - k0; } o.set(isR ? key(ni, c) : key(r, ni), v); }); return o; };
      sh.notes = mv(sh.notes); sh.links = mv(sh.links);
      if (sh.filter) { const f = shiftRange(sh.filter.range, isR, at, count); if (f) sh.filter.range = f; else sh.filter = null; }
      if (isR && sh.freeze.r > at && count > 0) sh.freeze.r += count;
      sh.spill.clear(); sh._mergeMap = null;
    });
  };

  /* ---------------------------------------------------------------- clipboard */
  W.copyRange = function (sh, g, cut) {
    this.ensureCalc(); const rows = [];
    for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const cell = sh.cells.get(key(r, c)); const sp = sh.spill.get(key(r, c)); row.push(cell ? Object.assign({}, cell, { val: cell.f != null ? this.value(sh, r, c) : undefined }) : (sp !== undefined ? { v: sp, spilled: true } : (sh.rowStyle[r] !== undefined || sh.colStyle[c] ? { s: sh.rowStyle[r] !== undefined ? sh.rowStyle[r] : sh.colStyle[c] } : null))); } rows.push(row); }
    return { sheet: sh.name, g: Object.assign({}, g), rows, cut: !!cut, text: this.rangeText(sh, g), merges: sh.merges.filter(m => m.r1 >= g.r1 && m.r2 <= g.r2 && m.c1 >= g.c1 && m.c2 <= g.c2).map(m => ({ r1: m.r1 - g.r1, c1: m.c1 - g.c1, r2: m.r2 - g.r1, c2: m.c2 - g.c1 })) };
  };
  W.rangeText = function (sh, g) { const out = []; this.ensureCalc(); for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 100000); r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const d = this.display(sh, r, c); row.push(d.text.replace(/[\t\r\n]/g, ' ')); } out.push(row.join('\t')); } return out.join('\n'); };
  W.rangeHTML = function (sh, g) {
    this.ensureCalc(); let h = '<table border="1" style="border-collapse:collapse">';
    for (let r = g.r1; r <= g.r2; r++) { h += '<tr>'; for (let c = g.c1; c <= g.c2; c++) { const d = this.display(sh, r, c); const st = d.st; let css = ''; if (st.b) css += 'font-weight:bold;'; if (st.i) css += 'font-style:italic;'; if (st.u) css += 'text-decoration:underline;'; if (st.color) css += 'color:' + st.color + ';'; if (st.bg) css += 'background:' + st.bg + ';'; if (st.ha) css += 'text-align:' + st.ha + ';'; else if (d.num) css += 'text-align:right;'; h += `<td style="${css}">${XL.esc(d.text)}</td>`; } h += '</tr>'; }
    return h + '</table>';
  };
  XL.esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  W.pasteClip = function (sh, r0, c0, clip, mode, opts) {
    opts = opts || {}; mode = mode || 'all'; const tr = !!opts.transpose; const R = clip.rows.length, C = clip.rows[0].length;
    const outR = tr ? C : R, outC = tr ? R : C; const src = this.sheet(clip.sheet) || sh;
    const dr = r0 - clip.g.r1, dc = c0 - clip.g.c1;
    this.tx(() => {
      if (clip.cut && mode === 'all' && !tr) {
        // move: remove sources first (collect), then place
      }
      const placed = [];
      for (let i = 0; i < outR; i++) for (let j = 0; j < outC; j++) {
        const rr = r0 + i, cc = c0 + j; if (rr >= XL.MAXR || cc >= XL.MAXC) continue;
        const s = tr ? clip.rows[j][i] : clip.rows[i][j]; const k = key(rr, cc); const old = sh.cells.get(k);
        if (opts.skipBlanks && (!s || (s.v === undefined && s.f == null))) continue;
        let nc = null;
        if (mode === 'formats') { if (s && s.s) nc = Object.assign({}, old || {}, { s: s.s }); else if (old) { nc = Object.assign({}, old); delete nc.s; } if (nc && nc.f == null && nc.v === undefined && !nc.s) nc = null; }
        else {
          const keepStyle = mode === 'values' || mode === 'formulas' || mode === 'valuesnf';
          let base = null;
          if (s) {
            if (s.f != null && mode !== 'values' && mode !== 'valuesnf') {
              let ast = s.ast; if (!clip.cut) ast = XL.shiftAst(ast, tr ? 0 : dr, tr ? 0 : dc); if (tr && !clip.cut) { ast = XL.shiftAst(s.ast, rr - (clip.g.r1 + j), cc - (clip.g.c1 + i)); }
              base = { f: XL.astToStr(ast), ast };
            } else { const v = s.f != null ? s.val : s.v; if (v !== undefined && v !== null || s.f != null) base = { v: v === undefined ? null : (isMat(v) ? v.mat[0][0] : v) }; }
          }
          let st = keepStyle ? (old ? old.s : 0) : (s ? s.s : 0); if (mode === 'valuesnf' && s && s.s) { const nf = this.styles[s.s].nf; if (nf) st = this.styleIdx(mergeStyle(this.styles[st || 0], { nf })); }
          if (base || st) { nc = Object.assign({}, base || {}); if (st) nc.s = st; if (base && base.v === null && !st) nc = null; } else nc = null;
          if (!base && old && keepStyle && old.s) nc = { s: old.s };
        }
        this.setCell(sh, k, nc);
      }
      if (mode === 'all' || mode === 'formats') { clip.merges.forEach(m => { if (!tr) { this.merge(sh, { r1: r0 + m.r1, c1: c0 + m.c1, r2: r0 + m.r2, c2: c0 + m.c2 }); } }); }
      if (clip.cut && mode === 'all' && !tr) {
        // clear sources not overlapped by destination, and retarget references
        const g = clip.g; const inDest = (r, c) => r >= r0 && r < r0 + R && c >= c0 && c < c0 + C && src === sh;
        for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (!inDest(r, c)) { const k = key(r, c); if (src.cells.has(k)) this.setCell(src, k, null); }
        this.sheets.forEach(s2 => s2.cells.forEach((cell, k) => { if (cell.f != null) { const na = XL.mapAst(cell.ast, n => { if ((n.t === 'ref') && (n.sheet || s2.name) === src.name && n.r >= g.r1 && n.r <= g.r2 && n.c >= g.c1 && n.c <= g.c2) return Object.assign({}, n, { r: n.r + dr, c: n.c + dc, sheet: n.sheet && n.sheet !== sh.name ? n.sheet : (src === sh ? n.sheet : sh.name) }); if (n.t === 'range' && (n.sheet || s2.name) === src.name && n.r1 >= g.r1 && n.r2 <= g.r2 && n.c1 >= g.c1 && n.c2 <= g.c2) return Object.assign({}, n, { r1: n.r1 + dr, r2: n.r2 + dr, c1: n.c1 + dc, c2: n.c2 + dc }); }); if (na !== cell.ast && XL.astToStr(na) !== cell.f) this.setCell(s2, k, Object.assign({}, cell, { ast: na, f: XL.astToStr(na) })); } }));
      }
    });
  };

  /* ---------------------------------------------------------------- fill */
  const SERIES = [['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']];
  function seriesNext(vals, n) { // vals: array of source raw values; returns array of n next values or null (copy)
    const nums = vals.filter(v => typeof v === 'number');
    if (vals.length && nums.length === vals.length) {
      if (vals.length === 1) return null;
      const d = vals[vals.length - 1] - vals[vals.length - 2]; let linear = true; for (let i = 1; i < vals.length; i++) if (Math.abs(vals[i] - vals[i - 1] - d) > 1e-9) linear = false;
      if (!linear) { const n2 = vals.length; const mx = (n2 - 1) / 2, my = vals.reduce((a, b) => a + b, 0) / n2; let sxy = 0, sxx = 0; vals.forEach((y, i) => { sxy += (i - mx) * (y - my); sxx += (i - mx) ** 2; }); const sl = sxy / sxx, ic = my - sl * mx; return Array.from({ length: n }, (_, i) => ic + sl * (n2 + i)); }
      return Array.from({ length: n }, (_, i) => vals[vals.length - 1] + d * (i + 1));
    }
    if (vals.length && vals.every(v => typeof v === 'string')) {
      for (const lst of SERIES) { const idx = vals.map(v => lst.findIndex(x => x.toLowerCase() === v.toLowerCase())); if (idx.every(i => i >= 0)) { const step = vals.length > 1 ? ((idx[1] - idx[0] + lst.length) % lst.length || 1) : 1; const last = idx[idx.length - 1]; return Array.from({ length: n }, (_, i) => { const w = lst[(last + step * (i + 1)) % lst.length]; const src = vals[0]; return src === src.toUpperCase() ? w.toUpperCase() : src === src.toLowerCase() ? w.toLowerCase() : w; }); } }
      const m = vals.map(v => /^(.*?)(\d+)$/.exec(v)); if (m.every(Boolean) && m.every(x => x[1] === m[0][1])) { const ns = m.map(x => +x[2]); const step = ns.length > 1 ? ns[ns.length - 1] - ns[ns.length - 2] : 1; const w = m[0][2].length; const pad = m[0][2][0] === '0'; return Array.from({ length: n }, (_, i) => m[0][1] + String(ns[ns.length - 1] + step * (i + 1)).padStart(pad ? w : 0, '0')); }
    }
    return null;
  }
  W.autoFill = function (sh, src, dest, opts) { // src: g, dest: g covering src+extension
    opts = opts || {}; this.ensureCalc();
    const down = dest.r2 > src.r2, up = dest.r1 < src.r1, right = dest.c2 > src.c2, left = dest.c1 < src.c1;
    this.tx(() => {
      const fillLine = (cells, positions, step) => { // cells: source cell objects array (in order), positions: [{r,c}] to fill
        const raws = cells.map(c => c && c.f == null ? c.v : undefined); const canSeries = !opts.copy && raws.every(v => v !== undefined && v !== null) && cells.every(c => c && c.f == null);
        let ser = canSeries ? seriesNext(raws, positions.length) : null;
        // single date/number with fill series option
        if (!ser && opts.series && raws.length === 1 && typeof raws[0] === 'number') ser = positions.map((_, i) => raws[0] + (opts.step || 1) * (i + 1));
        positions.forEach((p, i) => {
          const s = cells[i % cells.length]; const k = key(p.r, p.c); const old = sh.cells.get(k); let nc;
          if (ser) nc = { v: ser[i], s: (cells[cells.length - 1] || {}).s }; else if (!s) nc = null;
          else if (s.f != null) { const sr = p.srcR[i % cells.length], sc = p.srcC[i % cells.length]; const ast = XL.shiftAst(s.ast, p.r - sr, p.c - sc); nc = { f: XL.astToStr(ast), ast, s: s.s }; } else nc = Object.assign({}, s);
          if (nc && !nc.s) delete nc.s; if (nc && nc.v === undefined && nc.f == null && !nc.s) nc = null; this.setCell(sh, k, nc);
        });
      };
      if (down || up) for (let c = src.c1; c <= src.c2; c++) {
        const cells = []; for (let r = src.r1; r <= src.r2; r++) cells.push(sh.cells.get(key(r, c)) || null);
        if (down) { const pos = []; for (let r = src.r2 + 1; r <= dest.r2; r++) pos.push({ r, c, srcR: cells.map((_, i) => src.r1 + i), srcC: cells.map(() => c) }); fillLine(cells, pos); }
        else { const rev = cells.slice().reverse(); const pos = []; for (let r = src.r1 - 1; r >= dest.r1; r--) pos.push({ r, c, srcR: rev.map((_, i) => src.r2 - i), srcC: rev.map(() => c) }); fillLine(rev, pos); }
      } else if (right || left) for (let r = src.r1; r <= src.r2; r++) {
        const cells = []; for (let c = src.c1; c <= src.c2; c++) cells.push(sh.cells.get(key(r, c)) || null);
        if (right) { const pos = []; for (let c = src.c2 + 1; c <= dest.c2; c++) pos.push({ r, c, srcR: cells.map(() => r), srcC: cells.map((_, i) => src.c1 + i) }); fillLine(cells, pos); }
        else { const rev = cells.slice().reverse(); const pos = []; for (let c = src.c1 - 1; c >= dest.c1; c--) pos.push({ r, c, srcR: rev.map(() => r), srcC: rev.map((_, i) => src.c2 - i) }); fillLine(rev, pos); }
      }
    });
  };

  /* ---------------------------------------------------------------- sort / dedupe / find */
  W.sortRange = function (sh, g, keys, hasHeader) { // keys: [{col (absolute), asc}]
    this.ensureCalc(); const r1 = g.r1 + (hasHeader ? 1 : 0); if (r1 >= g.r2) return;
    this.tx(() => {
      const rows = []; for (let r = r1; r <= g.r2; r++) { const cells = []; for (let c = g.c1; c <= g.c2; c++) cells.push(sh.cells.get(key(r, c)) || null); rows.push({ r, cells, vals: keys.map(k => this.value(sh, r, k.col)) }); }
      const emptyRank = v => v === null || v === '' ? 1 : 0;
      rows.sort((a, b) => { for (let i = 0; i < keys.length; i++) { const x = a.vals[i], y = b.vals[i]; const ex = emptyRank(x), ey = emptyRank(y); if (ex || ey) { if (ex !== ey) return ex - ey; continue; } const c = XL.cmpVals(x, y) * (keys[i].asc ? 1 : -1); if (c) return c; } return a.r - b.r; });
      rows.forEach((row, i) => { const r = r1 + i; row.cells.forEach((cell, j) => { const c = g.c1 + j; let nc = cell; if (cell && cell.f != null) { const ast = XL.shiftAst(cell.ast, r - row.r, 0); nc = Object.assign({}, cell, { ast, f: XL.astToStr(ast) }); } this.setCell(sh, key(r, c), nc); }); });
    });
  };
  W.removeDuplicates = function (sh, g, cols, hasHeader) {
    this.ensureCalc(); const seen = new Set(); const del = []; const r1 = g.r1 + (hasHeader ? 1 : 0);
    for (let r = r1; r <= g.r2; r++) { const k = JSON.stringify(cols.map(c => { const v = this.value(sh, r, c); return typeof v === 'string' ? v.toLowerCase() : v; })); if (seen.has(k)) del.push(r); else seen.add(k); }
    if (!del.length) return 0;
    this.tx(() => { const keep = []; for (let r = r1; r <= g.r2; r++) if (!del.includes(r)) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(sh.cells.get(key(r, c)) || null); keep.push(row); }
      for (let r = r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) { const i = r - r1; this.setCell(sh, key(r, c), i < keep.length ? keep[i][c - g.c1] : null); } });
    return del.length;
  };
  W.findAll = function (q, o) {
    o = o || {}; this.ensureCalc(); const out = []; const ql = o.matchCase ? q : q.toLowerCase();
    const sheets = o.allSheets ? this.sheets : [o.sheet]; const re = o.wild !== false && /[*?]/.test(q) ? new RegExp((o.whole ? '^' : '') + q.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + (o.whole ? '$' : ''), o.matchCase ? '' : 'i') : null;
    sheets.forEach(sh => { const ks = Array.from(sh.cells.keys()).sort((a, b) => a - b); const byRow = o.byCol ? ks.sort((a, b) => (a % KC) - (b % KC) || a - b) : ks; byRow.forEach(k => { const { r, c } = unkey(k); if (o.range && !(r >= o.range.r1 && r <= o.range.r2 && c >= o.range.c1 && c <= o.range.c2)) return; const cell = sh.cells.get(k); let t = o.formulas ? this.rawText(sh, r, c) : this.display(sh, r, c).text; if (cell.f == null && cell.v === undefined) return; let hit; if (re) hit = re.test(t); else { const tt = o.matchCase ? t : t.toLowerCase(); hit = o.whole ? tt === ql : tt.includes(ql); } if (hit) out.push({ sheet: sh, r, c }); }); });
    return out;
  };
  W.replaceAt = function (sh, r, c, q, rep, o) {
    o = o || {}; const cell = sh.cells.get(key(r, c)); if (!cell) return; const raw = this.rawText(sh, r, c); const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\\\*/g, '.*').replace(/\\\?/g, '.'), 'g' + (o.matchCase ? '' : 'i')); const nt = o.whole ? rep : raw.replace(re, rep);
    this.setInput(sh, r, c, nt);
  };

  /* ---------------------------------------------------------------- conditional formatting */
  W.cfFor = function (sh, r, c) { const out = []; sh.cf.forEach(rule => { const g = rule.range; if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) out.push(rule); }); return out; };
  W._cfStats = function (sh, rule) {
    const ck = sh.name + ':' + rule.id; let st = this.cfCache.get(ck); if (st) return st;
    const g = rule.range, vals = [], all = []; const ext = this.extent(sh.name);
    for (let r = g.r1; r <= Math.min(g.r2, ext.r); r++) for (let c = g.c1; c <= Math.min(g.c2, ext.c); c++) { const v = this.value(sh, r, c); all.push(v); if (typeof v === 'number') vals.push(v); }
    vals.sort((a, b) => a - b); const n = vals.length; const sum = vals.reduce((a, b) => a + b, 0); const avg = n ? sum / n : 0;
    const counts = new Map(); all.forEach(v => { if (v !== null && v !== '') { const k = typeof v === 'string' ? v.toLowerCase() : v; counts.set(k, (counts.get(k) || 0) + 1); } });
    let sd = 0; if (n > 1) sd = Math.sqrt(vals.reduce((a, b) => a + (b - avg) ** 2, 0) / (n - 1));
    st = { min: vals[0], max: vals[n - 1], n, avg, sd, vals, counts }; this.cfCache.set(ck, st); return st;
  };
  W.evalCF = function (sh, r, c) {
    const rules = this.cfFor(sh, r, c); if (!rules.length) return null; const v = this.value(sh, r, c); const res = {}; let any = false;
    for (const rule of rules) {
      const o = rule; let hit = false;
      try {
        const num = typeof v === 'number', str = typeof v === 'string';
        switch (o.type) {
          case 'cell': { const a = XL.parseInput(String(o.v1)).v, b = o.v2 !== undefined ? XL.parseInput(String(o.v2)).v : null; const x = v === null ? 0 : v; const cmp = XL.cmpVals;
            hit = { '>': () => cmp(x, a) > 0, '<': () => cmp(x, a) < 0, '>=': () => cmp(x, a) >= 0, '<=': () => cmp(x, a) <= 0, '=': () => cmp(x, a) === 0, '<>': () => cmp(x, a) !== 0, between: () => cmp(x, a) >= 0 && cmp(x, b) <= 0, notbetween: () => cmp(x, a) < 0 || cmp(x, b) > 0 }[o.op || '>'](); break; }
          case 'text': { const t = (v === null ? '' : String(v)).toLowerCase(), q = String(o.v1).toLowerCase(); hit = o.op === 'begins' ? t.startsWith(q) : o.op === 'ends' ? t.endsWith(q) : o.op === 'notcontains' ? !t.includes(q) : t.includes(q); break; }
          case 'dup': { const s = this._cfStats(sh, o); const k = typeof v === 'string' ? v.toLowerCase() : v; hit = v !== null && v !== '' && (s.counts.get(k) > 1) === (o.op !== 'unique'); break; }
          case 'top': case 'bottom': { if (!num) break; const s = this._cfStats(sh, o); const nn = o.percent ? Math.max(1, Math.round(s.n * (o.v1 || 10) / 100)) : (o.v1 || 10); const th = o.type === 'top' ? s.vals[Math.max(0, s.n - nn)] : s.vals[Math.min(s.n - 1, nn - 1)]; hit = o.type === 'top' ? v >= th : v <= th; break; }
          case 'avg': { if (!num) break; const s = this._cfStats(sh, o); hit = o.op === 'below' ? v < s.avg : v > s.avg; break; }
          case 'blanks': hit = v === null || v === ''; break; case 'errors': hit = isErr(v); break;
          case 'formula': { const ast = o.ast || (o.ast = XL.parse(String(o.v1).replace(/^=/, ''))); const sa = XL.shiftAst(ast, r - o.range.r1, c - o.range.c1); const x = XL.evalFormula(sa, { wb: this, sheet: sh.name, r, c }); hit = !isErr(x) && (typeof x === 'number' ? x !== 0 : x === true); break; }
          case 'databar': { if (!num) break; const s = this._cfStats(sh, o); const lo = Math.min(0, s.min), hi = Math.max(0, s.max); res.bar = { pct: hi === lo ? 1 : (v - lo) / (hi - lo), color: o.color || '#638ec6', neg: v < 0 }; any = true; continue; }
          case 'scale': { if (!num) break; const s = this._cfStats(sh, o); const t = s.max === s.min ? 0.5 : (v - s.min) / (s.max - s.min); const cs = o.colors || ['#f8696b', '#ffeb84', '#63be7b']; res.bg = scaleColor(cs, t); any = true; continue; }
          case 'icons': { if (!num) break; const s = this._cfStats(sh, o); const t = s.max === s.min ? 0.5 : (v - s.min) / (s.max - s.min); const n3 = o.set === '5arrows' ? 5 : 3; const idx = Math.min(n3 - 1, Math.floor(t * n3 + 1e-9)); res.icon = { set: o.set || '3arrows', idx, n: n3 }; any = true; continue; }
        }
      } catch (e) { hit = false; }
      if (hit) { Object.assign(res, o.style || {}); any = true; if (o.stop) break; }
    }
    return any ? res : null;
  };
  function scaleColor(cs, t) {
    const p = h => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)]; const n = cs.length - 1; const i = Math.min(n - 1, Math.floor(t * n)); const f = t * n - i; const a = p(cs[i]), b = p(cs[i + 1]);
    return '#' + a.map((x, k) => Math.round(x + (b[k] - x) * f).toString(16).padStart(2, '0')).join('');
  }
  W.addCF = function (sh, rule) { rule.id = rule.id || XL.uid(); this.tx(() => { const old = sh.cf.slice(); sh.cf = sh.cf.concat([rule]); const nw = sh.cf.slice(); this._op(() => { sh.cf = old; }, () => { sh.cf = nw; }); }); };
  W.clearCF = function (sh, g) { this.tx(() => { const old = sh.cf.slice(); sh.cf = g ? sh.cf.filter(r => !(r.range.r1 <= g.r2 && r.range.r2 >= g.r1 && r.range.c1 <= g.c2 && r.range.c2 >= g.c1)) : []; const nw = sh.cf.slice(); this._op(() => { sh.cf = old; }, () => { sh.cf = nw; }); }); };

  /* ---------------------------------------------------------------- data validation */
  W.dvAt = function (sh, r, c) { for (let i = sh.dv.length - 1; i >= 0; i--) { const g = sh.dv[i].range; if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) return sh.dv[i]; } return null; };
  W.dvList = function (sh, rule) {
    if (Array.isArray(rule.list)) return rule.list.map(String);
    const src = String(rule.list || ''); if (src[0] === '=') { try { const ast = XL.parse(src.slice(1)); const rf = XL.evalRef(ast, { wb: this, sheet: sh.name, r: 0, c: 0 }); if (rf && rf.ref) { const out = []; const ext = this.extent(rf.sheet); for (let r = rf.r1; r <= Math.min(rf.r2, ext.r); r++) for (let c = rf.c1; c <= Math.min(rf.c2, ext.c); c++) { const v = this.value(rf.sheet, r, c); if (v !== null && v !== '') out.push(XL.formatValue(v, 'General').text); } return out; } } catch (e) { } return []; }
    return src.split(',').map(s => s.trim()).filter(s => s !== '');
  };
  W.dvCheck = function (sh, r, c, input) {
    const rule = this.dvAt(sh, r, c); if (!rule) return null; if (input === '' || input === null) return rule.allowBlank === false ? rule : null;
    const pi = XL.parseInput(input); const v = pi.kind === 'f' ? null : pi.v; let ok = true;
    const num = typeof v === 'number';
    const n1 = rule.v1 !== undefined && rule.v1 !== '' ? (isNaN(+rule.v1) ? (XL.parseDate(String(rule.v1)) || {}).v : +rule.v1) : null, n2 = rule.v2 !== undefined && rule.v2 !== '' ? (isNaN(+rule.v2) ? (XL.parseDate(String(rule.v2)) || {}).v : +rule.v2) : null;
    const cmp = (x) => ({ between: x >= n1 && x <= n2, notbetween: x < n1 || x > n2, '=': x === n1, '<>': x !== n1, '>': x > n1, '<': x < n1, '>=': x >= n1, '<=': x <= n1 }[rule.op || 'between']);
    switch (rule.type) {
      case 'list': ok = this.dvList(sh, rule).some(x => x.toLowerCase() === String(input).trim().toLowerCase()); break;
      case 'whole': ok = num && Number.isInteger(v) && cmp(v); break; case 'decimal': ok = num && cmp(v); break; case 'date': ok = num && cmp(v); break;
      case 'textLength': ok = cmp(String(input).length); break;
      case 'custom': try { const ast = XL.parse(String(rule.v1).replace(/^=/, '')); const x = XL.evalFormula(ast, { wb: this, sheet: sh.name, r, c }); ok = x === true || (typeof x === 'number' && x !== 0); } catch (e) { ok = true; } break;
    }
    return ok ? null : rule;
  };
  W.addDV = function (sh, rule) { this.tx(() => { const old = sh.dv.slice(); sh.dv = sh.dv.filter(x => !(x.range.r1 === rule.range.r1 && x.range.c1 === rule.range.c1 && x.range.r2 === rule.range.r2 && x.range.c2 === rule.range.c2)).concat(rule ? [rule] : []); const nw = sh.dv.slice(); this._op(() => { sh.dv = old; }, () => { sh.dv = nw; }); }); };
  W.clearDV = function (sh, g) { this.tx(() => { const old = sh.dv.slice(); sh.dv = sh.dv.filter(r => !(r.range.r1 <= g.r2 && r.range.r2 >= g.r1 && r.range.c1 <= g.c2 && r.range.c2 >= g.c1)); const nw = sh.dv.slice(); this._op(() => { sh.dv = old; }, () => { sh.dv = nw; }); }); };

  /* ---------------------------------------------------------------- filter */
  W.setFilter = function (sh, g) { this.tx(() => { const old = sh.filter ? clone(sh.filter) : null; const oh = new Set(sh.filterHidden); sh.filter = g ? { range: g, crit: {} } : null; sh.filterHidden = new Set(); const nw = sh.filter ? clone(sh.filter) : null; this._op(() => { sh.filter = old; sh.filterHidden = oh; }, () => { sh.filter = nw; sh.filterHidden = new Set(); }); }); };
  W.filterValues = function (sh, col) { this.ensureCalc(); const f = sh.filter; const vals = new Map(); for (let r = f.range.r1 + 1; r <= f.range.r2; r++) { const d = this.display(sh, r, col); const k = d.text; vals.set(k, (vals.get(k) || 0) + 1); } return Array.from(vals.entries()).sort((a, b) => XL.cmpVals(a[0], b[0])); };
  W.applyFilter = function (sh) {
    const f = sh.filter; if (!f) { sh.filterHidden = new Set(); this.dirty(); return; } this.ensureCalc(); const hid = new Set();
    for (let r = f.range.r1 + 1; r <= f.range.r2; r++) for (const c in f.crit) { const cr = f.crit[c]; const d = this.display(sh, r, +c); let ok = true; if (cr.set) ok = cr.set.includes(d.text); if (ok && cr.text) { const fn = XL.critFn(cr.text); ok = fn(d.v); } if (!ok) { hid.add(r); break; } }
    const old = sh.filterHidden; sh.filterHidden = hid; this._op(() => { sh.filterHidden = old; }, () => { sh.filterHidden = hid; }); this.dirty();
  };

  /* ---------------------------------------------------------------- names */
  W.defineName = function (name, ref, sheetName) { const l = name.toLowerCase(); const old = this.names[l]; const nw = { name, ref: ref.replace(/^=/, '') }; this.names[l] = nw; this._op(() => { if (old) this.names[l] = old; else delete this.names[l]; this.dirty(); }, () => { this.names[l] = nw; this.dirty(); }); this.dirty(); };
  W.deleteName = function (name) { const l = name.toLowerCase(); const old = this.names[l]; if (!old) return; delete this.names[l]; this._op(() => { this.names[l] = old; this.dirty(); }, () => { delete this.names[l]; this.dirty(); }); this.dirty(); };
  W.resolveName = function (s) { const n = this.names[s.toLowerCase()]; if (!n) return null; try { const ast = XL.parse(n.ref); if (ast.t === 'ref') return { sheet: ast.sheet, r1: ast.r, c1: ast.c, r2: ast.r, c2: ast.c }; if (ast.t === 'range') return { sheet: ast.sheet, r1: ast.r1, c1: ast.c1, r2: ast.r2, c2: ast.c2 }; } catch (e) { } return null; };

  /* ---------------------------------------------------------------- sizes */
  W.colWidth = (sh, c) => sh.hiddenC.has(c) ? 0 : (sh.colW[c] !== undefined ? sh.colW[c] : sh.defaultW);
  W.rowHeight = (sh, r) => (sh.hiddenR.has(r) || sh.filterHidden.has(r)) ? 0 : (sh.rowH[r] !== undefined ? sh.rowH[r] : sh.defaultH);

  /* ---------------------------------------------------------------- CSV */
  XL.parseCSV = function (text, delim) {
    text = text.replace(/^﻿/, ''); if (!delim) { const first = text.split('\n')[0] || ''; const cnt = d => (first.match(new RegExp(d === '|' ? '\\|' : d, 'g')) || []).length; delim = [',', '\t', ';', '|'].sort((a, b) => cnt(b) - cnt(a))[0]; }
    const rows = []; let row = [], cur = '', q = false;
    for (let i = 0; i < text.length; i++) { const ch = text[i]; if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; } else if (ch === '"' && cur === '') q = true; else if (ch === delim) { row.push(cur); cur = ''; } else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i + 1] === '\n') i++; row.push(cur); cur = ''; rows.push(row); row = []; } else cur += ch; }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); } return rows;
  };
  XL.toCSV = function (wb, sh, g, delim) {
    delim = delim || ','; wb.ensureCalc(); const out = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) { const d = wb.display(sh, r, c); let t = d.text; if (t.includes(delim) || /["\n\r]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"'; row.push(t); } out.push(row.join(delim)); } return out.join('\r\n');
  };
  W.loadRows = function (sh, rows, r0, c0) { r0 = r0 || 0; c0 = c0 || 0; this.tx(() => { rows.forEach((row, i) => row.forEach((v, j) => { if (v !== '') this.putInput(sh, r0 + i, c0 + j, v); })); }); };
})(window);
