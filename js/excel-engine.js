/* Excel formula engine + number formats — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL = global.XL || {};

  /* ---------------------------------------------------------------- addresses */
  XL.MAXR = 1048576; XL.MAXC = 16384;
  XL.colName = i => { let s = ''; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  XL.colIdx = s => { let n = 0; s = s.toUpperCase(); for (let i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n - 1; };
  XL.addr = (r, c, ar, ac) => (ac ? '$' : '') + XL.colName(c) + (ar ? '$' : '') + (r + 1);
  XL.parseAddr = function (s) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(String(s).trim()); if (!m) return null;
    const c = XL.colIdx(m[2]), r = +m[4] - 1; if (r < 0 || r >= XL.MAXR || c >= XL.MAXC) return null;
    return { r, c, ar: !!m[3], ac: !!m[1] };
  };
  // returns {r1,c1,r2,c2} normalised; supports A1, A1:B2, A:A, 1:1
  XL.parseRange = function (s) {
    s = String(s).trim().replace(/\$/g, ''); let m;
    if ((m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(s))) {
      const c1 = XL.colIdx(m[1]), r1 = +m[2] - 1, c2 = m[3] ? XL.colIdx(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
      return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
    }
    if ((m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(s))) { const a = XL.colIdx(m[1]), b = XL.colIdx(m[2]); return { r1: 0, r2: XL.MAXR - 1, c1: Math.min(a, b), c2: Math.max(a, b), fullCol: true }; }
    if ((m = /^(\d+):(\d+)$/.exec(s))) { const a = +m[1] - 1, b = +m[2] - 1; return { r1: Math.min(a, b), r2: Math.max(a, b), c1: 0, c2: XL.MAXC - 1, fullRow: true }; }
    return null;
  };
  XL.rangeStr = g => {
    if (g.fullCol) return XL.colName(g.c1) + ':' + XL.colName(g.c2);
    if (g.fullRow) return (g.r1 + 1) + ':' + (g.r2 + 1);
    return g.r1 === g.r2 && g.c1 === g.c2 ? XL.addr(g.r1, g.c1) : XL.addr(g.r1, g.c1) + ':' + XL.addr(g.r2, g.c2);
  };
  const quoteSheet = n => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'";
  XL.quoteSheet = quoteSheet;

  /* ---------------------------------------------------------------- dates */
  const D0 = Date.UTC(1899, 11, 30);
  XL.dateToSerial = d => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds()) - D0) / 864e5;
  XL.ymdToSerial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - D0) / 864e5);
  XL.serialToDate = s => new Date(D0 + Math.round(s * 864e5 / 1000) * 1000);   // UTC fields
  XL.serialParts = s => { const d = XL.serialToDate(s); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), h: d.getUTCHours(), mi: d.getUTCMinutes(), s: d.getUTCSeconds(), dow: d.getUTCDay() }; };
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  XL.parseNumber = function (s) {
    s = String(s).trim(); if (!s) return null;
    let m;
    if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s)) return { v: parseFloat(s) };
    if ((m = /^\(?\s*([-+]?)\s*([$€£¥₹]?)\s*([-+]?)(\d{1,3}(?:,\d{2,3})*(?:\.\d+)?|\d+(?:\.\d+)?|\.\d+)\s*(%?)\)?$/.exec(s)) && (m[2] || m[5] || /,/.test(m[4]) || /^\(/.test(s))) {
      let v = parseFloat(m[4].replace(/,/g, '')); if (m[1] === '-' || m[3] === '-' || /^\(.*\)$/.test(s)) v = -v;
      if (m[5]) return { v: v / 100, nf: m[4].includes('.') ? '0.00%' : '0%' };
      if (m[2]) return { v, nf: m[2] === '$' ? '$#,##0.00' : '"' + m[2] + '"#,##0.00' };
      return { v, nf: '#,##0' + (m[4].includes('.') ? '.' + '0'.repeat(m[4].split('.')[1].length) : '') };
    }
    if ((m = /^([-+]?\d+(?:\.\d+)?)\s*%$/.exec(s))) return { v: parseFloat(m[1]) / 100, nf: m[1].includes('.') ? '0.00%' : '0%' };
    return null;
  };
  XL.parseDate = function (s) {
    s = String(s).trim(); let m;
    const mon = n => { const i = MONTHS.findIndex(x => x.toLowerCase().startsWith(n.toLowerCase().slice(0, 3))); return i < 0 ? 0 : i + 1; };
    const ok = (y, mo, d) => { if (mo < 1 || mo > 12 || d < 1 || d > 31) return null; const dt = new Date(Date.UTC(y, mo - 1, d)); return dt.getUTCMonth() === mo - 1 ? XL.ymdToSerial(y, mo, d) : null; };
    const yy = y => y < 100 ? (y < 30 ? 2000 + y : 1900 + y) : y;
    let v = null, nf = null;
    if ((m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s))) { v = ok(+m[1], +m[2], +m[3]); nf = 'yyyy-mm-dd'; }
    else if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/.exec(s))) { let a = +m[1], b = +m[2]; if (a > 12 && b <= 12) v = ok(yy(+m[3]), b, a); else v = ok(yy(+m[3]), a, b); nf = 'm/d/yyyy'; }
    else if ((m = /^(\d{1,2})[- ]([A-Za-z]{3,9})[- ,]+(\d{2,4})$/.exec(s))) { v = ok(yy(+m[3]), mon(m[2]), +m[1]); nf = 'd-mmm-yy'; }
    else if ((m = /^([A-Za-z]{3,9})\.? (\d{1,2}),? (\d{2,4})$/.exec(s))) { v = ok(yy(+m[3]), mon(m[1]), +m[2]); nf = 'mmm d, yyyy'; }
    if (v != null) return { v, nf };
    if ((m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i.exec(s))) {
      let h = +m[1]; const mi = +m[2], se = +(m[3] || 0); if (m[4]) { const pm = /pm/i.test(m[4]); if (h === 12) h = pm ? 12 : 0; else if (pm) h += 12; }
      if (h < 24 && mi < 60 && se < 60) return { v: (h * 3600 + mi * 60 + se) / 86400, nf: m[4] ? 'h:mm AM/PM' : (m[3] ? 'h:mm:ss' : 'h:mm') };
    }
    return null;
  };

  /* ---------------------------------------------------------------- number formats */
  XL.BUILTIN_FMT = { 0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@' };
  function general(n) {
    if (!isFinite(n)) return '#NUM!';
    if (n === 0) return '0';
    const a = Math.abs(n);
    if (a >= 1e11 || a < 1e-9) { let s = n.toExponential(5).replace(/\.?0+e/, 'e'); return s.replace(/e([+-])(\d)$/, 'E$10$2').replace(/e([+-])/, 'E$1'); }
    let s = parseFloat(n.toPrecision(10)).toString(); if (s.includes('e')) s = n.toFixed(10).replace(/\.?0+$/, ''); return s;
  }
  XL.general = general;
  function rnd(x, d) { const f = Math.pow(10, d); return Math.round((x + (x >= 0 ? 1 : -1) * 1e-12 * Math.max(1, Math.abs(x))) * f) / f; }
  function splitSections(f) { const out = []; let cur = '', q = false, br = false; for (let i = 0; i < f.length; i++) { const c = f[i]; if (c === '"') q = !q; else if (!q && c === '[') br = true; else if (!q && c === ']') br = false; if (c === ';' && !q && !br) { out.push(cur); cur = ''; } else cur += c; } out.push(cur); return out; }
  const COLORS = { red: '#c00000', blue: '#0000ff', green: '#008000', black: '#000', white: '#fff', yellow: '#b8a000', magenta: '#c000c0', cyan: '#00a0a0' };
  const fmtCache = {};
  function isDateFmt(sec) { const s = sec.replace(/"[^"]*"|\\.|\[[^\]]*\]|_.|\*./g, ''); return /[ymdhs]/i.test(s.replace(/[eE][+-]/g, '')) && !/^[#0?,.%\s$€£¥]*$/.test(s) && !/[#0?]/.test(s.replace(/[ymdhsap\/:AMPM\s.,-]/gi, '')) || /^\[[hms]+\]/i.test(sec); }
  function fmtDate(sec, serial) {
    const toks = sec.match(/\[(?:h+|m+|s+)\]|"[^"]*"|\\.|_.|\*.|\[[^\]]*\]|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|\.0+|./gi) || [];
    const hasAP = toks.some(t => /^(AM\/PM|A\/P)$/i.test(t));
    let frac = serial - Math.floor(serial), total = Math.round(frac * 86400 * 1000) / 1000;
    let secs = Math.floor(total + 1e-9); let ms = total - secs;
    if (secs >= 86400) { serial = Math.floor(serial) + 1; secs -= 86400; }
    const p = XL.serialParts(Math.floor(serial)); const H = Math.floor(secs / 3600), Mi = Math.floor(secs % 3600 / 60), S = secs % 60;
    const out = [];
    toks.forEach((t, i) => {
      const lt = t.toLowerCase();
      if (/^\[(h+)\]$/i.test(t)) out.push(String(Math.floor(serial * 24 + secs / 3600 - (serial - Math.floor(serial)) * 24 * 0)).padStart(t.length - 2, '0'));
      else if (/^\[m+\]$/i.test(t)) out.push(String(Math.floor(serial * 1440)));
      else if (/^\[s+\]$/i.test(t)) out.push(String(Math.floor(serial * 86400)));
      else if (t[0] === '"') out.push(t.slice(1, -1)); else if (t[0] === '\\') out.push(t[1]); else if (t[0] === '_') out.push(' '); else if (t[0] === '*') out.push(''); else if (t[0] === '[') out.push('');
      else if (lt === 'yyyy') out.push(String(p.y).padStart(4, '0')); else if (lt === 'yy') out.push(String(p.y % 100).padStart(2, '0'));
      else if (lt === 'mmmmm') out.push(MONTHS[p.m - 1][0]); else if (lt === 'mmmm') out.push(MONTHS[p.m - 1]); else if (lt === 'mmm') out.push(MONTHS[p.m - 1].slice(0, 3));
      else if (lt === 'm' || lt === 'mm') {
        let j = i - 1, prev = ''; while (j >= 0 && !/^(h+|hh|s+|ss|yyyy|yy|d+|m+)$/i.test(toks[j])) j--; prev = j >= 0 ? toks[j].toLowerCase() : '';
        let k = i + 1, next = ''; while (k < toks.length && !/^(h+|hh|s+|ss|yyyy|yy|d+|m+)$/i.test(toks[k])) k++; next = k < toks.length ? toks[k].toLowerCase() : '';
        const isMin = /^h/.test(prev) || /^s/.test(next) || /^\[h/.test(prev);
        const val = isMin ? Mi : p.m; out.push(lt === 'mm' ? String(val).padStart(2, '0') : String(val));
      }
      else if (lt === 'dddd') out.push(DAYS[p.dow]); else if (lt === 'ddd') out.push(DAYS[p.dow].slice(0, 3)); else if (lt === 'dd') out.push(String(p.d).padStart(2, '0')); else if (lt === 'd') out.push(String(p.d));
      else if (lt === 'hh' || lt === 'h') { let h = hasAP ? (H % 12 || 12) : H; out.push(lt === 'hh' ? String(h).padStart(2, '0') : String(h)); }
      else if (lt === 'ss') out.push(String(Math.floor(S)).padStart(2, '0')); else if (lt === 's') out.push(String(Math.floor(S)));
      else if (lt === 'am/pm') out.push(H < 12 ? 'AM' : 'PM'); else if (lt === 'a/p') out.push(H < 12 ? 'A' : 'P');
      else if (/^\.0+$/.test(t)) out.push('.' + String(Math.round(ms * Math.pow(10, t.length - 1))).padStart(t.length - 1, '0'));
      else out.push(t);
    });
    return out.join('');
  }
  function fmtFraction(sec, n) {
    const m = /^(.*?)([#0?]*)\s*([#?0]+)\/([#?0]+|\d+)(.*)$/.exec(sec); if (!m) return null;
    const maxD = /^\d+$/.test(m[4]) ? +m[4] : Math.pow(10, m[4].length) - 1; const neg = n < 0; n = Math.abs(n);
    const whole = m[2] ? Math.floor(n) : 0; let f = n - whole, bn = 0, bd = 1, err = f;
    for (let d = 1; d <= maxD; d++) { const nn = Math.round(f * d); const e = Math.abs(f - nn / d); if (e < err - 1e-12) { err = e; bn = nn; bd = d; } if (e < 1e-10) break; }
    let w = whole; if (bn === bd) { w++; bn = 0; }
    const lit = s => s.replace(/"([^"]*)"|\\(.)|_./g, (x, a, b) => a !== undefined ? a : b !== undefined ? b : ' ');
    const frac = bn === 0 ? ' '.repeat(String(bd).length * 2 + 1) : bn + '/' + bd;
    return (neg ? '-' : '') + lit(m[1]) + (m[2] ? (w || bn === 0 ? w + ' ' : '') : '') + frac + lit(m[5]);
  }
  function fmtNumber(sec, n, neg1) {
    // sec: single section; n: number (sign handled by caller via neg1)
    const lit = s => s.replace(/"([^"]*)"|\\(.)|_(.)|\*(.)/g, (x, a, b, c, d) => a !== undefined ? a : b !== undefined ? b : c !== undefined ? ' ' : '');
    if (/[#0?]\s*[#0?]*\/[#0?\d]/.test(sec.replace(/"[^"]*"/g, ''))) { const r = fmtFraction(sec, neg1 ? -n : n); if (r != null) return r; }
    // scientific
    let m = /^(.*?)([#0?][#0?.,]*)E([+-])(0+)(.*)$/i.exec(sec.replace(/"[^"]*"/g, q => q.replace(/[#0?.,]/g, c => '\u0001' + c.charCodeAt(0) + '\u0002')));
    const unesc = s => s.replace(/\u0001(\d+)\u0002/g, (x, c) => String.fromCharCode(+c));
    if (m) {
      const dec = (m[2].split('.')[1] || '').length, minInt = (m[2].split('.')[0].match(/0/g) || []).length; const expDigits = m[4].length;
      let e = n === 0 ? 0 : Math.floor(Math.log10(n)); const eng = (m[2].split('.')[0].length > 1 && minInt === 0) || m[2].split('.')[0].length === 3;
      if (eng) e = Math.floor(e / 3) * 3; let mant = n / Math.pow(10, e); let ms = mant.toFixed(dec); if (parseFloat(ms) >= (eng ? 1000 : 10)) { e += eng ? 3 : 1; mant = n / Math.pow(10, e); ms = mant.toFixed(dec); }
      return (neg1 ? '-' : '') + lit(unesc(m[1])) + ms + 'E' + (e < 0 ? '-' : (m[3] === '+' ? '+' : '')) + String(Math.abs(e)).padStart(expDigits, '0') + lit(unesc(m[5]));
    }
    const masked = sec.replace(/"[^"]*"|\\.|_.|\*.|\[[^\]]*\]/g, q => '\u0003'.repeat(0) + '\u0004'.repeat(q.length));
    const first = masked.search(/[#0?]/); const lastIdx = (() => { for (let i = masked.length - 1; i >= 0; i--) if (/[#0?]/.test(masked[i])) return i; return -1; })();
    const pct = (masked.match(/%/g) || []).length; let val = n; for (let i = 0; i < pct; i++) val *= 100;
    if (first < 0) { // no digit placeholders: literal-only format
      return (neg1 ? '-' : '') + lit(sec);
    }
    // include commas/periods around placeholders
    let s0 = first, e0 = lastIdx;
    while (s0 > 0 && /[.,]/.test(masked[s0 - 1])) s0--;
    let trailingCommas = 0; while (e0 + 1 < masked.length && masked[e0 + 1] === ',') { e0++; trailingCommas++; }
    const run = masked.slice(s0, e0 + 1).replace(/,+$/, ''); const pre = sec.slice(0, s0), post = sec.slice(e0 + 1);
    val = val / Math.pow(1000, trailingCommas);
    const dot = run.indexOf('.'); const ip = dot < 0 ? run : run.slice(0, dot), fp = dot < 0 ? '' : run.slice(dot + 1);
    const dec = (fp.match(/[#0?]/g) || []).length; const thousands = /,/.test(ip);
    val = rnd(val, dec); let ds = Math.abs(val).toFixed(dec); let [is, fs = ''] = ds.split('.');
    const minInt = (ip.match(/0/g) || []).length; const qInt = (ip.match(/\?/g) || []).length;
    if (is === '0' && minInt === 0 && qInt === 0 && (ip.replace(/,/g, '') !== '')) is = '';
    if (is.length < minInt) is = is.padStart(minInt, '0'); else if (is.length < minInt + qInt) is = is.padStart(minInt + qInt, ' ');
    if (thousands) is = is.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    // frac: trim trailing optional (#) zeros
    let reqF = 0, qF = 0; for (let i = 0; i < fp.length; i++) { if (fp[i] === '0') reqF = i + 1; else if (fp[i] === '?') qF = i + 1; }
    while (fs.length > Math.max(reqF, 0) && fs.endsWith('0')) { if (fs.length <= qF) break; fs = fs.slice(0, -1); }
    if (qF > fs.length) fs = fs.padEnd(qF, ' ');
    const dotStr = dec > 0 && (fs.trim() !== '' || reqF > 0 || fp.includes('0')) ? '.' + fs : (dec > 0 && /^[#?]*$/.test(fp) && fs.trim() ? '.' + fs : '');
    const zero = parseFloat(ds) === 0; const sign = neg1 && !zero ? '-' : '';
    return lit(pre) + (neg1 && zero ? '' : '') + (sign ? '' : '') + (sign) + is + dotStr + lit(post);
  }
  XL.formatValue = function (v, fmt) {
    if (v === null || v === undefined) return { text: '' };
    if (typeof v === 'object' && v.err) return { text: v.err, err: true };
    if (typeof v === 'boolean') return { text: v ? 'TRUE' : 'FALSE' };
    fmt = fmt || 'General';
    if (typeof v === 'string') {
      if (fmt === 'General') return { text: v };
      const secs = fmtCache[fmt] || (fmtCache[fmt] = splitSections(fmt)); const tsec = secs[3] || (secs.length === 1 && /@/.test(secs[0]) ? secs[0] : null);
      if (tsec) return { text: tsec.replace(/"([^"]*)"|\\(.)|_.|\*./g, (x, a, b) => a !== undefined ? a : b !== undefined ? b : ' ').replace(/@/g, v) };
      return { text: v };
    }
    let n = v;
    if (fmt === 'General' || fmt === '') return { text: general(n) };
    const secs = fmtCache[fmt] || (fmtCache[fmt] = splitSections(fmt)); let sec, neg = n < 0 || Object.is(n, -0) && false;
    // conditions
    const cond = s => { const m = /^\[(<=|>=|<>|<|>|=)(-?[\d.]+)\]/.exec(s); return m ? { op: m[1], val: +m[2] } : null; };
    const test = (c, x) => c && ({ '<': x < c.val, '>': x > c.val, '<=': x <= c.val, '>=': x >= c.val, '=': x === c.val, '<>': x !== c.val }[c.op]);
    let usedNeg = false;
    if (secs.length === 1) sec = secs[0]; else if (secs.length === 2 && !cond(secs[0])) { if (n < 0) { sec = secs[1]; usedNeg = true; } else sec = secs[0]; } else if (cond(secs[0]) || cond(secs[1])) { const c0 = cond(secs[0]), c1 = cond(secs[1]); if (c0 && test(c0, n)) sec = secs[0]; else if (c1 && test(c1, n)) { sec = secs[1]; } else sec = secs[2] || secs[1] || secs[0]; usedNeg = n < 0 && sec !== secs[0] && sec === secs[1] && !c1; }
    else { if (n > 0) sec = secs[0]; else if (n < 0) { sec = secs[1]; usedNeg = true; } else sec = secs[2] !== undefined ? secs[2] : secs[0]; }
    let color = null; sec = sec.replace(/\[(Red|Blue|Green|Black|White|Yellow|Magenta|Cyan|Color\s*\d+)\]/i, (m, c) => { color = COLORS[c.toLowerCase()] || null; return ''; }).replace(/^\[(<=|>=|<>|<|>|=)-?[\d.]+\]/, '').replace(/\[\$([^\]-]*)(?:-[^\]]*)?\]/g, (m, c) => '"' + c + '"');
    if (/^\s*General\s*$/i.test(sec) || /^General$/i.test(sec.trim())) return { text: general(n), color };
    const abs = Math.abs(n);
    if (isDateFmt(sec)) { if (n < 0) return { text: '########', color }; return { text: fmtDate(sec, n), color, date: true }; }
    if (/^@$/.test(sec.trim())) return { text: general(n), color };
    return { text: fmtNumber(sec, abs, n < 0 && !usedNeg), color };
  };
  XL.isDateFormat = f => { if (!f || f === 'General') return false; const s = splitSections(f)[0].replace(/\[(Red|Blue|Green|Black|White|Yellow|Magenta|Cyan|Color\s*\d+)\]/i, ''); return isDateFmt(s); };

  /* ---------------------------------------------------------------- tokenizer + parser */
  const ERRS = ['#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!', '#SPILL!', '#CALC!'];
  const E = n => ({ err: n });
  XL.ERR = E;
  function tokenize(src) {
    const toks = []; let i = 0; const n = src.length;
    while (i < n) {
      const c = src[i], rest = src.slice(i); let m;
      if (/\s/.test(c)) { i++; continue; }
      if (c === '"') { let j = i + 1, s = ''; while (j < n) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; } if (j >= n) throw new Error('Unterminated string'); toks.push({ t: 'str', v: s }); i = j + 1; continue; }
      if (c === '#') { const e = ERRS.find(x => rest.toUpperCase().startsWith(x)); if (e) { toks.push({ t: 'err', v: e }); i += e.length; continue; } throw new Error('Bad token'); }
      if ((m = /^(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/.exec(rest)) && !/^[A-Za-z_]/.test(rest)) { toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      let sheet = null;
      if ((m = /^(?:'((?:[^']|'')+)'|([A-Za-z_\u00c0-\uffff][\w.\u00c0-\uffff]*))!/.exec(rest))) { sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]; i += m[0].length; }
      const r2 = src.slice(i);
      if ((m = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?(?![\w(.])/.exec(r2))) {
        const a = XL.parseAddr(m[1]), b = m[2] ? XL.parseAddr(m[2]) : null;
        if (a && (!m[2] || b)) { toks.push(b ? { t: 'range', sheet, a, b } : { t: 'ref', sheet, a }); i += m[0].length; continue; }
      }
      if ((m = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w(])/.exec(r2))) { toks.push({ t: 'range', sheet, a: { r: 0, c: XL.colIdx(m[1].replace('$', '')), ar: true, ac: m[1][0] === '$' }, b: { r: XL.MAXR - 1, c: XL.colIdx(m[2].replace('$', '')), ar: true, ac: m[2][0] === '$' }, fullCol: true }); i += m[0].length; continue; }
      if ((m = /^(\$?\d+):(\$?\d+)(?![\w(])/.exec(r2))) { toks.push({ t: 'range', sheet, a: { r: +m[1].replace('$', '') - 1, c: 0, ac: true, ar: m[1][0] === '$' }, b: { r: +m[2].replace('$', '') - 1, c: XL.MAXC - 1, ac: true, ar: m[2][0] === '$' }, fullRow: true }); i += m[0].length; continue; }
      if (sheet !== null) throw new Error('Bad reference');
      if ((m = /^[A-Za-z_\\\u00c0-\uffff][\w.\u00c0-\uffff]*/.exec(rest))) {
        const nm = m[0]; i += nm.length; const up = nm.toUpperCase();
        if (src[i] === '(') toks.push({ t: 'fn', v: up.replace(/^_XLFN\./, '').replace(/^_XLWS\./, '') });
        else if (up === 'TRUE' || up === 'FALSE') toks.push({ t: 'bool', v: up === 'TRUE' }); else toks.push({ t: 'name', v: nm });
        continue;
      }
      const two = src.substr(i, 2);
      if (['<=', '>=', '<>'].includes(two)) { toks.push({ t: 'op', v: two }); i += 2; continue; }
      if ('+-*/^&=<>%(),;{}:'.includes(c)) { toks.push({ t: 'op', v: c }); i++; continue; }
      throw new Error('Unexpected character ' + c);
    }
    return toks;
  }
  const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
  XL.parse = function (src) {
    const toks = tokenize(src); let p = 0;
    const peek = () => toks[p], isOp = v => toks[p] && toks[p].t === 'op' && toks[p].v === v;
    function expr(min) {
      let left = unary();
      for (;;) {
        const t = peek(); if (!t || t.t !== 'op' || !PREC[t.v] || PREC[t.v] < min) break;
        p++; const right = expr(PREC[t.v] + 1); left = { t: 'bin', op: t.v, a: left, b: right };
      }
      return left;
    }
    function unary() {
      if (isOp('-')) { p++; return { t: 'un', op: '-', a: unary() }; }
      if (isOp('+')) { p++; return unary(); }
      let x = primary(); while (isOp('%')) { p++; x = { t: 'pct', a: x }; } return x;
    }
    function primary() {
      const t = toks[p++]; if (!t) throw new Error('Unexpected end');
      switch (t.t) {
        case 'num': return { t: 'num', v: t.v }; case 'str': return { t: 'str', v: t.v }; case 'bool': return { t: 'bool', v: t.v }; case 'err': return { t: 'err', v: t.v };
        case 'ref': return { t: 'ref', sheet: t.sheet, r: t.a.r, c: t.a.c, ar: t.a.ar, ac: t.a.ac };
        case 'range': { const a = t.a, b = t.b; const nd = { t: 'range', sheet: t.sheet, r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c), ar1: a.r <= b.r ? a.ar : b.ar, ac1: a.c <= b.c ? a.ac : b.ac, ar2: a.r <= b.r ? b.ar : a.ar, ac2: a.c <= b.c ? b.ac : a.ac }; if (t.fullCol) nd.fullCol = true; if (t.fullRow) nd.fullRow = true; return nd; }
        case 'name': return { t: 'name', v: t.v };
        case 'fn': { p++; const args = []; if (isOp(')')) { p++; return { t: 'fn', name: t.v, args }; } for (;;) { if (isOp(',') || isOp(')')) args.push({ t: 'empty' }); else args.push(expr(1)); if (isOp(',') || isOp(';')) { p++; continue; } if (isOp(')')) { p++; break; } throw new Error('Expected , or )'); } return { t: 'fn', name: t.v, args }; }
        case 'op':
          if (t.v === '(') { const e = expr(1); if (!isOp(')')) throw new Error('Missing )'); p++; return { t: 'paren', a: e }; }
          if (t.v === '{') { const rows = [[]]; for (;;) { const x = unary(); rows[rows.length - 1].push(x); if (isOp(',')) { p++; continue; } if (isOp(';')) { p++; rows.push([]); continue; } if (isOp('}')) { p++; break; } throw new Error('Bad array'); } return { t: 'arr', rows }; }
      }
      throw new Error('Unexpected ' + (t.v || t.t));
    }
    const ast = expr(1); if (p < toks.length) throw new Error('Unexpected ' + (toks[p].v || toks[p].t)); return ast;
  };

  /* ---------------------------------------------------------------- AST utilities */
  XL.astToStr = function (n, cur) {
    const sh = n.sheet ? quoteSheet(n.sheet) + '!' : '';
    switch (n.t) {
      case 'num': return String(n.v); case 'str': return '"' + n.v.replace(/"/g, '""') + '"'; case 'bool': return n.v ? 'TRUE' : 'FALSE'; case 'err': return n.v; case 'empty': return '';
      case 'ref': return sh + XL.addr(n.r, n.c, n.ar, n.ac);
      case 'range': if (n.fullCol) return sh + (n.ac1 ? '$' : '') + XL.colName(n.c1) + ':' + (n.ac2 ? '$' : '') + XL.colName(n.c2); if (n.fullRow) return sh + (n.ar1 ? '$' : '') + (n.r1 + 1) + ':' + (n.ar2 ? '$' : '') + (n.r2 + 1); return sh + XL.addr(n.r1, n.c1, n.ar1, n.ac1) + ':' + XL.addr(n.r2, n.c2, n.ar2, n.ac2);
      case 'name': return n.v; case 'paren': return '(' + XL.astToStr(n.a) + ')';
      case 'un': return '-' + XL.astToStr(n.a); case 'pct': return XL.astToStr(n.a) + '%';
      case 'bin': return XL.astToStr(n.a) + n.op + XL.astToStr(n.b);
      case 'fn': return n.name + '(' + n.args.map(a => XL.astToStr(a)).join(',') + ')';
      case 'arr': return '{' + n.rows.map(r => r.map(a => XL.astToStr(a)).join(',')).join(';') + '}';
    }
    return '';
  };
  function mapAst(n, f) {
    const r = f(n); if (r !== undefined) return r;
    switch (n.t) {
      case 'paren': case 'un': case 'pct': return Object.assign({}, n, { a: mapAst(n.a, f) });
      case 'bin': return Object.assign({}, n, { a: mapAst(n.a, f), b: mapAst(n.b, f) });
      case 'fn': return Object.assign({}, n, { args: n.args.map(a => mapAst(a, f)) });
      case 'arr': return Object.assign({}, n, { rows: n.rows.map(r => r.map(a => mapAst(a, f))) });
    }
    return n;
  }
  XL.mapAst = mapAst;
  XL.shiftAst = (ast, dr, dc) => mapAst(ast, n => {
    if (n.t === 'ref') { const r = n.ar ? n.r : n.r + dr, c = n.ac ? n.c : n.c + dc; if (r < 0 || c < 0 || r >= XL.MAXR || c >= XL.MAXC) return { t: 'err', v: '#REF!' }; return Object.assign({}, n, { r, c }); }
    if (n.t === 'range') {
      if (n.fullCol) { const c1 = n.ac1 ? n.c1 : n.c1 + dc, c2 = n.ac2 ? n.c2 : n.c2 + dc; return Object.assign({}, n, { c1, c2 }); }
      if (n.fullRow) { const r1 = n.ar1 ? n.r1 : n.r1 + dr, r2 = n.ar2 ? n.r2 : n.r2 + dr; return Object.assign({}, n, { r1, r2 }); }
      const r1 = n.ar1 ? n.r1 : n.r1 + dr, c1 = n.ac1 ? n.c1 : n.c1 + dc, r2 = n.ar2 ? n.r2 : n.r2 + dr, c2 = n.ac2 ? n.c2 : n.c2 + dc;
      if (r1 < 0 || c1 < 0 || r2 >= XL.MAXR || c2 >= XL.MAXC) return { t: 'err', v: '#REF!' }; return Object.assign({}, n, { r1, c1, r2, c2 });
    }
  });
  // structural adjust: kind 'row'|'col'; at index; count>0 insert, <0 delete (|count| removed). target = sheet whose structure changed; own = sheet owning the formula
  XL.adjustAst = (ast, target, own, kind, at, count) => mapAst(ast, n => {
    if (n.t !== 'ref' && n.t !== 'range') return;
    if ((n.sheet || own) !== target) return;
    const isR = kind === 'row';
    const adj = (i, isEnd) => { // returns new index or -1 if deleted
      if (count > 0) return i >= at ? i + count : i;
      const k = -count; if (i < at) return i; if (i >= at + k) return i - k; return -1;
    };
    if (n.t === 'ref') { const v = isR ? n.r : n.c; const nv = adj(v); if (nv < 0) return { t: 'err', v: '#REF!' }; return Object.assign({}, n, isR ? { r: nv } : { c: nv }); }
    if (n.fullCol && isR || n.fullRow && !isR) return;
    let a = isR ? n.r1 : n.c1, b = isR ? n.r2 : n.c2;
    if (count > 0) { if (a >= at) a += count; if (b >= at) b += count; }
    else { const k = -count; if (a >= at && b < at + k) return { t: 'err', v: '#REF!' }; const na = a < at ? a : (a >= at + k ? a - k : at); const nb = b < at ? b : (b >= at + k ? b - k : at - 1); a = na; b = nb; }
    return Object.assign({}, n, isR ? { r1: a, r2: b } : { c1: a, c2: b });
  });
  XL.renameSheetAst = (ast, from, to) => mapAst(ast, n => { if ((n.t === 'ref' || n.t === 'range') && n.sheet === from) return Object.assign({}, n, { sheet: to }); });
  XL.collectRefs = function (ast, own) { const out = []; mapAst(ast, n => { if (n.t === 'ref') out.push({ sheet: n.sheet || own, r1: n.r, c1: n.c, r2: n.r, c2: n.c }); else if (n.t === 'range') out.push({ sheet: n.sheet || own, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 }); }); return out; };

  /* ---------------------------------------------------------------- evaluation helpers */
  const isErr = v => v && typeof v === 'object' && v.err !== undefined;
  const isMat = v => v && typeof v === 'object' && v.mat;
  const isRef = v => v && typeof v === 'object' && v.ref;
  XL.isErr = isErr; XL.isMat = isMat;
  function matOf(rows) { return { mat: rows }; }
  function dims(m) { return [m.mat.length, m.mat.length ? m.mat[0].length : 0]; }
  function toNum(v) {
    if (isErr(v)) throw v; if (v === null || v === undefined) return 0; if (typeof v === 'number') return v; if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') { if (v.trim() === '') throw E('#VALUE!'); const p = XL.parseNumber(v) || XL.parseDate(v); if (p) return p.v; throw E('#VALUE!'); }
    if (isMat(v)) return toNum(v.mat[0][0]); throw E('#VALUE!');
  }
  function toStr(v) { if (isErr(v)) throw v; if (v === null || v === undefined) return ''; if (typeof v === 'string') return v; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return general(v); if (isMat(v)) return toStr(v.mat[0][0]); return String(v); }
  function toBool(v) { if (isErr(v)) throw v; if (typeof v === 'boolean') return v; if (v === null || v === undefined) return false; if (typeof v === 'number') return v !== 0; if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; throw E('#VALUE!'); } if (isMat(v)) return toBool(v.mat[0][0]); throw E('#VALUE!'); }
  XL.toNum = toNum; XL.toStr = toStr;
  function cmpVals(a, b) { // Excel ordering: numbers < text < logical; empty -> treated per other
    const ra = a === null || a === undefined ? (typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0) : a, rb = b === null || b === undefined ? (typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0) : b;
    const ty = x => typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2;
    if (ty(ra) !== ty(rb)) return ty(ra) - ty(rb);
    if (typeof ra === 'string') { const x = ra.toLowerCase(), y = rb.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return ra < rb ? -1 : ra > rb ? 1 : 0;
  }
  XL.cmpVals = cmpVals;
  function lift2(a, b, f) {
    if (isMat(a) || isMat(b)) {
      const A = isMat(a) ? a.mat : [[a]], B = isMat(b) ? b.mat : [[b]]; const R = Math.max(A.length, B.length), C = Math.max(A[0].length, B[0].length); const out = [];
      for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) { const x = A.length === 1 ? A[0][A[0].length === 1 ? 0 : j] : A[i] ? A[i][A[0].length === 1 ? 0 : j] : undefined, y = B.length === 1 ? B[0][B[0].length === 1 ? 0 : j] : B[i] ? B[i][B[0].length === 1 ? 0 : j] : undefined; row.push(x === undefined || y === undefined ? E('#N/A') : f(x, y)); } out.push(row); }
      return matOf(out);
    }
    return f(a, b);
  }
  function binop(op, a, b) {
    return lift2(a, b, (x, y) => {
      try {
        if (isErr(x)) return x; if (isErr(y)) return y;
        switch (op) {
          case '+': return toNum(x) + toNum(y); case '-': return toNum(x) - toNum(y); case '*': return toNum(x) * toNum(y);
          case '/': { const d = toNum(y); if (d === 0) return E('#DIV/0!'); return toNum(x) / d; }
          case '^': { const r = Math.pow(toNum(x), toNum(y)); return isNaN(r) || !isFinite(r) ? E('#NUM!') : r; }
          case '&': return toStr(x) + toStr(y);
          case '=': return cmpVals(x, y) === 0; case '<>': return cmpVals(x, y) !== 0; case '<': return cmpVals(x, y) < 0; case '>': return cmpVals(x, y) > 0; case '<=': return cmpVals(x, y) <= 0; case '>=': return cmpVals(x, y) >= 0;
        }
      } catch (e) { if (isErr(e)) return e; throw e; }
    });
  }
  function mapMat(v, f) { if (isMat(v)) return matOf(v.mat.map(r => r.map(x => { try { return isErr(x) ? x : f(x); } catch (e) { if (isErr(e)) return e; throw e; } }))); return f(v); }

  /* criteria for SUMIF/COUNTIF */
  XL.critFn = function (c) {
    if (typeof c === 'number') return v => typeof v === 'number' && v === c;
    if (typeof c === 'boolean') return v => v === c;
    if (c === null || c === undefined) return v => v === null || v === '';
    c = String(c); let m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(c); const op = m[1] || '=', rest = m[2];
    const n = rest.trim() !== '' && !isNaN(+rest) ? +rest : null; const pd = n === null ? (XL.parseDate(rest) || null) : null; const nn = n !== null ? n : pd ? pd.v : null;
    if (nn !== null) return v => { if (typeof v === 'string') { const q = XL.parseNumber(v); if (q) v = q.v; else return op === '<>'; } if (typeof v !== 'number') return op === '<>'; return { '=': v === nn, '<>': v !== nn, '<': v < nn, '>': v > nn, '<=': v <= nn, '>=': v >= nn }[op]; };
    if (rest === '' && (op === '=' || op === '<>')) return v => op === '=' ? (v === null || v === '') : !(v === null || v === '');
    const re = new RegExp('^' + rest.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~?\*/g, s => s === '~*' ? '\\*' : '.*').replace(/~?\?/g, s => s === '~?' ? '\\?' : '.') + '$', 'i');
    if (op === '=') return v => typeof v === 'string' ? re.test(v) : (v !== null && v !== undefined && !isErr(v) && re.test(String(v)));
    if (op === '<>') return v => !(typeof v === 'string' ? re.test(v) : v !== null && !isErr(v) && re.test(String(v)));
    const lr = rest.toLowerCase(); return v => typeof v === 'string' && ({ '<': v.toLowerCase() < lr, '>': v.toLowerCase() > lr, '<=': v.toLowerCase() <= lr, '>=': v.toLowerCase() >= lr }[op]);
  };

  /* ---------------------------------------------------------------- evaluator */
  // ctx: { wb, sheet (name), r, c }
  function derefRef(rf, ctx) {
    const wb = ctx.wb; const ext = wb.extent(rf.sheet);
    const r2 = Math.min(rf.r2, Math.max(rf.r1, ext.r)), c2 = Math.min(rf.c2, Math.max(rf.c1, ext.c));
    if (rf.r1 === rf.r2 && rf.c1 === rf.c2) return wb.value(rf.sheet, rf.r1, rf.c1, ctx);
    const rows = [];
    for (let r = rf.r1; r <= r2; r++) { const row = []; for (let c = rf.c1; c <= c2; c++) row.push(wb.value(rf.sheet, r, c, ctx)); rows.push(row); }
    if (!rows.length) rows.push([null]); return matOf(rows);
  }
  function val(x, ctx) { return isRef(x) ? derefRef(x, ctx) : x; }
  function evalNode(n, ctx) {
    switch (n.t) {
      case 'num': case 'str': case 'bool': return n.v;
      case 'err': return E(n.v); case 'empty': return null;
      case 'paren': return val(evalNode(n.a, ctx), ctx);
      case 'ref': return { ref: true, sheet: n.sheet || ctx.sheet, r1: n.r, c1: n.c, r2: n.r, c2: n.c };
      case 'range': return { ref: true, sheet: n.sheet || ctx.sheet, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 };
      case 'name': {
        if (ctx.locals && ctx.locals[n.v.toLowerCase()] !== undefined) return ctx.locals[n.v.toLowerCase()];
        const d = ctx.wb.names && ctx.wb.names[n.v.toLowerCase()]; if (!d) return E('#NAME?');
        if (d.ast === undefined) { try { d.ast = XL.parse(d.ref.replace(/^=/, '')); } catch (e) { d.ast = null; } }
        if (!d.ast) return E('#NAME?'); return evalNode(d.ast, Object.assign({}, ctx, { sheet: d.sheet || ctx.sheet }));
      }
      case 'un': { const a = val(evalNode(n.a, ctx), ctx); return mapMat(a, x => -toNum(x)); }
      case 'pct': { const a = val(evalNode(n.a, ctx), ctx); return mapMat(a, x => toNum(x) / 100); }
      case 'bin': { const a = val(evalNode(n.a, ctx), ctx), b = val(evalNode(n.b, ctx), ctx); return binop(n.op, a, b); }
      case 'arr': return matOf(n.rows.map(r => r.map(x => { const v = val(evalNode(x, ctx), ctx); return isMat(v) ? v.mat[0][0] : v; })));
      case 'fn': return callFn(n, ctx);
    }
    return E('#VALUE!');
  }
  function callFn(n, ctx) {
    const f = FUNCS[n.name]; if (!f) return E('#NAME?');
    try {
      if (f.lazy) return f.fn(n.args, ctx);
      const args = n.args.map(a => f.refs ? evalNode(a, ctx) : val(evalNode(a, ctx), ctx));
      if (f.min !== undefined && args.length < f.min) return E('#VALUE!'); if (f.max !== undefined && args.length > f.max) return E('#VALUE!');
      if (f.ew) { if (args.some(isMat)) { const first = args.find(isMat); return lift(args, f.fn, ctx); } }
      return f.fn(args, ctx);
    } catch (e) { if (isErr(e)) return e; if (e && e.message === 'circ') return E('#REF!'); throw e; }
  }
  function lift(args, fn, ctx) {
    let R = 1, C = 1; args.forEach(a => { if (isMat(a)) { R = Math.max(R, a.mat.length); C = Math.max(C, a.mat[0].length); } });
    const out = []; for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < C; j++) { const aa = args.map(a => isMat(a) ? (a.mat[a.mat.length === 1 ? 0 : i] || [])[a.mat[0].length === 1 ? 0 : j] : a); try { row.push(fn(aa, ctx)); } catch (e) { if (isErr(e)) row.push(e); else throw e; } } out.push(row); }
    return matOf(out);
  }
  XL.evalFormula = function (ast, ctx) { let v; try { v = evalNode(ast, ctx); v = val(v, ctx); } catch (e) { if (isErr(e)) return e; throw e; } if (v === undefined) return null; return v; };
  XL.evalRef = (ast, ctx) => evalNode(ast, ctx);

  /* ---------------------------------------------------------------- functions */
  const FUNCS = XL.FUNCS = {};
  const def = (names, fn, o) => { (Array.isArray(names) ? names : [names]).forEach(n => { FUNCS[n] = Object.assign({ fn }, o || {}); }); };
  const ew = fn => def.bind(null); // placeholder
  function nums(args, opts) { // collect numbers from args per Excel rules
    opts = opts || {}; const out = [];
    const visit = (a, direct) => {
      if (isErr(a)) throw a;
      if (isMat(a)) { a.mat.forEach(r => r.forEach(x => visit(x, false))); return; }
      if (typeof a === 'number') out.push(a);
      else if (direct) { if (typeof a === 'boolean') out.push(a ? 1 : 0); else if (typeof a === 'string') out.push(toNum(a)); else if (a === null && opts.emptyZero) out.push(0); }
      else if (opts.text && typeof a === 'string') out.push(0); else if (opts.bool && typeof a === 'boolean') out.push(a ? 1 : 0);
    };
    args.forEach(a => visit(a, !isMat(a))); return out;
  }
  function flatten(v) { return isMat(v) ? [].concat(...v.mat) : [v]; }
  XL.flatten = flatten;
  const scalar = v => isMat(v) ? v.mat[0][0] : v;
  const N = (v, d) => (v === undefined || v === null && d !== undefined && false) ? d : toNum(v);
  const ewDef = (names, fn, o) => def(names, fn, Object.assign({ ew: true }, o || {}));

  // math
  def('SUM', a => nums(a).reduce((x, y) => x + y, 0));
  def('PRODUCT', a => nums(a).reduce((x, y) => x * y, 1));
  def('AVERAGE', a => { const n = nums(a); if (!n.length) return E('#DIV/0!'); return n.reduce((x, y) => x + y, 0) / n.length; });
  def('AVERAGEA', a => { const n = nums(a, { text: true, bool: true }); if (!n.length) return E('#DIV/0!'); return n.reduce((x, y) => x + y, 0) / n.length; });
  def('MIN', a => { const n = nums(a); return n.length ? Math.min(...n) : 0; }); def('MAX', a => { const n = nums(a); return n.length ? Math.max(...n) : 0; });
  def('COUNT', a => { let c = 0; a.forEach(x => flatten(x).forEach(v => { if (typeof v === 'number') c++; else if (!isMat(x) && (typeof v === 'boolean' || (typeof v === 'string' && (XL.parseNumber(v) || XL.parseDate(v))))) c++; })); return c; });
  def('COUNTA', a => { let c = 0; a.forEach(x => flatten(x).forEach(v => { if (v !== null && v !== undefined) c++; })); return c; });
  def('COUNTBLANK', a => { let c = 0; a.forEach(x => flatten(x).forEach(v => { if (v === null || v === '') c++; })); return c; });
  def('SUMSQ', a => nums(a).reduce((x, y) => x + y * y, 0));
  def('MEDIAN', a => { const n = nums(a).sort((x, y) => x - y); if (!n.length) return E('#NUM!'); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; });
  def('MODE', a => { const n = nums(a); const c = new Map(); n.forEach(x => c.set(x, (c.get(x) || 0) + 1)); let b = null, bc = 1; n.forEach(x => { if (c.get(x) > bc) { bc = c.get(x); b = x; } }); return b === null ? E('#N/A') : b; }); FUNCS['MODE.SNGL'] = FUNCS.MODE;
  const variance = (n, pop) => { if (n.length < (pop ? 1 : 2)) return E('#DIV/0!'); const m = n.reduce((x, y) => x + y, 0) / n.length; return n.reduce((x, y) => x + (y - m) * (y - m), 0) / (n.length - (pop ? 0 : 1)); };
  def(['VAR', 'VAR.S'], a => variance(nums(a), false)); def(['VARP', 'VAR.P'], a => variance(nums(a), true));
  def(['STDEV', 'STDEV.S'], a => { const v = variance(nums(a), false); return isErr(v) ? v : Math.sqrt(v); }); def(['STDEVP', 'STDEV.P'], a => { const v = variance(nums(a), true); return isErr(v) ? v : Math.sqrt(v); });
  def('LARGE', a => { const n = nums([a[0]]).sort((x, y) => y - x), k = toNum(a[1]); return k >= 1 && k <= n.length ? n[k - 1] : E('#NUM!'); });
  def('SMALL', a => { const n = nums([a[0]]).sort((x, y) => x - y), k = toNum(a[1]); return k >= 1 && k <= n.length ? n[k - 1] : E('#NUM!'); });
  const pctile = (n, k) => { n.sort((x, y) => x - y); if (!n.length || k < 0 || k > 1) return E('#NUM!'); const p = k * (n.length - 1), lo = Math.floor(p), hi = Math.ceil(p); return n[lo] + (n[hi] - n[lo]) * (p - lo); };
  def(['PERCENTILE', 'PERCENTILE.INC'], a => pctile(nums([a[0]]), toNum(a[1]))); def(['QUARTILE', 'QUARTILE.INC'], a => pctile(nums([a[0]]), toNum(a[1]) / 4));
  def(['RANK', 'RANK.EQ'], a => { const x = toNum(a[0]), n = nums([a[1]]), asc = a[2] !== undefined && toBool(a[2]); if (!n.includes(x)) return E('#N/A'); return 1 + n.filter(v => asc ? v < x : v > x).length; });
  const pairs = a => { const X = flatten(a[0]), Y = flatten(a[1]); const out = []; for (let i = 0; i < Math.min(X.length, Y.length); i++) if (typeof X[i] === 'number' && typeof Y[i] === 'number') out.push([X[i], Y[i]]); return out; };
  const regress = p => { const n = p.length; if (n < 2) throw E('#DIV/0!'); const mx = p.reduce((s, q) => s + q[0], 0) / n, my = p.reduce((s, q) => s + q[1], 0) / n; let sxy = 0, sxx = 0, syy = 0; p.forEach(q => { sxy += (q[0] - mx) * (q[1] - my); sxx += (q[0] - mx) ** 2; syy += (q[1] - my) ** 2; }); if (sxx === 0) throw E('#DIV/0!'); return { slope: sxy / sxx, icpt: my - sxy / sxx * mx, r: sxy / Math.sqrt(sxx * syy) }; };
  def('CORREL', a => regress(pairs(a)).r); def('PEARSON', a => regress(pairs(a)).r); def('RSQ', a => regress(pairs(a)).r ** 2);
  def('SLOPE', a => regress(pairs([a[1], a[0]])).slope); def('INTERCEPT', a => regress(pairs([a[1], a[0]])).icpt);
  def(['FORECAST', 'FORECAST.LINEAR'], a => { const r = regress(pairs([a[2], a[1]])); return r.icpt + r.slope * toNum(a[0]); });
  def('GEOMEAN', a => { const n = nums(a); return Math.exp(n.reduce((s, x) => s + Math.log(x), 0) / n.length); });
  def('DEVSQ', a => { const n = nums(a); const m = n.reduce((x, y) => x + y, 0) / n.length; return n.reduce((x, y) => x + (y - m) ** 2, 0); });
  ewDef('ABS', a => Math.abs(toNum(a[0]))); ewDef('SIGN', a => Math.sign(toNum(a[0]))); ewDef('INT', a => Math.floor(toNum(a[0])));
  ewDef('SQRT', a => { const x = toNum(a[0]); return x < 0 ? E('#NUM!') : Math.sqrt(x); }); ewDef('EXP', a => Math.exp(toNum(a[0])));
  ewDef('LN', a => { const x = toNum(a[0]); return x <= 0 ? E('#NUM!') : Math.log(x); }); ewDef('LOG10', a => { const x = toNum(a[0]); return x <= 0 ? E('#NUM!') : Math.log10(x); });
  ewDef('LOG', a => { const x = toNum(a[0]), b = a[1] === undefined ? 10 : toNum(a[1]); return x <= 0 || b <= 0 || b === 1 ? E('#NUM!') : Math.log(x) / Math.log(b); });
  def('PI', () => Math.PI);
  [['SIN', Math.sin], ['COS', Math.cos], ['TAN', Math.tan], ['ATAN', Math.atan], ['SINH', Math.sinh], ['COSH', Math.cosh], ['TANH', Math.tanh]].forEach(([n, f]) => ewDef(n, a => f(toNum(a[0]))));
  ewDef('ASIN', a => { const x = toNum(a[0]); return Math.abs(x) > 1 ? E('#NUM!') : Math.asin(x); }); ewDef('ACOS', a => { const x = toNum(a[0]); return Math.abs(x) > 1 ? E('#NUM!') : Math.acos(x); });
  ewDef('ATAN2', a => Math.atan2(toNum(a[1]), toNum(a[0]))); ewDef('DEGREES', a => toNum(a[0]) * 180 / Math.PI); ewDef('RADIANS', a => toNum(a[0]) * Math.PI / 180);
  const roundTo = (x, d, mode) => { const f = Math.pow(10, d); const y = x * f; const e = 1e-9 * Math.max(1, Math.abs(y)); let r; if (mode === 'up') r = y >= 0 ? Math.ceil(y - e) : Math.floor(y + e); else if (mode === 'down') r = y >= 0 ? Math.floor(y + e) : Math.ceil(y - e); else r = y >= 0 ? Math.floor(y + 0.5 + e) : -Math.floor(-y + 0.5 + e); return r / f; };
  ewDef('ROUND', a => roundTo(toNum(a[0]), toNum(a[1] === undefined ? 0 : a[1]), 'n')); ewDef('ROUNDUP', a => roundTo(toNum(a[0]), toNum(a[1] === undefined ? 0 : a[1]), 'up')); ewDef('ROUNDDOWN', a => roundTo(toNum(a[0]), toNum(a[1] === undefined ? 0 : a[1]), 'down'));
  ewDef('TRUNC', a => roundTo(toNum(a[0]), a[1] === undefined ? 0 : toNum(a[1]), 'down'));
  ewDef(['CEILING', 'CEILING.MATH'], a => { const x = toNum(a[0]), s = a[1] === undefined ? 1 : toNum(a[1]); return s === 0 ? 0 : Math.ceil(x / s - 1e-12) * s; });
  ewDef(['FLOOR', 'FLOOR.MATH'], a => { const x = toNum(a[0]), s = a[1] === undefined ? 1 : toNum(a[1]); return s === 0 ? 0 : Math.floor(x / s + 1e-12) * s; });
  ewDef('MROUND', a => { const x = toNum(a[0]), m = toNum(a[1]); return m === 0 ? 0 : Math.round(x / m) * m; });
  ewDef('EVEN', a => { const x = toNum(a[0]); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; }); ewDef('ODD', a => { const x = toNum(a[0]); let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r++; return x < 0 ? -r : r; });
  ewDef('MOD', a => { const x = toNum(a[0]), d = toNum(a[1]); if (d === 0) return E('#DIV/0!'); return x - d * Math.floor(x / d); });
  ewDef(['POWER'], a => { const r = Math.pow(toNum(a[0]), toNum(a[1])); return isNaN(r) ? E('#NUM!') : r; });
  ewDef('QUOTIENT', a => { const d = toNum(a[1]); return d === 0 ? E('#DIV/0!') : Math.trunc(toNum(a[0]) / d); });
  ewDef('FACT', a => { const x = Math.floor(toNum(a[0])); if (x < 0) return E('#NUM!'); let r = 1; for (let i = 2; i <= x; i++) r *= i; return r; });
  const gcd = (a, b) => b ? gcd(b, a % b) : a;
  def('GCD', a => nums(a).reduce((x, y) => gcd(x, Math.abs(y)), 0)); def('LCM', a => nums(a).reduce((x, y) => x * Math.abs(y) / (gcd(x, Math.abs(y)) || 1), 1));
  ewDef('COMBIN', a => { const n = toNum(a[0]), k = toNum(a[1]); if (k > n || k < 0) return E('#NUM!'); let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); });
  ewDef('PERMUT', a => { const n = toNum(a[0]), k = toNum(a[1]); let r = 1; for (let i = 0; i < k; i++) r *= (n - i); return r; });
  def('RAND', () => Math.random(), { volatile: true }); def('RANDBETWEEN', a => { const lo = Math.ceil(toNum(a[0])), hi = Math.floor(toNum(a[1])); return lo > hi ? E('#NUM!') : lo + Math.floor(Math.random() * (hi - lo + 1)); }, { volatile: true });
  def('RANDARRAY', a => { const R = a[0] === undefined ? 1 : toNum(a[0]), C = a[1] === undefined ? 1 : toNum(a[1]); return matOf(Array.from({ length: R }, () => Array.from({ length: C }, () => Math.random()))); }, { volatile: true });
  def('SUMPRODUCT', a => { const ms = a.map(x => isMat(x) ? x.mat : [[x]]); const R = ms[0].length, C = ms[0][0].length; let s = 0; for (const m of ms) if (m.length !== R || m[0].length !== C) return E('#VALUE!'); for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { let p = 1; for (const m of ms) { const v = m[i][j]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : (typeof v === 'boolean' ? +v : 0); } s += p; } return s; });
  def('SUBTOTAL', a => { const f = toNum(a[0]) % 100; const m = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' }[f]; return m ? FUNCS[m].fn(a.slice(1)) : E('#VALUE!'); });
  def('AGGREGATE', a => { const f = toNum(a[0]); const m = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP', 12: 'MEDIAN', 14: 'LARGE', 15: 'SMALL' }[f]; const rest = a.slice(2).map(x => isMat(x) ? matOf(x.mat.map(r => r.filter(v => !isErr(v)))) : x); return m ? FUNCS[m].fn(rest) : E('#VALUE!'); });
  // conditional aggregates
  function critPairs(a, from) { const out = []; for (let i = from; i + 1 < a.length; i += 2) { const rng = isMat(a[i]) ? a[i].mat : [[a[i]]]; out.push({ rng, fn: XL.critFn(isMat(a[i + 1]) ? a[i + 1].mat[0][0] : a[i + 1]) }); } return out; }
  function matchRows(cp) { const R = cp[0].rng.length, C = cp[0].rng[0].length; const hit = []; for (let i = 0; i < R; i++) for (let j = 0; j < C; j++) { if (cp.every(p => p.rng[i] && p.fn(p.rng[i][j]))) hit.push([i, j]); } return hit; }
  def('COUNTIF', a => matchRows(critPairs(a, 0)).length); def('COUNTIFS', a => matchRows(critPairs(a, 0)).length);
  const sumRange = (a, avgIdx, cp, sumArr, agg) => { const hit = matchRows(cp); const vals = hit.map(([i, j]) => (sumArr[i] || [])[j]).filter(v => typeof v === 'number'); return agg(vals); };
  def('SUMIF', a => { const cp = critPairs([a[0], a[1]], 0); const s = a[2] !== undefined ? (isMat(a[2]) ? a[2].mat : [[a[2]]]) : cp[0].rng; return sumRange(a, 0, cp, s, v => v.reduce((x, y) => x + y, 0)); });
  def('SUMIFS', a => sumRange(a, 0, critPairs(a, 1), isMat(a[0]) ? a[0].mat : [[a[0]]], v => v.reduce((x, y) => x + y, 0)));
  def('AVERAGEIF', a => { const cp = critPairs([a[0], a[1]], 0); const s = a[2] !== undefined ? (isMat(a[2]) ? a[2].mat : [[a[2]]]) : cp[0].rng; return sumRange(a, 0, cp, s, v => v.length ? v.reduce((x, y) => x + y, 0) / v.length : E('#DIV/0!')); });
  def('AVERAGEIFS', a => sumRange(a, 0, critPairs(a, 1), isMat(a[0]) ? a[0].mat : [[a[0]]], v => v.length ? v.reduce((x, y) => x + y, 0) / v.length : E('#DIV/0!')));
  def('MAXIFS', a => sumRange(a, 0, critPairs(a, 1), isMat(a[0]) ? a[0].mat : [[a[0]]], v => v.length ? Math.max(...v) : 0)); def('MINIFS', a => sumRange(a, 0, critPairs(a, 1), isMat(a[0]) ? a[0].mat : [[a[0]]], v => v.length ? Math.min(...v) : 0));
  // logical
  def('IF', (args, ctx) => { const c = val(evalNode(args[0], ctx), ctx); let b; try { b = isMat(c) ? c : toBool(c); } catch (e) { if (isErr(e)) return e; throw e; } if (isMat(b)) { return matOf(b.mat.map(r => r.map(x => { try { return toBool(x) ? (args[1] ? scalar(val(evalNode(args[1], ctx), ctx)) : true) : (args[2] ? scalar(val(evalNode(args[2], ctx), ctx)) : false); } catch (e) { if (isErr(e)) return e; throw e; } }))); } if (b) return args[1] ? (args[1].t === 'empty' ? 0 : val(evalNode(args[1], ctx), ctx)) : true; return args[2] ? (args[2].t === 'empty' ? 0 : val(evalNode(args[2], ctx), ctx)) : false; }, { lazy: true });
  def('IFS', (args, ctx) => { for (let i = 0; i + 1 < args.length; i += 2) { const c = val(evalNode(args[i], ctx), ctx); if (isErr(c)) return c; if (toBool(c)) return val(evalNode(args[i + 1], ctx), ctx); } return E('#N/A'); }, { lazy: true });
  def('IFERROR', (args, ctx) => { const v = val(evalNode(args[0], ctx), ctx); if (isErr(v)) return val(evalNode(args[1], ctx), ctx); if (isMat(v)) return matOf(v.mat.map(r => r.map(x => isErr(x) ? scalar(val(evalNode(args[1], ctx), ctx)) : x))); return v; }, { lazy: true });
  def('IFNA', (args, ctx) => { const v = val(evalNode(args[0], ctx), ctx); return isErr(v) && v.err === '#N/A' ? val(evalNode(args[1], ctx), ctx) : v; }, { lazy: true });
  def('CHOOSE', (args, ctx) => { const i = toNum(val(evalNode(args[0], ctx), ctx)); if (i < 1 || i >= args.length) return E('#VALUE!'); return val(evalNode(args[Math.floor(i)], ctx), ctx); }, { lazy: true });
  def('SWITCH', (args, ctx) => { const v = val(evalNode(args[0], ctx), ctx); for (let i = 1; i + 1 < args.length; i += 2) { if (cmpVals(v, val(evalNode(args[i], ctx), ctx)) === 0) return val(evalNode(args[i + 1], ctx), ctx); } return args.length % 2 === 0 ? val(evalNode(args[args.length - 1], ctx), ctx) : E('#N/A'); }, { lazy: true });
  const boolsOf = a => { const o = []; a.forEach(x => flatten(x).forEach(v => { if (isErr(v)) throw v; if (typeof v === 'boolean') o.push(v); else if (typeof v === 'number') o.push(v !== 0); else if (typeof v === 'string' && !isMat(x)) o.push(toBool(v)); })); return o; };
  def('AND', a => { const b = boolsOf(a); return b.length ? b.every(Boolean) : E('#VALUE!'); }); def('OR', a => { const b = boolsOf(a); return b.length ? b.some(Boolean) : E('#VALUE!'); });
  def('XOR', a => boolsOf(a).filter(Boolean).length % 2 === 1); ewDef('NOT', a => !toBool(a[0])); def('TRUE', () => true); def('FALSE', () => false);
  // info
  def('ISBLANK', a => { const v = a[0]; return isMat(v) ? matOf(v.mat.map(r => r.map(x => x === null))) : v === null; }, {});
  const isFn = (n, f) => def(n, a => isMat(a[0]) ? matOf(a[0].mat.map(r => r.map(f))) : f(a[0]));
  isFn('ISNUMBER', v => typeof v === 'number'); isFn('ISTEXT', v => typeof v === 'string'); isFn('ISNONTEXT', v => typeof v !== 'string'); isFn('ISLOGICAL', v => typeof v === 'boolean'); isFn('ISERROR', v => isErr(v)); isFn('ISERR', v => isErr(v) && v.err !== '#N/A'); isFn('ISNA', v => isErr(v) && v.err === '#N/A');
  isFn('ISEVEN', v => Math.floor(Math.abs(toNum(v))) % 2 === 0); isFn('ISODD', v => Math.floor(Math.abs(toNum(v))) % 2 === 1);
  def('NA', () => E('#N/A')); def('N', a => typeof a[0] === 'number' ? a[0] : typeof a[0] === 'boolean' ? +a[0] : 0); def('T', a => typeof a[0] === 'string' ? a[0] : '');
  def('TYPE', a => isMat(a[0]) ? 64 : isErr(a[0]) ? 16 : typeof a[0] === 'string' ? 2 : typeof a[0] === 'boolean' ? 4 : 1);
  def('ERROR.TYPE', a => { const i = ERRS.indexOf(isErr(a[0]) ? a[0].err : ''); return i < 0 ? E('#N/A') : i + 1; });
  // text
  ewDef('LEN', a => toStr(a[0]).length); ewDef('UPPER', a => toStr(a[0]).toUpperCase()); ewDef('LOWER', a => toStr(a[0]).toLowerCase());
  ewDef('PROPER', a => toStr(a[0]).toLowerCase().replace(/(^|[^a-zA-Z\u00c0-\uffff])([a-z\u00e0-\uffff])/g, (m, p, c) => p + c.toUpperCase())); ewDef('TRIM', a => toStr(a[0]).replace(/^ +| +$/g, '').replace(/ {2,}/g, ' '));
  ewDef('CLEAN', a => toStr(a[0]).replace(/[\x00-\x1f]/g, '')); ewDef('LEFT', a => { const n = a[1] === undefined ? 1 : toNum(a[1]); return n < 0 ? E('#VALUE!') : toStr(a[0]).slice(0, n); }); ewDef('RIGHT', a => { const s = toStr(a[0]), n = a[1] === undefined ? 1 : toNum(a[1]); return n < 0 ? E('#VALUE!') : n === 0 ? '' : s.slice(-n); });
  ewDef('MID', a => { const s = toStr(a[0]), st = toNum(a[1]), n = toNum(a[2]); return st < 1 || n < 0 ? E('#VALUE!') : s.substr(st - 1, n); });
  ewDef('REPT', a => { const n = toNum(a[1]); return n < 0 ? E('#VALUE!') : toStr(a[0]).repeat(n); }); ewDef('EXACT', a => toStr(a[0]) === toStr(a[1]));
  ewDef('CHAR', a => String.fromCharCode(toNum(a[0]))); ewDef('CODE', a => toStr(a[0]).charCodeAt(0) || E('#VALUE!')); ewDef('UNICODE', a => toStr(a[0]).codePointAt(0)); ewDef('UNICHAR', a => String.fromCodePoint(toNum(a[0])));
  ewDef('VALUE', a => { const v = a[0]; if (typeof v === 'number') return v; const p = XL.parseNumber(toStr(v)) || XL.parseDate(toStr(v)); return p ? p.v : E('#VALUE!'); });
  ewDef('NUMBERVALUE', a => { const s = toStr(a[0]).replace(new RegExp('\\' + (a[2] === undefined ? ',' : toStr(a[2])), 'g'), ''); const p = XL.parseNumber(s.replace(a[1] ? toStr(a[1]) : '.', '.')); return p ? p.v : E('#VALUE!'); });
  ewDef('FIND', a => { const i = toStr(a[1]).indexOf(toStr(a[0]), a[2] === undefined ? 0 : toNum(a[2]) - 1); return i < 0 ? E('#VALUE!') : i + 1; });
  ewDef('SEARCH', a => { const pat = toStr(a[0]); const re = new RegExp(pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~?\*/g, s => s === '~*' ? '\\*' : '.*').replace(/~?\?/g, s => s === '~?' ? '\\?' : '.'), 'i'); const s = toStr(a[1]), st = a[2] === undefined ? 0 : toNum(a[2]) - 1; const m = re.exec(s.slice(st)); return m ? m.index + st + 1 : E('#VALUE!'); });
  ewDef('SUBSTITUTE', a => { const s = toStr(a[0]), f = toStr(a[1]), t = toStr(a[2]); if (!f) return s; if (a[3] === undefined) return s.split(f).join(t); let n = toNum(a[3]), idx = -1; for (let i = 0; i < n; i++) { idx = s.indexOf(f, idx + 1); if (idx < 0) return s; } return s.slice(0, idx) + t + s.slice(idx + f.length); });
  ewDef('REPLACE', a => { const s = toStr(a[0]), st = toNum(a[1]) - 1, n = toNum(a[2]); return s.slice(0, st) + toStr(a[3]) + s.slice(st + n); });
  def(['CONCATENATE'], a => a.map(toStr).join('')); def('CONCAT', a => { let s = ''; a.forEach(x => flatten(x).forEach(v => s += toStr(v))); return s; });
  def('TEXTJOIN', a => { const d = toStr(a[0]), ig = toBool(a[1]); const out = []; a.slice(2).forEach(x => flatten(x).forEach(v => { if (isErr(v)) throw v; const s = toStr(v); if (!(ig && s === '')) out.push(s); })); return out.join(d); });
  ewDef('TEXT', a => XL.formatValue(typeof a[0] === 'string' && XL.parseNumber(a[0]) ? XL.parseNumber(a[0]).v : a[0], toStr(a[1])).text);
  ewDef('FIXED', a => { const d = a[1] === undefined ? 2 : toNum(a[1]); return XL.formatValue(toNum(a[0]), (a[2] && toBool(a[2]) ? '0' : '#,##0') + (d > 0 ? '.' + '0'.repeat(d) : '')).text; });
  ewDef('DOLLAR', a => XL.formatValue(toNum(a[0]), '$#,##0' + (a[1] === undefined || toNum(a[1]) > 0 ? '.' + '0'.repeat(a[1] === undefined ? 2 : toNum(a[1])) : '')).text);
  ewDef('TEXTBEFORE', a => { const s = toStr(a[0]), d = toStr(a[1]), i = s.indexOf(d); return i < 0 ? E('#N/A') : s.slice(0, i); }); ewDef('TEXTAFTER', a => { const s = toStr(a[0]), d = toStr(a[1]), i = s.indexOf(d); return i < 0 ? E('#N/A') : s.slice(i + d.length); });
  def('TEXTSPLIT', a => { const s = toStr(a[0]); const cd = a[1] ? toStr(a[1]) : null, rd = a[2] ? toStr(a[2]) : null; const rows = rd ? s.split(rd) : [s]; const m = rows.map(r => cd ? r.split(cd) : [r]); const w = Math.max(...m.map(r => r.length)); return matOf(m.map(r => { while (r.length < w) r.push(E('#N/A')); return r; })); });
  // lookup
  const asMat = v => isMat(v) ? v.mat : [[v]];
  const eqLookup = (x, y) => typeof x === 'string' && typeof y === 'string' ? x.toLowerCase() === y.toLowerCase() : (typeof x === typeof y && x === y);
  const wildRe = s => new RegExp('^' + s.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~?\*/g, m => m === '~*' ? '\\*' : '.*').replace(/~?\?/g, m => m === '~?' ? '\\?' : '.') + '$', 'i');
  function findIdx(look, arr, type) { // arr flat list
    if (isErr(look)) throw look;
    if (type === 0) { const wild = typeof look === 'string' && /[*?]/.test(look) ? wildRe(look) : null; for (let i = 0; i < arr.length; i++) { if (wild ? (typeof arr[i] === 'string' && wild.test(arr[i])) : eqLookup(look, arr[i])) return i; } return -1; }
    let best = -1; for (let i = 0; i < arr.length; i++) { const v = arr[i]; if (v === null || typeof v !== typeof look) continue; const c = cmpVals(v, look); if (type === 1 ? c <= 0 : c >= 0) { if (best < 0 || (type === 1 ? cmpVals(v, arr[best]) >= 0 : cmpVals(v, arr[best]) <= 0)) best = i; } } return best;
  }
  def('MATCH', a => { const arr = flatten(a[1]); const t = a[2] === undefined ? 1 : toNum(a[2]); const i = findIdx(scalar(a[0]), arr, t === 0 ? 0 : t > 0 ? 1 : -1); return i < 0 ? E('#N/A') : i + 1; });
  def('VLOOKUP', a => { const m = asMat(a[1]), ci = toNum(a[2]); if (ci < 1 || ci > m[0].length) return E('#REF!'); const exact = a[3] !== undefined && !toBool(a[3]); const i = findIdx(scalar(a[0]), m.map(r => r[0]), exact ? 0 : 1); return i < 0 ? E('#N/A') : m[i][ci - 1]; });
  def('HLOOKUP', a => { const m = asMat(a[1]), ri = toNum(a[2]); if (ri < 1 || ri > m.length) return E('#REF!'); const exact = a[3] !== undefined && !toBool(a[3]); const i = findIdx(scalar(a[0]), m[0], exact ? 0 : 1); return i < 0 ? E('#N/A') : m[ri - 1][i]; });
  def('LOOKUP', a => { const arr = flatten(a[1]); const i = findIdx(scalar(a[0]), arr, 1); if (i < 0) return E('#N/A'); const res = a[2] !== undefined ? flatten(a[2]) : arr; return res[i]; });
  def('XMATCH', a => { const arr = flatten(a[1]); const mm = a[2] === undefined ? 0 : toNum(a[2]); const i = mm === 0 ? findIdx(scalar(a[0]), arr, 0) : findIdx(scalar(a[0]), arr, mm > 0 ? 1 : -1); return i < 0 ? E('#N/A') : i + 1; });
  def('XLOOKUP', a => { const look = asMat(a[1]), ret = asMat(a[2]); const vert = look[0].length === 1 && look.length > 1 || (look.length === 1 && look[0].length === 1); const arr = flatten(a[1]); const mm = a[4] === undefined ? 0 : toNum(a[4]); const rev = a[5] !== undefined && toNum(a[5]) < 0; let a2 = arr; let i; if (rev) { const r = arr.slice().reverse(); i = findIdx(scalar(a[0]), r, mm === 0 ? 0 : mm > 0 ? 1 : -1); if (i >= 0) i = arr.length - 1 - i; } else i = findIdx(scalar(a[0]), arr, mm === 0 ? 0 : mm > 0 ? 1 : -1); if (i < 0) return a[3] !== undefined ? a[3] : E('#N/A'); if (look[0].length === 1 || look.length === 1 && look[0].length === 1) return ret.length === arr.length || look[0].length === 1 ? (ret[0].length === 1 ? ret[i][0] : matOf([ret[i]])) : ret[0][i]; return ret.length === 1 ? ret[0][i] : matOf(ret.map(r => [r[i]])); });
  def('INDEX', (a, ctx) => { const m = asMat(a[0]); let r = a[1] === undefined ? 0 : toNum(a[1]), c = a[2] === undefined ? 0 : toNum(a[2]); if (m.length === 1 && a[2] === undefined && r > 0) { c = r; r = 1; } else if (m[0].length === 1 && a[2] === undefined && r > 0) c = 1; if (r < 0 || c < 0 || r > m.length || c > m[0].length) return E('#REF!'); if (r === 0 && c === 0) return matOf(m); if (r === 0) return matOf(m.map(x => [x[c - 1]])); if (c === 0) return matOf([m[r - 1]]); return m[r - 1][c - 1]; });
  def('ROWS', a => asMat(a[0]).length); def('COLUMNS', a => asMat(a[0])[0].length);
  def('ROW', (a, ctx) => { if (a[0] === undefined) return ctx.r + 1; if (isRef(a[0])) return a[0].r1 === a[0].r2 ? a[0].r1 + 1 : matOf(Array.from({ length: a[0].r2 - a[0].r1 + 1 }, (_, i) => [a[0].r1 + i + 1])); return E('#VALUE!'); }, { refs: true });
  def('COLUMN', (a, ctx) => { if (a[0] === undefined) return ctx.c + 1; if (isRef(a[0])) return a[0].c1 === a[0].c2 ? a[0].c1 + 1 : matOf([Array.from({ length: a[0].c2 - a[0].c1 + 1 }, (_, i) => a[0].c1 + i + 1)]); return E('#VALUE!'); }, { refs: true });
  def('ADDRESS', a => { const r = toNum(a[0]) - 1, c = toNum(a[1]) - 1, t = a[2] === undefined ? 1 : toNum(a[2]); const s = XL.addr(r, c, t === 1 || t === 2, t === 1 || t === 3); return a[4] ? quoteSheet(toStr(a[4])) + '!' + s : s; });
  def('TRANSPOSE', a => { const m = asMat(a[0]); return matOf(m[0].map((_, j) => m.map(r => r[j]))); });
  def('INDIRECT', (a, ctx) => { const s = toStr(a[0]); let sheet = ctx.sheet, addr = s; const m = /^(?:'((?:[^']|'')+)'|([^!']+))!(.+)$/.exec(s); if (m) { sheet = (m[1] || m[2]).replace(/''/g, "'"); addr = m[3]; } const g = XL.parseRange(addr); if (!g || !ctx.wb.hasSheet(sheet)) return E('#REF!'); return derefRef({ sheet, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, ctx); });
  def('OFFSET', (a, ctx) => { const rf = a[0]; if (!isRef(rf)) return E('#VALUE!'); const dr = toNum(val(a[1], ctx)), dc = toNum(val(a[2], ctx)); const h = a[3] !== undefined ? toNum(val(a[3], ctx)) : rf.r2 - rf.r1 + 1, w = a[4] !== undefined ? toNum(val(a[4], ctx)) : rf.c2 - rf.c1 + 1; const r1 = rf.r1 + dr, c1 = rf.c1 + dc; if (r1 < 0 || c1 < 0 || h < 1 || w < 1) return E('#REF!'); return derefRef({ sheet: rf.sheet, r1, c1, r2: r1 + h - 1, c2: c1 + w - 1 }, ctx); }, { refs: true });
  // dynamic arrays
  def('SEQUENCE', a => { const R = toNum(a[0]), C = a[1] === undefined ? 1 : toNum(a[1]), s = a[2] === undefined ? 1 : toNum(a[2]), st = a[3] === undefined ? 1 : toNum(a[3]); if (R < 1 || C < 1) return E('#CALC!'); return matOf(Array.from({ length: R }, (_, i) => Array.from({ length: C }, (_, j) => s + (i * C + j) * st))); });
  def('SORT', a => { const m = asMat(a[0]).map(r => r.slice()); const idx = a[1] === undefined ? 1 : toNum(a[1]), desc = a[2] !== undefined && toNum(a[2]) < 0; const byCol = a[3] !== undefined && toBool(a[3]); if (byCol) { const t = m[0].map((_, j) => m.map(r => r[j])); t.sort((x, y) => cmpVals(x[idx - 1], y[idx - 1]) * (desc ? -1 : 1)); return matOf(t[0].map((_, j) => t.map(r => r[j]))); } m.sort((x, y) => cmpVals(x[idx - 1], y[idx - 1]) * (desc ? -1 : 1)); return matOf(m); });
  def('SORTBY', a => { const m = asMat(a[0]).map((r, i) => ({ r, k: asMat(a[1])[i] ? asMat(a[1])[i][0] : null })); const desc = a[2] !== undefined && toNum(a[2]) < 0; m.sort((x, y) => cmpVals(x.k, y.k) * (desc ? -1 : 1)); return matOf(m.map(x => x.r)); });
  def('UNIQUE', a => { const m = asMat(a[0]); const seen = new Set(), out = []; m.forEach(r => { const k = JSON.stringify(r.map(x => typeof x === 'string' ? x.toLowerCase() : x)); if (!seen.has(k)) { seen.add(k); out.push(r); } }); return matOf(out); });
  def('FILTER', a => { const m = asMat(a[0]), inc = asMat(a[1]); const out = []; if (inc[0].length === 1) m.forEach((r, i) => { if (inc[i] && inc[i][0] && toBool(inc[i][0])) out.push(r); }); else { const cols = m[0].map((_, j) => inc[0][j] && toBool(inc[0][j])); m.forEach(r => out.push(r.filter((_, j) => cols[j]))); } return out.length ? matOf(out) : (a[2] !== undefined ? a[2] : E('#CALC!')); });
  def('TAKE', a => { const m = asMat(a[0]); const n = toNum(a[1]); return matOf(n >= 0 ? m.slice(0, n) : m.slice(n)); });
  def('FLATTEN', a => matOf(flatten(a[0]).map(x => [x])));
  def('MMULT', a => { const A = asMat(a[0]), B = asMat(a[1]); if (A[0].length !== B.length) return E('#VALUE!'); return matOf(A.map(r => B[0].map((_, j) => r.reduce((s, x, k) => s + toNum(x) * toNum(B[k][j]), 0)))); });
  // dates
  def('TODAY', () => Math.floor(XL.dateToSerial(new Date())), { volatile: true }); def('NOW', () => XL.dateToSerial(new Date()), { volatile: true });
  ewDef('DATE', a => { let y = toNum(a[0]), m = toNum(a[1]), d = toNum(a[2]); if (y < 0 || y > 9999) return E('#NUM!'); if (y < 1900) y += 1900; return Math.round((Date.UTC(y, m - 1, d) - D0) / 864e5); });
  ewDef('TIME', a => ((toNum(a[0]) * 3600 + toNum(a[1]) * 60 + toNum(a[2])) % 86400) / 86400);
  ewDef('DATEVALUE', a => { const p = XL.parseDate(toStr(a[0])); return p ? Math.floor(p.v) : E('#VALUE!'); }); ewDef('TIMEVALUE', a => { const p = XL.parseDate(toStr(a[0])); return p ? p.v - Math.floor(p.v) : E('#VALUE!'); });
  ewDef('YEAR', a => XL.serialParts(toNum(a[0])).y); ewDef('MONTH', a => XL.serialParts(toNum(a[0])).m); ewDef('DAY', a => XL.serialParts(toNum(a[0])).d);
  ewDef('HOUR', a => XL.serialParts(toNum(a[0])).h); ewDef('MINUTE', a => XL.serialParts(toNum(a[0])).mi); ewDef('SECOND', a => XL.serialParts(toNum(a[0])).s);
  ewDef('WEEKDAY', a => { const d = XL.serialParts(Math.floor(toNum(a[0]))).dow, t = a[1] === undefined ? 1 : toNum(a[1]); return t === 1 ? d + 1 : t === 2 ? (d + 6) % 7 + 1 : (d + 6) % 7; });
  ewDef('WEEKNUM', a => { const p = XL.serialParts(Math.floor(toNum(a[0]))); const jan1 = XL.ymdToSerial(p.y, 1, 1); const j = XL.serialParts(jan1).dow; return Math.floor((Math.floor(toNum(a[0])) - jan1 + j) / 7) + 1; });
  ewDef('EDATE', a => { const p = XL.serialParts(Math.floor(toNum(a[0]))); const t = new Date(Date.UTC(p.y, p.m - 1 + toNum(a[1]), 1)); const dim = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate(); return XL.ymdToSerial(t.getUTCFullYear(), t.getUTCMonth() + 1, Math.min(p.d, dim)); });
  ewDef('EOMONTH', a => { const p = XL.serialParts(Math.floor(toNum(a[0]))); const t = new Date(Date.UTC(p.y, p.m + toNum(a[1]), 0)); return XL.ymdToSerial(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); });
  ewDef('DAYS', a => Math.floor(toNum(a[0])) - Math.floor(toNum(a[1])));
  ewDef('DATEDIF', a => { const s = Math.floor(toNum(a[0])), e = Math.floor(toNum(a[1])); if (s > e) return E('#NUM!'); const p = XL.serialParts(s), q = XL.serialParts(e); const u = toStr(a[2]).toUpperCase(); let months = (q.y - p.y) * 12 + q.m - p.m - (q.d < p.d ? 1 : 0);
    switch (u) { case 'D': return e - s; case 'M': return months; case 'Y': return Math.floor(months / 12); case 'YM': return months % 12; case 'MD': { let d = q.d - p.d; if (d < 0) d += new Date(Date.UTC(q.y, q.m - 1, 0)).getUTCDate(); return d; } case 'YD': { let y = q.y; let ps = XL.ymdToSerial(y, p.m, Math.min(p.d, 28)); if (ps > e) ps = XL.ymdToSerial(y - 1, p.m, Math.min(p.d, 28)); return e - ps; } } return E('#NUM!'); });
  const isWE = s => { const d = XL.serialParts(s).dow; return d === 0 || d === 6; };
  def('NETWORKDAYS', a => { let s = Math.floor(toNum(a[0])), e = Math.floor(toNum(a[1])); const sg = s <= e ? 1 : -1; const hol = a[2] ? nums([a[2]]).map(Math.floor) : []; let c = 0; for (let d = s; sg > 0 ? d <= e : d >= e; d += sg) if (!isWE(d) && !hol.includes(d)) c++; return c * sg; });
  def('WORKDAY', a => { let d = Math.floor(toNum(a[0])), n = toNum(a[1]); const hol = a[2] ? nums([a[2]]).map(Math.floor) : []; const sg = n < 0 ? -1 : 1; while (n !== 0) { d += sg; if (!isWE(d) && !hol.includes(d)) n -= sg; } return d; });
  ewDef('YEARFRAC', a => Math.abs(toNum(a[1]) - toNum(a[0])) / 365.25 > 0 ? (Math.abs(toNum(a[1]) - toNum(a[0])) / 365) : 0);
  // financial
  def('PMT', a => { const r = toNum(a[0]), n = toNum(a[1]), pv = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + fv) / n; return -(r * (pv * Math.pow(1 + r, n) + fv)) / ((1 + r * t) * (Math.pow(1 + r, n) - 1)); });
  def('FV', a => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), pv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + pmt * n); return -(pv * Math.pow(1 + r, n) + pmt * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r); });
  def('PV', a => { const r = toNum(a[0]), n = toNum(a[1]), pmt = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(fv + pmt * n); return -(fv + pmt * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r) / Math.pow(1 + r, n); });
  def('NPER', a => { const r = toNum(a[0]), pmt = toNum(a[1]), pv = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); if (r === 0) return -(pv + fv) / pmt; const z = pmt * (1 + r * t); return Math.log((z - fv * r) / (z + pv * r)) / Math.log(1 + r); });
  def('NPV', a => { const r = toNum(a[0]); const v = nums(a.slice(1)); return v.reduce((s, x, i) => s + x / Math.pow(1 + r, i + 1), 0); });
  def('IRR', a => { const v = nums([a[0]]); let g = a[1] === undefined ? 0.1 : toNum(a[1]); for (let i = 0; i < 100; i++) { let f = 0, d = 0; v.forEach((x, k) => { f += x / Math.pow(1 + g, k); d -= k * x / Math.pow(1 + g, k + 1); }); if (Math.abs(f) < 1e-9) return g; if (d === 0) break; g -= f / d; } return E('#NUM!'); });
  def('RATE', a => { const n = toNum(a[0]), pmt = toNum(a[1]), pv = toNum(a[2]), fv = a[3] === undefined ? 0 : toNum(a[3]), t = a[4] === undefined ? 0 : toNum(a[4]); let g = a[5] === undefined ? 0.1 : toNum(a[5]); for (let i = 0; i < 100; i++) { const f = pv * Math.pow(1 + g, n) + pmt * (1 + g * t) * (Math.pow(1 + g, n) - 1) / g + fv; const h = 1e-6; const g2 = g + h; const f2 = pv * Math.pow(1 + g2, n) + pmt * (1 + g2 * t) * (Math.pow(1 + g2, n) - 1) / g2 + fv; const d = (f2 - f) / h; if (Math.abs(f) < 1e-9) return g; g -= f / d; } return E('#NUM!'); });
  def('SLN', a => (toNum(a[0]) - toNum(a[1])) / toNum(a[2]));
  def('IPMT', a => { const r = toNum(a[0]), per = toNum(a[1]), n = toNum(a[2]), pv = toNum(a[3]); const pmt = FUNCS.PMT.fn([r, n, pv]); const bal = pv * Math.pow(1 + r, per - 1) + pmt * (Math.pow(1 + r, per - 1) - 1) / r; return -(bal * r); });
  def('PPMT', a => FUNCS.PMT.fn([a[0], a[2], a[3]]) - FUNCS.IPMT.fn(a));
  // misc
  def('LET', (args, ctx) => { const names = Object.assign({}, ctx.locals); const c2 = Object.assign({}, ctx); for (let i = 0; i + 1 < args.length; i += 2) { if (args[i].t !== 'name') return E('#NAME?'); const v = val(evalNode(args[i + 1], c2), c2); names[args[i].v.toLowerCase()] = v; c2.locals = Object.assign({}, names); } return val(evalNode(args[args.length - 1], c2), c2); }, { lazy: true });
  def('ROMAN', a => { let n = toNum(a[0]); const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; t.forEach(([v, r]) => { while (n >= v) { s += r; n -= v; } }); return s; });
  def('HYPERLINK', a => a[1] !== undefined ? a[1] : a[0]); def('BASE', a => toNum(a[0]).toString(toNum(a[1])).toUpperCase().padStart(a[2] === undefined ? 0 : toNum(a[2]), '0'));
  def('DECIMAL', a => parseInt(toStr(a[0]), toNum(a[1]))); def('BIN2DEC', a => parseInt(toStr(a[0]), 2)); def('DEC2BIN', a => toNum(a[0]).toString(2)); def('HEX2DEC', a => parseInt(toStr(a[0]), 16)); def('DEC2HEX', a => toNum(a[0]).toString(16).toUpperCase());
  // name resolution for LET locals: patch evalNode name
  const baseName = evalNode;
  XL.fn = { evalNode, val, def, nums, toNum, toStr, toBool, matOf, E, isMat, isErr, isRef, asMat, scalar, flatten, cmpVals };
  XL.FDOC = {
    Math: ['SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'PRODUCT', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'INT', 'MOD', 'ABS', 'POWER', 'SQRT', 'EXP', 'LN', 'LOG', 'LOG10', 'PI', 'CEILING', 'FLOOR', 'MROUND', 'TRUNC', 'SIGN', 'RAND', 'RANDBETWEEN', 'FACT', 'GCD', 'LCM', 'COMBIN', 'SIN', 'COS', 'TAN', 'ATAN2', 'DEGREES', 'RADIANS', 'SUBTOTAL', 'SEQUENCE'],
    Statistical: ['AVERAGE', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK', 'COUNTIF', 'COUNTIFS', 'MAX', 'MAXIFS', 'MIN', 'MINIFS', 'MEDIAN', 'MODE', 'STDEV', 'STDEV.P', 'VAR', 'VAR.P', 'LARGE', 'SMALL', 'RANK', 'PERCENTILE', 'QUARTILE', 'CORREL', 'SLOPE', 'INTERCEPT', 'FORECAST'],
    Logical: ['IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'SWITCH', 'TRUE', 'FALSE', 'LET'],
    Text: ['LEN', 'LEFT', 'RIGHT', 'MID', 'UPPER', 'LOWER', 'PROPER', 'TRIM', 'CLEAN', 'CONCAT', 'CONCATENATE', 'TEXTJOIN', 'TEXT', 'VALUE', 'FIND', 'SEARCH', 'SUBSTITUTE', 'REPLACE', 'REPT', 'EXACT', 'CHAR', 'CODE', 'FIXED', 'DOLLAR', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT'],
    'Lookup & Reference': ['VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'LOOKUP', 'MATCH', 'XMATCH', 'INDEX', 'CHOOSE', 'OFFSET', 'INDIRECT', 'ROW', 'COLUMN', 'ROWS', 'COLUMNS', 'ADDRESS', 'TRANSPOSE', 'SORT', 'SORTBY', 'UNIQUE', 'FILTER'],
    'Date & Time': ['TODAY', 'NOW', 'DATE', 'TIME', 'YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'WEEKNUM', 'EDATE', 'EOMONTH', 'DATEDIF', 'DAYS', 'NETWORKDAYS', 'WORKDAY', 'DATEVALUE', 'TIMEVALUE'],
    Financial: ['PMT', 'FV', 'PV', 'NPER', 'RATE', 'NPV', 'IRR', 'SLN', 'IPMT', 'PPMT'],
    Information: ['ISBLANK', 'ISNUMBER', 'ISTEXT', 'ISERROR', 'ISNA', 'ISEVEN', 'ISODD', 'NA', 'TYPE']
  };
  XL.FSIG = { SUM: 'SUM(number1, [number2], …)', IF: 'IF(logical_test, value_if_true, [value_if_false])', VLOOKUP: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])', XLOOKUP: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])', SUMIF: 'SUMIF(range, criteria, [sum_range])', COUNTIF: 'COUNTIF(range, criteria)', INDEX: 'INDEX(array, row_num, [column_num])', MATCH: 'MATCH(lookup_value, lookup_array, [match_type])', ROUND: 'ROUND(number, num_digits)', CONCAT: 'CONCAT(text1, [text2], …)', TEXT: 'TEXT(value, format_text)', DATE: 'DATE(year, month, day)', PMT: 'PMT(rate, nper, pv, [fv], [type])', IFERROR: 'IFERROR(value, value_if_error)', LEFT: 'LEFT(text, [num_chars])', RIGHT: 'RIGHT(text, [num_chars])', MID: 'MID(text, start_num, num_chars)', FILTER: 'FILTER(array, include, [if_empty])', SORT: 'SORT(array, [sort_index], [sort_order], [by_col])', UNIQUE: 'UNIQUE(array)', SEQUENCE: 'SEQUENCE(rows, [columns], [start], [step])' };
})(window);
