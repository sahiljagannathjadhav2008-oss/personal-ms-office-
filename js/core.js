/* Office Web - core utilities: DOM helpers, zip, dialogs, popups, ribbon, router
   Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const OA = (global.OA = global.OA || {});
  OA.DEV = 'Sahil Jadhav';
  OA.apps = {};

  /* ------------------------------------------------------------------ DOM */
  OA.$ = (s, r) => (r || document).querySelector(s);
  OA.$$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  OA.esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  OA.el = function (tag, attrs) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      const v = attrs[k];
      if (v == null || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k === 'html') e.innerHTML = v;
      else if (k === 'text') e.textContent = v;
      else if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? '' : v);
    }
    const add = k => { if (k == null || k === false) return; if (Array.isArray(k)) k.forEach(add); else e.appendChild(k.nodeType ? k : document.createTextNode(k)); };
    for (let i = 2; i < arguments.length; i++) add(arguments[i]);
    return e;
  };
  OA.debounce = (fn, ms) => { let t; return function () { clearTimeout(t); const a = arguments, c = this; t = setTimeout(() => fn.apply(c, a), ms); }; };
  OA.uid = () => Math.random().toString(36).slice(2, 9);
  OA.clamp = (v, a, b) => Math.min(b, Math.max(a, v));

  OA.colorToHex = function (c) {
    if (!c) return null;
    if (!(global.CSS && CSS.supports('color', c))) return null;
    const ctx = OA._cx || (OA._cx = document.createElement('canvas').getContext('2d'));
    ctx.fillStyle = '#000'; ctx.fillStyle = c;
    const v = ctx.fillStyle;
    if (v[0] === '#') return v.slice(1).toUpperCase();
    const m = v.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/);
    if (!m) return null;
    if (m[4] !== undefined && +m[4] === 0) return null;
    return [m[1], m[2], m[3]].map(n => (+n).toString(16).padStart(2, '0')).join('').toUpperCase();
  };

  OA.sanitizeHTML = function (html) {
    const doc = new DOMParser().parseFromString('<body>' + html, 'text/html');
    doc.querySelectorAll('script,iframe,object,embed,link,meta,style,form,input,button,textarea,select,base,noscript').forEach(n => n.remove());
    doc.body.querySelectorAll('*').forEach(n => {
      Array.from(n.attributes).forEach(a => {
        const nm = a.name.toLowerCase(), val = a.value.trim().toLowerCase();
        if (nm.startsWith('on')) n.removeAttribute(a.name);
        else if ((nm === 'href' || nm === 'src' || nm === 'xlink:href') && /^(javascript|vbscript):/.test(val)) n.removeAttribute(a.name);
        else if (nm === 'src' && /^data:(?!image\/)/.test(val)) n.removeAttribute(a.name);
      });
    });
    return doc.body.innerHTML;
  };

  /* ------------------------------------------------------------------ Icons */
  OA.icons = {
    undo: '<path d="M3 7v6h6"/><path d="M21 17a9 9 0 0 0-15-6.7L3 13"/>',
    redo: '<path d="M21 7v6h-6"/><path d="M3 17a9 9 0 0 1 15-6.7L21 13"/>',
    copy: '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
    cut: '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><line x1="20" y1="4" x2="8.12" y2="15.88"/><line x1="14.47" y1="14.48" x2="20" y2="20"/><line x1="8.12" y1="8.12" x2="12" y2="12"/>',
    paste: '<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>',
    alignLeft: '<line x1="17" y1="10" x2="3" y2="10"/><line x1="21" y1="6" x2="3" y2="6"/><line x1="21" y1="14" x2="3" y2="14"/><line x1="17" y1="18" x2="3" y2="18"/>',
    alignCenter: '<line x1="18" y1="10" x2="6" y2="10"/><line x1="21" y1="6" x2="3" y2="6"/><line x1="21" y1="14" x2="3" y2="14"/><line x1="18" y1="18" x2="6" y2="18"/>',
    alignRight: '<line x1="21" y1="10" x2="7" y2="10"/><line x1="21" y1="6" x2="3" y2="6"/><line x1="21" y1="14" x2="3" y2="14"/><line x1="21" y1="18" x2="7" y2="18"/>',
    justify: '<line x1="21" y1="10" x2="3" y2="10"/><line x1="21" y1="6" x2="3" y2="6"/><line x1="21" y1="14" x2="3" y2="14"/><line x1="21" y1="18" x2="3" y2="18"/>',
    list: '<line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><circle cx="3.5" cy="6" r="1"/><circle cx="3.5" cy="12" r="1"/><circle cx="3.5" cy="18" r="1"/>',
    listNum: '<line x1="10" y1="6" x2="21" y2="6"/><line x1="10" y1="12" x2="21" y2="12"/><line x1="10" y1="18" x2="21" y2="18"/><text x="1" y="8" font-size="7" fill="currentColor" stroke="none">1</text><text x="1" y="14" font-size="7" fill="currentColor" stroke="none">2</text><text x="1" y="20" font-size="7" fill="currentColor" stroke="none">3</text>',
    indent: '<polyline points="3 8 7 12 3 16"/><line x1="21" y1="12" x2="11" y2="12"/><line x1="21" y1="6" x2="11" y2="6"/><line x1="21" y1="18" x2="11" y2="18"/>',
    outdent: '<polyline points="7 8 3 12 7 16"/><line x1="21" y1="12" x2="11" y2="12"/><line x1="21" y1="6" x2="11" y2="6"/><line x1="21" y1="18" x2="11" y2="18"/>',
    link: '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>',
    image: '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/>',
    table: '<rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/>',
    download: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>',
    folder: '<path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>',
    file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>',
    print: '<polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>',
    search: '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>',
    plus: '<line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/>',
    minus: '<line x1="5" y1="12" x2="19" y2="12"/>',
    trash: '<polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6M14 11v6"/><path d="M9 6V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"/>',
    chart: '<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>',
    play: '<polygon points="5 3 19 12 5 21 5 3"/>',
    home: '<path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>',
    square: '<rect x="3" y="3" width="18" height="18" rx="2"/>',
    circle: '<circle cx="12" cy="12" r="9"/>',
    triangle: '<path d="M12 3L22 20H2z"/>',
    line: '<line x1="4" y1="20" x2="20" y2="4"/>',
    arrow: '<line x1="5" y1="12" x2="19" y2="12"/><polyline points="12 5 19 12 12 19"/>',
    textbox: '<polyline points="4 7 4 4 20 4 20 7"/><line x1="9" y1="20" x2="15" y2="20"/><line x1="12" y1="4" x2="12" y2="20"/>',
    layers: '<polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/>',
    eraser: '<path d="M20 20H7L3 16a1 1 0 0 1 0-1.4L14.6 3a1 1 0 0 1 1.4 0l5 5a1 1 0 0 1 0 1.4L11 20"/>',
    menu: '<line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/>',
    border: '<rect x="3" y="3" width="18" height="18"/><line x1="12" y1="3" x2="12" y2="21"/><line x1="3" y1="12" x2="21" y2="12"/>',
    merge: '<rect x="3" y="6" width="18" height="12"/><polyline points="8 10 5 12 8 14"/><polyline points="16 10 19 12 16 14"/><line x1="5" y1="12" x2="19" y2="12"/>',
    sort: '<line x1="4" y1="6" x2="13" y2="6"/><line x1="4" y1="12" x2="10" y2="12"/><line x1="4" y1="18" x2="7" y2="18"/><polyline points="17 4 17 20"/><polyline points="14 17 17 20 20 17"/>',
    filter: '<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>',
    pen: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',
    bucket: '<path d="M19 11l-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2a2 2 0 0 0 2.8 0L19 11z"/><path d="M5 2l5 5"/><path d="M21 15s2 2.3 2 3.5a2 2 0 0 1-4 0C19 17.3 21 15 21 15z"/>',
    calendar: '<rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>',
    pagebreak: '<path d="M6 9V3h12v6"/><path d="M6 15v6h12v-6"/><line x1="2" y1="12" x2="5" y2="12"/><line x1="9" y1="12" x2="11" y2="12"/><line x1="13" y1="12" x2="15" y2="12"/><line x1="19" y1="12" x2="22" y2="12"/>',
    eye: '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>',
    moon: '<path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/>',
    chevronDown: '<polyline points="6 9 12 15 18 9"/>',
    chevronUp: '<polyline points="18 15 12 9 6 15"/>',
    chevronLeft: '<polyline points="15 18 9 12 15 6"/>',
    chevronRight: '<polyline points="9 18 15 12 9 6"/>',
    x: '<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>',
    check: '<polyline points="20 6 9 17 4 12"/>',
    arrowLeft: '<line x1="19" y1="12" x2="5" y2="12"/><polyline points="12 19 5 12 12 5"/>',
    ruler: '<rect x="2" y="8" width="20" height="8" rx="1"/><line x1="6" y1="8" x2="6" y2="12"/><line x1="10" y1="8" x2="10" y2="12"/><line x1="14" y1="8" x2="14" y2="12"/><line x1="18" y1="8" x2="18" y2="12"/>',
    columns: '<rect x="3" y="3" width="18" height="18" rx="1"/><line x1="12" y1="3" x2="12" y2="21"/>',
    orientation: '<rect x="6" y="3" width="12" height="18" rx="1"/>',
    margin: '<rect x="3" y="3" width="18" height="18" rx="1"/><rect x="7" y="7" width="10" height="10" stroke-dasharray="2 2"/>',
    comment: '<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>',
    check2: '<path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/>',
    zoomIn: '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="11" y1="8" x2="11" y2="14"/><line x1="8" y1="11" x2="14" y2="11"/>',
    zoomOut: '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/><line x1="8" y1="11" x2="14" y2="11"/>',
    expand: '<polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/>',
    freeze: '<rect x="3" y="3" width="18" height="18" rx="1"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="9" y1="3" x2="9" y2="21"/>',
    sigma: '<path d="M18 7V4H6l6 8-6 8h12v-3"/>',
    fx: '<text x="3" y="17" font-size="13" font-style="italic" font-family="serif" fill="currentColor" stroke="none">fx</text>',
    symbol: '<text x="4" y="18" font-size="16" fill="currentColor" stroke="none">Ω</text>',
    smile: '<circle cx="12" cy="12" r="10"/><path d="M8 14s1.5 2 4 2 4-2 4-2"/><line x1="9" y1="9" x2="9.01" y2="9"/><line x1="15" y1="9" x2="15.01" y2="9"/>',
    hr: '<line x1="3" y1="12" x2="21" y2="12"/>',
    clipboard: '<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>',
    brush: '<path d="M9 11l-6 6a2.8 2.8 0 0 0 4 4l6-6"/><path d="M20 4l-9 9"/><path d="M14 8l3 3"/>',
    rotate: '<polyline points="23 4 23 10 17 10"/><path d="M20.5 15a9 9 0 1 1-2.1-9.4L23 10"/>',
    notes: '<path d="M4 4h16v13H8l-4 4z"/>',
    present: '<rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>',
    slide: '<rect x="3" y="5" width="18" height="14" rx="1"/>',
    star: '<polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/>',
    hex: '<polygon points="12 2 21 7 21 17 12 22 3 17 3 7"/>',
    diamond: '<polygon points="12 2 22 12 12 22 2 12"/>',
    lock: '<rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>',
    info: '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>',
    group: '<rect x="3" y="3" width="8" height="8"/><rect x="13" y="13" width="8" height="8"/>',
    flipH: '<line x1="12" y1="3" x2="12" y2="21"/><polygon points="7 7 7 17 2 12"/><polygon points="17 7 17 17 22 12"/>'
  };
  Object.assign(OA.icons, { users: OA.icons.group, book: OA.icons.notes, globe: OA.icons.circle, mail: OA.icons.file, edit: OA.icons.pen, quote: OA.icons.comment, bookmark: OA.icons.star, refresh: OA.icons.rotate, stamp: OA.icons.square, palette: OA.icons.bucket });
  OA.icon = (n, s) => `<svg class="ic" viewBox="0 0 24 24" width="${s || 18}" height="${s || 18}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${OA.icons[n] || ''}</svg>`;

  /* ------------------------------------------------------------------ ZIP */
  const crcT = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  OA.crc32 = u8 => { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = crcT[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const enc = new TextEncoder(), dec = new TextDecoder();
  OA.utf8 = s => enc.encode(s);
  OA.fromUtf8 = u => dec.decode(u);

  async function collect(stream) {
    const reader = stream.getReader(), parts = []; let len = 0;
    for (;;) { const { done, value } = await reader.read(); if (done) break; parts.push(value); len += value.length; }
    const out = new Uint8Array(len); let o = 0; parts.forEach(p => { out.set(p, o); o += p.length; });
    return out;
  }
  async function deflateRaw(u8) {
    const cs = new CompressionStream('deflate-raw'); const w = cs.writable.getWriter(); w.write(u8); w.close();
    return collect(cs.readable);
  }
  async function inflateRaw(u8) {
    const ds = new DecompressionStream('deflate-raw'); const w = ds.writable.getWriter(); w.write(u8); w.close();
    return collect(ds.readable);
  }
  // files: array of [name, string|Uint8Array] (order preserved). returns Uint8Array
  OA.zip = async function (files, opts) {
    const useDeflate = typeof CompressionStream !== 'undefined' && !(opts && opts.store);
    const chunks = [], central = []; let offset = 0;
    const d = new Date();
    const dosTime = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const dosDate = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const [name, data] of files) {
      const raw = typeof data === 'string' ? enc.encode(data) : data;
      const nameB = enc.encode(name), crc = OA.crc32(raw);
      let body = raw, method = 0;
      if (useDeflate && raw.length > 64 && !/\.(png|jpe?g|gif|webp)$/i.test(name)) {
        const c = await deflateRaw(raw);
        if (c.length < raw.length) { body = c; method = 8; }
      }
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, method, true);
      lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true); lh.setUint32(14, crc, true);
      lh.setUint32(18, body.length, true); lh.setUint32(22, raw.length, true); lh.setUint16(26, nameB.length, true); lh.setUint16(28, 0, true);
      chunks.push(new Uint8Array(lh.buffer), nameB, body);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, method, true);
      ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true);
      ch.setUint32(20, body.length, true); ch.setUint32(24, raw.length, true); ch.setUint16(28, nameB.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), nameB);
      offset += 30 + nameB.length + body.length;
    }
    let cdSize = 0; central.forEach(c => cdSize += c.length);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = chunks.concat(central, [new Uint8Array(end.buffer)]);
    let total = 0; all.forEach(c => total += c.length);
    const out = new Uint8Array(total); let o = 0; all.forEach(c => { out.set(c, o); o += c.length; });
    return out;
  };
  // returns {name: Uint8Array}
  OA.unzip = async function (buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let e = u8.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('Not a valid ZIP/Office file');
    const count = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
    const out = {};
    for (let i = 0; i < count; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
      const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), lo = dv.getUint32(p + 42, true);
      const name = dec.decode(u8.subarray(p + 46, p + 46 + nl));
      p += 46 + nl + xl + cl;
      if (name.endsWith('/')) continue;
      const lnl = dv.getUint16(lo + 26, true), lxl = dv.getUint16(lo + 28, true);
      const start = lo + 30 + lnl + lxl, data = u8.subarray(start, start + csize);
      out[name] = method === 0 ? data.slice() : method === 8 ? await inflateRaw(data) : null;
    }
    return out;
  };
  OA.xmlDoc = (u8or) => new DOMParser().parseFromString(typeof u8or === 'string' ? u8or : OA.fromUtf8(u8or), 'application/xml');

  /* ------------------------------------------------------------------ Files */
  OA.download = function (data, name, mime) {
    const blob = data instanceof Blob ? data : new Blob([data], { type: mime || 'application/octet-stream' });
    const a = OA.el('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
  };
  OA.pickFile = function (accept, multiple) {
    return new Promise(res => {
      const inp = OA.el('input', { type: 'file', accept: accept || '', style: { display: 'none' } });
      if (multiple) inp.multiple = true;
      inp.addEventListener('change', () => { res(multiple ? Array.from(inp.files) : inp.files[0] || null); inp.remove(); });
      inp.addEventListener('cancel', () => { res(null); inp.remove(); });
      document.body.appendChild(inp); inp.click();
    });
  };
  OA.readDataURL = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  OA.dataURLtoBytes = function (url) {
    const m = /^data:([^;,]+)?(;base64)?,(.*)$/s.exec(url);
    if (!m) return null;
    const bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]);
    const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return { mime: m[1] || 'application/octet-stream', bytes: u };
  };
  OA.bytesToDataURL = function (bytes, mime) {
    let s = ''; const CH = 0x8000;
    for (let i = 0; i < bytes.length; i += CH) s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
    return 'data:' + mime + ';base64,' + btoa(s);
  };
  OA.baseName = n => (n || '').replace(/\.[^.]+$/, '');
  OA.ext = n => ((n || '').match(/\.([^.]+)$/) || [, ''])[1].toLowerCase();

  /* ------------------------------------------------------------------ Toast / dialogs / popups */
  OA.toast = function (msg, ms) {
    let box = OA.$('#toasts'); if (!box) { box = OA.el('div', { id: 'toasts' }); document.body.appendChild(box); }
    const t = OA.el('div', { class: 'toast', text: msg }); box.appendChild(t);
    setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 300); }, ms || 2600);
  };

  OA.dialog = function (o) {
    const back = OA.el('div', { class: 'dlg-back' });
    const body = OA.el('div', { class: 'dlg-body' });
    if (typeof o.content === 'string') body.innerHTML = o.content; else if (o.content) body.appendChild(o.content);
    const foot = OA.el('div', { class: 'dlg-foot' });
    const box = OA.el('div', { class: 'dlg ' + (o.cls || ''), role: 'dialog', 'aria-modal': 'true', style: o.width ? { width: o.width + 'px' } : null },
      OA.el('div', { class: 'dlg-head' }, OA.el('span', { text: o.title || '' }), OA.el('button', { class: 'dlg-x', type: 'button', html: OA.icon('x', 16), 'aria-label': 'Close', onclick: () => api.close(false) })),
      body, (o.buttons && o.buttons.length) ? foot : null);
    back.appendChild(box);
    const api = {
      el: box, body,
      close(val) { if (api.closed) return; api.closed = true; document.removeEventListener('keydown', onKey, true); back.remove(); o.onClose && o.onClose(val); }
    };
    (o.buttons || []).forEach(b => {
      const btn = OA.el('button', { type: 'button', class: 'btn' + (b.primary ? ' primary' : ''), text: b.label });
      btn.addEventListener('click', e => { const r = b.onClick ? b.onClick(e, api) : undefined; if (r !== false) api.close(b.value !== undefined ? b.value : true); });
      foot.appendChild(btn); b._el = btn;
    });
    const onKey = e => {
      if (!document.body.contains(back)) return;
      const topmost = OA.$$('.dlg-back').pop() === back; if (!topmost) return;
      if (e.key === 'Escape') { e.stopPropagation(); api.close(false); }
      else if (e.key === 'Enter' && !(e.target && /^(TEXTAREA|BUTTON|SELECT)$/.test(e.target.tagName)) && o.buttons) {
        const p = o.buttons.find(b => b.primary); if (p) { e.preventDefault(); p._el.click(); }
      }
    };
    document.addEventListener('keydown', onKey, true);
    document.body.appendChild(back);
    OA.closePopup();
    setTimeout(() => { const f = box.querySelector('input:not([type=hidden]),select,textarea') || box.querySelector('.btn.primary'); f && f.focus(); if (f && f.select && f.tagName === 'INPUT' && f.type === 'text') f.select(); }, 30);
    return api;
  };
  OA.alert = (msg, title) => new Promise(res => OA.dialog({ title: title || 'Office Web', content: OA.el('div', { class: 'dlg-msg', html: msg }), buttons: [{ label: 'OK', primary: true }], onClose: () => res() }));
  OA.confirm = (msg, o) => { o = o || {}; return new Promise(res => OA.dialog({ title: o.title || 'Office Web', content: OA.el('div', { class: 'dlg-msg', html: msg }), buttons: [{ label: o.cancel || 'Cancel', value: false }, { label: o.ok || 'OK', primary: true, value: true }], onClose: v => res(v === true) })); };
  OA.prompt = (label, value, o) => {
    o = o || {};
    return new Promise(res => {
      const inp = OA.el('input', { type: o.type || 'text', value: value == null ? '' : value, placeholder: o.placeholder || '', class: 'inp wide' });
      let result = null;
      OA.dialog({
        title: o.title || 'Office Web', content: OA.el('div', null, OA.el('label', { class: 'dlg-lbl', text: label }), inp),
        buttons: [{ label: 'Cancel', value: false }, { label: 'OK', primary: true, onClick: () => { result = inp.value; } }],
        onClose: () => res(result)
      });
    });
  };
  // generic form dialog: fields [{id,label,type:'text|number|select|checkbox|color',value,options,min,max,step}] -> Promise<values|null>
  OA.form = function (title, fields, o) {
    o = o || {};
    return new Promise(res => {
      const wrap = OA.el('div', { class: 'form-grid' }); const inputs = {};
      if (o.intro) wrap.appendChild(OA.el('div', { class: 'form-intro', html: o.intro }));
      fields.forEach(f => {
        let inp;
        if (f.type === 'select') { inp = OA.el('select', { class: 'inp' }); f.options.forEach(op => { const v = Array.isArray(op) ? op[0] : op, l = Array.isArray(op) ? op[1] : op; const oe = OA.el('option', { value: v, text: l }); if (String(v) === String(f.value)) oe.selected = true; inp.appendChild(oe); }); }
        else if (f.type === 'checkbox') { inp = OA.el('input', { type: 'checkbox' }); inp.checked = !!f.value; }
        else if (f.type === 'textarea') { inp = OA.el('textarea', { class: 'inp', rows: f.rows || 4 }); inp.value = f.value == null ? '' : f.value; }
        else { inp = OA.el('input', { type: f.type || 'text', class: 'inp', value: f.value == null ? '' : f.value, min: f.min, max: f.max, step: f.step, placeholder: f.placeholder }); }
        inputs[f.id] = inp;
        wrap.appendChild(OA.el('label', { class: 'form-row' + (f.type === 'checkbox' ? ' chk' : '') }, f.type === 'checkbox' ? [inp, OA.el('span', { text: f.label })] : [OA.el('span', { text: f.label }), inp]));
      });
      let result = null;
      OA.dialog({
        title, content: wrap, width: o.width || 380,
        buttons: [{ label: 'Cancel', value: false }, {
          label: o.ok || 'OK', primary: true, onClick: () => {
            const out = {};
            fields.forEach(f => { const i = inputs[f.id]; out[f.id] = f.type === 'checkbox' ? i.checked : f.type === 'number' ? (i.value === '' ? null : +i.value) : i.value; });
            if (o.validate) { const msg = o.validate(out); if (msg) { OA.toast(msg); return false; } }
            result = out;
          }
        }],
        onClose: () => res(result)
      });
    });
  };

  let curPopup = null, lastClosed = { a: null, t: 0 };
  OA.closePopup = () => { if (curPopup) curPopup.close(); };
  OA.popup = function (anchor, content, opts) {
    opts = opts || {};
    if (anchor && anchor.nodeType && lastClosed.a === anchor && Date.now() - lastClosed.t < 250) { lastClosed = { a: null, t: 0 }; return null; }
    OA.closePopup();
    const box = OA.el('div', { class: 'popup ' + (opts.cls || '') });
    box.appendChild(content); document.body.appendChild(box);
    const place = () => {
      const r = anchor.nodeType ? anchor.getBoundingClientRect() : anchor;
      box.style.left = '0px'; box.style.top = '0px'; box.style.maxHeight = (innerHeight - 8) + 'px';
      const bw = box.offsetWidth, bh = box.offsetHeight;
      let x = r.left, y = anchor.nodeType ? r.bottom + 2 : r.top;
      if (opts.right) x = r.right - bw;
      if (x + bw > innerWidth - 4) x = Math.max(4, innerWidth - bw - 4);
      if (x < 4) x = 4;
      if (y + bh > innerHeight - 4) { if (anchor.nodeType && r.top - bh - 2 > 4) y = r.top - bh - 2; else y = Math.max(4, innerHeight - bh - 4); }
      box.style.left = x + 'px'; box.style.top = y + 'px';
    };
    place();
    const onDown = e => { if (!box.contains(e.target)) close(); };
    const onKey = e => { if (e.key === 'Escape') close(); };
    const onScroll = e => { if (!box.contains(e.target) && !opts.keepOnScroll) close(); };
    function close() {
      if (!curPopup || curPopup.box !== box) return;
      curPopup = null; lastClosed = { a: anchor, t: Date.now() };
      document.removeEventListener('pointerdown', onDown, true); document.removeEventListener('keydown', onKey, true);
      window.removeEventListener('resize', close); document.removeEventListener('scroll', onScroll, true);
      box.remove(); opts.onClose && opts.onClose();
    }
    curPopup = { box, close };
    setTimeout(() => { document.addEventListener('pointerdown', onDown, true); document.addEventListener('keydown', onKey, true); window.addEventListener('resize', close); document.addEventListener('scroll', onScroll, true); }, 0);
    return { close, box, place };
  };
  OA.menu = function (anchor, items, opts) {
    const list = OA.el('div', { class: 'menu' });
    (typeof items === 'function' ? items() : items).forEach(it => {
      if (!it) return;
      if (it === '-' || it.sep) { list.appendChild(OA.el('div', { class: 'msep' })); return; }
      if (it.header) { list.appendChild(OA.el('div', { class: 'mhead', text: it.header })); return; }
      const row = OA.el('button', { class: 'mi' + (it.disabled ? ' dis' : ''), type: 'button' });
      row.innerHTML = `<span class="mic">${it.icon ? OA.icon(it.icon, 16) : (it.checked ? '✓' : '')}</span><span class="mlb"${it.style ? ` style="${OA.esc(it.style)}"` : ''}>${OA.esc(it.label)}</span>${it.hint ? `<span class="mht">${OA.esc(it.hint)}</span>` : ''}`;
      row.addEventListener('mousedown', e => e.preventDefault());
      row.addEventListener('click', () => { if (it.disabled) return; OA.closePopup(); it.onClick && it.onClick(); });
      list.appendChild(row);
    });
    return OA.popup(anchor, list, opts);
  };

  /* Colour palette */
  const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  const STD = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];
  function mix(hex, t) { // t>0 lighten toward white, t<0 darken toward black
    const n = parseInt(hex, 16); let r = n >> 16, g = (n >> 8) & 255, b = n & 255;
    const f = (c) => Math.round(t >= 0 ? c + (255 - c) * t : c * (1 + t));
    return [f(r), f(g), f(b)].map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  OA.mixColor = mix;
  OA.colorPicker = function (anchor, o) {
    o = o || {};
    const wrap = OA.el('div', { class: 'cpick' });
    const pick = c => { OA.closePopup(); o.onPick && o.onPick(c); };
    if (o.noColor !== false) wrap.appendChild(OA.el('button', { type: 'button', class: 'cp-none', text: o.noColor || 'No Color', onclick: () => pick(null) }));
    wrap.appendChild(OA.el('div', { class: 'cp-h', text: 'Theme Colors' }));
    const g = OA.el('div', { class: 'cp-grid' });
    THEME.forEach(c => g.appendChild(sw(c)));
    [0.8, 0.6, 0.4, -0.25, -0.5].forEach(t => THEME.forEach(c => g.appendChild(sw(mix(c, t)))));
    wrap.appendChild(g);
    wrap.appendChild(OA.el('div', { class: 'cp-h', text: 'Standard Colors' }));
    const g2 = OA.el('div', { class: 'cp-grid' }); STD.forEach(c => g2.appendChild(sw(c))); wrap.appendChild(g2);
    const ci = OA.el('input', { type: 'color', value: '#' + (o.current || '4472C4'), class: 'cp-custom' });
    ci.addEventListener('change', () => pick(ci.value.slice(1).toUpperCase()));
    wrap.appendChild(OA.el('label', { class: 'cp-more' }, 'More Colors… ', ci));
    function sw(c) { return OA.el('button', { type: 'button', class: 'cp-sw', title: '#' + c, style: { background: '#' + c }, onclick: () => pick(c) }); }
    wrap.addEventListener('mousedown', e => { if (e.target.tagName !== 'INPUT') e.preventDefault(); });
    return OA.popup(anchor, wrap);
  };

  /* ------------------------------------------------------------------ Ribbon */
  // spec: {tabs:[{id,label,hidden,groups:[{label,items:[...]}]}], file:{label,onClick}}
  OA.buildRibbon = function (host, spec) {
    const ctl = {}, tabEls = {}, panels = {};
    const tabbar = OA.el('div', { class: 'rtabs' });
    const body = OA.el('div', { class: 'rbody' });
    const collapseBtn = OA.el('button', { class: 'rcollapse', type: 'button', title: 'Collapse the ribbon', html: OA.icon('chevronUp', 16) });
    let active = null;
    const small = () => matchMedia('(max-width: 760px)').matches;
    let collapsed = small();
    const applyCollapsed = () => { host.classList.toggle('collapsed', collapsed); collapseBtn.innerHTML = OA.icon(collapsed ? 'chevronDown' : 'chevronUp', 16); };
    collapseBtn.addEventListener('click', () => { collapsed = !collapsed; applyCollapsed(); });
    if (spec.file) {
      const fb = OA.el('button', { class: 'rtab file', type: 'button', text: spec.file.label || 'File', onclick: spec.file.onClick });
      tabbar.appendChild(fb);
    }
    function mk(it, group) {
      if (it.type === 'sep') return OA.el('span', { class: 'rsep' });
      if (it.type === 'row') { const rl = OA.el('div', { class: 'rline' }); it.items.forEach(x => { const n = mk(x, group); if (n) rl.appendChild(n); }); return rl; }
      let node;
      if (it.type === 'select') {
        const s = OA.el('select', { class: 'rsel', title: it.title || '', style: it.width ? { width: it.width + 'px' } : null });
        it.options.forEach(op => { const v = Array.isArray(op) ? op[0] : op, l = Array.isArray(op) ? op[1] : op; s.appendChild(OA.el('option', { value: v, text: l })); });
        s.addEventListener('change', () => it.onChange && it.onChange(s.value, s));
        node = OA.el('div', { class: 'rrow' }, it.label ? OA.el('span', { class: 'rlab', text: it.label }) : null, s); ctl[it.id] = s;
      } else if (it.type === 'combo') {
        const dl = OA.el('datalist', { id: 'dl-' + it.id + '-' + (OA.uid()) });
        it.options.forEach(op => dl.appendChild(OA.el('option', { value: op })));
        const i = OA.el('input', { class: 'rsel combo', type: 'text', title: it.title || '', style: { width: (it.width || 56) + 'px' }, list: dl.id, value: it.value || '', inputmode: it.inputmode || 'decimal' });
        const apply = () => it.onChange && it.onChange(i.value, i);
        i.addEventListener('change', apply); i.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); apply(); it.onEnter && it.onEnter(); } });
        node = OA.el('div', { class: 'rrow' }, it.label ? OA.el('span', { class: 'rlab', text: it.label }) : null, i, dl); ctl[it.id] = i;
      } else if (it.type === 'number') {
        const i = OA.el('input', { class: 'rsel num', type: 'number', title: it.title || '', min: it.min, max: it.max, step: it.step || 1, value: it.value, style: { width: (it.width || 64) + 'px' } });
        i.addEventListener('change', () => it.onChange && it.onChange(+i.value, i));
        node = OA.el('div', { class: 'rrow' }, it.label ? OA.el('span', { class: 'rlab', text: it.label }) : null, i); ctl[it.id] = i;
      } else if (it.type === 'check') {
        const i = OA.el('input', { type: 'checkbox' }); i.checked = !!it.value;
        i.addEventListener('change', () => it.onChange && it.onChange(i.checked));
        node = OA.el('label', { class: 'rchk', title: it.title || '' }, i, OA.el('span', { text: it.label })); ctl[it.id] = i;
      } else if (it.type === 'color') {
        let cur = it.default || null;
        const bar = OA.el('i', { class: 'cbar', style: { background: cur ? '#' + cur : 'transparent' } });
        const main = OA.el('button', { class: 'rb sm cmain', type: 'button', title: it.title || it.label, html: (it.icon ? OA.icon(it.icon) : `<span class="gl">${it.glyph || 'A'}</span>`) });
        main.appendChild(bar);
        const car = OA.el('button', { class: 'rb sm ccar', type: 'button', title: it.title, html: '<i class="caret"></i>' });
        main.addEventListener('mousedown', e => e.preventDefault()); car.addEventListener('mousedown', e => e.preventDefault());
        main.addEventListener('click', () => it.onPick(cur));
        car.addEventListener('click', () => OA.colorPicker(car, { current: cur, noColor: it.noColor, onPick: c => { cur = c; bar.style.background = c ? '#' + c : 'transparent'; it.onPick(c); } }));
        node = OA.el('div', { class: 'rsplit' }, main, car);
        ctl[it.id] = { set(c) { cur = c; bar.style.background = c ? '#' + c : 'transparent'; }, get: () => cur, el: main };
      } else if (it.type === 'gallery') {
        node = OA.el('div', { class: 'rgal' }); ctl[it.id] = node;
        it.items.forEach(g => {
          const b = OA.el('button', { class: 'gitem', type: 'button', title: g.label, 'data-k': g.key || g.label }, OA.el('span', { class: 'gs', style: g.style || '', text: g.sample || 'AaBbCc' }), OA.el('span', { class: 'gl2', text: g.label }));
          b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => g.onClick(g));
          node.appendChild(b);
        });
      } else if (it.type === 'custom') { node = it.el; if (it.id) ctl[it.id] = node; }
      else { // button
        const b = OA.el('button', { class: 'rb ' + (it.size === 'lg' ? 'lg' : 'sm') + (it.wide ? ' wide' : ''), type: 'button', title: it.title || it.label || '', 'data-id': it.id || '' });
        b.innerHTML = (it.icon ? OA.icon(it.icon, it.size === 'lg' ? 26 : 17) : it.glyph ? `<span class="gl" style="${it.gstyle || ''}">${it.glyph}</span>` : '') + (it.label ? `<span class="lb">${OA.esc(it.label)}</span>` : '') + (it.menu ? '<i class="caret"></i>' : '');
        b.addEventListener('mousedown', e => e.preventDefault());
        b.addEventListener('click', e => { if (it.menu) OA.menu(b, it.menu); else it.onClick && it.onClick(e, b); });
        node = b; if (it.id) ctl[it.id] = b;
      }
      return node;
    }
    spec.tabs.forEach(t => {
      const tb = OA.el('button', { class: 'rtab' + (t.contextual ? ' ctx' : ''), type: 'button', text: t.label, hidden: !!t.hidden, 'data-tab': t.id });
      tb.addEventListener('click', () => {
        if (active === t.id) { if (small()) { collapsed = !collapsed; applyCollapsed(); } return; }
        if (collapsed && small()) { collapsed = false; applyCollapsed(); }
        showTab(t.id);
      });
      tabEls[t.id] = tb; tabbar.appendChild(tb);
      const p = OA.el('div', { class: 'rpanel', hidden: true });
      t.groups.forEach(g => {
        const gi = OA.el('div', { class: 'gi' }); g.items.forEach(it => { const n = mk(it, g); if (n) gi.appendChild(n); });
        p.appendChild(OA.el('div', { class: 'rgroup' + (g.cls ? ' ' + g.cls : '') }, gi, OA.el('div', { class: 'glabel', text: g.label })));
      });
      panels[t.id] = p; body.appendChild(p);
    });
    tabbar.appendChild(OA.el('span', { class: 'rfill' }));
    if (spec.right) tabbar.appendChild(spec.right);
    tabbar.appendChild(collapseBtn);
    host.appendChild(tabbar); host.appendChild(body);
    function showTab(id) {
      active = id;
      Object.keys(tabEls).forEach(k => { tabEls[k].classList.toggle('on', k === id); panels[k].hidden = k !== id; });
    }
    applyCollapsed();
    showTab(spec.active || spec.tabs[0].id);
    return {
      ctl, showTab, tabEl: id => tabEls[id],
      setTabVisible(id, v) { tabEls[id].hidden = !v; if (!v && active === id) showTab(spec.active || spec.tabs[0].id); },
      get active() { return active; },
      setOn(id, on) { const c = ctl[id]; if (c && c.classList) c.classList.toggle('on', !!on); }
    };
  };

  /* Backstage (File menu) */
  OA.backstage = function (root, o) {
    const ov = OA.el('div', { class: 'backstage', hidden: true });
    const side = OA.el('div', { class: 'bs-side' });
    const main = OA.el('div', { class: 'bs-main' });
    const back = OA.el('button', { class: 'bs-back', type: 'button', html: OA.icon('arrowLeft', 22), title: 'Back', onclick: () => api.close() });
    side.appendChild(back);
    const btns = {};
    const api = {
      open(id) { ov.hidden = false; this.show(id || o.items.find(i => i.render).id); },
      close() { ov.hidden = true; },
      show(id) {
        const it = o.items.find(i => i.id === id); if (!it) return;
        if (!it.render) { api.close(); it.action && it.action(); return; }
        Object.keys(btns).forEach(k => btns[k].classList.toggle('on', k === id));
        main.innerHTML = ''; const p = OA.el('div', { class: 'bs-panel' }); main.appendChild(p);
        p.appendChild(OA.el('h2', { text: it.title || it.label })); it.render(p, api);
      }
    };
    o.items.forEach(it => {
      if (it.sep) { side.appendChild(OA.el('div', { class: 'bs-sep' })); return; }
      const b = OA.el('button', { class: 'bs-item', type: 'button', html: (it.icon ? OA.icon(it.icon, 18) : '') + `<span>${OA.esc(it.label)}</span>`, onclick: () => api.show(it.id) });
      btns[it.id] = b; side.appendChild(b);
    });
    side.appendChild(OA.el('div', { class: 'bs-dev', text: 'Developer: ' + OA.DEV }));
    ov.appendChild(side); ov.appendChild(main); root.appendChild(ov);
    api.el = ov;
    return api;
  };
  OA.bigActions = function (items) {
    const w = OA.el('div', { class: 'bigacts' });
    items.forEach(it => w.appendChild(OA.el('button', { class: 'bigact', type: 'button', onclick: it.onClick }, OA.el('span', { class: 'bi', html: OA.icon(it.icon || 'download', 28) }), OA.el('span', { class: 'bt' }, OA.el('b', { text: it.label }), OA.el('small', { text: it.desc || '' })))));
    return w;
  };

  /* ------------------------------------------------------------------ App shell */
  OA.appShell = function (root, a) {
    root.classList.add('app-shell');
    const title = OA.el('div', { class: 'titlebar' });
    const home = OA.el('button', { class: 'tb-btn', type: 'button', title: 'Back to Office Home', html: OA.icon('home', 18), onclick: () => OA.go('home') });
    const logo = OA.el('span', { class: 'tb-logo', html: OA.appLogo(a) });
    const name = OA.el('input', { class: 'tb-name', type: 'text', value: a.defaultName, spellcheck: 'false', 'aria-label': 'File name' });
    const right = OA.el('div', { class: 'tb-right' });
    const warn = OA.el('span', { class: 'tb-note', html: OA.icon('info', 14) + '<span>Not saved – export to keep</span>', title: 'Data is not saved after you close this website. Use File → Export to save your file.' });
    title.append(home, logo, OA.el('span', { class: 'tb-app', text: a.name }), name, OA.el('span', { class: 'tb-fill' }), warn, right);
    const ribbon = OA.el('div', { class: 'ribbon' });
    const content = OA.el('div', { class: 'app-content' });
    const status = OA.el('div', { class: 'statusbar' });
    root.append(title, ribbon, content, status);
    return { title, name, right, ribbon, content, status, warn };
  };
  OA.appLogo = function (a, size) {
    return `<svg viewBox="0 0 32 32" width="${size || 22}" height="${size || 22}"><rect width="32" height="32" rx="5" fill="${a.color}"/><text x="16" y="23.5" text-anchor="middle" font-family="Segoe UI,Arial,sans-serif" font-weight="700" font-size="20" fill="#fff">${a.letter}</text></svg>`;
  };

  /* ------------------------------------------------------------------ Router / home */
  OA.registerApp = a => { OA.apps[a.id] = a; };
  let current = 'home';
  OA.current = () => current;
  function showView(id) {
    OA.closePopup();
    current = id;
    OA.$('#home').hidden = id !== 'home';
    Object.values(OA.apps).forEach(a => { if (a.root) a.root.hidden = a.id !== id; });
    let color = '#f3f3f3';
    if (id !== 'home') {
      const a = OA.apps[id];
      if (!a.root) {
        a.root = OA.el('section', { class: 'app-view app-' + id });
        a.root.style.setProperty('--accent', a.color);
        OA.$('#apps').appendChild(a.root);
        a.api = a.create(a.root);
      }
      a.root.hidden = false; a.api.onShow && a.api.onShow();
      OA.touched = true; color = a.color; document.title = a.name + ' – Office Web';
    } else document.title = 'Office Web – Word, Excel, PowerPoint';
    const tc = OA.$('meta[name=theme-color]'); if (tc) tc.setAttribute('content', color);
  }
  OA.go = function (id) { if (id === current) return; history.pushState({ oa: id }, ''); showView(id); };
  OA.open = function (id, how) { // how: undefined (blank) | {template} | {file}
    OA.go(id);
    const a = OA.apps[id];
    if (how && how.template && a.api.loadTemplate) a.api.loadTemplate(how.template);
    if (how && how.file) a.api.openFile(how.file);
  };
  OA.openFile = async function (file) {
    const e = OA.ext(file.name);
    const a = Object.values(OA.apps).find(x => x.accept.includes(e));
    if (!a) {
      if (['doc', 'xls', 'ppt'].includes(e)) return OA.alert(`<b>.${e}</b> is the old binary Office format and can't be opened here. Please open it in Office/LibreOffice and save as <b>.${e}x</b> first.`);
      return OA.alert('Unsupported file type: .' + OA.esc(e));
    }
    OA.touched = true; OA.open(a.id, { file });
  };
  OA.showExit = function (fromBack) {
    const a = OA.apps[current];
    const body = OA.el('div', { class: 'exit-msg' },
      OA.el('div', { class: 'exit-ic', html: OA.icon('info', 30) }),
      OA.el('p', null, OA.el('b', { text: 'Data will not be saved after you close this website.' })),
      OA.el('p', { text: 'Please export (save) your file before leaving, using File → Export / Save As.' }),
      OA.el('p', { class: 'hi', text: 'Website band karne ke baad data save nahi rahega, isliye file export kar lijiye.' }));
    const buttons = [{
      label: 'Stay on page', primary: true, value: 'stay', onClick: () => { if (fromBack) { history.pushState({ oa: 'home' }, ''); showView(current === 'home' ? 'home' : current); } }
    }];
    if (a && a.api && a.api.exportDefault) buttons.unshift({ label: 'Export file now', value: 'exp', onClick: () => { a.api.exportDefault(); return fromBack ? (history.pushState({ oa: current }, ''), true) : true; } });
    buttons.push({ label: 'Close anyway', value: 'leave', onClick: () => { OA.allowLeave = true; try { window.close(); } catch (e) { } if (fromBack) history.back(); else location.replace('about:blank'); } });
    OA.dialog({ title: 'Close Office Web?', content: body, buttons, cls: 'exit-dlg', width: 440 });
  };
  OA.start = function () {
    history.replaceState({ oa: 'exit' }, '');
    history.pushState({ oa: 'home' }, '');
    window.addEventListener('popstate', e => {
      const s = e.state;
      if (s && s.oa === 'exit') { if (OA.allowLeave) return; showView('home'); OA.showExit(true); }
      else if (s && s.oa) showView(s.oa);
    });
    window.addEventListener('beforeunload', e => { if (OA.touched && !OA.allowLeave) { e.preventDefault(); e.returnValue = 'Data will not be saved after closing this website. Please export your file first.'; return e.returnValue; } });
    ['dragover', 'drop'].forEach(ev => window.addEventListener(ev, e => {
      if (!e.dataTransfer || !Array.from(e.dataTransfer.types || []).includes('Files')) return;
      if (e.target.closest && e.target.closest('[data-own-drop]')) return;
      e.preventDefault();
      if (ev === 'drop' && e.dataTransfer.files[0]) OA.openFile(e.dataTransfer.files[0]);
    }));
    renderHome();
    showView('home');
    if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('sw.js').catch(() => { });
  };
  function renderHome() {
    const h = OA.$('#home'); h.innerHTML = '';
    const hour = new Date().getHours();
    const greet = hour < 12 ? 'Good morning' : hour < 17 ? 'Good afternoon' : 'Good evening';
    const top = OA.el('header', { class: 'home-top' },
      OA.el('div', { class: 'home-brand', html: `<svg viewBox="0 0 32 32" width="28" height="28"><rect width="32" height="32" rx="7" fill="#d83b01"/><path d="M9 7l11-3v24l-11-3z" fill="#fff"/></svg><span>Office Web</span>` }),
      OA.el('span', { class: 'tb-fill' }),
      OA.el('button', { class: 'btn ghost', type: 'button', text: 'Open file…', onclick: async () => { const f = await OA.pickFile('.docx,.xlsx,.csv,.tsv,.pptx,.txt,.html,.htm'); if (f) OA.openFile(f); } }),
      OA.el('button', { class: 'btn ghost', type: 'button', text: 'Exit', onclick: () => OA.showExit(false) }));
    const hero = OA.el('div', { class: 'home-hero' }, OA.el('h1', { text: greet }), OA.el('p', { text: 'Create documents, spreadsheets and presentations – right in your browser, on phone, tablet or PC.' }));
    const note = OA.el('div', { class: 'home-note', html: OA.icon('info', 16) + '<span><b>Your files are not stored on this site.</b> After you close the website the data is gone – use <b>File → Export</b> to save .docx / .xlsx / .pptx to your device.</span>' });
    const grid = OA.el('div', { class: 'tiles' });
    Object.values(OA.apps).forEach(a => {
      const card = OA.el('div', { class: 'tile', style: { '--accent': a.color } });
      const main = OA.el('button', { class: 'tile-main', type: 'button', onclick: () => OA.open(a.id), 'aria-label': 'New blank ' + a.name },
        OA.el('span', { class: 'tile-ic', html: OA.appLogo(a, 56) }),
        OA.el('span', { class: 'tile-tx' }, OA.el('b', { text: a.name }), OA.el('small', { text: a.desc })));
      const row = OA.el('div', { class: 'tile-row' }, OA.el('button', { class: 'chip primary', type: 'button', text: '+ Blank', onclick: () => OA.open(a.id) }));
      (a.templates || []).forEach(t => row.appendChild(OA.el('button', { class: 'chip', type: 'button', text: t.label, title: t.desc || '', onclick: () => OA.open(a.id, { template: t.id }) })));
      row.appendChild(OA.el('button', { class: 'chip', type: 'button', text: 'Open…', onclick: async () => { const f = await OA.pickFile(a.accept.map(x => '.' + x).join(',')); if (f) OA.openFile(f); } }));
      card.append(main, row); grid.appendChild(card);
    });
    const foot = OA.el('footer', { class: 'home-foot' }, OA.el('span', null, 'Developer: ', OA.el('b', { text: OA.DEV })));
    h.append(top, OA.el('main', { class: 'home-main' }, hero, note, grid), foot);
  }
})(typeof window !== 'undefined' ? window : globalThis);
