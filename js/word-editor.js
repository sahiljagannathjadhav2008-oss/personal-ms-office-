/* Word editor core (contenteditable based) — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA;
  const BLOCKQ = 'p,h1,h2,h3,h4,h5,h6,li,blockquote,div,pre,td,th';
  const NESTQ = 'p,h1,h2,h3,h4,h5,h6,div,blockquote,pre,table';
  const FONTS = ['Calibri', 'Calibri Light', 'Arial', 'Arial Black', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Courier New', 'Garamond', 'Georgia', 'Impact', 'Lucida Console', 'Palatino Linotype', 'Segoe UI', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Nirmala UI', 'Mangal', 'Noto Sans Devanagari'];
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const TW = 15; // twips per px

  OA.wordFonts = FONTS; OA.wordSizes = SIZES;

  OA.createWordEditor = function (host, cb) {
    cb = cb || {};
    const el = OA.el;
    /* ---------------------------------------------------------------- DOM */
    const nav = el('div', { class: 'w-nav', hidden: true });
    const navTitle = el('span', { text: 'Navigation' });
    const navClose = el('button', { type: 'button', html: OA.icon('x', 16), title: 'Close', onclick: () => ed.hideNav() });
    const findBox = el('div', { class: 'w-find', hidden: true });
    const navBody = el('div', { class: 'w-nav-b' });
    nav.append(el('div', { class: 'w-nav-h' }, navTitle, navClose), findBox, navBody);
    const scroll = el('div', { class: 'w-scroll' });
    const stage = el('div', { class: 'w-stage' });
    const page = el('div', { class: 'w-page', contenteditable: 'true', spellcheck: 'true', 'data-ph': 'Start typing…' });
    const hfHost = el('div', { class: 'w-hf', 'data-noexport': '1', style: { position: 'absolute', left: 0, top: 0, right: 0, pointerEvents: 'none' } });
    stage.append(page, hfHost); scroll.appendChild(stage);
    const cmtPane = el('div', { class: 'w-cmts', hidden: true });
    const ctxbar = el('div', { class: 'w-ctxbar' });
    const linkb = el('div', { class: 'w-linkb' });
    host.append(nav, scroll, cmtPane); document.body.append(ctxbar, linkb);
    page.style.whiteSpace = 'pre-wrap'; page.style.tabSize = '36pt';

    const layout = { name: 'A4', w: 11906, h: 16838, landscape: false, margins: { t: 1440, r: 1440, b: 1440, l: 1440 }, cols: 1, header: '', footer: '', pageNum: 'none', pageColor: null, view: 'print' };
    let zoom = 1, saved = null, fp = null, readOnly = false;

    const ed = { page, layout, scroll, stage, nav, FONTS, SIZES };

    /* ---------------------------------------------------------------- selection helpers */
    const gs = () => window.getSelection();
    const inPage = n => n && (n === page || page.contains(n));
    function curRange() { const s = gs(); return s.rangeCount && inPage(s.getRangeAt(0).commonAncestorContainer) ? s.getRangeAt(0) : null; }
    document.addEventListener('selectionchange', () => { const r = curRange(); if (r) saved = r.cloneRange(); scheduleState(); });
    function restore() {
      const wasFocused = document.activeElement === page, sv = saved && saved.cloneRange();
      if (!wasFocused) { page.focus({ preventScroll: true }); if (sv && inPage(sv.commonAncestorContainer)) { const s0 = gs(); s0.removeAllRanges(); s0.addRange(sv); saved = sv.cloneRange(); } }
      if (saved && inPage(saved.commonAncestorContainer)) { const s = gs(); const cur = s.rangeCount ? s.getRangeAt(0) : null; if (!cur || !inPage(cur.commonAncestorContainer)) { s.removeAllRanges(); s.addRange(saved); } }
      else { const s = gs(); if (!s.rangeCount || !inPage(s.getRangeAt(0).commonAncestorContainer)) { const r = document.createRange(); r.selectNodeContents(page); r.collapse(false); s.removeAllRanges(); s.addRange(r); } }
    }
    ed.restore = restore;
    function setRange(r) { const s = gs(); s.removeAllRanges(); s.addRange(r); saved = r.cloneRange(); }
    function elOf(n) { return n && n.nodeType === 3 ? n.parentNode : n; }
    function closestBlock(n) { n = elOf(n); while (n && n !== page) { if (n.matches && n.matches(BLOCKQ) && n.tagName !== 'DIV') return n; if (n.tagName === 'DIV' && n.parentNode === page) return n; n = n.parentNode; } return null; }
    function selBlocks() {
      const r = curRange(); if (!r) return [];
      if (r.collapsed) { const b = closestBlock(r.startContainer); return b ? [b] : []; }
      const out = [];
      page.querySelectorAll(BLOCKQ).forEach(b => {
        if (b.classList.contains('pb') || b.closest('.pb')) return;
        if (!r.intersectsNode(b)) return;
        if (b.querySelector(NESTQ)) return;
        out.push(b);
      });
      if (!out.length) { const b = closestBlock(r.startContainer); if (b) out.push(b); }
      return out;
    }
    ed.selBlocks = selBlocks;
    function domPath(node) { const p = []; while (node && node !== page) { p.unshift(Array.prototype.indexOf.call(node.parentNode.childNodes, node)); node = node.parentNode; } return p; }
    function fromPath(p) { let n = page; for (const i of p) { if (!n.childNodes[i]) break; n = n.childNodes[i]; } return n; }
    function selPath() { const r = curRange(); if (!r) return null; return { s: domPath(r.startContainer), so: r.startOffset, e: domPath(r.endContainer), eo: r.endOffset }; }
    function applySelPath(sp) {
      if (!sp) return;
      try {
        const r = document.createRange(); const sn = fromPath(sp.s), en = fromPath(sp.e);
        const lim = n => n.nodeType === 3 ? n.length : n.childNodes.length;
        r.setStart(sn, Math.min(sp.so, lim(sn))); r.setEnd(en, Math.min(sp.eo, lim(en))); setRange(r);
      } catch (e) { }
    }

    /* ---------------------------------------------------------------- history (snapshot based) */
    const imgIds = new Map(), imgById = [];
    function ser() {
      return page.innerHTML.replace(/src="(data:[^"]+)"/g, (m, u) => { let id = imgIds.get(u); if (id == null) { id = imgById.length; imgById.push(u); imgIds.set(u, id); } return 'src="@' + id + '"'; })
        .replace(/ class="(csel|sel)"/g, '').replace(/(class="[^"]*?)\s*\b(csel|sel)\b\s*([^"]*")/g, '$1$3').replace(/ class=""/g, '');
    }
    function deser(h) { return h.replace(/src="@(\d+)"/g, (m, i) => 'src="' + imgById[+i] + '"'); }
    const hist = { s: [], i: -1 }; let pend = null;
    function record() {
      pend = null;
      const h = ser(); const top = hist.s[hist.i];
      if (top && top.h === h) { top.sel = selPath(); return; }
      hist.s.length = hist.i + 1; hist.s.push({ h, sel: selPath() });
      if (hist.s.length > 120) hist.s.shift();
      hist.i = hist.s.length - 1; cb.onHistory && cb.onHistory();
    }
    function flush() { if (pend) { clearTimeout(pend); record(); } }
    function schedule() { clearTimeout(pend); pend = setTimeout(record, 500); }
    function load(st) {
      page.innerHTML = deser(st.h); applySelPath(st.sel); page.focus({ preventScroll: true });
      changed(true);
    }
    ed.undo = function () { flush(); if (hist.i > 0) { hist.i--; load(hist.s[hist.i]); } cb.onHistory && cb.onHistory(); };
    ed.redo = function () { flush(); if (hist.i < hist.s.length - 1) { hist.i++; load(hist.s[hist.i]); } cb.onHistory && cb.onHistory(); };
    ed.canUndo = () => hist.i > 0 || !!pend; ed.canRedo = () => hist.i < hist.s.length - 1;
    // run a mutation with history around it
    function mutate(fn) {
      restore(); flush(); const r = fn(); normalize(); record(); changed(true); return r;
    }
    ed.mutate = mutate;

    /* ---------------------------------------------------------------- normalisation / change notifications */
    function cleanWs(root) {
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); const dead = [];
      while (w.nextNode()) { const n = w.currentNode; if (/^\s*$/.test(n.nodeValue) && n.parentNode && /^(UL|OL|TABLE|TBODY|THEAD|TFOOT|TR|COLGROUP)$/.test(n.parentNode.tagName) || (n.parentNode === root && /^\s*$/.test(n.nodeValue))) dead.push(n); }
      dead.forEach(n => n.remove());
    }
    function normalize() {
      let loose = [];
      const flushLoose = before => {
        if (!loose.length) return;
        if (loose.every(n => n.nodeType === 3 && !n.nodeValue.trim())) { loose = []; return; }
        const p = document.createElement('p'); loose[0].parentNode.insertBefore(p, loose[0]); loose.forEach(n => p.appendChild(n)); loose = [];
      };
      Array.from(page.childNodes).forEach(n => {
        if (n.nodeType === 1 && (/^(P|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|PRE|HR|DIV|FIGURE)$/.test(n.tagName))) flushLoose(); else if (n.nodeType === 1 && n.tagName === 'BR' && !loose.length) { const p = document.createElement('p'); n.replaceWith(p); p.appendChild(n); }
        else loose.push(n);
      });
      flushLoose();
      if (!page.firstChild) page.innerHTML = '<p><br></p>';
      page.querySelectorAll('font[size="7"]').forEach(f => { });
    }
    let stateRaf = 0;
    function scheduleState() { if (stateRaf) return; stateRaf = requestAnimationFrame(() => { stateRaf = 0; updateState(); }); }
    const statsDeb = OA.debounce(() => { refreshLayout(); stats(); if (navMode) refreshNav(); }, 250);
    function changed(now) {
      cb.onChange && cb.onChange();
      if (now) { refreshLayout(); stats(); if (navMode) refreshNav(); } else statsDeb();
      page.classList.toggle('empty', page.childNodes.length === 1 && page.firstChild.tagName === 'P' && !page.textContent && !page.querySelector('img,table'));
      scheduleState();
    }
    ed.changed = changed;
    page.addEventListener('input', e => { normalize(); schedule(); changed(); });
    page.addEventListener('beforeinput', e => {
      if (e.inputType === 'historyUndo') { e.preventDefault(); ed.undo(); } else if (e.inputType === 'historyRedo') { e.preventDefault(); ed.redo(); }
    });

    /* ---------------------------------------------------------------- state for ribbon */
    function qs(c) { try { return document.queryCommandState(c); } catch (e) { return false; } }
    function updateState() {
      const r = curRange(); const st = {};
      if (!r) { cb.onState && cb.onState(null); return; }
      const n = elOf(r.startContainer), cs = n ? getComputedStyle(n) : null;
      st.b = qs('bold'); st.i = qs('italic'); st.u = qs('underline'); st.s = qs('strikeThrough'); st.sub = qs('subscript'); st.sup = qs('superscript');
      st.ul = qs('insertUnorderedList'); st.ol = qs('insertOrderedList');
      st.align = qs('justifyCenter') ? 'center' : qs('justifyRight') ? 'right' : qs('justifyFull') ? 'justify' : 'left';
      if (cs) {
        st.font = (cs.fontFamily.split(',')[0] || '').replace(/["']/g, '').trim(); st.size = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
        st.color = OA.colorToHex(cs.color);
      }
      const b = closestBlock(r.startContainer);
      st.block = !b ? 'p' : b.classList.contains('title') ? 'title' : b.classList.contains('subtitle') ? 'subtitle' : b.tagName === 'BLOCKQUOTE' ? 'quote' : /^H[1-6]$/.test(b.tagName) ? b.tagName.toLowerCase() : 'p';
      st.inTable = !!(n && n.closest && n.closest('table.dt'));
      st.link = n && n.closest ? n.closest('a') : null;
      st.img = selImg;
      st.lineHeight = b ? parseFloat(getComputedStyle(b).lineHeight) / parseFloat(getComputedStyle(b).fontSize) : 1.15;
      cb.onState && cb.onState(st);
      linkBubble(st.link, r);
      ctxBar(r);
    }
    ed.state = updateState;

    /* ---------------------------------------------------------------- basic commands */
    function exec(name, val) { document.execCommand('styleWithCSS', false, true); return document.execCommand(name, false, val); }
    ed.cmd = function (name, val) { return mutate(() => exec(name, val)); };
    ed.focus = () => { restore(); };

    function setFontSize(pt) {
      mutate(() => {
        const r = curRange(); if (!r) return;
        if (r.collapsed) {
          const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt'; sp.textContent = '​';
          r.insertNode(sp); const nr = document.createRange(); nr.setStart(sp.firstChild, 1); nr.collapse(true); setRange(nr); return;
        }
        document.execCommand('styleWithCSS', false, false); document.execCommand('fontSize', false, '7'); document.execCommand('styleWithCSS', false, true);
        const made = [];
        page.querySelectorAll('font[size="7"]').forEach(f => {
          const sp = document.createElement('span'); sp.style.fontSize = pt + 'pt';
          while (f.firstChild) sp.appendChild(f.firstChild); f.replaceWith(sp); made.push(sp);
          sp.querySelectorAll('[style*="font-size"]').forEach(c => { c.style.fontSize = ''; if (!c.getAttribute('style')) c.removeAttribute('style'); });
        });
        if (made.length) { const nr = document.createRange(); nr.setStartBefore(made[0]); nr.setEndAfter(made[made.length - 1]); setRange(nr); }
      });
    }
    ed.setFontSize = setFontSize;
    ed.stepFont = function (dir) {
      const r = curRange(); const n = r && elOf(r.startContainer); const cur = n ? parseFloat(getComputedStyle(n).fontSize) * 0.75 : 11;
      let t; if (dir > 0) t = SIZES.find(s => s > cur + 0.01) || cur + 4; else { t = SIZES.slice().reverse().find(s => s < cur - 0.01) || Math.max(1, cur - 1); }
      setFontSize(t);
    };
    ed.setFont = f => ed.cmd('fontName', f);
    ed.setColor = c => ed.cmd('foreColor', c ? '#' + c : '#000000');
    ed.setHighlight = c => ed.cmd('hiliteColor', c ? '#' + c : 'transparent');
    ed.clearFormat = function () { mutate(() => { exec('removeFormat'); exec('unlink'); selBlocks().forEach(b => { if (!/^(LI|TD|TH)$/.test(b.tagName)) { b.removeAttribute('style'); b.removeAttribute('class'); if (/^H[1-6]|BLOCKQUOTE$/.test(b.tagName)) exec('formatBlock', 'p'); } }); }); };
    ed.changeCase = function (mode) {
      mutate(() => {
        const r = curRange(); if (!r || r.collapsed) return;
        const w = document.createTreeWalker(page, NodeFilter.SHOW_TEXT); const nodes = [];
        while (w.nextNode()) if (r.intersectsNode(w.currentNode)) nodes.push(w.currentNode);
        const f = { upper: s => s.toUpperCase(), lower: s => s.toLowerCase(), title: s => s.toLowerCase().replace(/(^|[\s(\-"'])(\S)/g, (m, a, b) => a + b.toUpperCase()), sentence: s => s.toLowerCase().replace(/(^\s*|[.!?]\s+)(\S)/g, (m, a, b) => a + b.toUpperCase()), toggle: s => Array.from(s).map(c => c === c.toUpperCase() ? c.toLowerCase() : c.toUpperCase()).join('') }[mode];
        nodes.forEach(n => { const s = n === r.startContainer ? r.startOffset : 0, e = n === r.endContainer ? r.endOffset : n.length; n.nodeValue = n.nodeValue.slice(0, s) + f(n.nodeValue.slice(s, e)) + n.nodeValue.slice(e); });
      });
    };
    ed.format = ['bold', 'italic', 'underline', 'strikeThrough', 'subscript', 'superscript'].reduce((o, k) => { o[k] = () => ed.cmd(k); return o; }, {});
    ed.align = a => ed.cmd({ left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' }[a]);

    /* paragraph level */
    ed.setBlockStyle = function (name) {
      mutate(() => {
        const map = { normal: 'p', h1: 'h1', h2: 'h2', h3: 'h3', h4: 'h4', h5: 'h5', h6: 'h6', title: 'h1', subtitle: 'h2', quote: 'blockquote', nospace: 'p', listpara: 'p' };
        const tag = map[name] || 'p';
        const keepList = selBlocks().some(b => b.tagName === 'LI');
        if (!keepList) exec('formatBlock', tag);
        selBlocks().forEach(b => {
          if (b.tagName === 'LI') return;
          b.className = ''; if (name !== 'normal' || true) b.removeAttribute('style');
          if (name === 'title') b.classList.add('title'); if (name === 'subtitle') b.classList.add('subtitle');
          if (name === 'nospace') { b.style.margin = '0'; b.style.lineHeight = '1'; }
        });
      });
    };
    ed.setLineSpacing = v => mutate(() => selBlocks().forEach(b => { b.style.lineHeight = v; }));
    ed.setSpacing = (before, after) => mutate(() => selBlocks().forEach(b => { if (before != null) b.style.marginTop = before + 'pt'; if (after != null) b.style.marginBottom = after + 'pt'; }));
    ed.setIndent = (left, right, first) => mutate(() => selBlocks().forEach(b => {
      if (left != null) b.style.marginLeft = left + 'pt'; if (right != null) b.style.marginRight = right + 'pt';
      if (first != null) b.style.textIndent = first + 'pt';
    }));
    ed.indent = function (dir) {
      mutate(() => {
        const bl = selBlocks();
        if (bl.some(b => b.tagName === 'LI')) { exec(dir > 0 ? 'indent' : 'outdent'); return; }
        bl.forEach(b => { const cur = parseFloat(b.style.marginLeft) || 0; const v = Math.max(0, cur + dir * 36); b.style.marginLeft = v ? v + 'pt' : ''; if (!b.getAttribute('style')) b.removeAttribute('style'); });
      });
    };
    ed.list = function (kind, type) {
      mutate(() => {
        const r = curRange(); let n = r && elOf(r.startContainer); let l = n && n.closest && n.closest('ul,ol');
        const want = kind === 'ol' ? 'ol' : 'ul';
        if (!l || l.tagName.toLowerCase() !== want) { exec(want === 'ol' ? 'insertOrderedList' : 'insertUnorderedList'); n = elOf(curRange().startContainer); l = n && n.closest && n.closest('ul,ol'); }
        else if (!type) { exec(want === 'ol' ? 'insertOrderedList' : 'insertUnorderedList'); return; }
        if (l && type) l.style.listStyleType = type;
      });
    };
    ed.sortSel = function (desc) {
      mutate(() => {
        const bl = selBlocks(); if (bl.length < 2) return;
        const par = bl[0].parentNode; if (!bl.every(b => b.parentNode === par)) return;
        const sorted = bl.slice().sort((a, b) => a.textContent.localeCompare(b.textContent, undefined, { sensitivity: 'base' }) * (desc ? -1 : 1));
        const marker = document.createComment(''); par.insertBefore(marker, bl[0]); sorted.forEach(b => par.insertBefore(b, marker)); marker.remove();
      });
    };
    ed.selectAll = () => { restore(); const r = document.createRange(); r.selectNodeContents(page); setRange(r); };
    ed.cut = () => { restore(); if (!document.execCommand('cut')) OA.toast('Use Ctrl+X to cut'); };
    ed.copy = () => { restore(); if (!document.execCommand('copy')) OA.toast('Use Ctrl+C to copy'); };
    ed.paste = async function (plain) {
      restore();
      try {
        const items = await navigator.clipboard.read(); let done = false;
        for (const it of items) {
          if (!plain && it.types.includes('text/html')) { const t = await (await it.getType('text/html')).text(); pasteHTML(t); done = true; break; }
          if (it.types.includes('text/plain')) { const t = await (await it.getType('text/plain')).text(); pasteText(t); done = true; break; }
          const im = it.types.find(t => t.startsWith('image/')); if (im) { ed.insertImageBlob(await it.getType(im)); done = true; break; }
        }
        if (!done) throw 0;
      } catch (e) {
        try { const t = await navigator.clipboard.readText(); pasteText(t); } catch (e2) { OA.toast('Your browser blocked clipboard access — press Ctrl+V to paste.'); }
      }
    };
    function cleanPasted(html) {
      html = html.replace(/<!--[\s\S]*?-->/g, '').replace(/<\/?o:p>/gi, '').replace(/<\/?(meta|link|style|script|title)[^>]*>/gi, '');
      const m = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html); if (m) html = m[1];
      const d = new DOMParser().parseFromString('<body>' + OA.sanitizeHTML(html), 'text/html');
      d.body.querySelectorAll('[class]').forEach(n => { if (/Mso/i.test(n.className)) n.removeAttribute('class'); else n.removeAttribute('class'); });
      d.body.querySelectorAll('[style]').forEach(n => {
        const s = n.style; ['mso-style-name', 'mso-line-height-rule'].forEach(p => s.removeProperty(p));
        const keep = ['font-weight', 'font-style', 'text-decoration', 'color', 'background-color', 'font-family', 'font-size', 'text-align', 'vertical-align'];
        const out = keep.map(k => s.getPropertyValue(k) ? k + ':' + s.getPropertyValue(k) : '').filter(Boolean).join(';');
        if (out) n.setAttribute('style', out); else n.removeAttribute('style');
      });
      d.body.querySelectorAll('span').forEach(s => { if (!s.attributes.length) s.replaceWith(...s.childNodes); });
      d.body.querySelectorAll('table').forEach(t => { t.classList.add('dt'); t.setAttribute('data-b', t.getAttribute('border') && t.getAttribute('border') !== '0' ? 'all' : t.getAttribute('data-b') || 'all'); t.removeAttribute('width'); t.removeAttribute('border'); t.style.removeProperty('width'); });
      d.body.querySelectorAll('td,th').forEach(c => { c.removeAttribute('width'); c.style.removeProperty('width'); });
      return d.body.innerHTML;
    }
    function pasteHTML(h) { mutate(() => { document.execCommand('insertHTML', false, cleanPasted(h)); }); }
    function pasteText(t) { mutate(() => { document.execCommand('insertText', false, t); }); }
    page.addEventListener('paste', e => {
      const cd = e.clipboardData; if (!cd) return;
      const files = Array.from(cd.files || []).filter(f => f.type.startsWith('image/'));
      const html = cd.getData('text/html');
      if (files.length && !html) { e.preventDefault(); files.forEach(f => ed.insertImageBlob(f)); return; }
      if (html && !ed._plain) { e.preventDefault(); pasteHTML(html); return; }
      if (ed._plain) { e.preventDefault(); pasteText(cd.getData('text/plain')); ed._plain = false; }
    });
    page.addEventListener('cut', () => setTimeout(() => { record(); changed(); }, 0));
    page.addEventListener('dragover', e => { if (e.dataTransfer && Array.from(e.dataTransfer.types).includes('Files')) e.preventDefault(); });
    page.addEventListener('drop', e => {
      const fs = e.dataTransfer && Array.from(e.dataTransfer.files || []).filter(f => f.type.startsWith('image/'));
      if (fs && fs.length) { e.preventDefault(); const r = document.caretRangeFromPoint && document.caretRangeFromPoint(e.clientX, e.clientY); if (r) setRange(r); fs.forEach(f => ed.insertImageBlob(f)); }
    });

    /* ---------------------------------------------------------------- insert things */
    ed.insertHTML = h => mutate(() => { document.execCommand('insertHTML', false, h); });
    ed.insertText = t => mutate(() => { document.execCommand('insertText', false, t); });
    ed.pageBreak = () => ed.insertHTML('<div class="pb" contenteditable="false"></div><p><br></p>');
    ed.hr = () => ed.insertHTML('<hr><p><br></p>');
    ed.insertDate = fmt => ed.insertText(fmt || new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' }));
    ed.insertTime = () => ed.insertText(new Date().toLocaleTimeString());
    ed.insertLink = function (url, text) {
      if (!/^[a-z][a-z0-9+.-]*:/i.test(url) && !url.startsWith('#')) url = (/@/.test(url) && !/\s/.test(url) ? 'mailto:' : 'https://') + url;
      mutate(() => {
        const r = curRange();
        const a0 = r && elOf(r.startContainer).closest && elOf(r.startContainer).closest('a');
        if (a0) { a0.setAttribute('href', url); if (text) a0.textContent = text; return; }
        if (r && r.collapsed) document.execCommand('insertHTML', false, `<a href="${OA.esc(url)}">${OA.esc(text || url)}</a>`);
        else exec('createLink', url);
      });
    };
    ed.removeLink = () => mutate(() => {
      const r = curRange(); const a = r && elOf(r.startContainer).closest && elOf(r.startContainer).closest('a');
      if (a && r.collapsed) { const rr = document.createRange(); rr.selectNodeContents(a); setRange(rr); }
      exec('unlink');
    });
    ed.currentLink = () => { const r = curRange(); const a = r && elOf(r.startContainer).closest && elOf(r.startContainer).closest('a'); return a ? { href: a.getAttribute('href'), text: a.textContent } : null; };
    ed.selectedText = () => { const r = curRange(); return r ? r.toString() : ''; };
    ed.insertBookmarkLink = id => ed.insertLink('#' + id);

    ed.insertImageURL = function (url, alt) {
      return new Promise(res => {
        const im = new Image();
        im.onload = () => {
          let w = im.naturalWidth, h = im.naturalHeight; const maxW = (layout.w - layout.margins.l - layout.margins.r) / TW;
          if (w > maxW) { h = h * maxW / w; w = maxW; }
          ed.insertHTML(`<img src="${url}" alt="${OA.esc(alt || '')}" style="width:${Math.round(w)}px;height:${Math.round(h)}px">`);
          res(true);
        };
        im.onerror = () => { OA.toast('Could not load that image'); res(false); };
        im.src = url;
      });
    };
    ed.insertImageBlob = async function (b) { ed.insertImageURL(await OA.readDataURL(b), b.name || ''); };
    ed.insertImageFile = async function () { const fs = await OA.pickFile('image/*', true); for (const f of (fs || [])) await ed.insertImageBlob(f); };
    ed.insertSVGShape = function (kind) {
      const sh = { rect: '<rect x="4" y="4" width="112" height="72" fill="#4472c4" stroke="#2f528f" stroke-width="2"/>', round: '<rect x="4" y="4" width="112" height="72" rx="14" fill="#4472c4" stroke="#2f528f" stroke-width="2"/>', ellipse: '<ellipse cx="60" cy="40" rx="56" ry="36" fill="#4472c4" stroke="#2f528f" stroke-width="2"/>', triangle: '<polygon points="60,4 116,76 4,76" fill="#4472c4" stroke="#2f528f" stroke-width="2"/>', diamond: '<polygon points="60,2 118,40 60,78 2,40" fill="#4472c4" stroke="#2f528f" stroke-width="2"/>', star: '<polygon points="60,4 74,30 104,34 82,54 88,78 60,64 32,78 38,54 16,34 46,30" fill="#ffc000" stroke="#bc8c00" stroke-width="2"/>', arrow: '<polygon points="4,28 76,28 76,6 118,40 76,74 76,52 4,52" fill="#4472c4" stroke="#2f528f" stroke-width="2"/>', line: '<line x1="4" y1="76" x2="116" y2="4" stroke="#2f528f" stroke-width="3"/>' }[kind] || '';
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="240" height="160" viewBox="0 0 120 80">${sh}</svg>`;
      ed.insertImageURL('data:image/svg+xml;base64,' + btoa(svg), 'Shape');
    };
    ed.insertTable = function (rows, cols, opts) {
      opts = opts || {};
      const total = (layout.w - layout.margins.l - layout.margins.r) / TW; const cw = Math.floor(total / cols);
      let h = `<table class="dt" data-b="${opts.border || 'all'}" style="width:${cw * cols}px"><colgroup>${('<col style="width:' + cw + 'px">').repeat(cols)}</colgroup><tbody>`;
      for (let r = 0; r < rows; r++) { h += '<tr>' + '<td><br></td>'.repeat(cols) + '</tr>'; }
      h += '</tbody></table><p><br></p>';
      ed.insertHTML(h);
    };

    /* ---------------------------------------------------------------- tables */
    function tblMap(t) {
      const m = [];
      Array.from(t.rows).forEach((tr, ri) => {
        m[ri] = m[ri] || []; let ci = 0;
        Array.from(tr.cells).forEach(td => {
          while (m[ri][ci]) ci++;
          for (let r = 0; r < td.rowSpan; r++) { m[ri + r] = m[ri + r] || []; for (let c = 0; c < td.colSpan; c++) m[ri + r][ci + c] = td; }
          ci += td.colSpan;
        });
      });
      return m;
    }
    function cellPos(td) { const t = td.closest('table'); const m = tblMap(t); for (let r = 0; r < m.length; r++) for (let c = 0; c < (m[r] || []).length; c++) if (m[r][c] === td) return { t, m, r, c }; return null; }
    function activeCell() { const r = curRange(); const n = r && elOf(r.startContainer); return n && n.closest ? n.closest('td,th') : null; }
    function selCells() { const s = Array.from(page.querySelectorAll('td.csel,th.csel')); if (s.length) return s; const c = activeCell(); return c ? [c] : []; }
    function clearCsel() { page.querySelectorAll('.csel').forEach(c => { c.classList.remove('csel'); if (!c.getAttribute('class')) c.removeAttribute('class'); }); }
    function newCell(like) { const td = document.createElement(like && like.tagName === 'TH' ? 'th' : 'td'); td.innerHTML = '<br>'; if (like) { const b = like.style.backgroundColor; if (like.style.textAlign) td.style.textAlign = like.style.textAlign; } return td; }
    function ensureCols(t) {
      let cg = t.querySelector('colgroup'); const m = tblMap(t); const n = Math.max(0, ...m.map(r => r.length));
      if (cg && cg.children.length === n) return cg;
      if (cg) cg.remove();
      cg = document.createElement('colgroup'); const w = new Array(n).fill(0);
      m.forEach(row => row.forEach((c, i) => { if (c.colSpan === 1 && c.getBoundingClientRect) { const ww = c.getBoundingClientRect().width / (zoomFix()); if (ww > w[i]) w[i] = ww; } }));
      const known = w.filter(Boolean).reduce((a, b) => a + b, 0), unk = w.filter(x => !x).length, tot = t.getBoundingClientRect().width / zoomFix() || 400;
      for (let i = 0; i < n; i++) { const col = document.createElement('col'); col.style.width = Math.round(w[i] || Math.max(30, (tot - known) / Math.max(1, unk))) + 'px'; cg.appendChild(col); }
      t.insertBefore(cg, t.firstChild);
      t.style.width = Array.from(cg.children).reduce((a, c) => a + parseFloat(c.style.width), 0) + 'px';
      t.querySelectorAll('td,th').forEach(c => c.style.removeProperty('width'));
      return cg;
    }
    function zoomFix() { return 1; }
    function syncTableWidth(t) { const cg = t.querySelector('colgroup'); if (cg) t.style.width = Array.from(cg.children).reduce((a, c) => a + (parseFloat(c.style.width) || 0), 0) + 'px'; }
    ed.tableOp = function (op, arg) {
      const cells = selCells(); const td = cells[0];
      if (!td) { OA.toast('Click inside a table first'); return; }
      mutate(() => {
        const p = cellPos(td); if (!p) return; const { t, m, r, c } = p; const rs = td.rowSpan, cs = td.colSpan, ncols = Math.max(...m.map(x => x.length));
        const trs = Array.from(t.rows);
        const colOf = (row, cell) => m[row].indexOf(cell);
        switch (op) {
          case 'rowAbove': case 'rowBelow': {
            const after = op === 'rowBelow', ref = after ? r + rs - 1 : r, tr = document.createElement('tr'), seen = new Set();
            for (let cc = 0; cc < ncols; cc++) {
              const cell = m[ref][cc]; const spans = after ? (m[ref + 1] && m[ref + 1][cc] === cell) : (ref > 0 && m[ref - 1][cc] === cell);
              if (spans) { if (!seen.has(cell)) { cell.rowSpan++; seen.add(cell); } } else { const n = newCell(cell); if (cell.colSpan > 1 && !seen.has(cell)) { n.colSpan = cell.colSpan; seen.add(cell); cc += cell.colSpan - 1; } tr.appendChild(n); }
            }
            const refTr = trs[ref]; after ? refTr.after(tr) : refTr.before(tr); break;
          }
          case 'colLeft': case 'colRight': {
            const after = op === 'colRight', ref = after ? c + cs - 1 : c; const cg = ensureCols(t);
            const base = cg.children[ref]; const nc = document.createElement('col'); nc.style.width = (base ? base.style.width : '80px');
            after ? base.after(nc) : base.before(nc);
            const seen = new Set();
            trs.forEach((tr, ri) => {
              const cell = m[ri][ref]; const spans = after ? m[ri][ref + 1] === cell : (ref > 0 && m[ri][ref - 1] === cell);
              if (spans) { if (!seen.has(cell)) { cell.colSpan++; seen.add(cell); } return; }
              const n = newCell(cell); const target = after ? ref + 1 : ref; let before = null;
              for (let k = target; k < m[ri].length; k++) { const x = m[ri][k]; if (x && x.parentNode === tr) { before = x; break; } }
              tr.insertBefore(n, before);
            });
            syncTableWidth(t); break;
          }
          case 'delRow': {
            if (trs.length - rs < 1) { t.remove(); break; }
            for (let rr = r + rs - 1; rr >= r; rr--) {
              const tr = trs[rr]; const row = m[rr];
              Array.from(tr.cells).forEach(cell => {
                if (cell.rowSpan > 1 && trs[rr + 1] && rr + 1 < m.length) {
                  const col = colOf(rr, cell); const cp = newCell(cell); cp.innerHTML = cell.innerHTML; cp.colSpan = cell.colSpan; cp.rowSpan = cell.rowSpan - 1; cp.setAttribute('style', cell.getAttribute('style') || '');
                  let before = null; for (let k = col + 1; k < (m[rr + 1] || []).length; k++) { const x = m[rr + 1][k]; if (x && x.parentNode === trs[rr + 1]) { before = x; break; } }
                  trs[rr + 1].insertBefore(cp, before);
                }
              });
              const seen = new Set();
              for (let cc = 0; cc < row.length; cc++) { const cell = row[cc]; if (cell.parentNode !== tr && !seen.has(cell)) { cell.rowSpan = Math.max(1, cell.rowSpan - 1); seen.add(cell); } }
              tr.remove();
            }
            break;
          }
          case 'delCol': {
            const cg = t.querySelector('colgroup'); if (ncols - cs < 1) { t.remove(); break; }
            const seen = new Set();
            trs.forEach((tr, ri) => { for (let k = c; k < c + cs; k++) { const cell = m[ri][k]; if (!cell || seen.has(cell + ':' + k)) continue; seen.add(cell + ':' + k); if (cell.colSpan > 1) { if (!seen.has(cell)) { cell.colSpan = Math.max(1, cell.colSpan - 1); seen.add(cell); } } else if (cell.parentNode) cell.remove(); } });
            if (cg && cg.children.length === ncols) { for (let k = c + cs - 1; k >= c; k--) cg.children[k].remove(); syncTableWidth(t); }
            break;
          }
          case 'delTable': t.remove(); break;
          case 'merge': {
            const sel = cells; if (sel.length < 2) { OA.toast('Select two or more cells (drag across cells) to merge'); return; }
            let r1 = 1e9, r2 = -1, c1 = 1e9, c2 = -1;
            sel.forEach(x => { const q = cellPos(x); r1 = Math.min(r1, q.r); c1 = Math.min(c1, q.c); r2 = Math.max(r2, q.r + x.rowSpan - 1); c2 = Math.max(c2, q.c + x.colSpan - 1); });
            const first = m[r1][c1]; const seen = new Set([first]); let html = '';
            for (let rr = r1; rr <= r2; rr++) for (let cc = c1; cc <= c2; cc++) { const x = m[rr][cc]; if (x && !seen.has(x)) { seen.add(x); if (x.textContent.trim() || x.querySelector('img')) html += '<br>' + x.innerHTML.replace(/^(<br>)+$/, ''); x.remove(); } }
            if (html) first.innerHTML = (first.textContent.trim() || first.querySelector('img') ? first.innerHTML : '') + html.replace(/^<br>/, first.textContent.trim() ? '<br>' : '');
            first.rowSpan = r2 - r1 + 1; first.colSpan = c2 - c1 + 1; clearCsel(); break;
          }
          case 'split': {
            if (rs === 1 && cs === 1) { OA.toast('This cell is not merged'); return; }
            td.rowSpan = 1; td.colSpan = 1;
            for (let rr = 0; rr < rs; rr++) for (let cc = 0; cc < cs; cc++) {
              if (!rr && !cc) continue; const tr = trs[r + rr]; const n = newCell(td); let before = null;
              for (let k = c + cc + 1; k < (m[r + rr] || []).length; k++) { const x = m[r + rr][k]; if (x && x !== td && x.parentNode === tr) { before = x; break; } }
              tr.insertBefore(n, before);
            }
            break;
          }
          case 'shade': cells.forEach(x => { x.style.backgroundColor = arg ? '#' + arg : ''; if (!x.getAttribute('style')) x.removeAttribute('style'); }); break;
          case 'valign': cells.forEach(x => x.style.verticalAlign = arg); break;
          case 'calign': cells.forEach(x => x.style.textAlign = arg); break;
          case 'border': t.setAttribute('data-b', arg); break;
          case 'talign': t.style.marginLeft = arg === 'left' ? '' : arg === 'center' ? 'auto' : 'auto'; t.style.marginRight = arg === 'left' ? '' : arg === 'center' ? 'auto' : '0'; if (arg === 'right') t.style.marginLeft = 'auto'; break;
          case 'autofit': { const cg = t.querySelector('colgroup'); if (cg) cg.remove(); t.style.width = ''; t.style.tableLayout = 'auto'; break; }
          case 'fixed': { t.style.tableLayout = 'fixed'; ensureCols(t); break; }
          case 'fitwin': { const cg = ensureCols(t); const total = (layout.w - layout.margins.l - layout.margins.r) / TW; const cur = Array.from(cg.children).reduce((a, x) => a + parseFloat(x.style.width), 0); Array.from(cg.children).forEach(x => x.style.width = Math.floor(parseFloat(x.style.width) * total / cur) + 'px'); syncTableWidth(t); break; }
          case 'distribute': { const cg = ensureCols(t); const cur = Array.from(cg.children).reduce((a, x) => a + parseFloat(x.style.width), 0); const w = Math.floor(cur / cg.children.length); Array.from(cg.children).forEach(x => x.style.width = w + 'px'); syncTableWidth(t); break; }
          case 'rowheight': trs.slice(r, r + rs).forEach(tr => Array.from(tr.cells).forEach(x => x.style.height = arg ? arg + 'px' : '')); break;
          case 'colwidth': { const cg = ensureCols(t); for (let k = c; k < c + cs; k++) cg.children[k].style.width = arg + 'px'; syncTableWidth(t); break; }
          case 'header': { const tr = trs[0]; Array.from(tr.cells).forEach(x => { x.style.fontWeight = arg ? '700' : ''; }); break; }
          case 'style': applyTableStyle(t, arg); break;
          case 'sort': {
            const rows = trs.slice(1); const col = c; const key = x => { const v = (x.cells[col] ? x.cells[col].textContent : '').trim(); return isNaN(+v) || v === '' ? v.toLowerCase() : +v; };
            rows.sort((a, b) => { const ka = key(a), kb = key(b); return (ka > kb ? 1 : ka < kb ? -1 : 0) * (arg ? -1 : 1); }); rows.forEach(x => x.parentNode.appendChild(x)); break;
          }
        }
      });
      if (op !== 'merge') scheduleState();
    };
    function applyTableStyle(t, k) {
      const trs = Array.from(t.rows); t.setAttribute('data-b', k === 'plain' ? 'none' : k === 'minimal' ? 'outer' : 'all');
      t.querySelectorAll('td,th').forEach(c => { c.style.backgroundColor = ''; c.style.color = ''; c.style.fontWeight = ''; c.style.borderColor = ''; if (!c.getAttribute('style')) c.removeAttribute('style'); });
      const hdr = { blue: ['#4472c4', '#fff'], green: ['#70ad47', '#fff'], orange: ['#ed7d31', '#fff'], dark: ['#44546a', '#fff'], grayhead: ['#d9d9d9', '#000'] }[k];
      if (hdr) Array.from(trs[0].cells).forEach(c => { c.style.backgroundColor = hdr[0]; c.style.color = hdr[1]; c.style.fontWeight = '700'; });
      if (k === 'banded' || hdr) trs.forEach((tr, i) => { if (i > 0 && i % 2 === 0) Array.from(tr.cells).forEach(c => c.style.backgroundColor = '#e9eef8'); });
      if (k === 'banded') Array.from(trs[0].cells).forEach(c => { c.style.fontWeight = '700'; c.style.backgroundColor = '#cfd9ee'; });
    }
    ed.selectTablePart = function (what) {
      const td = activeCell(); if (!td) return; const p = cellPos(td); clearCsel();
      const { m, r, c } = p; const n = Math.max(...m.map(x => x.length));
      if (what === 'row') m[r].forEach(x => x.classList.add('csel'));
      else if (what === 'col') m.forEach(row => row[c] && row[c].classList.add('csel'));
      else if (what === 'table') m.forEach(row => row.forEach(x => x.classList.add('csel')));
      else td.classList.add('csel');
    };
    // drag-select cells + column resize
    let cdrag = null, rdrag = null;
    page.addEventListener('mousedown', e => {
      const td = e.target.closest && e.target.closest('td,th');
      if (rdrag === 'hover') { startResize(e); return; }
      if (!e.shiftKey && !(e.target.closest && e.target.closest('.csel'))) clearCsel();
      if (td && e.button === 0) { if (e.shiftKey && cdrag) { markRect(cdrag, td); e.preventDefault(); } else cdrag = td; } else cdrag = null;
    });
    page.addEventListener('mouseover', e => {
      if (!cdrag || !(e.buttons & 1)) return;
      const td = e.target.closest && e.target.closest('td,th'); if (!td || td === cdrag || td.closest('table') !== cdrag.closest('table')) return;
      markRect(cdrag, td); gs().removeAllRanges();
    });
    function markRect(a, b) {
      clearCsel(); const pa = cellPos(a), pb = cellPos(b); if (!pa || !pb) return;
      let r1 = Math.min(pa.r, pb.r), r2 = Math.max(pa.r + a.rowSpan - 1, pb.r + b.rowSpan - 1), c1 = Math.min(pa.c, pb.c), c2 = Math.max(pa.c + a.colSpan - 1, pb.c + b.colSpan - 1);
      let grow = true;
      while (grow) { grow = false; for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) { const x = pa.m[r][c]; if (!x) continue; const q = cellPos(x); if (q.r < r1) { r1 = q.r; grow = true; } if (q.c < c1) { c1 = q.c; grow = true; } if (q.r + x.rowSpan - 1 > r2) { r2 = q.r + x.rowSpan - 1; grow = true; } if (q.c + x.colSpan - 1 > c2) { c2 = q.c + x.colSpan - 1; grow = true; } } }
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) pa.m[r][c] && pa.m[r][c].classList.add('csel');
    }
    page.addEventListener('mousemove', e => {
      if (rdrag && rdrag !== 'hover') return;
      const td = e.target.closest && e.target.closest('td,th');
      if (td && !(e.buttons & 1)) { const rc = td.getBoundingClientRect(); if (e.clientX >= rc.right - 4 && e.clientX <= rc.right + 2) { page.style.cursor = 'col-resize'; rdrag = 'hover'; return; } }
      if (rdrag === 'hover') { rdrag = null; page.style.cursor = ''; }
    });
    function startResize(e) {
      const td = e.target.closest('td,th'); if (!td) { rdrag = null; return; }
      e.preventDefault(); const p = cellPos(td); const t = p.t; flush(); const cg = ensureCols(t); const ci = p.c + td.colSpan - 1; const col = cg.children[ci]; const w0 = parseFloat(col.style.width) || td.getBoundingClientRect().width; const x0 = e.clientX; const z = zoom;
      rdrag = {};
      const mv = ev => { col.style.width = Math.max(20, Math.round(w0 + (ev.clientX - x0) / z)) + 'px'; syncTableWidth(t); };
      const up = () => { document.removeEventListener('mousemove', mv); document.removeEventListener('mouseup', up); rdrag = null; page.style.cursor = ''; record(); changed(); };
      document.addEventListener('mousemove', mv); document.addEventListener('mouseup', up);
    }

    /* ---------------------------------------------------------------- key handling */
    page.addEventListener('keydown', e => {
      const ctrl = e.ctrlKey || e.metaKey;
      if (ctrl && !e.altKey) {
        const k = e.key.toLowerCase();
        if (k === 'z' && !e.shiftKey) { e.preventDefault(); ed.undo(); return; }
        if (k === 'y' || (k === 'z' && e.shiftKey)) { e.preventDefault(); ed.redo(); return; }
        if (k === 'k') { e.preventDefault(); cb.onShortcut && cb.onShortcut('link'); return; }
        if (k === 'f') { e.preventDefault(); ed.showNav('find'); return; }
        if (k === 'h') { e.preventDefault(); ed.showNav('replace'); return; }
        if (k === 'e') { e.preventDefault(); ed.align('center'); return; }
        if (k === 'l') { e.preventDefault(); ed.align('left'); return; }
        if (k === 'r') { e.preventDefault(); ed.align('right'); return; }
        if (k === 'j') { e.preventDefault(); ed.align('justify'); return; }
        if (k === ']') { e.preventDefault(); ed.stepFont(1); return; }
        if (k === '[') { e.preventDefault(); ed.stepFont(-1); return; }
        if (k === 'enter') { e.preventDefault(); ed.pageBreak(); return; }
        if (k === 's') { e.preventDefault(); cb.onShortcut && cb.onShortcut('save'); return; }
        if (k === 'p') { e.preventDefault(); ed.print(); return; }
        if (k === '=' && e.shiftKey) { e.preventDefault(); ed.cmd('superscript'); return; }
        if (k === '=') { e.preventDefault(); ed.cmd('subscript'); return; }
        if (k === 'b' || k === 'i' || k === 'u') { e.preventDefault(); ed.cmd({ b: 'bold', i: 'italic', u: 'underline' }[k]); return; }
        if (k === 'd' && false) return;
      }
      if (ctrl && e.altKey && /^[1-6]$/.test(e.key)) { e.preventDefault(); ed.setBlockStyle('h' + e.key); return; }
      if (ctrl && e.altKey && e.code === 'Digit0') { e.preventDefault(); ed.setBlockStyle('normal'); return; }
      if (e.key === 'Tab') {
        const td = activeCell();
        if (td) {
          e.preventDefault(); const cells = Array.from(td.closest('table').querySelectorAll('td,th')); const i = cells.indexOf(td);
          if (e.shiftKey) { if (i > 0) caretTo(cells[i - 1]); }
          else if (i < cells.length - 1) caretTo(cells[i + 1]);
          else { ed.tableOp('rowBelow'); const nc = Array.from(td.closest('table').querySelectorAll('td,th')); caretTo(nc[i + 1]); }
          return;
        }
        const r = curRange(); const li = r && elOf(r.startContainer).closest && elOf(r.startContainer).closest('li');
        if (li) { e.preventDefault(); ed.indent(e.shiftKey ? -1 : 1); return; }
        e.preventDefault(); if (!e.shiftKey) ed.insertText('\t'); return;
      }
      if (e.key === 'Escape') { clearCsel(); hideImgSel(); }
      if ((e.key === 'Delete' || e.key === 'Backspace') && page.querySelector('.csel')) {
        e.preventDefault(); mutate(() => page.querySelectorAll('.csel').forEach(c => { c.innerHTML = '<br>'; })); return;
      }
      if (e.key === 'Enter' && !e.shiftKey && !ctrl) {
        // exit title/heading style continues as normal paragraph (handled by browser); make new paragraph after heading plain
        const r = curRange(); const b = r && closestBlock(r.startContainer);
        if (b && /^H[1-6]$/.test(b.tagName) && r.collapsed && isAtEnd(b, r)) {
          e.preventDefault(); const p = document.createElement('p'); p.innerHTML = '<br>'; flush(); b.after(p); caretTo(p); normalize(); record(); changed(true); return;
        }
      }
    });
    function isAtEnd(b, r) { const t = document.createRange(); t.selectNodeContents(b); t.setStart(r.endContainer, r.endOffset); return t.toString() === ''; }
    function caretTo(node) { const r = document.createRange(); r.selectNodeContents(node); r.collapse(false); setRange(r); const s = gs(); s.removeAllRanges(); s.addRange(r); page.focus({ preventScroll: true }); node.scrollIntoView({ block: 'nearest' }); }
    ed.caretTo = caretTo;
    // auto-format: "1. " / "- " / "* " lists, "---" rule
    page.addEventListener('input', e => {
      if (e.inputType !== 'insertText' || e.data !== ' ') return;
      const r = curRange(); if (!r || !r.collapsed) return; const b = closestBlock(r.startContainer); if (!b || b.tagName !== 'P') return;
      const txt = b.textContent.replace(/ /g, ' ');
      const m = /^(\d+[.)]|[-*•])\s$/.exec(txt);
      if (m) { flush(); const ordered = /\d/.test(m[1]); b.innerHTML = '<br>'; const rr = document.createRange(); rr.setStart(b, 0); rr.collapse(true); setRange(rr); exec(ordered ? 'insertOrderedList' : 'insertUnorderedList'); record(); changed(true); }
      const hm = /^(#{1,3})\s$/.exec(txt); if (hm) { flush(); b.innerHTML = '<br>'; const rr = document.createRange(); rr.setStart(b, 0); rr.collapse(true); setRange(rr); exec('formatBlock', 'h' + hm[1].length); record(); changed(true); }
    });

    /* ---------------------------------------------------------------- images: select + resize */
    let selImg = null, handle = null;
    function hideImgSel() { if (selImg) selImg.classList.remove('sel'); selImg = null; if (handle) { handle.remove(); handle = null; } cb.onState && updateState(); }
    function showImgSel(img) {
      if (selImg && selImg !== img) selImg.classList.remove('sel');
      selImg = img; img.classList.add('sel');
      if (!handle) {
        handle = el('div', { class: 'w-imgh', 'data-noexport': '1', style: { cursor: 'nwse-resize', touchAction: 'none' } });
        stage.appendChild(handle);
        handle.addEventListener('pointerdown', ev => {
          ev.preventDefault(); handle.setPointerCapture(ev.pointerId); flush();
          const w0 = selImg.getBoundingClientRect().width / zoom, h0 = selImg.getBoundingClientRect().height / zoom, x0 = ev.clientX;
          const mv = e2 => { const w = Math.max(16, w0 + (e2.clientX - x0) / zoom); selImg.style.width = Math.round(w) + 'px'; selImg.style.height = Math.round(w * h0 / w0) + 'px'; selImg.removeAttribute('width'); selImg.removeAttribute('height'); posHandle(); };
          const up = () => { handle.removeEventListener('pointermove', mv); handle.removeEventListener('pointerup', up); record(); changed(); };
          handle.addEventListener('pointermove', mv); handle.addEventListener('pointerup', up);
        });
      }
      posHandle(); scheduleState();
    }
    function posHandle() {
      if (!handle || !selImg) return;
      const a = selImg.getBoundingClientRect(), s = stage.getBoundingClientRect();
      handle.style.left = ((a.right - s.left) / zoom - 6) + 'px'; handle.style.top = ((a.bottom - s.top) / zoom - 6) + 'px';
    }
    page.addEventListener('click', e => {
      const t = e.target;
      if (t.tagName === 'IMG') { showImgSel(t); }
      else if (selImg) hideImgSel();
      const a = t.closest && t.closest('a');
      if (a && (e.ctrlKey || e.metaKey)) { e.preventDefault(); window.open(a.href, '_blank', 'noopener'); }
      else if (a) e.preventDefault();
      const tocE = t.closest && t.closest('.toc [data-t]');
      if (tocE) { const h = page.querySelector('#' + CSS.escape(tocE.getAttribute('data-t'))); if (h) h.scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    });
    page.addEventListener('keyup', () => { if (selImg && !page.contains(selImg)) hideImgSel(); });
    ed.imageOp = function (op, arg) {
      const img = selImg; if (!img) { OA.toast('Click a picture first'); return; }
      mutate(() => {
        const set = (d, ml, mr, fl) => { img.style.display = d || ''; img.style.marginLeft = ml || ''; img.style.marginRight = mr || ''; img.style.float = fl || ''; if (fl) { img.style.margin = fl === 'left' ? '0 12px 6px 0' : '0 0 6px 12px'; } img.dataset.align = op; };
        switch (op) {
          case 'inline': set(); delete img.dataset.align; break;
          case 'center': set('block', 'auto', 'auto'); break;
          case 'right': set('block', 'auto', '0'); break;
          case 'left': set('block', '0', 'auto'); break;
          case 'floatL': set('', '', '', 'left'); img.dataset.align = 'left'; break;
          case 'floatR': set('', '', '', 'right'); img.dataset.align = 'right'; break;
          case 'size': { const ar = img.getBoundingClientRect(); if (arg.w) img.style.width = arg.w + 'px'; if (arg.h) img.style.height = arg.h + 'px'; break; }
          case 'alt': img.setAttribute('alt', arg); break;
          case 'border': img.style.border = arg ? '2px solid #' + arg : ''; break;
          case 'rotate': { const cur = parseFloat((img.style.transform.match(/rotate\((-?[\d.]+)deg/) || [])[1]) || 0; img.style.transform = 'rotate(' + (cur + arg) + 'deg)'; break; }
          case 'flipH': img.style.transform = (img.style.transform.includes('scaleX(-1)') ? img.style.transform.replace('scaleX(-1)', '') : img.style.transform + ' scaleX(-1)').trim(); break;
          case 'opacity': img.style.opacity = arg; break;
          case 'delete': hideImgSel(); img.remove(); return;
        }
      });
      if (selImg) posHandle();
    };
    ed.imageInfo = () => selImg ? { w: Math.round(selImg.getBoundingClientRect().width / zoom), h: Math.round(selImg.getBoundingClientRect().height / zoom), alt: selImg.getAttribute('alt') || '' } : null;

    /* ---------------------------------------------------------------- floating helpers (link bubble, mini toolbar) */
    function linkBubble(a, r) {
      if (!a || !r || !r.collapsed || !page.contains(a)) { linkb.classList.remove('on'); return; }
      const rc = a.getBoundingClientRect();
      linkb.innerHTML = '';
      const href = a.getAttribute('href') || '';
      linkb.append(el('a', { href, target: '_blank', rel: 'noopener', text: href.length > 48 ? href.slice(0, 46) + '…' : href }),
        el('button', { type: 'button', text: 'Edit', onclick: () => cb.onShortcut && cb.onShortcut('link') }),
        el('button', { type: 'button', text: 'Remove', onclick: () => ed.removeLink() }));
      linkb.classList.add('on');
      linkb.style.left = Math.max(4, Math.min(rc.left, innerWidth - 340)) + 'px'; linkb.style.top = Math.min(innerHeight - 50, rc.bottom + 6) + 'px';
    }
    const fine = matchMedia('(pointer: fine)').matches;
    ctxbar.append(...[['B', 'bold', 'font-weight:700'], ['I', 'italic', 'font-style:italic'], ['U', 'underline', 'text-decoration:underline'], ['A+', 'grow'], ['A−', 'shrink'], ['🔗', 'link']].map(([t, c, s]) => {
      const b = el('button', { type: 'button', text: t, style: s || '', title: c }); b.addEventListener('mousedown', e => e.preventDefault());
      b.addEventListener('click', () => { if (c === 'grow') ed.stepFont(1); else if (c === 'shrink') ed.stepFont(-1); else if (c === 'link') cb.onShortcut && cb.onShortcut('link'); else ed.cmd(c); });
      return b;
    }));
    function ctxBar(r) {
      if (!fine || !r || r.collapsed || readOnly || !page.contains(r.commonAncestorContainer)) { ctxbar.classList.remove('on'); return; }
      const rc = r.getBoundingClientRect(); if (!rc.width && !rc.height) { ctxbar.classList.remove('on'); return; }
      ctxbar.classList.add('on'); const w = ctxbar.offsetWidth;
      ctxbar.style.left = Math.max(4, Math.min(innerWidth - w - 4, rc.left + rc.width / 2 - w / 2)) + 'px'; ctxbar.style.top = Math.max(4, rc.top - 42) + 'px';
    }
    scroll.addEventListener('scroll', () => { ctxbar.classList.remove('on'); linkb.classList.remove('on'); posHandle(); }, { passive: true });
    page.addEventListener('mousedown', () => { ctxbar.classList.remove('on'); });

    /* format painter */
    ed.copyFormat = function () {
      const r = curRange(); if (!r) return; const n = elOf(r.startContainer); const cs = getComputedStyle(n);
      fp = { b: qs('bold'), i: qs('italic'), u: qs('underline'), s: qs('strikeThrough'), font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(), size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2, color: OA.colorToHex(cs.color), bg: OA.colorToHex(cs.backgroundColor) };
      OA.toast('Format copied — select text to apply'); cb.onFormatPainter && cb.onFormatPainter(true);
    };
    page.addEventListener('mouseup', () => {
      if (!fp) return; const r = curRange(); if (!r || r.collapsed) return; const f = fp; fp = null; cb.onFormatPainter && cb.onFormatPainter(false);
      mutate(() => {
        ['bold:b', 'italic:i', 'underline:u', 'strikeThrough:s'].forEach(x => { const [c, k] = x.split(':'); if (qs(c) !== !!f[k]) exec(c); });
        exec('fontName', f.font); exec('foreColor', f.color ? '#' + f.color : '#000');
        if (f.bg) exec('hiliteColor', '#' + f.bg); else exec('hiliteColor', 'transparent');
      });
      setFontSize(f.size);
    });

    /* ---------------------------------------------------------------- layout, zoom, stats */
    const PX = v => v / TW;
    function applyLayout() {
      let w = layout.w, h = layout.h; if (layout.landscape) { const t = w; w = h; h = t; }
      const web = layout.view === 'web';
      page.style.width = web ? 'auto' : PX(w) + 'px'; page.style.minHeight = web ? '70vh' : PX(h) + 'px';
      page.style.maxWidth = web ? '980px' : 'none';
      page.style.padding = web ? '24px 32px' : `${PX(layout.margins.t)}px ${PX(layout.margins.r)}px ${PX(layout.margins.b)}px ${PX(layout.margins.l)}px`;
      page.style.columnCount = layout.cols > 1 ? layout.cols : ''; page.style.columnGap = layout.cols > 1 ? '47px' : '';
      page.style.backgroundColor = layout.pageColor ? '#' + layout.pageColor : '#fff';
      stage.style.width = web ? '100%' : PX(w) + 'px';
      if (web) { stage.style.maxWidth = '980px'; } else stage.style.maxWidth = '';
      stage.style.zoom = zoom;
      refreshLayout();
    }
    function pageDims() { let w = layout.w, h = layout.h; if (layout.landscape) { const t = w; w = h; h = t; } return { w: PX(w), h: PX(h) }; }
    function refreshLayout() {
      const { h } = pageDims(); const web = layout.view === 'web';
      const inner = h - PX(layout.margins.t) - PX(layout.margins.b);
      const contentH = page.scrollHeight - PX(layout.margins.t) - PX(layout.margins.b);
      const brk = page.querySelectorAll('.pb').length;
      ed.pageDims = pageDims;
      ed.pageCount = web ? 1 : Math.max(1, Math.ceil(Math.max(1, contentH) / inner) + brk);
      // header / footer overlays
      hfHost.innerHTML = '';
      if (!web && (layout.header || layout.footer || layout.pageNum !== 'none')) {
        const total = Math.max(1, Math.round(page.offsetHeight / h));
        for (let i = 0; i < total; i++) {
          const top = i * h;
          if (layout.header || layout.pageNum.startsWith('header')) hfHost.appendChild(hf(top + 24, 'header', i + 1));
          if (layout.footer || layout.pageNum.startsWith('footer')) hfHost.appendChild(hf(top + h - 44, 'footer', i + 1));
        }
      }
      posHandle();
      cb.onStats && cb.onStats(ed.stats());
    }
    function hf(top, kind, n) {
      const d = el('div', { style: { position: 'absolute', left: PX(layout.margins.l) + 'px', right: PX(layout.margins.r) + 'px', top: top + 'px', fontSize: '10pt', color: '#666', display: 'flex', justifyContent: 'space-between', fontFamily: 'Calibri, Carlito, Arial, sans-serif', overflow: 'hidden' } });
      const txt = kind === 'header' ? layout.header : layout.footer; const pn = layout.pageNum.startsWith(kind);
      const al = layout.pageNum.endsWith('right') ? 'right' : 'center';
      d.style.flexDirection = 'column'; d.style.alignItems = 'center';
      if (txt) d.appendChild(el('div', { text: txt }));
      if (pn) d.appendChild(el('div', { text: String(n), style: { alignSelf: al === 'right' ? 'flex-end' : 'center' } }));
      return d;
    }
    ed.setLayout = function (patch) { Object.assign(layout, patch); if (patch.margins) layout.margins = Object.assign({}, layout.margins, patch.margins); applyLayout(); cb.onChange && cb.onChange(); cb.onLayout && cb.onLayout(layout); };
    ed.setZoom = function (z) { zoom = OA.clamp(z, 0.25, 4); stage.style.zoom = zoom; posHandle(); cb.onZoom && cb.onZoom(zoom); };
    ed.getZoom = () => zoom;
    ed.fitWidth = function () { const { w } = pageDims(); ed.setZoom((scroll.clientWidth - 24) / (layout.view === 'web' ? 980 : w)); };
    ed.fitPage = function () { const { w, h } = pageDims(); ed.setZoom(Math.min((scroll.clientWidth - 24) / w, (scroll.clientHeight - 24) / h)); };
    ed.setView = v => { layout.view = v; applyLayout(); cb.onLayout && cb.onLayout(layout); };
    scroll.addEventListener('wheel', e => { if (e.ctrlKey) { e.preventDefault(); ed.setZoom(zoom * (e.deltaY < 0 ? 1.1 : 0.9)); } }, { passive: false });
    // pinch zoom on touch
    let pinch = null;
    scroll.addEventListener('touchstart', e => { if (e.touches.length === 2) pinch = { d: Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY), z: zoom }; }, { passive: true });
    scroll.addEventListener('touchmove', e => { if (pinch && e.touches.length === 2) { const d = Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY); ed.setZoom(pinch.z * d / pinch.d); } }, { passive: true });
    scroll.addEventListener('touchend', () => { pinch = null; });
    new ResizeObserver(() => refreshLayout()).observe(page);

    ed.stats = function () {
      const r = curRange(); const t = page.innerText || ''; const clean = t.replace(/​/g, '');
      const words = (clean.match(/\S+/g) || []).length;
      const s = { words, chars: clean.replace(/\n/g, '').length, charsNoSpace: clean.replace(/\s/g, '').length, paragraphs: Array.from(page.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')).filter(p => p.textContent.trim()).length, pages: ed.pageCount || 1, lines: Math.max(1, Math.round(page.scrollHeight / (11 * 1.15 * 1.333))) };
      if (r && !r.collapsed) { const st = r.toString(); s.selWords = (st.match(/\S+/g) || []).length; }
      return s;
    };
    function stats() { cb.onStats && cb.onStats(ed.stats()); }
    ed.currentPage = function () { const { h } = pageDims(); return OA.clamp(Math.floor((scroll.scrollTop / zoom + scroll.clientHeight / zoom / 3) / (h + 0)) + 1, 1, ed.pageCount || 1); };
    scroll.addEventListener('scroll', OA.debounce(() => stats(), 100), { passive: true });

    /* ---------------------------------------------------------------- content in/out */
    ed.setContent = function (html, meta) {
      page.innerHTML = OA.sanitizeHTML(html || '<p><br></p>'); cleanWs(page); normalize();
      page.querySelectorAll('table').forEach(t => { if (!t.classList.contains('dt')) t.classList.add('dt'); if (!t.getAttribute('data-b')) t.setAttribute('data-b', 'all'); });
      if (meta) {
        if (meta.page) { const nm = Object.keys(OA.wordIO.PAGES).find(k => { const p = OA.wordIO.PAGES[k]; const a = [meta.page.w, meta.page.h].sort(), b = [p.w, p.h].sort(); return Math.abs(a[0] - b[0]) < 40 && Math.abs(a[1] - b[1]) < 40; }); layout.name = nm || 'Custom'; layout.landscape = !!meta.page.landscape; const lw = meta.page.landscape ? Math.max(meta.page.w, meta.page.h) : Math.min(meta.page.w, meta.page.h); const lh = meta.page.landscape ? Math.min(meta.page.w, meta.page.h) : Math.max(meta.page.w, meta.page.h); layout.w = Math.min(lw, lh); layout.h = Math.max(lw, lh); }
        if (meta.margins) layout.margins = meta.margins; layout.cols = meta.cols || 1; layout.header = meta.header || ''; layout.footer = meta.footer || ''; layout.pageNum = meta.pageNum || 'none'; layout.pageColor = meta.pageColor || null;
      }
      hist.s = []; hist.i = -1; saved = null; hideImgSel(); applyLayout(); record(); changed(true);
      scroll.scrollTop = 0; cb.onLayout && cb.onLayout(layout);
    };
    ed.getHTML = () => { const c = page.cloneNode(true); c.querySelectorAll('[data-noexport]').forEach(n => n.remove()); c.querySelectorAll('.csel,.sel').forEach(n => { n.classList.remove('csel', 'sel'); if (!n.getAttribute('class')) n.removeAttribute('class'); }); return c.innerHTML; };
    ed.getText = () => page.innerText;
    ed.isEmpty = () => !page.textContent.trim() && !page.querySelector('img,table,hr');
    ed.exportDocx = async function (title) {
      flush(); clearCsel(); hideImgSel();
      const z = zoom; stage.style.zoom = 1;
      const web = layout.view === 'web'; const cwas = { w: page.style.width, p: page.style.padding, mw: page.style.maxWidth };
      if (web) { page.style.maxWidth = 'none'; }
      try { return await OA.wordIO.toDocx(page, { page: { w: layout.w, h: layout.h }, landscape: layout.landscape, margins: layout.margins, cols: layout.cols, header: layout.header, footer: layout.footer, pageNum: layout.pageNum, pageColor: layout.pageColor, title }); }
      finally { stage.style.zoom = z; page.style.maxWidth = cwas.mw; }
    };
    ed.setReadOnly = function (v) { readOnly = v; page.contentEditable = v ? 'false' : 'true'; };
    ed.setSpellcheck = v => { page.spellcheck = v; page.blur(); page.focus(); };
    ed.toggleMarks = function () { page.classList.toggle('marks'); return page.classList.contains('marks'); };
    ed.print = function () {
      flush(); clearCsel(); hideImgSel(); let w = layout.w, h = layout.h; if (layout.landscape) { const t = w; w = h; h = t; }
      const st = document.createElement('style'); st.id = 'w-print-style';
      const m = layout.margins; const hdr = layout.header ? `@top-center{content:"${layout.header.replace(/"/g, '\\"')}"; font:10pt Calibri,Arial}` : ''; const ftxt = layout.footer ? `"${layout.footer.replace(/"/g, '\\"')}"` : '';
      const pn = layout.pageNum !== 'none' ? (layout.pageNum.startsWith('footer') ? `@bottom-${layout.pageNum.endsWith('right') ? 'right' : 'center'}{content:${ftxt ? ftxt + ' " – " ' : ''}counter(page); font:10pt Calibri,Arial}` : `@top-${layout.pageNum.endsWith('right') ? 'right' : 'center'}{content:counter(page); font:10pt Calibri,Arial}`) : (ftxt ? `@bottom-center{content:${ftxt}; font:10pt Calibri,Arial}` : '');
      st.textContent = `@media print{@page{size:${PX(w) * 0.75}pt ${PX(h) * 0.75}pt;margin:${m.t / 20}pt ${m.r / 20}pt ${m.b / 20}pt ${m.l / 20}pt;${hdr}${pn}} .w-page{column-count:${layout.cols};} .w-hf{display:none!important}}`;
      document.head.appendChild(st); document.body.classList.add('printing');
      const done = () => { document.body.classList.remove('printing'); st.remove(); window.removeEventListener('afterprint', done); };
      window.addEventListener('afterprint', done); setTimeout(() => window.print(), 50); setTimeout(() => { if (document.body.classList.contains('printing') && !document.hasFocus()) { } }, 1000);
    };

    /* ---------------------------------------------------------------- navigation / find */
    let navMode = null, matches = [], mi = -1;
    const hl = (name, ranges) => { try { if (window.CSS && CSS.highlights) { if (ranges && ranges.length) CSS.highlights.set(name, new Highlight(...ranges)); else CSS.highlights.delete(name); } } catch (e) { } };
    ed.showNav = function (mode) {
      navMode = mode; nav.hidden = false; navTitle.textContent = mode === 'headings' ? 'Navigation' : mode === 'replace' ? 'Find and Replace' : 'Find';
      findBox.hidden = mode === 'headings'; if (mode !== 'headings') buildFind(mode === 'replace'); refreshNav(); cb.onNav && cb.onNav(true);
      if (mode !== 'headings') setTimeout(() => { const i = findBox.querySelector('input[type=text]'); if (i) { const t = ed.selectedText(); if (t && t.length < 80 && !/\n/.test(t)) i.value = t; i.focus(); i.select(); runFind(); } }, 30);
    };
    ed.hideNav = function () { nav.hidden = true; navMode = null; hl('find'); hl('findcur'); matches = []; cb.onNav && cb.onNav(false); };
    ed.toggleNav = m => { if (!nav.hidden && navMode === m) ed.hideNav(); else ed.showNav(m); };
    function refreshNav() {
      if (navMode !== 'headings') return;
      navBody.innerHTML = ''; const hs = Array.from(page.querySelectorAll('h1,h2,h3,h4')).filter(h => h.textContent.trim());
      if (!hs.length) { navBody.appendChild(el('div', { class: 'empty', text: 'Create an outline: apply Heading styles to your text and they will appear here.' })); return; }
      hs.forEach(h => navBody.appendChild(el('a', { class: 'nv l' + h.tagName[1], text: h.textContent.trim(), onclick: () => { h.scrollIntoView({ behavior: 'smooth', block: 'start' }); caretTo(h); } })));
    }
    let fq, fr, fcase, fword, fcnt;
    function buildFind(withReplace) {
      findBox.innerHTML = '';
      fq = el('input', { type: 'text', placeholder: 'Find…', 'aria-label': 'Find' }); fr = el('input', { type: 'text', placeholder: 'Replace with…', 'aria-label': 'Replace with' });
      fcase = el('input', { type: 'checkbox' }); fword = el('input', { type: 'checkbox' }); fcnt = el('span', { class: 'cnt', text: '' });
      const prev = el('button', { type: 'button', text: '▲', title: 'Previous', onclick: () => step(-1) }), next = el('button', { type: 'button', text: '▼', title: 'Next', onclick: () => step(1) });
      findBox.append(el('div', { class: 'row' }, fq, prev, next));
      if (withReplace) findBox.append(el('div', { class: 'row' }, fr, el('button', { type: 'button', text: 'Replace', onclick: () => replaceOne() }), el('button', { type: 'button', text: 'All', onclick: () => replaceAll() })));
      findBox.append(el('div', { class: 'row' }, el('label', null, fcase, ' Match case'), el('label', null, fword, ' Whole word')), el('div', { class: 'row' }, fcnt));
      fq.addEventListener('input', runFind); fcase.addEventListener('change', runFind); fword.addEventListener('change', runFind);
      fq.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey ? -1 : 1); } });
    }
    function esc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
    function runFind() {
      matches = []; mi = -1; const q = fq ? fq.value : ''; hl('find'); hl('findcur');
      if (!q) { fcnt && (fcnt.textContent = ''); return; }
      const re = new RegExp((fword.checked ? '\\b' : '') + esc(q) + (fword.checked ? '\\b' : ''), 'g' + (fcase.checked ? '' : 'i'));
      const w = document.createTreeWalker(page, NodeFilter.SHOW_TEXT);
      while (w.nextNode()) { const n = w.currentNode; let m; re.lastIndex = 0; while ((m = re.exec(n.nodeValue)) && m[0]) { const r = document.createRange(); r.setStart(n, m.index); r.setEnd(n, m.index + m[0].length); matches.push(r); } }
      hl('find', matches); fcnt.textContent = matches.length ? matches.length + ' result' + (matches.length > 1 ? 's' : '') : 'No results';
      if (matches.length) { mi = -1; step(1, true); }
    }
    function step(d, silent) {
      if (!matches.length) return; mi = (mi + d + matches.length) % matches.length; const r = matches[mi];
      hl('findcur', [r]); fcnt.textContent = (mi + 1) + ' of ' + matches.length;
      const rc = r.getBoundingClientRect(), sc = scroll.getBoundingClientRect();
      if (rc.top < sc.top + 20 || rc.bottom > sc.bottom - 20) scroll.scrollTop += rc.top - sc.top - sc.height / 3;
      if (!window.CSS || !CSS.highlights) { setRange(r); }
    }
    function replaceOne() {
      if (!matches.length) return; if (mi < 0) mi = 0; const r = matches[mi];
      mutate(() => { setRange(r); document.execCommand('insertText', false, fr.value); }); runFind();
    }
    function replaceAll() {
      if (!matches.length) return; const n = matches.length;
      mutate(() => { matches.slice().reverse().forEach(r => { setRange(r); document.execCommand('insertText', false, fr.value); }); });
      runFind(); OA.toast('Replaced ' + n + ' occurrence' + (n > 1 ? 's' : ''));
    }
    ed.findNext = () => step(1); ed.findPrev = () => step(-1);

    /* ---------------------------------------------------------------- TOC, footnote, comments, bookmarks */
    ed.insertTOC = function (levels) {
      levels = levels || 3;
      const hs = Array.from(page.querySelectorAll('h1,h2,h3,h4')).filter(h => h.textContent.trim() && +h.tagName[1] <= levels && !h.closest('.toc'));
      if (!hs.length) { OA.alert('No headings found. Apply <b>Heading 1/2/3</b> styles first, then insert the table of contents.'); return; }
      const { h: ph } = pageDims();
      let html = '<div class="toc"><div class="toc-h">Table of Contents</div>';
      hs.forEach((h, i) => { if (!h.id) h.id = 'hd-' + OA.uid(); const pg = Math.floor(h.offsetTop / ph) + 1; html += `<div class="t${h.tagName[1]}" data-t="${h.id}">${OA.esc(h.textContent.trim())}  ${pg}</div>`; });
      html += '</div><p><br></p>';
      mutate(() => { page.querySelectorAll('.toc').forEach(t => t.remove()); document.execCommand('insertHTML', false, html); });
    };
    ed.updateTOC = function () { if (page.querySelector('.toc')) ed.insertTOC(3); else OA.toast('No table of contents in this document'); };
    ed.removeTOC = () => mutate(() => page.querySelectorAll('.toc').forEach(t => t.remove()));
    ed.insertFootnote = async function () {
      const t = await OA.prompt('Footnote text:', '', { title: 'Insert Footnote' }); if (!t) return;
      mutate(() => {
        let box = page.querySelector('.fnote'); const n = (box ? box.children.length : 0) + 1;
        document.execCommand('insertHTML', false, `<sup>${n}</sup>`);
        if (!box) { box = document.createElement('div'); box.className = 'fnote'; page.appendChild(box); }
        const p = document.createElement('p'); p.innerHTML = `<sup>${n}</sup> ${OA.esc(t)}`; p.style.fontSize = '9pt'; p.style.margin = '0'; box.appendChild(p);
      });
    };
    ed.addComment = async function () {
      const r = curRange(); if (!r || r.collapsed) { OA.toast('Select some text first, then add a comment'); return; }
      const t = await OA.prompt('Comment:', '', { title: 'New Comment' }); if (!t) return;
      restore();
      mutate(() => { const sp = document.createElement('span'); sp.className = 'cmt'; sp.setAttribute('data-c', t); sp.setAttribute('data-by', 'You'); sp.title = t; try { r.surroundContents(sp); } catch (e) { exec('hiliteColor', '#fff2ac'); } });
      showComments(true);
    };
    function renderComments() {
      cmtPane.innerHTML = ''; cmtPane.appendChild(el('div', { class: 'w-nav-h' }, el('span', { text: 'Comments' }), el('button', { type: 'button', html: OA.icon('x', 16), onclick: () => showComments(false) })));
      const cs = Array.from(page.querySelectorAll('.cmt')); const b = el('div', { class: 'w-nav-b' });
      if (!cs.length) b.appendChild(el('div', { class: 'empty', text: 'No comments. Select text and use Review → New Comment.' }));
      cs.forEach(c => b.appendChild(el('div', { class: 'cm' }, el('small', { text: 'On: “' + c.textContent.slice(0, 40) + '”' }), el('div', { text: c.getAttribute('data-c') }), el('button', { type: 'button', text: 'Delete', onclick: () => { mutate(() => c.replaceWith(...c.childNodes)); renderComments(); } }), ' ', el('button', { type: 'button', style: { color: '#2b579a' }, text: 'Go to', onclick: () => c.scrollIntoView({ block: 'center', behavior: 'smooth' }) }))));
      cmtPane.appendChild(b);
    }
    function showComments(v) { cmtPane.hidden = !v; if (v) renderComments(); cb.onComments && cb.onComments(v); }
    ed.toggleComments = () => showComments(cmtPane.hidden);
    ed.deleteAllComments = () => mutate(() => page.querySelectorAll('.cmt').forEach(c => c.replaceWith(...c.childNodes)));
    ed.addBookmark = function (name) { const r = curRange(); const b = closestBlock(r && r.startContainer); if (b) mutate(() => { b.id = name; }); };
    ed.readAloud = function () {
      if (!('speechSynthesis' in window)) { OA.toast('Read aloud is not supported in this browser'); return; }
      if (speechSynthesis.speaking) { speechSynthesis.cancel(); return; }
      const r = curRange(); const t = r && !r.collapsed ? r.toString() : page.innerText; speechSynthesis.speak(new SpeechSynthesisUtterance(t));
    };
    ed.findInPage = () => ed.showNav('find');
    ed.symbolPicker = function (anchor) {
      const set = '©®™°±×÷≠≈≤≥∞∑∏√∫∂∆π µ α β γ δ θ λ σ ω Ω → ← ↑ ↓ ↔ ⇒ ⇔ • ◦ ▪ ■ □ ● ○ ★ ☆ ✓ ✗ ✔ ✘ ♠ ♣ ♥ ♦ ♪ ♫ ☺ ☻ € £ ¥ ¢ ₹ ₩ ₽ § ¶ † ‡ ‰ « » “ ” ‘ ’ – — … ¼ ½ ¾ ² ³ ¹ ⁰ ¿ ¡ ß æ Æ ø Ø å Å ñ Ñ ü Ü ö Ö ä Ä é É è È ç Ç अ आ इ ई उ ऊ क ख ग घ च छ ज झ त थ द ध न प फ ब भ म य र ल व श ष स ह ज्ञ'.split(/\s+/).join('').split('');
      const g = el('div', { class: 'symgrid' }); set.filter(c => c.trim()).forEach(c => { const b = el('button', { type: 'button', text: c, onclick: () => { OA.closePopup(); ed.insertText(c); } }); b.addEventListener('mousedown', e => e.preventDefault()); g.appendChild(b); });
      OA.popup(anchor, el('div', { style: { padding: '8px', width: '320px' } }, g));
    };
    ed.imgSelected = () => !!selImg;
    ed.docStats = () => ed.stats();

    /* ---------------------------------------------------------------- init */
    document.execCommand('defaultParagraphSeparator', false, 'p');
    page.addEventListener('focus', () => document.execCommand('defaultParagraphSeparator', false, 'p'));
    page.addEventListener('focus', () => { if (!page.firstChild) page.innerHTML = '<p><br></p>'; });
    page.addEventListener('mouseup', scheduleState); page.addEventListener('keyup', scheduleState);
    page.innerHTML = '<p><br></p>'; applyLayout(); record(); changed(true);
    return ed;
  };
})();
