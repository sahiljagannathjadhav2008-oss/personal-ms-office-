/* Word: .docx export + import (own OOXML writer/reader, no external libs). Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, W = (OA.wordIO = {});
  const xe = s => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const NSW = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'TABLE', 'BLOCKQUOTE', 'PRE', 'HR', 'FIGURE', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'ADDRESS', 'CENTER', 'DL', 'DT', 'DD', 'TBODY', 'THEAD', 'TFOOT', 'TR']);
  W.PAGES = { A4: { w: 11906, h: 16838 }, Letter: { w: 12240, h: 15840 }, Legal: { w: 12240, h: 20160 }, A5: { w: 8391, h: 11906 }, A3: { w: 16838, h: 23811 } };

  function pxOf(v, base) {
    if (v == null || v === '') return null;
    const m = /^(-?[\d.]+)\s*(px|pt|in|cm|mm|em|rem|%)?$/.exec(String(v).trim());
    if (!m) return null; const n = parseFloat(m[1]);
    switch (m[2]) { case 'pt': return n * 4 / 3; case 'in': return n * 96; case 'cm': return n * 96 / 2.54; case 'mm': return n * 96 / 25.4; case 'em': case 'rem': return n * (base || 16); case '%': return null; default: return n; }
  }
  const FONTSZ = { 1: 8, 2: 10, 3: 12, 4: 14, 5: 18, 6: 24, 7: 36 };
  const firstFont = f => { const m = /^\s*["']?([^,"']+)/.exec(f || ''); return m ? m[1].trim() : null; };

  /* ================================================================= EXPORT */
  async function loadImg(src) { return new Promise((res, rej) => { const i = new Image(); i.crossOrigin = 'anonymous'; i.onload = () => res(i); i.onerror = rej; i.src = src; }); }
  async function prepImages(root, st, maxW) {
    const imgs = Array.from(root.querySelectorAll('img'));
    let n = 0;
    for (const img of imgs) {
      try {
        const src = img.currentSrc || img.src; if (!src) continue;
        let bytes, ext;
        const d = src.startsWith('data:') ? OA.dataURLtoBytes(src) : null;
        if (d && /^image\/(png|jpeg|jpg|gif)$/.test(d.mime)) { bytes = d.bytes; ext = d.mime.split('/')[1].replace('jpg', 'jpeg'); }
        else {
          const im = img.complete && img.naturalWidth ? img : await loadImg(src);
          const c = document.createElement('canvas'); c.width = im.naturalWidth || 300; c.height = im.naturalHeight || 150;
          c.getContext('2d').drawImage(im, 0, 0);
          const url = c.toDataURL('image/png'); bytes = OA.dataURLtoBytes(url).bytes; ext = 'png';
        }
        const r = img.getBoundingClientRect();
        let w = r.width || img.naturalWidth || 240, h = r.height || img.naturalHeight || 160;
        if (w > maxW) { h = h * maxW / w; w = maxW; }
        const id = ++n;
        st.imgs.set(img, { rid: 'rIdImg' + id, path: 'media/image' + id + '.' + ext, bytes, ext, w, h, name: 'Picture ' + id, alt: img.getAttribute('alt') || '' });
      } catch (e) { /* skip image that cannot be read */ }
    }
  }

  function runStyle(el, s) {
    s = Object.assign({}, s);
    const t = el.tagName;
    if (t === 'B' || t === 'STRONG') s.b = 1;
    if (t === 'I' || t === 'EM' || t === 'CITE') s.i = 1;
    if (t === 'U' || t === 'INS') s.u = 1;
    if (t === 'S' || t === 'STRIKE' || t === 'DEL') s.s = 1;
    if (t === 'SUB') s.va = 'subscript';
    if (t === 'SUP') s.va = 'superscript';
    if (t === 'CODE' || t === 'KBD' || t === 'PRE' || t === 'SAMP') s.font = 'Consolas';
    if (t === 'MARK') s.bg = 'FFFF00';
    if (t === 'A') { s.u = 1; if (!s.color) s.color = '0563C1'; }
    if (t === 'FONT') {
      if (el.getAttribute('color')) s.color = OA.colorToHex(el.getAttribute('color')) || s.color;
      if (el.getAttribute('face')) s.font = firstFont(el.getAttribute('face'));
      if (el.getAttribute('size') && FONTSZ[el.getAttribute('size')]) s.sz = FONTSZ[el.getAttribute('size')] * 2;
    }
    const cs = el.style;
    if (cs && cs.length) {
      const fw = cs.fontWeight; if (fw) s.b = (fw === 'bold' || fw === 'bolder' || parseInt(fw) >= 600) ? 1 : 0;
      if (cs.fontStyle) s.i = cs.fontStyle === 'italic' || cs.fontStyle === 'oblique' ? 1 : 0;
      const td = cs.textDecorationLine || cs.textDecoration;
      if (td) { if (td.includes('none')) { s.u = 0; s.s = 0; } if (td.includes('underline')) s.u = 1; if (td.includes('line-through')) s.s = 1; }
      if (cs.color) s.color = OA.colorToHex(cs.color) || s.color;
      if (cs.backgroundColor) { const b = OA.colorToHex(cs.backgroundColor); if (b) s.bg = b; }
      if (cs.fontFamily) s.font = firstFont(cs.fontFamily) || s.font;
      if (cs.fontSize) { const px = pxOf(cs.fontSize, 16); if (px) s.sz = Math.round(px * 0.75 * 2); }
      if (cs.verticalAlign === 'sub') s.va = 'subscript'; else if (cs.verticalAlign === 'super') s.va = 'superscript';
      if (cs.textTransform === 'uppercase') s.caps = 1;
    }
    return s;
  }
  function rPr(s) {
    let x = '';
    if (s.font) x += `<w:rFonts w:ascii="${xe(s.font)}" w:hAnsi="${xe(s.font)}" w:cs="${xe(s.font)}" w:eastAsia="${xe(s.font)}"/>`;
    if (s.b) x += '<w:b/><w:bCs/>';
    if (s.i) x += '<w:i/><w:iCs/>';
    if (s.caps) x += '<w:caps/>';
    if (s.s) x += '<w:strike/>';
    if (s.color) x += `<w:color w:val="${s.color}"/>`;
    if (s.sz) x += `<w:sz w:val="${s.sz}"/><w:szCs w:val="${s.sz}"/>`;
    if (s.u) x += '<w:u w:val="single"/>';
    if (s.bg) x += `<w:shd w:val="clear" w:color="auto" w:fill="${s.bg}"/>`;
    if (s.va) x += `<w:vertAlign w:val="${s.va}"/>`;
    return x ? `<w:rPr>${x}</w:rPr>` : '';
  }
  function textRuns(text, s, pre) {
    if (!pre) text = text.replace(/[\r\n]+/g, ' ');
    text = text.replace(/ /g, ' ');
    if (!text) return '';
    const pr = rPr(s); let out = '';
    text.split(/(\t|\n)/).forEach(part => {
      if (part === '\t') out += `<w:r>${pr}<w:tab/></w:r>`;
      else if (part === '\n') out += `<w:r>${pr}<w:br/></w:r>`;
      else if (part) out += `<w:r>${pr}<w:t xml:space="preserve">${xe(part)}</w:t></w:r>`;
    });
    return out;
  }
  function drawing(info, st) {
    const cx = Math.round(info.w * 9525), cy = Math.round(info.h * 9525), id = ++st.docPr;
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="${xe(info.name)}" descr="${xe(info.alt)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${xe(info.path.split('/').pop())}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${info.rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  }
  function inline(nodes, st, s, pre) {
    let out = '';
    nodes.forEach(n => {
      if (n.nodeType === 3) { out += textRuns(n.nodeValue, s, pre); return; }
      if (n.nodeType !== 1) return;
      const t = n.tagName;
      if (n.hasAttribute('data-noexport')) return;
      if (t === 'BR') { out += `<w:r>${rPr(s)}<w:br/></w:r>`; return; }
      if (t === 'IMG') { const info = st.imgs.get(n); if (info) out += drawing(info, st); return; }
      if (t === 'SCRIPT' || t === 'STYLE') return;
      const ns = runStyle(n, s);
      const kids = inline(Array.from(n.childNodes), st, ns, pre || t === 'PRE');
      if (t === 'A' && n.getAttribute('href')) {
        const href = n.getAttribute('href');
        let rid = st.links.get(href);
        if (!rid) { rid = 'rIdLnk' + (st.links.size + 1); st.links.set(href, rid); }
        out += `<w:hyperlink r:id="${rid}" w:history="1">${kids}</w:hyperlink>`;
      } else out += kids;
    });
    return out;
  }
  function pProps(n, inh) {
    const p = { jc: inh.jc, indL: 0, indR: 0, first: 0, hang: 0 };
    const t = n.tagName, cs = n.style || {};
    if (/^H[1-6]$/.test(t)) p.pStyle = n.classList.contains('title') ? 'Title' : n.classList.contains('subtitle') ? 'Subtitle' : 'Heading' + t[1];
    else if (n.classList && n.classList.contains('title')) p.pStyle = 'Title';
    else if (n.classList && n.classList.contains('subtitle')) p.pStyle = 'Subtitle';
    else if (t === 'BLOCKQUOTE') { p.pStyle = 'Quote'; }
    const al = cs.textAlign || n.getAttribute('align') || (t === 'CENTER' ? 'center' : '');
    if (al) p.jc = al === 'center' ? 'center' : al === 'right' || al === 'end' ? 'right' : al === 'justify' ? 'both' : 'left';
    const ml = pxOf(cs.marginLeft) || pxOf(cs.paddingLeft); if (ml) p.indL = Math.round(ml * 15);
    const mr = pxOf(cs.marginRight); if (mr) p.indR = Math.round(mr * 15);
    const ti = pxOf(cs.textIndent); if (ti) { if (ti > 0) p.first = Math.round(ti * 15); else p.hang = Math.round(-ti * 15); }
    const mt = pxOf(cs.marginTop, 16), mb = pxOf(cs.marginBottom, 16);
    if (mt != null) p.before = Math.round(mt * 15);
    if (mb != null) p.after = Math.round(mb * 15);
    if (cs.lineHeight) {
      const lh = cs.lineHeight;
      if (/^[\d.]+$/.test(lh)) p.line = Math.round(parseFloat(lh) * 240);
      else if (/%$/.test(lh)) p.line = Math.round(parseFloat(lh) / 100 * 240);
      else { const px = pxOf(lh, 16); if (px) p.line = Math.round(px * 15); p.exact = 1; }
    }
    if (cs.pageBreakBefore === 'always' || cs.breakBefore === 'page') p.pbb = 1;
    if (n.classList && n.classList.contains('cb-hr')) p.hr = 1;
    return p;
  }
  function paraXml(pp, runs, num, inhIndL) {
    let x = '';
    if (pp.pStyle) x += `<w:pStyle w:val="${pp.pStyle}"/>`;
    if (pp.pbb) x += '<w:pageBreakBefore/>';
    if (num) x += `<w:numPr><w:ilvl w:val="${num.lvl}"/><w:numId w:val="${num.id}"/></w:numPr>`;
    if (pp.hr) x += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="808080"/></w:pBdr>';
    if (pp.before != null || pp.after != null || pp.line) {
      x += '<w:spacing' + (pp.before != null ? ` w:before="${pp.before}"` : '') + (pp.after != null ? ` w:after="${pp.after}"` : '') + (pp.line ? ` w:line="${pp.line}" w:lineRule="${pp.exact ? 'exact' : 'auto'}"` : '') + '/>';
    }
    const left = (pp.indL || 0) + (inhIndL || 0);
    if (!num && (left || pp.indR || pp.first || pp.hang)) x += `<w:ind w:left="${left}"` + (pp.indR ? ` w:right="${pp.indR}"` : '') + (pp.first ? ` w:firstLine="${pp.first}"` : '') + (pp.hang ? ` w:hanging="${pp.hang}"` : '') + '/>';
    if (pp.jc) x += `<w:jc w:val="${pp.jc}"/>`;
    return `<w:p>${x ? '<w:pPr>' + x + '</w:pPr>' : ''}${runs}</w:p>`;
  }
  function hasBlockChild(n) { return Array.from(n.children).some(c => BLOCK.has(c.tagName)); }

  function blocks(parent, st, inh) {
    let out = '', buf = [];
    const flush = () => {
      if (!buf.length) return;
      const onlyWs = buf.every(n => n.nodeType === 3 && !n.nodeValue.trim());
      if (!onlyWs) out += paragraphFrom(buf, st, { jc: inh.jc }, inh);
      buf = [];
    };
    for (const n of parent.childNodes) {
      if (n.nodeType === 3) { buf.push(n); continue; }
      if (n.nodeType !== 1) continue;
      if (n.hasAttribute('data-noexport')) continue;
      if (!BLOCK.has(n.tagName)) { buf.push(n); continue; }
      flush(); out += blockEl(n, st, inh);
    }
    flush();
    return out;
  }
  function paragraphFrom(nodes, st, pp, inh, num, s0) {
    nodes = nodes.slice();
    while (nodes.length && nodes[nodes.length - 1].nodeType === 1 && nodes[nodes.length - 1].tagName === 'BR') nodes.pop();
    if (!pp.jc) { const im = nodes.find(n => n.nodeType === 1 && n.tagName === 'IMG'); if (im && nodes.filter(n => n.nodeType === 1 || (n.nodeValue || '').trim()).length === 1 && im.dataset.align) pp.jc = im.dataset.align === 'center' ? 'center' : im.dataset.align === 'right' ? 'right' : 'left'; }
    return paraXml(pp, inline(nodes, st, s0 || {}, !!(inh && inh.pre)), num, inh && inh.indL);
  }
  function blockEl(n, st, inh) {
    const t = n.tagName;
    if (t === 'HR') return paraXml({ hr: 1 }, '');
    if (n.classList.contains('pb')) return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    if (t === 'TABLE') return tableXml(n, st, inh);
    if (t === 'UL' || t === 'OL') return listXml(n, st, 0, null, inh);
    if (t === 'TBODY' || t === 'THEAD' || t === 'TFOOT' || t === 'TR') return blocks(n, st, inh);
    const pp = pProps(n, inh);
    const cur = Object.assign({}, inh, { jc: pp.jc || inh.jc });
    if (t === 'BLOCKQUOTE') cur.indL = (inh.indL || 0) + 720;
    if (t === 'PRE') cur.pre = true;
    const s0 = runStyle(n, {});
    if (t === 'PRE') s0.font = 'Consolas';
    if (hasBlockChild(n)) return blocks(n, st, cur);
    const myInd = Object.assign({}, cur);
    if (t === 'BLOCKQUOTE') { pp.indL = 0; }
    return paragraphFrom(Array.from(n.childNodes), st, pp, myInd, null, s0);
  }

  function newNum(st, kind) {
    const abs = { disc: 0, circle: 0, square: 0, ul: 0, decimal: 1, 'lower-alpha': 2, 'lower-latin': 2, 'upper-alpha': 3, 'upper-latin': 3, 'lower-roman': 4, 'upper-roman': 5 }[kind];
    const id = st.nums.length + 1; st.nums.push(abs == null ? 1 : abs); return id;
  }
  function listXml(el, st, lvl, numId, inh) {
    if (lvl === 0 || numId == null) {
      const kind = el.tagName === 'UL' ? 'ul' : (el.style.listStyleType || ({ a: 'lower-alpha', A: 'upper-alpha', i: 'lower-roman', I: 'upper-roman' }[el.getAttribute('type')]) || 'decimal');
      numId = newNum(st, kind);
    }
    let out = '';
    Array.from(el.children).forEach(li => {
      if (li.tagName === 'UL' || li.tagName === 'OL') { out += listXml(li, st, Math.min(lvl + 1, 8), null, inh); return; }
      if (li.tagName !== 'LI') return;
      const pp = pProps(li, inh);
      const inl = []; let nested = '';
      Array.from(li.childNodes).forEach(c => {
        if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) nested += listXml(c, st, Math.min(lvl + 1, 8), null, inh);
        else if (c.nodeType === 1 && (c.tagName === 'P' || c.tagName === 'DIV')) { if (inl.length) inl.push(document.createElement('br')); Array.from(c.childNodes).forEach(k => inl.push(k)); }
        else inl.push(c);
      });
      out += paragraphFrom(inl, st, { jc: pp.jc || inh.jc, pStyle: 'ListParagraph' }, { indL: 0 }, { lvl, id: numId }, runStyle(li, {})) + nested;
    });
    return out;
  }

  function tableXml(tb, st, inh) {
    const rows = Array.from(tb.rows);
    if (!rows.length) return '';
    const bmode = tb.getAttribute('data-b') || 'all';
    const trect = tb.getBoundingClientRect();
    const tw = Math.min(Math.round((trect.width || 600) * 15), st.textW);
    let ncols = 0; rows.forEach(r => { let c = 0; Array.from(r.cells).forEach(x => c += x.colSpan || 1); ncols = Math.max(ncols, c); });
    const widths = new Array(ncols).fill(0);
    rows.forEach(r => { let ci = 0; Array.from(r.cells).forEach(x => { const sp = x.colSpan || 1; if (sp === 1) { const w = x.getBoundingClientRect().width; if (w > widths[ci]) widths[ci] = w; } ci += sp; }); });
    const known = widths.filter(w => w > 0).reduce((a, b) => a + b, 0), unk = widths.filter(w => !w).length;
    const total = (trect.width || 600);
    widths.forEach((w, i) => { if (!w) widths[i] = Math.max(30, (total - known) / Math.max(1, unk)); });
    const scale = tw / (widths.reduce((a, b) => a + b, 0) * 15 || 1);
    const gw = widths.map(w => Math.max(300, Math.round(w * 15 * scale)));
    const bd = (v) => `<w:${v} w:val="single" w:sz="4" w:space="0" w:color="${bmode === 'all' || bmode === 'outer' ? '808080' : 'auto'}"/>`;
    let x = `<w:tbl><w:tblPr><w:tblW w:w="${gw.reduce((a, b) => a + b, 0)}" w:type="dxa"/>`;
    if (bmode !== 'none') x += `<w:tblBorders>${bd('top')}${bd('left')}${bd('bottom')}${bd('right')}${bmode === 'all' ? bd('insideH') + bd('insideV') : ''}</w:tblBorders>`;
    x += '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>' + gw.map(w => `<w:gridCol w:w="${w}"/>`).join('') + '</w:tblGrid>';
    const carry = new Array(ncols).fill(0);
    rows.forEach(r => {
      x += '<w:tr>'; let ci = 0;
      const cells = Array.from(r.cells); let k = 0;
      while (ci < ncols) {
        if (carry[ci] > 0) { x += `<w:tc><w:tcPr><w:tcW w:w="${gw[ci]}" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>`; carry[ci]--; ci++; continue; }
        const c = cells[k++]; if (!c) break;
        const sp = Math.min(c.colSpan || 1, ncols - ci), rs = c.rowSpan || 1;
        let w = 0; for (let j = 0; j < sp; j++) w += gw[ci + j];
        let pr = `<w:tcW w:w="${w}" w:type="dxa"/>`;
        if (sp > 1) pr += `<w:gridSpan w:val="${sp}"/>`;
        if (rs > 1) { pr += '<w:vMerge w:val="restart"/>'; for (let j = 0; j < sp; j++) carry[ci + j] = rs - 1; }
        const bg = OA.colorToHex(c.style.backgroundColor || c.getAttribute('bgcolor') || '');
        if (bg) pr += `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
        const va = c.style.verticalAlign; if (va === 'middle' || va === 'bottom') pr += `<w:vAlign w:val="${va === 'middle' ? 'center' : 'bottom'}"/>`;
        const hs = c.tagName === 'TH' ? { b: 1 } : {};
        let inner = hasBlockChild(c) ? blocks(c, st, { jc: c.style.textAlign ? undefined : inh.jc, th: c.tagName === 'TH' }) : '';
        if (!hasBlockChild(c)) {
          const pp = pProps(c, {}); if (!pp.jc && c.tagName === 'TH') pp.jc = 'center';
          inner = paragraphFrom(Array.from(c.childNodes), st, Object.assign(pp, { before: 0, after: 0 }), { indL: 0 }, null, Object.assign(runStyle(c, {}), hs));
        }
        if (!inner) inner = '<w:p/>';
        x += `<w:tc><w:tcPr>${pr}</w:tcPr>${inner}</w:tc>`;
        ci += sp;
      }
      x += '</w:tr>';
    });
    return x + '</w:tbl><w:p><w:pPr><w:spacing w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>';
  }

  function absNum(id, fmt0) {
    const cyc = ['decimal', 'lowerLetter', 'lowerRoman'];
    let lv = '';
    for (let i = 0; i < 9; i++) {
      let fmt, txt, font = '';
      if (fmt0 === 'bullet') { fmt = 'bullet'; txt = ['•', 'o', '▪'][i % 3]; font = '<w:rPr><w:rFonts w:ascii="' + (i % 3 === 1 ? 'Courier New' : 'Arial') + '" w:hAnsi="' + (i % 3 === 1 ? 'Courier New' : 'Arial') + '" w:cs="Arial"/></w:rPr>'; }
      else { fmt = i === 0 ? fmt0 : cyc[i % 3]; txt = `%${i + 1}.`; }
      lv += `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr>${font}</w:lvl>`;
    }
    return `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${lv}</w:abstractNum>`;
  }
  function numberingXml(st) {
    const fm = ['bullet', 'decimal', 'lowerLetter', 'upperLetter', 'lowerRoman', 'upperRoman'];
    let x = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${NSW}>` + fm.map((f, i) => absNum(i, f)).join('');
    st.nums.forEach((abs, i) => { x += `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${abs}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`; });
    if (!st.nums.length) x += '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>';
    return x + '</w:numbering>';
  }
  function stylesXml() {
    const hd = (id, name, sz, color, extra, before, after) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="${after}"/><w:outlineLvl w:val="${id.slice(-1) - 1}"/></w:pPr><w:rPr>${extra || ''}<w:color w:val="${color}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${NSW}>
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Nirmala UI"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="hi-IN"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
${hd('Heading1', 'heading 1', 32, '2F5496', '<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>', 240, 80)}
${hd('Heading2', 'heading 2', 26, '2F5496', '<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>', 200, 80)}
${hd('Heading3', 'heading 3', 24, '1F3763', '<w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/>', 160, 80)}
${hd('Heading4', 'heading 4', 22, '2F5496', '<w:i/><w:iCs/>', 80, 40)}
${hd('Heading5', 'heading 5', 22, '2F5496', '', 80, 40)}
${hd('Heading6', 'heading 6', 22, '1F3763', '', 80, 40)}
<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr><w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr><w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="0"/><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`;
  }

  W.toDocx = async function (root, o) {
    o = Object.assign({ page: W.PAGES.A4, landscape: false, margins: { t: 1440, r: 1440, b: 1440, l: 1440 }, cols: 1, header: '', footer: '', pageNum: 'none', title: 'Document', pageColor: null }, o || {});
    let pw = o.page.w, ph = o.page.h; if (o.landscape) { const t = pw; pw = ph; ph = t; }
    const st = { imgs: new Map(), links: new Map(), nums: [], docPr: 0, textW: pw - o.margins.l - o.margins.r };
    await prepImages(root, st, st.textW / 15);
    let body = blocks(root, st, {});
    if (!body) body = '<w:p/>';
    const files = [], rels = [];
    let hdrRef = '', ftrRef = '';
    const pageField = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const hp = (jc, inner) => `<w:p><w:pPr><w:spacing w:after="0"/><w:jc w:val="${jc}"/></w:pPr>${inner}</w:p>`;
    const hasHdr = o.header || o.pageNum === 'header-right' || o.pageNum === 'header-center';
    const hasFtr = o.footer || o.pageNum === 'footer-right' || o.pageNum === 'footer-center';
    if (hasHdr) {
      let inner = o.header ? `<w:r><w:t xml:space="preserve">${xe(o.header)}</w:t></w:r>` : '';
      let jc = o.header ? 'center' : (o.pageNum === 'header-right' ? 'right' : 'center');
      let x = '';
      if (o.pageNum.startsWith('header')) { if (o.header) { x = hp('center', inner) + hp(o.pageNum === 'header-right' ? 'right' : 'center', pageField); } else x = hp(jc, pageField); }
      else x = hp('center', inner);
      files.push(['word/header1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${NSW}>${x}</w:hdr>`]);
      rels.push(['rIdHdr1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml']); hdrRef = '<w:headerReference w:type="default" r:id="rIdHdr1"/>';
    }
    if (hasFtr) {
      let x = '';
      if (o.footer) x += hp('center', `<w:r><w:t xml:space="preserve">${xe(o.footer)}</w:t></w:r>`);
      if (o.pageNum.startsWith('footer')) x += hp(o.pageNum === 'footer-right' ? 'right' : 'center', pageField);
      files.push(['word/footer1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${NSW}>${x}</w:ftr>`]);
      rels.push(['rIdFtr1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer', 'footer1.xml']); ftrRef = '<w:footerReference w:type="default" r:id="rIdFtr1"/>';
    }
    const sect = `<w:sectPr>${hdrRef}${ftrRef}<w:pgSz w:w="${pw}" w:h="${ph}"${o.landscape ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${o.margins.t}" w:right="${o.margins.r}" w:bottom="${o.margins.b}" w:left="${o.margins.l}" w:header="708" w:footer="708" w:gutter="0"/>${o.cols > 1 ? `<w:cols w:num="${o.cols}" w:space="708"/>` : ''}</w:sectPr>`;
    const bg = o.pageColor ? `<w:background w:color="${o.pageColor}"/>` : '';
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NSW}>${bg}<w:body>${body}${sect}</w:body></w:document>`;
    st.imgs.forEach(info => { files.push(['word/' + info.path, info.bytes]); rels.push([info.rid, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', info.path]); });
    const relXml = rels.map(r => `<Relationship Id="${r[0]}" Type="${r[1]}" Target="${r[2]}"/>`).join('') + Array.from(st.links).map(([h, id]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xe(h)}" TargetMode="External"/>`).join('');
    const exts = new Set(); st.imgs.forEach(i => exts.add(i.ext));
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${Array.from(exts).map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('')}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${hasHdr ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}${hasFtr ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const all = [
      ['[Content_Types].xml', ct],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`],
      ['word/document.xml', doc],
      ['word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${relXml}</Relationships>`],
      ['word/styles.xml', stylesXml()],
      ['word/numbering.xml', numberingXml(st)],
      ['word/settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings ${NSW}><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`],
      ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xe(o.title)}</dc:title><dc:creator>Office Web</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`],
      ['docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Office Web</Application></Properties>`]
    ].concat(files);
    return OA.zip(all);
  };

  /* ================================================================= IMPORT */
  const HL = { yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF' };
  const ln = (n, name) => { const r = []; if (!n) return r; for (const c of n.childNodes) if (c.nodeType === 1 && c.localName === name) r.push(c); return r; };
  const l1 = (n, name) => ln(n, name)[0] || null;
  const wv = (n) => n ? (n.getAttribute('w:val') ?? n.getAttribute('val')) : null;
  const esch = s => OA.esc(s);

  W.fromDocx = async function (buf) {
    const z = await OA.unzip(buf);
    if (!z['word/document.xml']) throw new Error('This is not a valid .docx file');
    const rels = {};
    if (z['word/_rels/document.xml.rels']) Array.from(OA.xmlDoc(z['word/_rels/document.xml.rels']).getElementsByTagName('Relationship')).forEach(r => { rels[r.getAttribute('Id')] = { type: r.getAttribute('Type'), target: r.getAttribute('Target'), ext: r.getAttribute('TargetMode') === 'External' }; });
    const numFmt = {};
    if (z['word/numbering.xml']) {
      const nd = OA.xmlDoc(z['word/numbering.xml']); const abs = {};
      Array.from(nd.getElementsByTagName('w:abstractNum')).forEach(a => { const m = {}; ln(a, 'lvl').forEach(l => { m[l.getAttribute('w:ilvl')] = wv(l1(l, 'numFmt')); }); abs[a.getAttribute('w:abstractNumId')] = m; });
      Array.from(nd.getElementsByTagName('w:num')).forEach(n => { numFmt[n.getAttribute('w:numId')] = abs[wv(l1(n, 'abstractNumId'))] || {}; });
    }
    const styles = {};
    if (z['word/styles.xml']) Array.from(OA.xmlDoc(z['word/styles.xml']).getElementsByTagName('w:style')).forEach(s => {
      const id = s.getAttribute('w:styleId'), nm = (wv(l1(s, 'name')) || '').toLowerCase(), pp = l1(s, 'pPr');
      styles[id] = { name: nm, based: wv(l1(s, 'basedOn')), num: pp && l1(pp, 'numPr') ? { id: wv(l1(l1(pp, 'numPr'), 'numId')), lvl: wv(l1(l1(pp, 'numPr'), 'ilvl')) } : null, outline: pp && l1(pp, 'outlineLvl') ? +wv(l1(pp, 'outlineLvl')) : null, rpr: l1(s, 'rPr') };
    });
    const styleTag = id => {
      let s = styles[id], guard = 0;
      while (s && guard++ < 6) {
        const nm = s.name;
        let m = /^heading (\d)/.exec(nm); if (m) return 'h' + Math.min(6, +m[1]);
        if (nm === 'title') return 'title'; if (nm === 'subtitle') return 'subtitle'; if (nm === 'quote' || nm === 'intense quote') return 'quote';
        if (s.outline != null && s.outline < 6) return 'h' + (s.outline + 1);
        s = styles[s.based];
      }
      return null;
    };
    const meta = { page: null, margins: null };
    const media = async (rid) => {
      const r = rels[rid]; if (!r || r.ext) return null;
      let p = r.target.replace(/^\//, ''); p = p.startsWith('word/') ? p : 'word/' + p; p = p.replace(/\/\.\//g, '/');
      while (/[^/]+\/\.\.\//.test(p)) p = p.replace(/[^/]+\/\.\.\//, '');
      const b = z[p]; if (!b) return null;
      const e = OA.ext(p), mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', svg: 'image/svg+xml', webp: 'image/webp', bmp: 'image/bmp' }[e];
      return mime ? OA.bytesToDataURL(b, mime) : null;
    };
    const doc = OA.xmlDoc(z['word/document.xml']);
    const body = doc.getElementsByTagName('w:body')[0];

    async function runs(p, ctx) {
      let html = '';
      for (const r of p.childNodes) {
        if (r.nodeType !== 1) continue;
        const nm = r.localName;
        if (nm === 'r') html += await run(r, ctx);
        else if (nm === 'hyperlink') {
          const inner = await runs(r, ctx), id = r.getAttribute('r:id'), rel = id && rels[id];
          html += rel && rel.ext ? `<a href="${esch(rel.target)}">${inner}</a>` : inner;
        } else if (nm === 'fldSimple' || nm === 'smartTag' || nm === 'sdt' || nm === 'ins' || nm === 'sdtContent') html += await runs(nm === 'sdt' ? (l1(r, 'sdtContent') || r) : r, ctx);
        else if (nm === 'AlternateContent') { const ch = l1(r, 'Choice') || l1(r, 'Fallback'); if (ch) html += await runs(ch, ctx); }
      }
      return html;
    }
    async function run(r, ctx) {
      const pr = l1(r, 'rPr'); let css = '', open = '', close = '';
      if (pr) {
        const on = n => { const e = l1(pr, n); return e && !['0', 'false', 'none'].includes(wv(e)); };
        if (on('b')) { open += '<b>'; close = '</b>' + close; }
        if (on('i')) { open += '<i>'; close = '</i>' + close; }
        if (on('u')) { open += '<u>'; close = '</u>' + close; }
        if (on('strike') || on('dstrike')) { open += '<s>'; close = '</s>' + close; }
        const va = wv(l1(pr, 'vertAlign')); if (va === 'superscript') { open += '<sup>'; close = '</sup>' + close; } else if (va === 'subscript') { open += '<sub>'; close = '</sub>' + close; }
        const col = wv(l1(pr, 'color')); if (col && col !== 'auto' && /^[0-9a-f]{6}$/i.test(col)) css += `color:#${col};`;
        const hl = wv(l1(pr, 'highlight')); if (hl && HL[hl]) css += `background-color:${HL[hl]};`;
        const shd = l1(pr, 'shd'); if (shd && !hl) { const f = shd.getAttribute('w:fill'); if (f && f !== 'auto' && /^[0-9a-f]{6}$/i.test(f)) css += `background-color:#${f};`; }
        const rf = l1(pr, 'rFonts'); if (rf) { const f = rf.getAttribute('w:ascii') || rf.getAttribute('w:hAnsi') || rf.getAttribute('w:cs'); if (f) css += `font-family:'${f.replace(/'/g, '')}';`; }
        const sz = wv(l1(pr, 'sz')); if (sz) css += `font-size:${(+sz / 2)}pt;`;
        if (on('caps')) css += 'text-transform:uppercase;';
      }
      let inner = '';
      for (const c of r.childNodes) {
        if (c.nodeType !== 1) continue;
        switch (c.localName) {
          case 't': inner += esch(c.textContent); break;
          case 'tab': inner += '<span style="display:inline-block;width:2em"></span>'; break;
          case 'br': inner += wv(c) === 'page' || c.getAttribute('w:type') === 'page' ? '</span><div class="pb"></div><span>' : '<br>'; break;
          case 'noBreakHyphen': inner += '-'; break;
          case 'drawing': case 'pict': case 'object': {
            const bl = c.getElementsByTagName('a:blip')[0], vi = c.getElementsByTagName('v:imagedata')[0];
            const rid = bl ? (bl.getAttribute('r:embed') || bl.getAttribute('r:link')) : vi ? vi.getAttribute('r:id') : null;
            const url = rid ? await media(rid) : null;
            if (url) {
              const ex = c.getElementsByTagName('wp:extent')[0]; let sz = '';
              if (ex) sz = ` style="width:${Math.round(+ex.getAttribute('cx') / 9525)}px;height:auto"`;
              inner += `<img src="${url}"${sz} alt="">`;
            }
            break;
          }
          case 'sym': break;
        }
      }
      if (!inner) return '';
      inner = inner.replace(/<span><\/span>/g, '');
      return css ? `<span style="${css}">${open}${inner}${close}</span>` : open + inner + close;
    }
    function pStyleCss(p) {
      const pp = l1(p, 'pPr'); let css = '';
      if (!pp) return { css };
      const j = wv(l1(pp, 'jc')); if (j) { const m = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[j]; if (m) css += `text-align:${m};`; }
      const ind = l1(pp, 'ind');
      if (ind) {
        const g = a => +(ind.getAttribute('w:' + a) || ind.getAttribute('w:' + (a === 'left' ? 'start' : a === 'right' ? 'end' : a)) || 0);
        const L = g('left'), R = g('right'), F = g('firstLine'), H = g('hanging');
        if (L) css += `margin-left:${Math.round(L / 15)}px;`; if (R) css += `margin-right:${Math.round(R / 15)}px;`;
        if (F) css += `text-indent:${Math.round(F / 15)}px;`; else if (H) css += `text-indent:-${Math.round(H / 15)}px;`;
      }
      const sp = l1(pp, 'spacing');
      if (sp) {
        const b = sp.getAttribute('w:before'), a = sp.getAttribute('w:after'), l = sp.getAttribute('w:line'), lr = sp.getAttribute('w:lineRule');
        if (b != null) css += `margin-top:${(+b / 20)}pt;`; if (a != null) css += `margin-bottom:${(+a / 20)}pt;`;
        if (l && (!lr || lr === 'auto')) css += `line-height:${(+l / 240).toFixed(2)};`;
      }
      if (l1(pp, 'pageBreakBefore')) css += 'break-before:page;';
      return { css };
    }
    const hlTag = { h1: 'h1', h2: 'h2', h3: 'h3', h4: 'h4', h5: 'h5', h6: 'h6' };
    async function paragraph(p) {
      const pp = l1(p, 'pPr'); const sid = pp && wv(l1(pp, 'pStyle'));
      let tag = sid ? styleTag(sid) : null;
      const { css } = pStyleCss(p);
      let inner = await runs(p, {});
      inner = inner.replace(/<span><\/span>/g, '');
      let cls = '', t = 'p';
      if (tag && hlTag[tag]) t = tag; else if (tag === 'title') { t = 'h1'; cls = ' class="title"'; } else if (tag === 'subtitle') { t = 'h2'; cls = ' class="subtitle"'; } else if (tag === 'quote') t = 'blockquote';
      if (!inner) inner = '<br>';
      return `<${t}${cls}${css ? ` style="${css}"` : ''}>${inner}</${t}>`;
    }
    function listInfo(p) {
      const pp = l1(p, 'pPr'); if (!pp) return null;
      let np = l1(pp, 'numPr'), id, lvl;
      if (np) { id = wv(l1(np, 'numId')); lvl = wv(l1(np, 'ilvl')) || '0'; }
      else { const sid = wv(l1(pp, 'pStyle')); const st = sid && styles[sid]; if (st && st.num) { id = st.num.id; lvl = st.num.lvl || '0'; } }
      if (id == null || id === '0') return null;
      const fmt = (numFmt[id] || {})[lvl] || 'bullet';
      return { id, lvl: +lvl, ordered: fmt !== 'bullet' && fmt !== 'none', fmt };
    }
    const LT = { decimal: '', lowerLetter: 'lower-alpha', upperLetter: 'upper-alpha', lowerRoman: 'lower-roman', upperRoman: 'upper-roman' };
    async function table(tbl) {
      const props = l1(tbl, 'tblPr'); let bordered = false;
      if (props) { const b = l1(props, 'tblBorders'); if (b && ln(b, 'top').some(x => !['nil', 'none'].includes(wv(x)))) bordered = true; const ts = wv(l1(props, 'tblStyle')); if (ts && /grid|table/i.test(ts) && !/normal/i.test(ts)) bordered = true; }
      const trs = ln(tbl, 'tr'); const grid = []; const owner = [];
      let html = `<table class="dt" data-b="${bordered ? 'all' : 'none'}"><tbody>`;
      for (let ri = 0; ri < trs.length; ri++) {
        html += '<tr>'; let ci = 0;
        for (const tc of ln(trs[ri], 'tc')) {
          const pr = l1(tc, 'tcPr'); const gs = pr && l1(pr, 'gridSpan') ? +wv(l1(pr, 'gridSpan')) : 1;
          const vm = pr && l1(pr, 'vMerge'); const vmv = vm ? (wv(vm) || 'continue') : null;
          if (vmv === 'continue') { const o = owner[ci]; if (o) o.rs++; ci += gs; continue; }
          let st = '';
          const shd = pr && l1(pr, 'shd'); if (shd) { const f = shd.getAttribute('w:fill'); if (f && f !== 'auto' && /^[0-9a-f]{6}$/i.test(f)) st += `background-color:#${f};`; }
          const va = pr && wv(l1(pr, 'vAlign')); if (va) st += `vertical-align:${va === 'center' ? 'middle' : va};`;
          const tw = pr && l1(pr, 'tcW'); if (tw && tw.getAttribute('w:type') === 'dxa') st += `width:${Math.round(+tw.getAttribute('w:w') / 15)}px;`;
          const inner = await blockContent(tc);
          const cell = { rs: 1, id: 'c' + ri + '_' + ci };
          for (let j = 0; j < gs; j++) owner[ci + j] = vmv === 'restart' ? cell : null;
          html += `<td${gs > 1 ? ` colspan="${gs}"` : ''}${st ? ` style="${st}"` : ''}{{RS:${cell.id}}}>${inner || '<br>'}</td>`;
          grid.push(cell); ci += gs;
        }
        html += '</tr>';
      }
      html += '</tbody></table>';
      grid.forEach(c => { html = html.replace(`{{RS:${c.id}}}`, c.rs > 1 ? ` rowspan="${c.rs}"` : ''); });
      return html.replace(/\{\{RS:[^}]+\}\}/g, '');
    }
    async function blockContent(parent) {
      let html = ''; let stack = [];
      const closeTo = (n) => { while (stack.length > n) { const l = stack.pop(); html += `</li></${l.tag}>`; } };
      for (const n of parent.childNodes) {
        if (n.nodeType !== 1) continue;
        const nm = n.localName;
        if (nm === 'p') {
          const li = listInfo(n);
          if (!li) { closeTo(0); html += await paragraph(n); continue; }
          const inner = (await runs(n, {})).replace(/<span><\/span>/g, '') || '<br>';
          const tag = li.ordered ? 'ol' : 'ul';
          const lt = li.ordered && LT[li.fmt] ? ` style="list-style-type:${LT[li.fmt]}"` : '';
          const openList = () => { html += `<${tag}${lt}><li>`; stack.push({ tag, id: li.id, lvl: stack.length }); };
          if (!stack.length) { openList(); html += inner; }
          else {
            const top = stack.length - 1;
            if (li.lvl > top) { while (stack.length <= li.lvl) openList(); html += inner; }
            else {
              if (li.lvl < top) closeTo(li.lvl + 1);
              const cur = stack[li.lvl];
              if (li.lvl === 0 && (cur.id !== li.id || cur.tag !== tag)) { closeTo(0); openList(); html += inner; }
              else html += `</li><li>${inner}`;
            }
          }
        } else if (nm === 'tbl') { closeTo(0); html += await table(n) + '<p><br></p>'; }
        else if (nm === 'sdt') { closeTo(0); html += await blockContent(l1(n, 'sdtContent') || n); }
        else if (nm === 'sectPr') {
          const pg = l1(n, 'pgSz'), mg = l1(n, 'pgMar');
          if (pg) meta.page = { w: +pg.getAttribute('w:w'), h: +pg.getAttribute('w:h'), landscape: pg.getAttribute('w:orient') === 'landscape' };
          if (mg) meta.margins = { t: +mg.getAttribute('w:top'), r: +mg.getAttribute('w:right'), b: +mg.getAttribute('w:bottom'), l: +mg.getAttribute('w:left') };
          const cols = l1(n, 'cols'); if (cols && cols.getAttribute('w:num')) meta.cols = +cols.getAttribute('w:num');
          const hr = ln(n, 'headerReference')[0], fr = ln(n, 'footerReference')[0];
          meta.hdrRef = hr ? hr.getAttribute('r:id') : null; meta.ftrRef = fr ? fr.getAttribute('r:id') : null;
        }
      }
      closeTo(0);
      return html;
    }
    let html = await blockContent(body);
    const partText = async rid => {
      const r = rels[rid]; if (!r) return { text: '', page: false };
      const f = z['word/' + r.target.replace(/^\/?(word\/)?/, '')]; if (!f) return { text: '', page: false };
      const d = OA.xmlDoc(f); const paras = Array.from(d.getElementsByTagName('w:p')).map(p => Array.from(p.getElementsByTagName('w:t')).map(t => t.textContent).join('')).filter(Boolean);
      const xml = OA.fromUtf8(f);
      return { text: paras.join(' '), page: /PAGE/.test(xml) };
    };
    if (meta.hdrRef) { const h = await partText(meta.hdrRef); meta.header = h.text; if (h.page) meta.pageNum = 'header-right'; }
    if (meta.ftrRef) { const f = await partText(meta.ftrRef); meta.footer = f.text; if (f.page) meta.pageNum = 'footer-center'; }
    const bgc = doc.getElementsByTagName('w:background')[0]; if (bgc && bgc.getAttribute('w:color') && bgc.getAttribute('w:color') !== 'auto') meta.pageColor = bgc.getAttribute('w:color');
    return { html: OA.sanitizeHTML(html), meta };
  };
})();
