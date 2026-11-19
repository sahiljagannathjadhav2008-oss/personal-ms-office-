/* Word app UI — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, el = OA.el;

  const TEMPLATES = {
    letter: { name: 'Letter', html: '<p style="text-align:right">[Your Name]<br>[Street Address]<br>[City, State ZIP]<br>[Date]</p><p>[Recipient Name]<br>[Company]<br>[Address]</p><p>Dear [Recipient],</p><p>I am writing to … Replace this text with the body of your letter. Keep paragraphs short and clear.</p><p>Thank you for your time and consideration.</p><p>Sincerely,<br><br>[Your Name]</p>' },
    resume: { name: 'Resume', html: '<h1 class="title" style="text-align:center">Your Name</h1><p style="text-align:center">City, Country · +91 00000 00000 · you@example.com · linkedin.com/in/you</p><h2>Summary</h2><p>Results-driven professional with experience in … Write 2–3 lines about yourself.</p><h2>Experience</h2><p><b>Job Title</b> — Company <span style="float:right"></span><br><i>Jan 2022 – Present</i></p><ul><li>Achievement with measurable result.</li><li>Responsibility or project you led.</li></ul><h2>Education</h2><p><b>Degree</b> — University, Year</p><h2>Skills</h2><p>Skill 1 · Skill 2 · Skill 3 · Skill 4</p>' },
    report: { name: 'Report', html: '<h1 class="title">Report Title</h1><h2 class="subtitle">Subtitle or short description</h2><p>Prepared by: [Name] &nbsp;|&nbsp; Date: [Date]</p><h1>1. Introduction</h1><p>Describe the purpose and scope of this report.</p><h1>2. Findings</h1><h2>2.1 Key point</h2><p>Explain your findings here.</p><table class="dt" data-b="all" style="width:600px"><colgroup><col style="width:200px"><col style="width:200px"><col style="width:200px"></colgroup><tbody><tr><td><b>Item</b></td><td><b>Value</b></td><td><b>Notes</b></td></tr><tr><td><br></td><td><br></td><td><br></td></tr><tr><td><br></td><td><br></td><td><br></td></tr></tbody></table><p><br></p><h1>3. Conclusion</h1><p>Summarise the main points and next steps.</p>' },
    notes: { name: 'Meeting Notes', html: '<h1>Meeting Notes</h1><p><b>Date:</b> [Date]<br><b>Attendees:</b> [Names]<br><b>Topic:</b> [Topic]</p><h2>Agenda</h2><ol><li>Item one</li><li>Item two</li></ol><h2>Discussion</h2><p><br></p><h2>Action Items</h2><ul><li>[Owner] – task – due date</li></ul>' }
  };

  function htmlToMd(root) {
    const walk = (n, ctx) => {
      if (n.nodeType === 3) return n.nodeValue.replace(/\s+/g, ' ');
      if (n.nodeType !== 1) return '';
      const t = n.tagName, kids = () => Array.from(n.childNodes).map(c => walk(c, ctx)).join('');
      if (n.classList.contains('pb')) return '\n\n---\n\n';
      switch (t) {
        case 'H1': return '\n\n# ' + kids().trim() + '\n\n'; case 'H2': return '\n\n## ' + kids().trim() + '\n\n'; case 'H3': return '\n\n### ' + kids().trim() + '\n\n';
        case 'H4': case 'H5': case 'H6': return '\n\n#### ' + kids().trim() + '\n\n';
        case 'P': case 'DIV': return '\n\n' + kids().trim() + '\n\n'; case 'BR': return '  \n';
        case 'B': case 'STRONG': return '**' + kids() + '**'; case 'I': case 'EM': return '*' + kids() + '*'; case 'S': case 'STRIKE': case 'DEL': return '~~' + kids() + '~~';
        case 'CODE': return '`' + kids() + '`'; case 'PRE': return '\n\n```\n' + n.textContent + '\n```\n\n';
        case 'A': return '[' + kids() + '](' + (n.getAttribute('href') || '') + ')';
        case 'IMG': return '![' + (n.alt || '') + '](' + (n.src.startsWith('data:') ? 'image' : n.src) + ')';
        case 'HR': return '\n\n---\n\n'; case 'BLOCKQUOTE': return '\n\n> ' + kids().trim().replace(/\n/g, '\n> ') + '\n\n';
        case 'UL': case 'OL': { let i = 0; return '\n\n' + Array.from(n.children).map(li => { const sub = walk(li, ctx).trim().replace(/\n/g, '\n  '); return (t === 'OL' ? (++i) + '. ' : '- ') + sub; }).join('\n') + '\n\n'; }
        case 'TABLE': { const rows = Array.from(n.rows).map(r => '| ' + Array.from(r.cells).map(c => c.textContent.trim().replace(/\|/g, '\\|')).join(' | ') + ' |'); if (!rows.length) return ''; const nc = n.rows[0].cells.length; rows.splice(1, 0, '|' + ' --- |'.repeat(nc)); return '\n\n' + rows.join('\n') + '\n\n'; }
        default: return kids();
      }
    };
    return walk(root, {}).replace(/\n{3,}/g, '\n\n').trim() + '\n';
  }
  function mdToHtml(md) {
    const esc = OA.esc; const inl = s => esc(s).replace(/!\[([^\]]*)\]\(([^)]+)\)/g, '<img alt="$1" src="$2">').replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>').replace(/\*\*(.+?)\*\*/g, '<b>$1</b>').replace(/(^|[^*])\*(?!\s)(.+?)\*/g, '$1<i>$2</i>').replace(/~~(.+?)~~/g, '<s>$1</s>').replace(/`([^`]+)`/g, '<code>$1</code>');
    const lines = md.replace(/\r/g, '').split('\n'); let out = '', i = 0;
    while (i < lines.length) {
      const l = lines[i]; let m;
      if (/^```/.test(l)) { let b = ''; i++; while (i < lines.length && !/^```/.test(lines[i])) b += lines[i++] + '\n'; i++; out += '<pre>' + esc(b) + '</pre>'; continue; }
      if ((m = /^(#{1,6})\s+(.*)$/.exec(l))) { out += `<h${m[1].length}>${inl(m[2])}</h${m[1].length}>`; i++; continue; }
      if (/^(-{3,}|\*{3,})\s*$/.test(l)) { out += '<hr>'; i++; continue; }
      if (/^\s*[-*+]\s+/.test(l)) { out += '<ul>'; while (i < lines.length && /^\s*[-*+]\s+/.test(lines[i])) out += '<li>' + inl(lines[i++].replace(/^\s*[-*+]\s+/, '')) + '</li>'; out += '</ul>'; continue; }
      if (/^\s*\d+[.)]\s+/.test(l)) { out += '<ol>'; while (i < lines.length && /^\s*\d+[.)]\s+/.test(lines[i])) out += '<li>' + inl(lines[i++].replace(/^\s*\d+[.)]\s+/, '')) + '</li>'; out += '</ol>'; continue; }
      if (/^>\s?/.test(l)) { let b = ''; while (i < lines.length && /^>\s?/.test(lines[i])) b += lines[i++].replace(/^>\s?/, '') + ' '; out += '<blockquote>' + inl(b.trim()) + '</blockquote>'; continue; }
      if (/^\|.*\|\s*$/.test(l) && /^\|[\s:|-]+\|\s*$/.test(lines[i + 1] || '')) {
        const row = s => s.trim().replace(/^\||\|$/g, '').split('|').map(c => c.trim()); const hd = row(l); i += 2; let rows = '';
        while (i < lines.length && /^\|.*\|\s*$/.test(lines[i])) rows += '<tr>' + row(lines[i++]).map(c => '<td>' + inl(c) + '</td>').join('') + '</tr>';
        out += `<table class="dt" data-b="all"><tbody><tr>${hd.map(c => '<td><b>' + inl(c) + '</b></td>').join('')}</tr>${rows}</tbody></table><p><br></p>`; continue;
      }
      if (!l.trim()) { i++; continue; }
      let b = l; i++; while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|\s*[-*+]\s|\s*\d+[.)]\s|>|\|)/.test(lines[i])) b += '<br>' + inl(lines[i++]).replace(/^/, ''); out += '<p>' + (b === l ? inl(b) : inl(l) + b.slice(l.length)) + '</p>';
    }
    return out;
  }

  OA.registerApp({
    id: 'word', name: 'Word', letter: 'W', color: '#2b579a', accept: ['docx', 'txt', 'html', 'htm', 'md'], defaultName: 'Document1',
    desc: 'Create and edit documents: formatting, styles, tables, pictures, lists, find & replace, export to .docx / PDF.',
    templates: Object.keys(TEMPLATES).map(k => ({ id: k, label: TEMPLATES[k].name })),
    create(root) {
      const a = OA.apps.word; const sh = OA.appShell(root, a);
      let ed, rb, bs, dirty = false;
      const baseName = () => (sh.name.value.trim() || 'Document1').replace(/\.(docx|html?|txt|md)$/i, '');
      const sbWords = el('span'), sbPage = el('span'), sbLang = el('span', { class: 'sb-hide-s', text: 'English (India)' });
      const zoomLbl = el('button', { type: 'button', text: '100%', title: 'Zoom level – click to fit width', onclick: () => ed.fitWidth() });
      const zoomRange = el('input', { type: 'range', min: 25, max: 300, value: 100, 'aria-label': 'Zoom' });
      zoomRange.addEventListener('input', () => ed.setZoom(zoomRange.value / 100));
      const viewBtns = ['print', 'web'].map(v => el('button', { type: 'button', title: v === 'print' ? 'Print Layout' : 'Web Layout', html: OA.icon(v === 'print' ? 'file' : 'columns', 15), onclick: () => ed.setView(v), 'data-v': v }));
      sh.status.append(sbPage, sbWords, sbLang, el('span', { class: 'sb-fill' }), el('span', { class: 'sb-dev sb-hide-s', text: 'Developer: ' + OA.DEV }), ...viewBtns,
        el('button', { type: 'button', html: OA.icon('zoomOut', 15), title: 'Zoom out', onclick: () => ed.setZoom(ed.getZoom() - 0.1) }), zoomRange,
        el('button', { type: 'button', html: OA.icon('zoomIn', 15), title: 'Zoom in', onclick: () => ed.setZoom(ed.getZoom() + 0.1) }), zoomLbl);

      const cb = {
        onChange() { dirty = true; },
        onStats(s) { if (!ed) return; sbWords.textContent = (s.selWords ? s.selWords + ' of ' : '') + s.words + ' word' + (s.words === 1 ? '' : 's'); sbPage.textContent = 'Page ' + ed.currentPage() + ' of ' + s.pages; },
        onZoom(z) { zoomLbl.textContent = Math.round(z * 100) + '%'; zoomRange.value = Math.round(z * 100); },
        onLayout(l) { viewBtns.forEach(b => b.classList.toggle('on', b.dataset.v === l.view)); },
        onState(s) { if (rb) syncState(s); }, onShortcut(k) { if (k === 'link') linkDialog(); else if (k === 'save') doExport('docx'); },
        onNav(v) { rb && rb.setOn('nav', v); }, onComments(v) { rb && rb.setOn('cmts', v); },
        onFormatPainter(v) { rb && rb.setOn('fp', v); }
      };
      ed = OA.createWordEditor(sh.content, cb);
      const P = ed.page;
      const hi = n => ed.setBlockStyle(n);

      /* ------------ dialogs */
      async function linkDialog() {
        const cur = ed.currentLink(); const sel = ed.selectedText();
        const v = await OA.form('Insert Hyperlink', [{ id: 'text', label: 'Text to display', value: cur ? cur.text : sel }, { id: 'url', label: 'Address (URL / email)', value: cur ? cur.href : 'https://' }], { ok: 'OK' });
        if (v && v.url && v.url !== 'https://') ed.insertLink(v.url, v.text);
      }
      async function marginsDialog() {
        const m = ed.layout.margins; const cm = v => +(v / 567).toFixed(2);
        const v = await OA.form('Custom Margins (cm)', [{ id: 't', label: 'Top', type: 'number', value: cm(m.t), step: 0.1, min: 0 }, { id: 'b', label: 'Bottom', type: 'number', value: cm(m.b), step: 0.1, min: 0 }, { id: 'l', label: 'Left', type: 'number', value: cm(m.l), step: 0.1, min: 0 }, { id: 'r', label: 'Right', type: 'number', value: cm(m.r), step: 0.1, min: 0 }]);
        if (v) ed.setLayout({ margins: { t: Math.round(v.t * 567), b: Math.round(v.b * 567), l: Math.round(v.l * 567), r: Math.round(v.r * 567) } });
      }
      async function paraDialog() {
        const v = await OA.form('Paragraph', [{ id: 'l', label: 'Indent left (pt)', type: 'number', value: 0 }, { id: 'r', label: 'Indent right (pt)', type: 'number', value: 0 }, { id: 'f', label: 'First line (pt)', type: 'number', value: 0 }, { id: 'b', label: 'Spacing before (pt)', type: 'number', value: 0 }, { id: 'a', label: 'Spacing after (pt)', type: 'number', value: 8 }, { id: 'ls', label: 'Line spacing', type: 'select', value: '1.15', options: ['1', '1.15', '1.5', '2', '2.5', '3'] }]);
        if (v) { ed.setIndent(v.l, v.r, v.f); ed.setSpacing(v.b, v.a); ed.setLineSpacing(v.ls); }
      }
      async function customTable() {
        const v = await OA.form('Insert Table', [{ id: 'c', label: 'Number of columns', type: 'number', value: 3, min: 1, max: 20 }, { id: 'r', label: 'Number of rows', type: 'number', value: 3, min: 1, max: 100 }]);
        if (v) ed.insertTable(Math.max(1, v.r | 0), Math.max(1, v.c | 0));
      }
      function tablePicker(anchor) {
        const w = el('div', { class: 'gridpick' }); const t = el('div', { class: 'gp-t', text: 'Insert Table' }); const g = el('div', { class: 'gp-g' });
        const R = 8, C = 10; g.style.gridTemplateColumns = `repeat(${C}, 20px)`;
        for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
          const cell = el('div', { class: 'gp-c' }); cell.dataset.r = r + 1; cell.dataset.c = c + 1; g.appendChild(cell);
        }
        const hover = (r, c) => { Array.from(g.children).forEach(x => x.classList.toggle('on', +x.dataset.r <= r && +x.dataset.c <= c)); t.textContent = r && c ? c + ' × ' + r + ' Table' : 'Insert Table'; };
        g.addEventListener('mouseover', e => { if (e.target.dataset.r) hover(+e.target.dataset.r, +e.target.dataset.c); });
        g.addEventListener('click', e => { if (e.target.dataset.r) { OA.closePopup(); ed.insertTable(+e.target.dataset.r, +e.target.dataset.c); } });
        w.append(t, g, el('button', { class: 'mi', type: 'button', onclick: () => { OA.closePopup(); customTable(); } }, el('span', { class: 'mlb', text: 'Insert Table…' })));
        w.addEventListener('mousedown', e => e.preventDefault());
        OA.popup(anchor, w);
      }
      async function headerDialog(kind) {
        const v = await OA.form(kind === 'header' ? 'Header' : 'Footer', [{ id: 't', label: 'Text', value: ed.layout[kind] }]);
        if (v) ed.setLayout({ [kind]: v.t });
      }
      async function imageSizeDialog() {
        const i = ed.imageInfo(); if (!i) return OA.toast('Click a picture first');
        const v = await OA.form('Picture Size', [{ id: 'w', label: 'Width (px)', type: 'number', value: i.w, min: 8 }, { id: 'h', label: 'Height (px)', type: 'number', value: i.h, min: 8 }, { id: 'lock', label: 'Lock aspect ratio', type: 'checkbox', value: true }]);
        if (!v) return; if (v.lock && v.w !== i.w) v.h = Math.round(v.w * i.h / i.w); ed.imageOp('size', { w: v.w, h: v.h });
      }
      function docStatsDialog() {
        const s = ed.stats(); OA.dialog({ title: 'Word Count', buttons: [{ label: 'Close', primary: true }], content: el('div', { class: 'kv', html: `<span>Pages</span><span>${s.pages}</span><span>Words</span><span>${s.words}</span><span>Characters (no spaces)</span><span>${s.charsNoSpace}</span><span>Characters (with spaces)</span><span>${s.chars}</span><span>Paragraphs</span><span>${s.paragraphs}</span><span>Lines (approx.)</span><span>${s.lines}</span>` }) });
      }

      /* ------------ ribbon */
      const fontSel = { type: 'combo', id: 'font', options: OA.wordFonts, width: 130, value: 'Calibri', title: 'Font', inputmode: 'text', onChange: v => v && ed.setFont(v) };
      const sizeSel = { type: 'combo', id: 'size', options: OA.wordSizes.map(String), width: 48, value: '11', title: 'Font size', onChange: v => { const n = parseFloat(v); if (n > 0 && n < 1000) ed.setFontSize(n); } };
      const B = (id, label, icon, onClick, o) => Object.assign({ id, label, icon, title: label, onClick, size: 'sm' }, o || {});
      const L = (id, label, icon, onClick, o) => Object.assign({ id, label, icon, title: label, onClick, size: 'lg' }, o || {});
      const G = (id, g, title, onClick, style) => ({ id, glyph: g, title, onClick, size: 'sm', gstyle: style || '' });
      const styleGal = { type: 'gallery', id: 'styles', items: [
        { label: 'Normal', key: 'normal', sample: 'AaBbCc', onClick: () => hi('normal') },
        { label: 'No Spacing', key: 'nospace', sample: 'AaBbCc', onClick: () => hi('nospace') },
        { label: 'Heading 1', key: 'h1', sample: 'AaBbCc', style: 'font-size:16px;color:#2f5496', onClick: () => hi('h1') },
        { label: 'Heading 2', key: 'h2', sample: 'AaBbCc', style: 'font-size:14px;color:#2f5496', onClick: () => hi('h2') },
        { label: 'Heading 3', key: 'h3', sample: 'AaBbCc', style: 'font-size:13px;color:#1f3763', onClick: () => hi('h3') },
        { label: 'Title', key: 'title', sample: 'AaBb', style: 'font-size:20px', onClick: () => hi('title') },
        { label: 'Subtitle', key: 'subtitle', sample: 'AaBbCc', style: 'color:#5a5a5a', onClick: () => hi('subtitle') },
        { label: 'Quote', key: 'quote', sample: 'AaBbCc', style: 'font-style:italic;color:#404040', onClick: () => hi('quote') }] };
      const pasteMenu = [{ label: 'Paste', icon: 'paste', onClick: () => ed.paste(false) }, { label: 'Paste as Plain Text', onClick: () => ed.paste(true) }];
      const bulletMenu = [{ header: 'Bullet Library' }, ...[['disc', '● Disc'], ['circle', '○ Circle'], ['square', '■ Square']].map(([t, l]) => ({ label: l, onClick: () => ed.list('ul', t) })), { label: 'None (remove list)', onClick: () => ed.cmd('insertUnorderedList') }];
      const numMenu = [{ header: 'Numbering Library' }, ...[['decimal', '1, 2, 3'], ['lower-alpha', 'a, b, c'], ['upper-alpha', 'A, B, C'], ['lower-roman', 'i, ii, iii'], ['upper-roman', 'I, II, III']].map(([t, l]) => ({ label: l, onClick: () => ed.list('ol', t) }))];
      const lineMenu = ['1.0', '1.15', '1.5', '2.0', '2.5', '3.0'].map(v => ({ label: v, onClick: () => ed.setLineSpacing(v) })).concat([{ sep: true }, { label: 'Add Space Before Paragraph', onClick: () => ed.setSpacing(12, null) }, { label: 'Add Space After Paragraph', onClick: () => ed.setSpacing(null, 12) }, { label: 'Remove Space Before/After', onClick: () => ed.setSpacing(0, 0) }, { sep: true }, { label: 'Line Spacing Options…', onClick: paraDialog }]);
      const caseMenu = [['Sentence case.', 'sentence'], ['lowercase', 'lower'], ['UPPERCASE', 'upper'], ['Capitalize Each Word', 'title'], ['tOGGLE cASE', 'toggle']].map(([l, m]) => ({ label: l, onClick: () => ed.changeCase(m) }));
      const shapeMenu = [['Rectangle', 'rect'], ['Rounded Rectangle', 'round'], ['Oval', 'ellipse'], ['Triangle', 'triangle'], ['Diamond', 'diamond'], ['Star', 'star'], ['Right Arrow', 'arrow'], ['Line', 'line']].map(([l, k]) => ({ label: l, onClick: () => ed.insertSVGShape(k) }));
      const marginMenu = [['Normal', 1440, 1440, 1440, 1440], ['Narrow', 720, 720, 720, 720], ['Moderate', 1440, 1080, 1440, 1080], ['Wide', 1440, 2880, 1440, 2880]].map(([l, t, r, b, lf]) => ({ label: l + '  (' + (t / 567).toFixed(1) + ' / ' + (lf / 567).toFixed(1) + ' cm)', onClick: () => ed.setLayout({ margins: { t, r, b, l: lf } }) })).concat([{ sep: true }, { label: 'Custom Margins…', onClick: marginsDialog }]);
      const sizeMenu = () => Object.keys(OA.wordIO.PAGES).map(k => ({ label: k + (k === 'A4' ? '  (21 × 29.7 cm)' : k === 'Letter' ? '  (8.5 × 11 in)' : ''), checked: ed.layout.name === k, onClick: () => ed.setLayout({ name: k, w: OA.wordIO.PAGES[k].w, h: OA.wordIO.PAGES[k].h }) }));
      const colMenu = [[1, 'One'], [2, 'Two'], [3, 'Three']].map(([n, l]) => ({ label: l, onClick: () => ed.setLayout({ cols: n }) }));
      const pnMenu = [['none', 'Remove Page Numbers'], ['header-center', 'Top of Page – Center'], ['header-right', 'Top of Page – Right'], ['footer-center', 'Bottom of Page – Center'], ['footer-right', 'Bottom of Page – Right']].map(([k, l]) => ({ label: l, onClick: () => ed.setLayout({ pageNum: k }) }));
      const tblStyles = [['Plain (no borders)', 'plain'], ['Grid', 'grid'], ['Minimal (outer border)', 'minimal'], ['Blue header', 'blue'], ['Green header', 'green'], ['Orange header', 'orange'], ['Dark header', 'dark'], ['Gray header, banded', 'grayhead'], ['Banded rows', 'banded']].map(([l, k]) => ({ label: l, onClick: () => ed.tableOp('style', k) }));
      const T = (op, arg) => () => ed.tableOp(op, arg);
      const cover = () => ed.insertHTML('<h1 class="title" style="text-align:center;margin-top:160pt">Document Title</h1><h2 class="subtitle" style="text-align:center">Subtitle</h2><p style="text-align:center">Author · Date</p><div class="pb" contenteditable="false"></div><p><br></p>');

      const spec = {
        file: { onClick: () => bs.open('info') },
        tabs: [
          { id: 'home', label: 'Home', groups: [
            { label: 'Clipboard', items: [L('paste', 'Paste', 'paste', () => ed.paste(false), { menu: pasteMenu }), { type: 'row', items: [B('cut', 'Cut', 'cut', () => ed.cut()), B('copy', 'Copy', 'copy', () => ed.copy()), B('fp', 'Format Painter', 'brush', () => ed.copyFormat())] }] },
            { label: 'Font', items: [{ type: 'row', items: [fontSel, sizeSel, B('grow', 'Increase Font Size (Ctrl+])', null, () => ed.stepFont(1), { glyph: 'A▲', gstyle: 'font-size:11px' }), B('shrink', 'Decrease Font Size (Ctrl+[)', null, () => ed.stepFont(-1), { glyph: 'A▼', gstyle: 'font-size:10px' }), B('case', 'Change Case', null, null, { glyph: 'Aa', menu: caseMenu }), B('clear', 'Clear All Formatting', 'eraser', () => ed.clearFormat())] },
              { type: 'row', items: [G('b', 'B', 'Bold (Ctrl+B)', () => ed.cmd('bold'), 'font-weight:700'), G('i', 'I', 'Italic (Ctrl+I)', () => ed.cmd('italic'), 'font-style:italic;font-family:serif'), G('u', 'U', 'Underline (Ctrl+U)', () => ed.cmd('underline'), 'text-decoration:underline'), G('s', 'ab', 'Strikethrough', () => ed.cmd('strikeThrough'), 'text-decoration:line-through'), G('sub', 'x₂', 'Subscript (Ctrl+=)', () => ed.cmd('subscript')), G('sup', 'x²', 'Superscript (Ctrl+Shift+=)', () => ed.cmd('superscript')),
                { type: 'color', id: 'hl', glyph: 'ab', title: 'Text Highlight Color', default: 'FFFF00', noColor: 'No Color', onPick: c => ed.setHighlight(c) }, { type: 'color', id: 'fc', glyph: 'A', title: 'Font Color', default: 'FF0000', noColor: 'Automatic', onPick: c => ed.setColor(c) }] }] },
            { label: 'Paragraph', items: [{ type: 'row', items: [B('ul', 'Bullets', 'list', () => ed.list('ul'), { menu: bulletMenu }), B('ol', 'Numbering', 'listNum', () => ed.list('ol'), { menu: numMenu }), B('outd', 'Decrease Indent', 'outdent', () => ed.indent(-1)), B('ind', 'Increase Indent', 'indent', () => ed.indent(1)), B('sort', 'Sort A→Z', 'sort', () => ed.sortSel(false)), B('marks', 'Show/Hide ¶', null, () => rb.setOn('marks', ed.toggleMarks()), { glyph: '¶' })] },
              { type: 'row', items: [B('al', 'Align Left (Ctrl+L)', 'alignLeft', () => ed.align('left')), B('ac', 'Center (Ctrl+E)', 'alignCenter', () => ed.align('center')), B('ar', 'Align Right (Ctrl+R)', 'alignRight', () => ed.align('right')), B('aj', 'Justify (Ctrl+J)', 'justify', () => ed.align('justify')), B('ls', 'Line and Paragraph Spacing', null, null, { glyph: '↕≡', menu: lineMenu }), B('pd', 'Paragraph Settings…', null, paraDialog, { glyph: '⚙' })] }] },
            { label: 'Styles', cls: 'wide', items: [styleGal] },
            { label: 'Editing', items: [{ type: 'row', items: [B('find', 'Find (Ctrl+F)', 'search', () => ed.showNav('find'), { label: 'Find' })] }, { type: 'row', items: [B('repl', 'Replace (Ctrl+H)', 'search', () => ed.showNav('replace'), { label: 'Replace' })] }, { type: 'row', items: [B('sa', 'Select All', null, () => ed.selectAll(), { glyph: '▣', label: 'Select All' })] }] }] },
          { id: 'insert', label: 'Insert', groups: [
            { label: 'Pages', items: [L('cover', 'Cover Page', 'file', cover), L('pbk', 'Page Break', 'pagebreak', () => ed.pageBreak())] },
            { label: 'Tables', items: [L('tbl', 'Table', 'table', (e, b) => tablePicker(b))] },
            { label: 'Illustrations', items: [L('pic', 'Pictures', 'image', () => ed.insertImageFile()), L('shp', 'Shapes', 'triangle', null, { menu: shapeMenu })] },
            { label: 'Links', items: [L('link', 'Link', 'link', linkDialog), B('bkm', 'Bookmark', null, async () => { const n = await OA.prompt('Bookmark name:', '', { title: 'Bookmark' }); if (n) ed.addBookmark(n.replace(/\s+/g, '_')); }, { glyph: '🔖', label: 'Bookmark' })] },
            { label: 'Header & Footer', items: [B('hdr', 'Header', null, () => headerDialog('header'), { glyph: '▔', label: 'Header' }), B('ftr', 'Footer', null, () => headerDialog('footer'), { glyph: '▁', label: 'Footer' }), B('pn', 'Page Number', null, null, { glyph: '#', label: 'Page Number', menu: pnMenu })] },
            { label: 'Text', items: [B('date', 'Date & Time', null, () => ed.insertDate(), { glyph: '📅', label: 'Date' }), B('time', 'Time', null, () => ed.insertTime(), { glyph: '🕒', label: 'Time' }), B('fn', 'Footnote', null, () => ed.insertFootnote(), { glyph: 'ab¹', label: 'Footnote' })] },
            { label: 'Symbols', items: [B('sym', 'Symbol', 'symbol', (e, b) => ed.symbolPicker(b), { label: 'Symbol' }), B('hr', 'Horizontal Line', 'hr', () => ed.hr(), { label: 'Line' })] },
            { label: 'Comments', items: [L('cmt', 'Comment', 'comment', () => ed.addComment())] }] },
          { id: 'layout', label: 'Layout', groups: [
            { label: 'Page Setup', items: [L('mar', 'Margins', 'margin', null, { menu: marginMenu }), L('ori', 'Orientation', 'orientation', null, { menu: () => [{ label: 'Portrait', checked: !ed.layout.landscape, onClick: () => ed.setLayout({ landscape: false }) }, { label: 'Landscape', checked: ed.layout.landscape, onClick: () => ed.setLayout({ landscape: true }) }] }), L('siz', 'Size', 'file', null, { menu: sizeMenu }), L('col', 'Columns', 'columns', null, { menu: colMenu })] },
            { label: 'Page Background', items: [{ type: 'color', id: 'pgc', icon: 'bucket', title: 'Page Color', noColor: 'No Color', onPick: c => ed.setLayout({ pageColor: c }) }, { type: 'custom', el: el('span', { class: 'rlab', text: 'Page Color' }) }] },
            { label: 'Paragraph', items: [{ type: 'number', id: 'indl', label: 'Indent Left (pt)', value: 0, min: 0, width: 60, onChange: v => ed.setIndent(v, null, null) }, { type: 'number', id: 'indr', label: 'Right (pt)', value: 0, min: 0, width: 60, onChange: v => ed.setIndent(null, v, null) }] },
            { label: 'Spacing', items: [{ type: 'number', id: 'spb', label: 'Before (pt)', value: 0, min: 0, width: 60, onChange: v => ed.setSpacing(v, null) }, { type: 'number', id: 'spa', label: 'After (pt)', value: 8, min: 0, width: 60, onChange: v => ed.setSpacing(null, v) }] }] },
          { id: 'refs', label: 'References', groups: [
            { label: 'Table of Contents', items: [L('toc', 'Table of Contents', 'list', () => ed.insertTOC(3)), B('utoc', 'Update Table', null, () => ed.updateTOC(), { glyph: '↻', label: 'Update' }), B('rtoc', 'Remove Table', null, () => ed.removeTOC(), { glyph: '✕', label: 'Remove' })] },
            { label: 'Footnotes', items: [L('fn2', 'Insert Footnote', 'notes', () => ed.insertFootnote())] }] },
          { id: 'review', label: 'Review', groups: [
            { label: 'Proofing', items: [{ type: 'check', id: 'spell', label: 'Spell check', value: true, onChange: v => ed.setSpellcheck(v) }, L('wc', 'Word Count', 'check2', docStatsDialog)] },
            { label: 'Speech', items: [L('ra', 'Read Aloud', 'play', () => ed.readAloud())] },
            { label: 'Comments', items: [L('nc', 'New Comment', 'comment', () => ed.addComment()), B('cmts', 'Show Comments', 'eye', () => ed.toggleComments(), { label: 'Show' }), B('dc', 'Delete All Comments', 'trash', () => ed.deleteAllComments(), { label: 'Delete all' })] }] },
          { id: 'view', label: 'View', groups: [
            { label: 'Views', items: [L('vp', 'Print Layout', 'file', () => ed.setView('print')), L('vw', 'Web Layout', 'columns', () => ed.setView('web'))] },
            { label: 'Show', items: [{ type: 'check', id: 'navc', label: 'Navigation Pane', value: false, onChange: v => v ? ed.showNav('headings') : ed.hideNav() }, { type: 'check', id: 'marksc', label: 'Formatting marks', value: false, onChange: () => ed.toggleMarks() }, { type: 'check', id: 'ro', label: 'Read-only', value: false, onChange: v => ed.setReadOnly(v) }] },
            { label: 'Zoom', items: [L('z100', '100%', 'zoomIn', () => ed.setZoom(1)), L('zfw', 'Page Width', 'expand', () => ed.fitWidth()), L('zfp', 'One Page', 'file', () => ed.fitPage()), B('zi', 'Zoom In', 'zoomIn', () => ed.setZoom(ed.getZoom() + 0.1), { label: 'In' }), B('zo', 'Zoom Out', 'zoomOut', () => ed.setZoom(ed.getZoom() - 0.1), { label: 'Out' })] }] },
          { id: 'tbl', label: 'Table Layout', contextual: true, hidden: true, groups: [
            { label: 'Table', items: [L('tsel', 'Select', 'table', null, { menu: [{ label: 'Select Cell', onClick: () => ed.selectTablePart('cell') }, { label: 'Select Row', onClick: () => ed.selectTablePart('row') }, { label: 'Select Column', onClick: () => ed.selectTablePart('col') }, { label: 'Select Table', onClick: () => ed.selectTablePart('table') }] }), L('tb', 'Borders', 'border', null, { menu: [{ label: 'All Borders', onClick: T('border', 'all') }, { label: 'Outside Borders', onClick: T('border', 'outer') }, { label: 'No Border', onClick: T('border', 'none') }] })] },
            { label: 'Rows & Columns', items: [{ type: 'row', items: [B('ra', 'Insert Above', null, T('rowAbove'), { glyph: '↑▭', label: 'Above' }), B('rbw', 'Insert Below', null, T('rowBelow'), { glyph: '↓▭', label: 'Below' })] }, { type: 'row', items: [B('cl', 'Insert Left', null, T('colLeft'), { glyph: '←▯', label: 'Left' }), B('cr', 'Insert Right', null, T('colRight'), { glyph: '→▯', label: 'Right' })] }, L('del', 'Delete', 'trash', null, { menu: [{ label: 'Delete Cells…(row)', onClick: T('delRow') }, { label: 'Delete Rows', onClick: T('delRow') }, { label: 'Delete Columns', onClick: T('delCol') }, { label: 'Delete Table', onClick: T('delTable') }] })] },
            { label: 'Merge', items: [L('mg', 'Merge Cells', 'merge', T('merge')), L('sp', 'Split Cells', 'columns', T('split'))] },
            { label: 'Cell Size', items: [L('af', 'AutoFit', 'expand', null, { menu: [{ label: 'AutoFit Contents', onClick: T('autofit') }, { label: 'AutoFit Window', onClick: T('fitwin') }, { label: 'Fixed Column Width', onClick: T('fixed') }] }), L('dist', 'Distribute Columns', 'columns', T('distribute')), { type: 'number', id: 'rh', label: 'Height', value: 24, min: 10, width: 60, onChange: v => ed.tableOp('rowheight', v) }, { type: 'number', id: 'cw', label: 'Width', value: 100, min: 20, width: 60, onChange: v => ed.tableOp('colwidth', v) }] },
            { label: 'Alignment', items: [{ type: 'row', items: [B('tl', 'Align Left', 'alignLeft', T('calign', 'left')), B('tc', 'Center', 'alignCenter', T('calign', 'center')), B('tr', 'Align Right', 'alignRight', T('calign', 'right'))] }, { type: 'row', items: [B('vt', 'Align Top', null, T('valign', 'top'), { glyph: '⤒' }), B('vm', 'Align Middle', null, T('valign', 'middle'), { glyph: '↕' }), B('vb', 'Align Bottom', null, T('valign', 'bottom'), { glyph: '⤓' })] }, { type: 'row', items: [B('ta', 'Table left', null, T('talign', 'left'), { glyph: '⇤' }), B('tm', 'Table center', null, T('talign', 'center'), { glyph: '↔' }), B('trr', 'Table right', null, T('talign', 'right'), { glyph: '⇥' })] }] },
            { label: 'Table Styles', items: [L('tst', 'Table Styles', 'table', null, { menu: tblStyles }), { type: 'color', id: 'shade', icon: 'bucket', title: 'Shading', noColor: 'No Color', onPick: c => ed.tableOp('shade', c) }, B('hr1', 'Header Row Bold', null, T('header', true), { glyph: 'H', label: 'Header' })] },
            { label: 'Data', items: [L('ts', 'Sort column', 'sort', T('sort', false))] }] },
          { id: 'pic', label: 'Picture Format', contextual: true, hidden: true, groups: [
            { label: 'Size', items: [L('psz', 'Size…', 'expand', imageSizeDialog)] },
            { label: 'Arrange', items: [{ type: 'row', items: [B('pi', 'In line with text', null, () => ed.imageOp('inline'), { glyph: '▭', label: 'Inline' }), B('pl', 'Align left', 'alignLeft', () => ed.imageOp('left')), B('pc', 'Center', 'alignCenter', () => ed.imageOp('center')), B('pr', 'Align right', 'alignRight', () => ed.imageOp('right'))] }, { type: 'row', items: [B('pfl', 'Wrap text – picture left', null, () => ed.imageOp('floatL'), { glyph: '◧', label: 'Wrap left' }), B('pfr', 'Wrap text – picture right', null, () => ed.imageOp('floatR'), { glyph: '◨', label: 'Wrap right' })] }] },
            { label: 'Adjust', items: [{ type: 'row', items: [B('rl', 'Rotate left 90°', 'rotate', () => ed.imageOp('rotate', -90), { label: 'Rotate −90' }), B('rr', 'Rotate right 90°', 'rotate', () => ed.imageOp('rotate', 90), { label: 'Rotate +90' }), B('fh', 'Flip horizontal', 'flipH', () => ed.imageOp('flipH'), { label: 'Flip' })] }, { type: 'row', items: [{ type: 'color', id: 'pbc', icon: 'border', title: 'Picture Border', noColor: 'No Border', onPick: c => ed.imageOp('border', c) }, B('op', 'Transparency', null, async () => { const v = await OA.form('Transparency', [{ id: 'o', label: 'Opacity (%)', type: 'number', value: 100, min: 10, max: 100 }]); if (v) ed.imageOp('opacity', v.o / 100); }, { glyph: '◐', label: 'Opacity' })] }] },
            { label: 'Accessibility', items: [L('alt', 'Alt Text', 'info', async () => { const i = ed.imageInfo(); const t = await OA.prompt('Alt text (describe the picture):', i ? i.alt : '', { title: 'Alt Text' }); if (t != null) ed.imageOp('alt', t); })] },
            { label: 'Delete', items: [L('pdl', 'Delete picture', 'trash', () => ed.imageOp('delete'))] }] }
        ]
      };
      const right = el('div', { class: 'rtab-right' });
      spec.right = right;
      rb = OA.buildRibbon(sh.ribbon, spec);
      const ctl = rb.ctl;
      const setOn = rb.setOn;

      function syncState(s) {
        if (!s) return;
        ['b', 'i', 'u', 's', 'sub', 'sup', 'ul', 'ol'].forEach(k => setOn(k, s[k]));
        setOn('al', s.align === 'left'); setOn('ac', s.align === 'center'); setOn('ar', s.align === 'right'); setOn('aj', s.align === 'justify');
        if (document.activeElement !== ctl.font && s.font) ctl.font.value = s.font;
        if (document.activeElement !== ctl.size && s.size) ctl.size.value = String(s.size);
        if (ctl.styles) Array.from(ctl.styles.children).forEach(b => b.classList.toggle('on', b.dataset.k === s.block));
        rb.setTabVisible('tbl', !!s.inTable); rb.setTabVisible('pic', !!s.img);
        if (s.img && rb.active !== 'pic' && !rb._pic) { rb.showTab('pic'); rb._pic = true; } if (!s.img) rb._pic = false;
        if (s.inTable && rb.active !== 'tbl' && !rb._tbl) { rb.showTab('tbl'); rb._tbl = true; } if (!s.inTable) rb._tbl = false;
      }

      /* ------------ file I/O */
      const stamp = () => baseName();
      async function doExport(kind) {
        const n = stamp();
        try {
          if (kind === 'docx') { OA.download(await ed.exportDocx(n), n + '.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'); OA.toast('Exported ' + n + '.docx'); }
          else if (kind === 'html') { OA.download('<!DOCTYPE html><html><head><meta charset="utf-8"><title>' + OA.esc(n) + '</title><style>body{font-family:Calibri,Arial,sans-serif;max-width:800px;margin:2em auto;line-height:1.4}table{border-collapse:collapse}td,th{border:1px solid #888;padding:4px 8px}img{max-width:100%}</style></head><body>' + ed.getHTML() + '</body></html>', n + '.html', 'text/html'); OA.toast('Exported ' + n + '.html'); }
          else if (kind === 'txt') { OA.download(ed.getText(), n + '.txt', 'text/plain'); OA.toast('Exported ' + n + '.txt'); }
          else if (kind === 'md') { const d = document.createElement('div'); d.innerHTML = ed.getHTML(); OA.download(htmlToMd(d), n + '.md', 'text/markdown'); OA.toast('Exported ' + n + '.md'); }
          else if (kind === 'pdf') { OA.toast('In the print dialog choose "Save as PDF"'); ed.print(); }
          dirty = false;
        } catch (e) { console.error(e); OA.alert('Export failed: ' + OA.esc(e.message)); }
      }
      async function openFile(file) {
        try {
          const ext = OA.ext(file.name); sh.name.value = OA.baseName(file.name);
          if (ext === 'docx') { const r = await OA.wordIO.fromDocx(new Uint8Array(await file.arrayBuffer())); ed.setContent(r.html, r.meta); }
          else if (ext === 'html' || ext === 'htm') { const t = await file.text(); const m = /<body[^>]*>([\s\S]*)<\/body>/i.exec(t); ed.setContent(m ? m[1] : t); }
          else if (ext === 'md') ed.setContent(mdToHtml(await file.text()));
          else ed.setContent((await file.text()).replace(/\r/g, '').split('\n').map(l => '<p>' + (OA.esc(l) || '<br>') + '</p>').join(''));
          dirty = false; OA.toast('Opened ' + file.name);
        } catch (e) { console.error(e); OA.alert('Could not open this file: ' + OA.esc(e.message)); }
      }
      function newDoc(tpl) {
        ed.setContent(tpl ? TEMPLATES[tpl].html : '<p><br></p>'); sh.name.value = 'Document1'; dirty = false;
      }

      /* ------------ backstage */
      bs = OA.backstage(root, { items: [
        { id: 'info', label: 'Info', icon: 'info', render(p) {
          p.appendChild(el('p', { class: 'bs-p', text: 'Your document lives only in this browser tab. Export it to keep it.' }));
          const s = ed.stats();
          p.appendChild(el('div', { class: 'kv', html: `<span>Name</span><span>${OA.esc(baseName())}</span><span>Pages</span><span>${s.pages}</span><span>Words</span><span>${s.words}</span><span>Page size</span><span>${ed.layout.name}${ed.layout.landscape ? ' (landscape)' : ''}</span>` }));
          p.appendChild(el('div', { style: { marginTop: '18px' } }, el('button', { class: 'btn primary', type: 'button', text: 'Export as .docx', onclick: () => doExport('docx') })));
        } },
        { id: 'new', label: 'New', icon: 'plus', render(p, api) {
          p.appendChild(OA.bigActions([{ label: 'Blank document', desc: 'Start with an empty page', icon: 'file', onClick: async () => { if (!dirty || await OA.confirm('Discard the current document and start a new one?', { ok: 'New document' })) { newDoc(); api.close(); } } },
            ...Object.keys(TEMPLATES).map(k => ({ label: TEMPLATES[k].name, desc: 'Template', icon: 'file', onClick: async () => { if (!dirty || await OA.confirm('Discard the current document?', { ok: 'Continue' })) { newDoc(k); api.close(); } } }))]));
        } },
        { id: 'open', label: 'Open', icon: 'folder', render(p, api) {
          p.appendChild(el('p', { class: 'bs-p', text: 'Open a .docx, .html, .md or .txt file from your device.' }));
          p.appendChild(el('button', { class: 'btn primary', type: 'button', text: 'Browse…', onclick: async () => { const f = await OA.pickFile('.docx,.html,.htm,.md,.txt'); if (f) { api.close(); openFile(f); } } }));
        } },
        { id: 'export', label: 'Export / Save As', icon: 'download', render(p) {
          p.appendChild(el('p', { class: 'bs-p', text: 'Choose a format to download. Data is not stored on this website.' }));
          p.appendChild(OA.bigActions([{ label: 'Word Document (.docx)', desc: 'Opens in Microsoft Word, LibreOffice, Google Docs', icon: 'download', onClick: () => doExport('docx') }, { label: 'PDF', desc: 'Use "Save as PDF" in the print dialog', icon: 'print', onClick: () => doExport('pdf') }, { label: 'Web Page (.html)', desc: 'Single HTML file', icon: 'download', onClick: () => doExport('html') }, { label: 'Plain Text (.txt)', desc: 'Text only, no formatting', icon: 'download', onClick: () => doExport('txt') }, { label: 'Markdown (.md)', desc: 'Headings, lists, tables, links', icon: 'download', onClick: () => doExport('md') }]));
        } },
        { id: 'print', label: 'Print', icon: 'print', action: () => ed.print() },
        { sep: true },
        { id: 'home', label: 'Close (Office Home)', icon: 'home', action: () => OA.go('home') }
      ] });
      sh.title.querySelector('.tb-right').append(el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('undo', 16), title: 'Undo (Ctrl+Z)', onclick: () => ed.undo() }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('redo', 16), title: 'Redo (Ctrl+Y)', onclick: () => ed.redo() }), el('button', { class: 'tb-btn2', type: 'button', html: OA.icon('download', 16), title: 'Export .docx (Ctrl+S)', onclick: () => doExport('docx') }));

      ed.setZoom(window.innerWidth < 760 ? Math.max(0.4, (window.innerWidth - 16) / 794) : 1);
      return {
        onShow() { setTimeout(() => { ed.setZoom(ed.getZoom()); P.focus({ preventScroll: true }); }, 50); },
        openFile, exportDefault: () => doExport('docx'), loadTemplate: k => newDoc(k)
      };
    }
  });
})();
