/* Word extras 1: Review (track changes, compare, thesaurus, translate), Mailings, References — Developer: Sahil Jadhav */
(function () {
  'use strict';
  const OA = window.OA, el = OA.el;
  OA.wordExtras = OA.wordExtras || [];

  const THES = `big: large, huge, enormous, vast, substantial|small: little, tiny, minor, compact, slight|good: great, fine, excellent, superb, positive|bad: poor, awful, terrible, harmful, inferior|happy: glad, joyful, cheerful, pleased, content|sad: unhappy, sorrowful, gloomy, downcast, miserable|fast: quick, rapid, swift, speedy, brisk|slow: gradual, unhurried, sluggish, leisurely|important: significant, crucial, vital, essential, key|easy: simple, effortless, straightforward, painless|hard: difficult, tough, challenging, demanding, firm|begin: start, commence, initiate, launch, open|end: finish, conclude, stop, terminate, complete|show: display, present, reveal, demonstrate, exhibit|help: assist, aid, support, back, serve|use: employ, utilize, apply, operate|make: create, build, produce, form, construct|get: obtain, acquire, receive, gain, fetch|give: provide, offer, grant, supply, donate|say: state, tell, declare, remark, mention|think: believe, consider, suppose, reckon, ponder|look: see, watch, observe, glance, gaze|buy: purchase, acquire, procure|old: aged, ancient, elderly, vintage, former|new: fresh, modern, novel, recent, original|beautiful: pretty, lovely, attractive, gorgeous, stunning|ugly: unattractive, hideous, unsightly|smart: clever, intelligent, bright, sharp, wise|stupid: foolish, silly, dull, unwise|angry: mad, furious, irate, upset, annoyed|clean: spotless, neat, tidy, pure, fresh|dirty: filthy, grubby, soiled, unclean|rich: wealthy, affluent, prosperous, well-off|poor: needy, broke, impoverished, penniless|strong: powerful, sturdy, robust, tough, solid|weak: feeble, frail, fragile, delicate|brave: courageous, bold, fearless, daring|afraid: scared, fearful, frightened, anxious|tired: weary, exhausted, sleepy, fatigued|funny: amusing, humorous, comical, hilarious|problem: issue, difficulty, trouble, obstacle, challenge|idea: concept, thought, notion, plan, suggestion|answer: reply, response, solution, resolution|question: query, inquiry, doubt, problem|change: alter, modify, adjust, transform, vary|increase: grow, rise, expand, boost, raise|decrease: reduce, lower, shrink, decline, cut|choose: select, pick, opt, elect, decide|keep: retain, hold, save, preserve, maintain|need: require, want, demand, lack|want: desire, wish, crave, need, seek|like: enjoy, love, prefer, fancy, appreciate|hate: detest, despise, loathe, dislike|talk: speak, chat, converse, discuss, communicate|walk: stroll, stride, march, wander, hike|run: sprint, dash, jog, race, rush|eat: consume, devour, dine, munch|drink: sip, gulp, swallow, sup|fix: repair, mend, restore, correct|break: smash, shatter, crack, destroy|try: attempt, endeavor, strive, test|ask: inquire, request, query, question|tell: inform, notify, advise, instruct|part: piece, portion, section, segment, share|whole: entire, complete, full, total|many: numerous, several, countless, plenty|few: some, scarce, rare, handful|often: frequently, regularly, commonly, usually|always: constantly, forever, continually, invariably|never: not ever, at no time, rarely|very: extremely, highly, really, truly, greatly|also: too, additionally, moreover, besides, furthermore|but: however, yet, although, still, nevertheless|because: since, as, due to, owing to|so: therefore, thus, hence, consequently|get up: rise, arise, stand|job: work, task, occupation, career, position|boss: manager, supervisor, leader, chief|friend: companion, buddy, ally, pal, mate|enemy: foe, rival, opponent, adversary|house: home, residence, dwelling, abode|car: vehicle, automobile, auto|road: street, path, route, highway, lane|city: town, metropolis, municipality|country: nation, state, land, homeland|world: earth, globe, planet, universe|time: period, moment, era, duration, age|place: location, spot, site, position, area|way: method, manner, approach, path, route|result: outcome, effect, consequence, product|reason: cause, motive, basis, grounds|example: instance, sample, case, illustration|danger: risk, threat, hazard, peril|safe: secure, protected, harmless, sound|quiet: silent, calm, peaceful, hushed|loud: noisy, deafening, booming, blaring|bright: shiny, brilliant, radiant, luminous, vivid|dark: dim, gloomy, shadowy, black, murky|hot: warm, burning, scorching, boiling|cold: chilly, cool, freezing, icy, frosty|wet: damp, moist, soaked, soggy|dry: arid, parched, dehydrated|full: packed, loaded, crowded, complete|empty: vacant, blank, hollow, bare|right: correct, accurate, proper, true|wrong: incorrect, mistaken, false, erroneous|true: real, genuine, accurate, factual|false: untrue, fake, incorrect, bogus|main: primary, chief, principal, central|enough: sufficient, adequate, ample|almost: nearly, approximately, virtually|maybe: perhaps, possibly, perchance|really: truly, genuinely, actually|quickly: rapidly, swiftly, promptly, hastily|suddenly: abruptly, unexpectedly, instantly|carefully: cautiously, attentively, thoroughly|nice: pleasant, kind, agreeable, lovely|great: excellent, wonderful, superb, terrific|interesting: fascinating, engaging, intriguing, absorbing|boring: dull, tedious, monotonous, uninteresting|improve: enhance, better, upgrade, refine, advance|develop: grow, evolve, create, build, advance|create: make, produce, generate, design, form|explain: clarify, describe, illustrate, justify|describe: portray, depict, outline, detail|suggest: propose, recommend, advise, hint|decide: determine, resolve, settle, conclude|discover: find, uncover, detect, learn|understand: comprehend, grasp, realize, follow|remember: recall, recollect, retain, memorize|forget: overlook, neglect, omit|learn: study, master, absorb, discover|teach: instruct, educate, train, coach|write: compose, pen, draft, record|read: peruse, scan, study, browse|send: dispatch, mail, forward, transmit|receive: get, accept, obtain, collect|move: shift, relocate, transfer, go|stay: remain, linger, wait, stop`;
  const thes = {}; THES.split('|').forEach(l => { const [w, s] = l.split(':'); thes[w.trim()] = s.split(',').map(x => x.trim()); });
  Object.keys(thes).forEach(k => thes[k].forEach(s => { if (!thes[s]) thes[s] = [k].concat(thes[k].filter(x => x !== s)); }));

  function csvParse(t, d) { return (window.XL && XL.parseCSV) ? XL.parseCSV(t, d) : t.split(/\r?\n/).map(l => l.split(d)); }
  function wordDiff(a, b) { // returns array of [op, text]; op: 0 same, 1 ins, -1 del
    const A = a.split(/(\s+)/).filter(x => x !== ''), B = b.split(/(\s+)/).filter(x => x !== '');
    const n = A.length, m = B.length; if (n * m > 4e6) return [[-1, a], [1, b]];
    const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const out = []; let i = 0, j = 0; const push = (op, t) => { const l = out[out.length - 1]; if (l && l[0] === op) l[1] += t; else out.push([op, t]); };
    while (i < n && j < m) { if (A[i] === B[j]) { push(0, A[i]); i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) push(-1, A[i++]); else push(1, B[j++]); }
    while (i < n) push(-1, A[i++]); while (j < m) push(1, B[j++]);
    return out;
  }

  OA.wordExtras.push(function (ed, H, spec) {
    const P = ed.page, { B, L } = H;
    const sel = () => { const s = getSelection(); return s.rangeCount && P.contains(s.getRangeAt(0).commonAncestorContainer) ? s.getRangeAt(0) : null; };
    const fire = () => P.dispatchEvent(new Event('input', { bubbles: true }));
    const X = ed.x = ed.x || {};
    X.insInline = html => ed.mutate(() => {
      const s = getSelection(); if (!s.rangeCount) return; const r = s.getRangeAt(0); r.deleteContents();
      const t = document.createElement('template'); t.innerHTML = html; const frag = t.content; const last = frag.lastChild; r.insertNode(frag);
      if (last) { const nr = document.createRange(); nr.setStartAfter(last); nr.collapse(true); s.removeAllRanges(); s.addRange(nr); }
    });
    X.insBlock = html => ed.mutate(() => {
      const s = getSelection(); let n = s.rangeCount ? s.getRangeAt(0).endContainer : null; if (n && n.nodeType === 3) n = n.parentNode;
      while (n && n.parentNode !== P) n = n.parentNode; if (!n || !P.contains(n)) n = P.lastElementChild;
      const t = document.createElement('div'); t.innerHTML = html; const nodes = Array.from(t.childNodes); let ref = n;
      const empty = n && n.tagName === 'P' && !n.textContent.trim() && !n.querySelector('img');
      nodes.forEach(x => { ref.after(x); ref = x; }); if (empty && nodes.length) n.remove();
      if (!ref.nextSibling || ref.nextSibling.nodeType !== 1) { const p = document.createElement('p'); p.innerHTML = '<br>'; ref.after(p); }
    });
    const tabAdd = (id, groups, before) => { const t = spec.tabs.find(x => x.id === id); if (t) t.groups.push(...groups); };

    /* ======================================================== TRACK CHANGES */
    const tc = { on: false, author: 'You' };
    const stamp = () => new Date().toLocaleString();
    const mkIns = txt => { const i = document.createElement('ins'); i.className = 'tc-ins'; i.dataset.by = tc.author; i.dataset.at = stamp(); i.title = tc.author + ' · ' + i.dataset.at; i.textContent = txt; return i; };
    function textNodesIn(r) {
      const out = []; const root = r.commonAncestorContainer.nodeType === 3 ? r.commonAncestorContainer.parentNode : r.commonAncestorContainer;
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT); while (w.nextNode()) { const n = w.currentNode; if (r.intersectsNode(n) && n.nodeValue) out.push(n); } return out;
    }
    function markDeleted(r, back) {
      const nodes = textNodesIn(r), sc = r.startContainer, so = r.startOffset, ec = r.endContainer, eo = r.endOffset; let first = null, last = null;
      const caret = document.createRange(); caret.setStart(sc, so); caret.collapse(true);
      nodes.forEach(n => {
        const s = n === sc ? so : 0, e = n === ec ? eo : n.nodeValue.length; if (s >= e) return;
        let mid = n; if (e < n.nodeValue.length) n.splitText(e); if (s > 0) mid = n.splitText(s);
        const p = mid.parentNode; if (p.closest && p.closest('del.tc-del')) return;
        if (p.closest && p.closest('ins.tc-ins')) { const a = mid.previousSibling, b = mid.nextSibling; mid.remove(); if (!first) first = a || null; last = b || a || last; return; }
        const d = document.createElement('del'); d.className = 'tc-del'; d.dataset.by = tc.author; d.dataset.at = stamp(); d.title = tc.author + ' · ' + d.dataset.at;
        p.insertBefore(d, mid); d.appendChild(mid); if (!first) first = d; last = d;
      });
      const s = getSelection(), nr = document.createRange();
      if (back && first && first.parentNode) nr.setStartBefore(first); else if (last && last.parentNode) nr.setStartAfter(last); else nr.setStart(caret.startContainer, caret.startOffset);
      nr.collapse(true); s.removeAllRanges(); s.addRange(nr);
    }
    P.addEventListener('beforeinput', e => {
      if (!tc.on) return; const r = sel(); if (!r) return; const t = e.inputType; const s = getSelection();
      const doInsert = txt => {
        if (!r.collapsed) markDeleted(r, false);
        const r2 = s.getRangeAt(0); const host = r2.startContainer.nodeType === 3 ? r2.startContainer.parentNode : r2.startContainer;
        const ins = host.closest && host.closest('ins.tc-ins');
        if (ins) { document.execCommand('insertText', false, txt); return; }
        const i = mkIns(txt); r2.insertNode(i); const nr = document.createRange(); nr.setStart(i.firstChild, i.firstChild.length); nr.collapse(true); s.removeAllRanges(); s.addRange(nr);
      };
      if (t === 'insertText' && e.data) { e.preventDefault(); doInsert(e.data); fire(); }
      else if (t === 'insertFromPaste' && e.dataTransfer && e.dataTransfer.getData('text/plain')) { e.preventDefault(); doInsert(e.dataTransfer.getData('text/plain')); fire(); }
      else if (/^delete/.test(t)) {
        e.preventDefault(); let rr = r;
        if (r.collapsed) {
          const back = /Backward/.test(t), unit = /Word/.test(t) ? 'word' : /Line/.test(t) ? 'lineboundary' : 'character';
          s.modify('extend', back ? 'backward' : 'forward', unit); rr = s.getRangeAt(0);
          if (rr.collapsed) return;
        }
        markDeleted(rr, !/Forward/.test(t)); fire();
      }
    });
    function changes() { return Array.from(P.querySelectorAll('ins.tc-ins,del.tc-del')); }
    function resolve(n, accept) {
      const isIns = n.tagName === 'INS'; if (isIns === accept) n.replaceWith(...n.childNodes); else n.remove();
    }
    X.setTrack = v => { tc.on = !!v; P.classList.toggle('tc-on', tc.on); OA.toast('Track Changes ' + (tc.on ? 'on' : 'off')); };
    X.trackAuthor = async () => { const n = await OA.prompt('Your name (shown on tracked changes):', tc.author, { title: 'User Name' }); if (n) tc.author = n; };
    X.showMarkup = v => P.classList.toggle('tc-hide', !v);
    X.accept = (all, acc = true) => ed.mutate(() => {
      let list; if (all) list = changes(); else { const r = sel(); list = changes().filter(n => r && (r.intersectsNode(n) || n.contains(r.startContainer))); if (!list.length) { const nx = nextChange(1); list = nx ? [nx] : []; } }
      list.forEach(n => resolve(n, acc)); if (!list.length) OA.toast('No tracked changes'); return list.length;
    });
    function nextChange(dir) {
      const cs = changes(); if (!cs.length) return null; const r = sel(); if (!r) return cs[0];
      if (dir > 0) return cs.find(n => (r.comparePoint(n, 0) > 0) || (r.compareBoundaryPoints(Range.START_TO_START, rangeOf(n)) < 0)) || cs[0];
      const rev = cs.slice().reverse(); return rev.find(n => r.compareBoundaryPoints(Range.START_TO_START, rangeOf(n)) > 0) || rev[0];
    }
    const rangeOf = n => { const r = document.createRange(); r.selectNode(n); return r; };
    X.gotoChange = dir => { const n = nextChange(dir); if (!n) return OA.toast('No tracked changes'); const r = document.createRange(); r.selectNodeContents(n); const s = getSelection(); s.removeAllRanges(); s.addRange(r); n.scrollIntoView({ block: 'center' }); };
    X.reviewPane = () => {
      const cs = changes(); const body = el('div', { style: { maxHeight: '55vh', overflow: 'auto', minWidth: '300px' } });
      if (!cs.length) body.appendChild(el('div', { class: 'dlg-msg', text: 'No tracked changes in this document.' }));
      cs.forEach(n => body.appendChild(el('div', { style: { padding: '6px 0', borderBottom: '1px solid #eee', fontSize: '13px' } }, el('b', { text: (n.tagName === 'INS' ? 'Inserted' : 'Deleted') + ' · ' + (n.dataset.by || '') }), el('div', { style: { color: '#666', fontSize: '11px' }, text: n.dataset.at || '' }), el('div', { text: n.textContent.slice(0, 120) }))));
      OA.dialog({ title: 'Reviewing Pane (' + cs.length + ' changes)', content: body, buttons: [{ label: 'Accept all', onClick: () => { X.accept(true, true); } }, { label: 'Reject all', onClick: () => { X.accept(true, false); } }, { label: 'Close', primary: true }] });
    };

    /* ======================================================== COMPARE */
    X.compare = async () => {
      const f = await OA.pickFile('.docx,.html,.htm,.txt,.md'); if (!f) return;
      let other = ''; const ext = OA.ext(f.name);
      try {
        if (ext === 'docx') { const r = await OA.wordIO.fromDocx(new Uint8Array(await f.arrayBuffer())); const d = document.createElement('div'); d.innerHTML = r.html; other = blockText(d); }
        else if (ext === 'html' || ext === 'htm') { const d = document.createElement('div'); d.innerHTML = await f.text(); other = blockText(d); }
        else other = (await f.text()).replace(/\r/g, '');
      } catch (e) { return OA.alert('Could not read that file: ' + OA.esc(e.message)); }
      const ok = await OA.confirm('Compare the current document (original) with <b>' + OA.esc(f.name) + '</b> (revised)? The result replaces the current content with tracked changes (use Undo to go back).', { ok: 'Compare' }); if (!ok) return;
      const cur = blockText(P).split('\n'), oth = other.split('\n');
      const diff = wordDiff(cur.join('\n'), oth.join('\n')); let html = '<p>';
      diff.forEach(([op, t]) => {
        const parts = t.split('\n');
        parts.forEach((seg, i) => { if (i > 0) html += '</p><p>'; if (!seg) return; const e = OA.esc(seg); html += op === 0 ? e : op === 1 ? `<ins class="tc-ins" data-by="Compare" data-at="${stamp()}">${e}</ins>` : `<del class="tc-del" data-by="Compare" data-at="${stamp()}">${e}</del>`; });
      });
      html += '</p>'; ed.mutate(() => { P.innerHTML = html.replace(/<p><\/p>/g, '<p><br></p>'); });
      OA.toast(changes().length + ' differences marked');
    };
    function blockText(root) { return Array.from(root.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')).filter(b => !b.closest('.toc,.fnote') && !(b.parentNode.closest && b.parentNode.closest('li') && b.tagName === 'P' && false)).map(b => b.textContent).join('\n'); }
    X.restrict = async () => {
      const v = await OA.form('Restrict Editing', [{ id: 'm', label: 'Allow only this type of editing', type: 'select', value: 'read', options: [['none', 'No restrictions'], ['read', 'No changes (Read only)'], ['track', 'Tracked changes only']] }]);
      if (!v) return; ed.setReadOnly(v.m === 'read'); if (v.m === 'track') X.setTrack(true);
      const ro = document.querySelector('.app-word input#ro, .app-word [data-id=ro] input'); if (ro) ro.checked = v.m === 'read';
      OA.toast(v.m === 'none' ? 'Editing unrestricted' : v.m === 'read' ? 'Document is read-only' : 'Only tracked changes recommended — Track Changes is on');
    };

    /* ======================================================== THESAURUS / TRANSLATE */
    X.thesaurus = async () => {
      const r = sel(); let w = (r && !r.collapsed ? r.toString() : '').trim().toLowerCase();
      if (!w) { const v = await OA.prompt('Look up a word:', '', { title: 'Thesaurus' }); if (!v) return; w = v.trim().toLowerCase(); }
      const list = thes[w] || thes[w.replace(/(ing|ed|es|s)$/, '')] || thes[w.replace(/(ing|ed)$/, 'e')];
      const body = el('div', { style: { minWidth: '260px' } }, el('div', { class: 'dlg-lbl', text: 'Synonyms for “' + w + '”' }));
      if (!list) body.appendChild(el('div', { class: 'dlg-msg', text: 'No entry in the built-in offline thesaurus.' }));
      else list.forEach(s => body.appendChild(el('button', { class: 'btn', type: 'button', style: { margin: '3px' }, text: s, onclick: function () { dlg.close(); if (r && !r.collapsed) { ed.insertText(s); } else { navigator.clipboard && navigator.clipboard.writeText(s); OA.toast('Copied “' + s + '”'); } } })));
      const dlg = OA.dialog({ title: 'Thesaurus', content: body, buttons: [{ label: 'Close', primary: true }] });
    };
    X.translate = async () => {
      const r = sel(); const txt = r && !r.collapsed ? r.toString() : P.innerText.slice(0, 500);
      const v = await OA.form('Translate', [{ id: 'to', label: 'Translate to', type: 'select', value: 'hi', options: [['hi', 'Hindi'], ['mr', 'Marathi'], ['gu', 'Gujarati'], ['bn', 'Bengali'], ['ta', 'Tamil'], ['te', 'Telugu'], ['en', 'English'], ['es', 'Spanish'], ['fr', 'French'], ['de', 'German'], ['it', 'Italian'], ['pt', 'Portuguese'], ['ru', 'Russian'], ['ja', 'Japanese'], ['ko', 'Korean'], ['zh-CN', 'Chinese'], ['ar', 'Arabic']] }, { id: 'text', label: 'Text (max ~500 characters)', type: 'textarea', value: txt.slice(0, 500) }], { ok: 'Translate', width: 420 });
      if (!v || !v.text.trim()) return;
      OA.toast('Translating…');
      try {
        const rs = await fetch('https://api.mymemory.translated.net/get?q=' + encodeURIComponent(v.text) + '&langpair=Autodetect|' + v.to); const j = await rs.json();
        const out = j.responseData && j.responseData.translatedText; if (!out) throw new Error('No result');
        OA.dialog({ title: 'Translation', content: el('div', { style: { whiteSpace: 'pre-wrap', maxWidth: '420px' }, text: out }), buttons: [{ label: 'Close' }, { label: 'Insert after selection', primary: true, onClick: () => { ed.insertText(' ' + out); } }] });
      } catch (e) { OA.alert('Translation needs an internet connection (uses the free MyMemory service).<br><small>' + OA.esc(e.message) + '</small>'); }
    };

    /* ======================================================== MAIL MERGE */
    const mm = { fields: [], rows: [], idx: 0, preview: false };
    const mfields = () => Array.from(P.querySelectorAll('span.mf'));
    function renderFields() { mfields().forEach(s => { const f = s.dataset.f; s.textContent = mm.preview && mm.rows[mm.idx] ? (mm.rows[mm.idx][f] == null ? '' : mm.rows[mm.idx][f]) : '«' + f + '»'; s.classList.toggle('mfv', mm.preview); }); }
    function setData(rows2d) {
      rows2d = rows2d.filter(r => r.some(c => String(c).trim() !== '')); if (rows2d.length < 2) return false;
      mm.fields = rows2d[0].map(x => String(x).trim()); mm.rows = rows2d.slice(1).map(r => { const o = {}; mm.fields.forEach((f, i) => o[f] = r[i] == null ? '' : String(r[i])); return o; }); mm.idx = 0; return true;
    }
    X.mmSelect = async (anchor) => {
      OA.menu(anchor, [
        { label: 'Type a New List / Paste CSV…', onClick: async () => { const v = await OA.form('Mailing list', [{ id: 't', label: 'First line = field names (comma or tab separated)', type: 'textarea', rows: 8, value: 'Name,Address,City\nAsha Patil,12 MG Road,Pune\nRavi Kumar,4 Park Street,Mumbai' }], { width: 460 }); if (v && setData(csvParse(v.t, /\t/.test(v.t.split('\n')[0]) ? '\t' : ','))) OA.toast(mm.rows.length + ' recipients loaded'); else if (v) OA.alert('Need a header line and at least one record.'); } },
        { label: 'Use an Existing List… (.csv / .xlsx)', onClick: async () => {
          const f = await OA.pickFile('.csv,.tsv,.txt,.xlsx'); if (!f) return;
          try {
            let rows;
            if (OA.ext(f.name) === 'xlsx') { const wb = await XL.fromXlsx(new Uint8Array(await f.arrayBuffer())); const sh = wb.sheets[0], e = wb.extent(sh.name); rows = []; for (let r = 0; r <= e.r; r++) { const row = []; for (let c = 0; c <= e.c; c++) { let d = wb.display(sh, r, c); if (d && typeof d === 'object') d = d.text; row.push(d == null ? '' : d); } rows.push(row); } }
            else { const t = await f.text(); rows = csvParse(t, OA.ext(f.name) === 'tsv' || /\t/.test(t.split('\n')[0]) ? '\t' : ','); }
            if (setData(rows)) OA.toast(mm.rows.length + ' recipients loaded from ' + f.name); else OA.alert('The file needs a header row and at least one record.');
          } catch (e) { OA.alert('Could not read the list: ' + OA.esc(e.message)); }
        } },
        { label: 'Use the First Table in This Document', onClick: () => { const t = P.querySelector('table'); if (!t) return OA.toast('No table in the document'); const rows = Array.from(t.rows).map(r => Array.from(r.cells).map(c => c.textContent.trim())); if (setData(rows)) OA.toast(mm.rows.length + ' recipients loaded'); else OA.alert('Table needs a header row and data rows.'); } },
        { sep: true }, { label: 'Edit Recipient List…', onClick: editList }]);
    };
    async function editList() {
      if (!mm.fields.length) return OA.toast('Choose recipients first');
      const t = [mm.fields.join('\t')].concat(mm.rows.map(r => mm.fields.map(f => r[f]).join('\t'))).join('\n');
      const v = await OA.form('Edit Recipient List', [{ id: 't', label: 'Tab-separated, first line = fields', type: 'textarea', rows: 12, value: t }], { width: 520 }); if (v) setData(csvParse(v.t, '\t')); renderFields();
    }
    X.mmField = (anchor) => { if (!mm.fields.length) return OA.toast('Select Recipients first'); OA.menu(anchor, mm.fields.map(f => ({ label: f, onClick: () => X.insInline(`<span class="mf" data-f="${OA.esc(f)}" contenteditable="false">«${OA.esc(f)}»</span>&nbsp;`) }))); };
    X.mmBlock = () => { if (!mm.fields.length) return OA.toast('Select Recipients first'); const f = n => mm.fields.find(x => x.toLowerCase() === n); const fl = [f('name') || f('first name') || mm.fields[0], f('address') || f('street'), f('city'), f('state'), f('zip'), f('pincode')].filter(Boolean); X.insBlock('<p>' + fl.map(x => `<span class="mf" data-f="${OA.esc(x)}" contenteditable="false">«${OA.esc(x)}»</span>`).join('<br>') + '</p>'); };
    X.mmGreeting = () => { if (!mm.fields.length) return OA.toast('Select Recipients first'); const f = mm.fields.find(x => /name/i.test(x)) || mm.fields[0]; X.insBlock(`<p>Dear <span class="mf" data-f="${OA.esc(f)}" contenteditable="false">«${OA.esc(f)}»</span>,</p>`); };
    X.mmHighlight = () => P.classList.toggle('mf-hl');
    X.mmPreview = v => { if (!mm.rows.length) { OA.toast('Select Recipients first'); return false; } mm.preview = v == null ? !mm.preview : v; renderFields(); return mm.preview; };
    X.mmNav = d => { if (!mm.rows.length) return OA.toast('Select Recipients first'); mm.idx = d === 'first' ? 0 : d === 'last' ? mm.rows.length - 1 : OA.clamp(mm.idx + d, 0, mm.rows.length - 1); if (!mm.preview) mm.preview = true; renderFields(); OA.toast('Record ' + (mm.idx + 1) + ' of ' + mm.rows.length); };
    X.mmFind = async () => { const q = await OA.prompt('Find recipient containing:', '', { title: 'Find Recipient' }); if (!q) return; const i = mm.rows.findIndex(r => Object.values(r).some(v => String(v).toLowerCase().includes(q.toLowerCase()))); if (i < 0) return OA.toast('Not found'); mm.idx = i; mm.preview = true; renderFields(); };
    X.mmRules = (anchor) => OA.menu(anchor, [{ label: 'Skip records where a field is empty…', onClick: async () => { const f = await OA.prompt('Skip records where this field is empty:', mm.fields[0] || '', { title: 'Rules' }); if (f) mm.skip = f; OA.toast(f ? 'Rule set: skip empty ' + f : 'No rule'); } }, { label: 'Clear rules', onClick: () => { mm.skip = null; } }]);
    function mergedHtml() {
      const rows = mm.rows.filter(r => !mm.skip || String(r[mm.skip] || '').trim()); const base = document.createElement('div'); base.innerHTML = ed.getHTML(); const parts = [];
      rows.forEach(r => { const d = base.cloneNode(true); d.querySelectorAll('span.mf').forEach(s => s.replaceWith(document.createTextNode(r[s.dataset.f] == null ? '' : r[s.dataset.f]))); parts.push(d.innerHTML); });
      return parts.join('<div class="pb" contenteditable="false"></div>');
    }
    X.mmFinish = anchor => {
      if (!mm.rows.length) return OA.toast('Select Recipients first'); if (!mfields().length) return OA.toast('Insert at least one merge field first');
      OA.menu(anchor, [
        { label: 'Edit Individual Documents (new content)', onClick: async () => { if (await OA.confirm('Replace the document with the merged letters (' + mm.rows.length + ' pages)? Use Undo to restore the template.', { ok: 'Merge' })) { mm.preview = false; ed.mutate(() => { P.innerHTML = mergedHtml(); }); OA.toast('Merged ' + mm.rows.length + ' records'); } } },
        { label: 'Download merged .docx', onClick: async () => { const keep = P.innerHTML, wasPrev = mm.preview; ed.mutate(() => { P.innerHTML = mergedHtml(); }); try { const n = 'MergedLetters'; OA.download(await ed.exportDocx(n), n + '.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'); } finally { ed.mutate(() => { P.innerHTML = keep; }); mm.preview = wasPrev; renderFields(); } } },
        { label: 'Print Documents…', onClick: () => { const keep = P.innerHTML; ed.mutate(() => { P.innerHTML = mergedHtml(); }); ed.print(); setTimeout(() => { ed.mutate(() => { P.innerHTML = keep; }); renderFields(); }, 1500); } }]);
    };
    X.envelope = async () => {
      const v = await OA.form('Envelopes', [{ id: 'to', label: 'Delivery address', type: 'textarea', rows: 4, value: '' }, { id: 'from', label: 'Return address', type: 'textarea', rows: 3, value: '' }, { id: 'size', label: 'Envelope size', type: 'select', value: 'c5', options: [['c5', 'C5 (22.9 × 16.2 cm)'], ['dl', 'DL (22 × 11 cm)'], ['10', 'Size 10 (9.5 × 4.1 in)']] }, { id: 'where', label: 'Add', type: 'select', value: 'top', options: [['top', 'To the start of this document'], ['here', 'At the cursor']] }], { ok: 'Add to Document', width: 440 });
      if (!v) return; const dim = { c5: [866, 612], dl: [832, 416], '10': [912, 394] }[v.size];
      const nl = s => OA.esc(s).replace(/\n/g, '<br>');
      const html = `<div class="env" contenteditable="false" style="width:${dim[0] - 40}px;height:${dim[1] - 30}px"><div class="env-from">${nl(v.from)}</div><div class="env-to">${nl(v.to)}</div></div><div class="pb" contenteditable="false"></div><p><br></p>`;
      if (v.where === 'here') X.insBlock(html); else ed.mutate(() => { const t = document.createElement('div'); t.innerHTML = html; Array.from(t.childNodes).reverse().forEach(n => P.insertBefore(n, P.firstChild)); });
    };
    X.labels = async () => {
      const v = await OA.form('Labels', [{ id: 'addr', label: 'Address / label text', type: 'textarea', rows: 4, value: '' }, { id: 'layout', label: 'Label sheet', type: 'select', value: '3x10', options: [['2x7', '2 × 7 (Avery 5163-style, 14 per page)'], ['3x10', '3 × 10 (Avery 5160-style, 30 per page)'], ['2x4', '2 × 4 (shipping, 8 per page)'], ['1x1', 'Single label']] }, { id: 'same', label: 'Use the same text on every label', type: 'checkbox', value: true }, { id: 'recip', label: 'Use mail merge fields (one label per recipient)', type: 'checkbox', value: false }], { ok: 'Add to Document', width: 460 });
      if (!v) return; const [c, rws] = v.layout.split('x').map(Number); let cells = c * rws; const nl = s => OA.esc(s).replace(/\n/g, '<br>');
      let items = [];
      if (v.recip) { if (!mm.rows.length) return OA.toast('Select Recipients first'); items = mm.rows.map(r => nl(v.addr.replace(/«([^»]+)»/g, (m, f) => r[f] == null ? '' : r[f]))); if (!/«/.test(v.addr)) OA.toast('Tip: write fields as «Name» in the label text'); cells = Math.ceil(items.length / c) * c; }
      else for (let i = 0; i < cells; i++) items.push(v.same ? nl(v.addr) : (i === 0 ? nl(v.addr) : ''));
      let h = '<table class="dt labels" data-b="none" style="width:100%">'; for (let r = 0; r < cells / c; r++) { h += '<tr>'; for (let k = 0; k < c; k++) h += `<td style="height:${rws > 7 ? 70 : 96}px">${items[r * c + k] || '<br>'}</td>`; h += '</tr>'; } h += '</table><p><br></p>';
      X.insBlock(h);
    };

    /* ======================================================== CAPTIONS / CROSS-REFS / TOF */
    const LABELS = ['Figure', 'Table', 'Equation'];
    function renumber() { const cnt = {}; P.querySelectorAll('p.caption').forEach(p => { const sp = p.querySelector('.cap'); if (!sp) return; const l = sp.dataset.l; cnt[l] = (cnt[l] || 0) + 1; sp.textContent = l + ' ' + cnt[l]; if (!p.id) p.id = 'cap-' + OA.uid(); }); }
    X.caption = async () => {
      const r = sel(); const b = r && ed.selBlocks && ed.selBlocks()[0];
      const v = await OA.form('Insert Caption', [{ id: 'l', label: 'Label', type: 'select', value: 'Figure', options: LABELS }, { id: 't', label: 'Caption text', value: '' }, { id: 'pos', label: 'Position', type: 'select', value: 'below', options: [['below', 'Below selected item'], ['above', 'Above selected item']] }]); if (!v) return;
      const html = `<p class="caption"><span class="cap" data-l="${v.l}">${v.l} 0</span>${v.t ? ': ' + OA.esc(v.t) : ''}</p>`;
      ed.mutate(() => {
        let target = null; const s = getSelection(); const node = s.rangeCount ? s.getRangeAt(0).startContainer : null; const el0 = node && (node.nodeType === 3 ? node.parentNode : node);
        target = el0 && el0.closest ? (el0.closest('table') || el0.closest('p,h1,h2,h3,li')) : null; if (!target || !P.contains(target)) target = P.lastElementChild;
        const t = document.createElement('div'); t.innerHTML = html; const p = t.firstChild;
        if (target.tagName === 'TABLE' || target.tagName === 'P' || /^H/.test(target.tagName)) { v.pos === 'above' ? target.before(p) : target.after(p); } else target.after(p);
        renumber();
      });
    };
    X.tof = async () => {
      const v = await OA.form('Table of Figures', [{ id: 'l', label: 'Caption label', type: 'select', value: 'Figure', options: LABELS }]); if (!v) return; renumber();
      const caps = Array.from(P.querySelectorAll('p.caption')).filter(p => p.querySelector('.cap').dataset.l === v.l); if (!caps.length) return OA.alert('No “' + v.l + '” captions found. Use References → Insert Caption first.');
      const { h: ph } = ed.pageDims();
      let html = `<div class="toc tof"><div class="toc-h">Table of ${v.l}s</div>` + caps.map(p => `<div class="t1" data-t="${p.id}">${OA.esc(p.textContent)}  ${Math.floor(p.offsetTop / ph) + 1}</div>`).join('') + '</div><p><br></p>';
      X.insBlock(html);
    };
    X.xref = async () => {
      renumber(); const items = [];
      P.querySelectorAll('h1,h2,h3,h4').forEach(h => { if (h.textContent.trim() && !h.closest('.toc')) { if (!h.id) h.id = 'hd-' + OA.uid(); items.push({ k: 'Heading', id: h.id, t: h.textContent.trim() }); } });
      P.querySelectorAll('p.caption').forEach(p => items.push({ k: p.querySelector('.cap').dataset.l, id: p.id, t: p.textContent }));
      P.querySelectorAll('[id]:not(h1):not(h2):not(h3):not(h4):not(.caption)').forEach(b => { if (!/^(hd-|cap-|fn)/.test(b.id) && !b.closest('.toc')) items.push({ k: 'Bookmark', id: b.id, t: b.id }); });
      if (!items.length) return OA.alert('Nothing to refer to yet. Add headings, captions or bookmarks first.');
      const v = await OA.form('Cross-reference', [{ id: 'k', label: 'Reference type', type: 'select', value: items[0].k, options: Array.from(new Set(items.map(i => i.k))) }, { id: 'what', label: 'Insert reference to', type: 'select', value: 'text', options: [['text', 'Heading / caption text'], ['label', 'Only label and number (captions)'], ['page', 'Page number']] }]); if (!v) return;
      const list = items.filter(i => i.k === v.k); const pick = await OA.form('Cross-reference', [{ id: 'i', label: 'For which ' + v.k.toLowerCase(), type: 'select', value: 0, options: list.map((x, i) => [i, x.t.slice(0, 60)]) }]); if (!pick) return;
      const it = list[+pick.i]; const tgt = document.getElementById(it.id); const { h: ph } = ed.pageDims();
      const txt = v.what === 'page' ? 'page ' + (Math.floor(tgt.offsetTop / ph) + 1) : v.what === 'label' && it.k !== 'Heading' && it.k !== 'Bookmark' ? it.t.split(':')[0] : it.t;
      X.insInline(`<a href="#${it.id}" class="xref">${OA.esc(txt)}</a>&nbsp;`);
    };

    /* ======================================================== CITATIONS / BIBLIOGRAPHY */
    const src = X.sources = X.sources || { list: [], style: 'APA' };
    const fmtAuthor = a => a || 'Unknown';
    const cite = s => { const y = s.year || 'n.d.'; return src.style === 'MLA' ? `(${fmtAuthor(s.author).split(',')[0]}${s.pages ? ' ' + s.pages : ''})` : src.style === 'Chicago' ? `(${fmtAuthor(s.author).split(',')[0]} ${y}${s.pages ? ', ' + s.pages : ''})` : `(${fmtAuthor(s.author).split(',')[0]}, ${y}${s.pages ? ', p. ' + s.pages : ''})`; };
    const biblio = s => { const a = fmtAuthor(s.author), y = s.year || 'n.d.', t = OA.esc(s.title || ''), pub = OA.esc(s.pub || ''), u = s.url ? ' ' + OA.esc(s.url) : '';
      if (src.style === 'MLA') return `${OA.esc(a)}. <i>${t}</i>. ${pub}${pub ? ', ' : ''}${y}.${u}`;
      if (src.style === 'Chicago') return `${OA.esc(a)}. ${y}. <i>${t}</i>. ${pub}.${u}`;
      return `${OA.esc(a)} (${y}). <i>${t}</i>. ${pub}${pub ? '.' : ''}${u}`; };
    async function sourceForm(s) { return OA.form(s ? 'Edit Source' : 'Create Source', [{ id: 'type', label: 'Type of source', type: 'select', value: (s || {}).type || 'Book', options: ['Book', 'Journal Article', 'Website', 'Report', 'Conference Proceedings', 'Film'] }, { id: 'author', label: 'Author (Last, First)', value: (s || {}).author || '' }, { id: 'title', label: 'Title', value: (s || {}).title || '' }, { id: 'year', label: 'Year', value: (s || {}).year || '' }, { id: 'pub', label: 'Publisher / Journal / Site', value: (s || {}).pub || '' }, { id: 'url', label: 'URL (optional)', value: (s || {}).url || '' }, { id: 'pages', label: 'Pages (optional, for citations)', value: (s || {}).pages || '' }], { width: 420 }); }
    X.styleMenu = anchor => OA.menu(anchor, ['APA', 'MLA', 'Chicago'].map(s => ({ label: s, checked: src.style === s, onClick: () => { src.style = s; OA.toast('Citation style: ' + s); X.updateCites(); } })));
    X.addSource = async () => { const v = await sourceForm(); if (!v) return null; v.id = 's' + OA.uid(); src.list.push(v); return v; };
    X.manageSources = () => {
      const render = () => { body.innerHTML = ''; if (!src.list.length) body.appendChild(el('div', { class: 'dlg-msg', text: 'No sources yet.' })); src.list.forEach(s => body.appendChild(el('div', { style: { display: 'flex', gap: '6px', alignItems: 'center', padding: '5px 0', borderBottom: '1px solid #eee' } }, el('div', { style: { flex: 1, fontSize: '13px' }, html: biblio(s) }), el('button', { class: 'btn', type: 'button', text: 'Edit', onclick: async () => { const v = await sourceForm(s); if (v) { Object.assign(s, v); render(); } } }), el('button', { class: 'btn', type: 'button', text: 'Delete', onclick: () => { src.list.splice(src.list.indexOf(s), 1); render(); } })))); };
      const body = el('div', { style: { minWidth: '420px', maxHeight: '50vh', overflow: 'auto' } }); render();
      OA.dialog({ title: 'Source Manager', content: body, width: 520, buttons: [{ label: 'New source…', onClick: async () => { await X.addSource(); render(); return false; } }, { label: 'Close', primary: true }] });
    };
    X.insertCitation = anchor => {
      const items = src.list.map(s => ({ label: cite(s) + ' — ' + (s.title || '').slice(0, 30), onClick: () => X.insInline(`<span class="cite" data-s="${s.id}">${OA.esc(cite(s))}</span>&nbsp;`) }));
      OA.menu(anchor, [{ label: 'Add New Source…', onClick: async () => { const s = await X.addSource(); if (s) X.insInline(`<span class="cite" data-s="${s.id}">${OA.esc(cite(s))}</span>&nbsp;`); } }, ...(items.length ? [{ sep: true }, ...items] : [])]);
    };
    X.updateCites = () => ed.mutate(() => { P.querySelectorAll('span.cite').forEach(c => { const s = src.list.find(x => x.id === c.dataset.s); if (s) c.textContent = cite(s); }); const b = P.querySelector('.biblio'); if (b) buildBib(b); });
    function buildBib(box) { const ttl = src.style === 'MLA' ? 'Works Cited' : src.style === 'Chicago' ? 'Bibliography' : 'References'; const list = src.list.slice().sort((a, b) => (a.author || '').localeCompare(b.author || '')); box.innerHTML = `<div class="biblio-h">${ttl}</div>` + list.map(s => `<p class="bib" style="padding-left:36pt;text-indent:-36pt">${biblio(s)}</p>`).join(''); }
    X.insertBib = () => { if (!src.list.length) return OA.alert('Add at least one source first (References → Manage Sources).'); const box = document.createElement('div'); box.className = 'biblio'; buildBib(box); X.insBlock(box.outerHTML); };

    /* ======================================================== INDEX / ENDNOTES / UPDATE FIELDS */
    X.markEntry = async () => { const r = sel(); const t = r && !r.collapsed ? r.toString().trim() : ''; const e = await OA.prompt('Main entry:', t, { title: 'Mark Index Entry' }); if (!e) return; if (r && !r.collapsed) ed.mutate(() => { const sp = document.createElement('span'); sp.className = 'xe'; sp.dataset.x = e; try { sel().surroundContents(sp); } catch (er) { } }); else X.insInline(`<span class="xe" data-x="${OA.esc(e)}"></span>`); };
    X.insertIndex = () => {
      const xs = Array.from(P.querySelectorAll('span.xe')); if (!xs.length) return OA.alert('No entries marked. Select text and use References → Mark Entry.'); const { h: ph } = ed.pageDims(); const map = {};
      xs.forEach(x => { const k = x.dataset.x; (map[k] = map[k] || new Set()).add(Math.floor(x.offsetTop / ph) + 1); });
      const keys = Object.keys(map).sort((a, b) => a.localeCompare(b)); let last = '', html = '<div class="indexbox"><div class="toc-h">Index</div>';
      keys.forEach(k => { const L0 = k[0].toUpperCase(); if (L0 !== last) { html += `<div class="ix-l"><b>${OA.esc(L0)}</b></div>`; last = L0; } html += `<div class="ix">${OA.esc(k)}, ${Array.from(map[k]).sort((a, b) => a - b).join(', ')}</div>`; });
      P.querySelectorAll('.indexbox').forEach(n => n.remove()); X.insBlock(html + '</div>');
    };
    X.insertEndnote = async () => {
      const t = await OA.prompt('Endnote text:', '', { title: 'Insert Endnote' }); if (!t) return; const rom = n => ['i', 'ii', 'iii', 'iv', 'v', 'vi', 'vii', 'viii', 'ix', 'x', 'xi', 'xii', 'xiii', 'xiv', 'xv', 'xvi', 'xvii', 'xviii', 'xix', 'xx'][n - 1] || String(n);
      ed.mutate(() => { let box = P.querySelector('.enote'); const n = (box ? box.querySelectorAll('p').length : 0) + 1; document.execCommand('insertHTML', false, `<sup class="en">${rom(n)}</sup>`); if (!box) { box = document.createElement('div'); box.className = 'enote'; box.innerHTML = '<div class="enote-h">Endnotes</div>'; P.appendChild(box); } const p = document.createElement('p'); p.innerHTML = `<sup>${rom(n)}</sup> ${OA.esc(t)}`; p.style.cssText = 'font-size:9pt;margin:0'; box.appendChild(p); });
    };
    X.updateFields = () => { renumber(); if (P.querySelector('.toc:not(.tof)')) ed.updateTOC(); X.updateCites(); OA.toast('Fields updated'); };

    /* ======================================================== RIBBON */
    const MM = [L('mmsel', 'Select Recipients', 'users', (e, b) => X.mmSelect(b))];
    tabAdd('review', [
      { label: 'Proofing', items: [L('thes', 'Thesaurus', 'book', X.thesaurus), L('trans', 'Translate', 'globe', X.translate)] },
      { label: 'Tracking', items: [{ type: 'check', id: 'tc', label: 'Track Changes', value: false, onChange: v => X.setTrack(v) }, { type: 'row', items: [B('tcn', 'Change tracking name', null, X.trackAuthor, { glyph: '👤', label: 'User name' }), B('tcm', 'Reviewing Pane', null, X.reviewPane, { glyph: '☰', label: 'Reviewing Pane' })] }, { type: 'check', id: 'tcs', label: 'Show Markup', value: true, onChange: v => X.showMarkup(v) }] },
      { label: 'Changes', items: [L('acc', 'Accept', 'check2', () => X.accept(false, true), { menu: [{ label: 'Accept This Change', onClick: () => X.accept(false, true) }, { label: 'Accept All Changes', onClick: () => X.accept(true, true) }] }), L('rej', 'Reject', 'x', () => X.accept(false, false), { menu: [{ label: 'Reject This Change', onClick: () => X.accept(false, false) }, { label: 'Reject All Changes', onClick: () => X.accept(true, false) }] }), { type: 'row', items: [B('pch', 'Previous change', null, () => X.gotoChange(-1), { glyph: '◀', label: 'Previous' }), B('nch', 'Next change', null, () => X.gotoChange(1), { glyph: '▶', label: 'Next' })] }] },
      { label: 'Compare', items: [L('cmp', 'Compare', 'columns', X.compare)] },
      { label: 'Protect', items: [L('rst', 'Restrict Editing', 'lock', X.restrict)] }]);
    spec.tabs.splice(spec.tabs.findIndex(t => t.id === 'review'), 0, { id: 'mail', label: 'Mailings', groups: [
      { label: 'Create', items: [L('env', 'Envelopes', 'mail', X.envelope), L('lbl', 'Labels', 'table', X.labels)] },
      { label: 'Start Mail Merge', items: [...MM, L('mme', 'Edit Recipient List', 'edit', editList)] },
      { label: 'Write & Insert Fields', items: [L('mmf', 'Insert Merge Field', 'plus', (e, b) => X.mmField(b)), { type: 'row', items: [B('mmb', 'Address Block', null, X.mmBlock, { glyph: '▤', label: 'Address Block' }), B('mmg', 'Greeting Line', null, X.mmGreeting, { glyph: '✉', label: 'Greeting Line' })] }, { type: 'row', items: [B('mmh', 'Highlight Merge Fields', null, X.mmHighlight, { glyph: '▨', label: 'Highlight' }), B('mmr', 'Rules', null, (e, b) => X.mmRules(b), { glyph: '⚙', label: 'Rules' })] }] },
      { label: 'Preview Results', items: [L('mmp', 'Preview Results', 'eye', () => X.mmPreview()), { type: 'row', items: [B('mm1', 'First record', null, () => X.mmNav('first'), { glyph: '⏮' }), B('mm2', 'Previous record', null, () => X.mmNav(-1), { glyph: '◀' }), B('mm3', 'Next record', null, () => X.mmNav(1), { glyph: '▶' }), B('mm4', 'Last record', null, () => X.mmNav('last'), { glyph: '⏭' })] }, B('mmfd', 'Find Recipient', null, X.mmFind, { glyph: '🔍', label: 'Find Recipient' })] },
      { label: 'Finish', items: [L('mmfin', 'Finish & Merge', 'check2', (e, b) => X.mmFinish(b))] }] });
    const refs = spec.tabs.find(t => t.id === 'refs'); refs.groups.push(
      { label: 'Endnotes', items: [L('en', 'Insert Endnote', 'notes', X.insertEndnote)] },
      { label: 'Citations & Bibliography', items: [L('cit', 'Insert Citation', 'quote', (e, b) => X.insertCitation(b)), { type: 'row', items: [B('srcm', 'Manage Sources', null, X.manageSources, { glyph: '📚', label: 'Manage Sources' }), B('sty', 'Style', null, (e, b) => X.styleMenu(b), { glyph: '§', label: 'Style' })] }, L('bib', 'Bibliography', 'book', X.insertBib)] },
      { label: 'Captions', items: [L('cap', 'Insert Caption', 'comment', X.caption), { type: 'row', items: [B('tof', 'Insert Table of Figures', null, X.tof, { glyph: '▤', label: 'Table of Figures' }), B('xrf', 'Cross-reference', null, X.xref, { glyph: '↗', label: 'Cross-reference' })] }] },
      { label: 'Index', items: [L('mke', 'Mark Entry', 'bookmark', X.markEntry), L('idx', 'Insert Index', 'list', X.insertIndex)] },
      { label: 'Update', items: [L('upf', 'Update Fields', 'refresh', X.updateFields)] });
  });
})();
