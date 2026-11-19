/* Excel .xlsx writer/reader (own OOXML, no libs) + chart data helper — Developer: Sahil Jadhav */
(function (global) {
  'use strict';
  const XL = global.XL, OA = global.OA, KC = XL.MAXC, key = XL.key;
  const xe = s => String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XFN = new Set(['XLOOKUP', 'XMATCH', 'SEQUENCE', 'SORT', 'SORTBY', 'FILTER', 'UNIQUE', 'CONCAT', 'TEXTJOIN', 'IFS', 'SWITCH', 'MAXIFS', 'MINIFS', 'LET', 'TEXTBEFORE', 'TEXTAFTER', 'TEXTSPLIT', 'TAKE', 'RANDARRAY', 'DAYS', 'IFNA', 'NUMBERVALUE', 'UNICHAR', 'UNICODE', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'MODE.SNGL', 'PERCENTILE.INC', 'QUARTILE.INC', 'RANK.EQ', 'FORECAST.LINEAR', 'CEILING.MATH', 'FLOOR.MATH', 'XOR', 'BASE', 'DECIMAL', 'ERROR.TYPE'.slice(0, 0)].filter(Boolean));
  const ARR = new Set(['SEQUENCE', 'SORT', 'SORTBY', 'FILTER', 'UNIQUE', 'TEXTSPLIT', 'TAKE', 'RANDARRAY', 'TRANSPOSE', 'MMULT']);
  function xlFormula(ast) {
    return XL.astToStr(XL.mapAst(ast, n => n.t === 'fn' && XFN.has(n.name) ? Object.assign({}, n, { name: '_xlfn.' + n.name, args: n.args.map(a => XL.mapAst(a, m => undefined)) }) : undefined)).replace(/_xlfn\._xlfn\./g, '_xlfn.');
  }
  const fixF = f => f.replace(/\b(XLOOKUP|XMATCH|SEQUENCE|SORTBY|SORT|FILTER|UNIQUE|CONCAT|TEXTJOIN|IFS|SWITCH|MAXIFS|MINIFS|LET|TEXTBEFORE|TEXTAFTER|TEXTSPLIT|TAKE|RANDARRAY|DAYS|IFNA|NUMBERVALUE|UNICHAR|UNICODE|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|MODE\.SNGL|PERCENTILE\.INC|QUARTILE\.INC|RANK\.EQ|FORECAST\.LINEAR|CEILING\.MATH|FLOOR\.MATH|XOR|BASE|DECIMAL)\(/g, (m, n, off, s) => s.slice(Math.max(0, off - 6), off) === '_xlfn.' ? m : '_xlfn.' + m);
  const argb = h => 'FF' + String(h || '#000000').replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6);
  const BSTYLE = { thin: 'thin', medium: 'medium', thick: 'thick', dashed: 'dashed', dotted: 'dotted', double: 'double', hair: 'hair' };

  /* ---------------------------------------------------------------- chart data */
  XL.chartData = function (wb, ch) {
    wb.ensureCalc();
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ch.ref); const shn = m ? (m[1] ? m[1].replace(/''/g, "'") : m[2]) : null; const sh = wb.sheet(shn) || wb.sheets[0]; const g = XL.parseRange(m ? m[3] : ch.ref); if (!g) return { cats: [], series: [] };
    const ext = wb.extent(sh.name); g.r2 = Math.min(g.r2, Math.max(g.r1, ext.r)); g.c2 = Math.min(g.c2, Math.max(g.c1, ext.c));
    const V = (r, c) => wb.value(sh, r, c); const D = (r, c) => wb.display(sh, r, c).text; const isN = (r, c) => typeof V(r, c) === 'number';
    const byRow = !!ch.rowsAsSeries; let r1 = g.r1, c1 = g.c1;
    const q = x => XL.quoteSheet(sh.name) + '!' + x;
    let hasHdr, hasCat;
    if (!byRow) {
      hasHdr = g.r2 > g.r1 && !isN(g.r1, Math.min(g.c1 + 1, g.c2)) ? true : (g.r2 > g.r1 && V(g.r1, g.c1 + 1) === null);
      hasCat = g.c2 > g.c1 && !isN(g.r1 + (hasHdr ? 1 : 0), g.c1) ;
      if (g.c2 === g.c1) hasCat = false;
      const dr1 = g.r1 + (hasHdr ? 1 : 0), dc1 = g.c1 + (hasCat ? 1 : 0); const cats = []; for (let r = dr1; r <= g.r2; r++) cats.push(hasCat ? D(r, g.c1) : String(r - dr1 + 1));
      const series = []; for (let c = dc1; c <= g.c2; c++) { const vals = []; for (let r = dr1; r <= g.r2; r++) { const v = V(r, c); vals.push(typeof v === 'number' ? v : (v !== null && !isNaN(+v) && v !== '' ? +v : 0)); } series.push({ name: hasHdr ? D(g.r1, c) : 'Series' + (c - dc1 + 1), vals, nameRef: hasHdr ? q(XL.addr(g.r1, c, true, true)) : null, catRef: hasCat ? q(XL.addr(dr1, g.c1, true, true) + ':' + XL.addr(g.r2, g.c1, true, true)) : null, valRef: q(XL.addr(dr1, c, true, true) + ':' + XL.addr(g.r2, c, true, true)) }); }
      return { cats, series };
    }
    hasHdr = g.c2 > g.c1 && !isN(g.r1, g.c1) ; hasCat = g.r2 > g.r1 && !isN(g.r1, g.c1 + (hasHdr ? 1 : 0));
    // by rows: first column = series names, first row = categories
    const nameCol = hasHdr, catRow = g.r2 > g.r1 && !isN(g.r1, g.c1 + (nameCol ? 1 : 0));
    const dc1 = g.c1 + (nameCol ? 1 : 0), dr1 = g.r1 + (catRow ? 1 : 0); const cats = []; for (let c = dc1; c <= g.c2; c++) cats.push(catRow ? D(g.r1, c) : String(c - dc1 + 1));
    const series = []; for (let r = dr1; r <= g.r2; r++) { const vals = []; for (let c = dc1; c <= g.c2; c++) { const v = V(r, c); vals.push(typeof v === 'number' ? v : 0); } series.push({ name: nameCol ? D(r, g.c1) : 'Series' + (r - dr1 + 1), vals, nameRef: nameCol ? q(XL.addr(r, g.c1, true, true)) : null, catRef: catRow ? q(XL.addr(g.r1, dc1, true, true) + ':' + XL.addr(g.r1, g.c2, true, true)) : null, valRef: q(XL.addr(r, dc1, true, true) + ':' + XL.addr(r, g.c2, true, true)) }); }
    return { cats, series };
  };

  /* ---------------------------------------------------------------- WRITE */
  function posToCell(sh, wb, x, y) { let c = 0, acc = 0; while (c < 16384) { const w = wb.colWidth(sh, c); if (acc + w > x) break; acc += w; c++; } const co = Math.round((x - acc)); let r = 0, ac = 0; while (r < 1048576) { const h = wb.rowHeight(sh, r); if (ac + h > y) break; ac += h; r++; } return { c, r, co, ro: Math.round(y - ac) }; }
  function chartXml(wb, ch) {
    const d = XL.chartData(wb, ch); const t = ch.type;
    const col = i => ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E'][i % 8];
    const strCache = arr => `<c:strCache><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${xe(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
    const numCache = arr => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
    const ser = d.series.map((s, i) => {
      const fill = `<c:spPr>${t === 'line' || t === 'scatter' ? `<a:ln w="28575"><a:solidFill><a:srgbClr val="${col(i)}"/></a:solidFill></a:ln>` : `<a:solidFill><a:srgbClr val="${col(i)}"/></a:solidFill>`}</c:spPr>`;
      const tx = `<c:tx>${s.nameRef ? `<c:strRef><c:f>${xe(s.nameRef)}</c:f>${strCache([s.name])}</c:strRef>` : `<c:v>${xe(s.name)}</c:v>`}</c:tx>`;
      const catXml = t === 'scatter' ? `<c:xVal><c:numRef><c:f>${xe(s.catRef || s.valRef)}</c:f>${numCache(d.cats.map((x, k) => isNaN(+x) ? k + 1 : +x))}</c:numRef></c:xVal>` : `<c:cat>${s.catRef ? `<c:strRef><c:f>${xe(s.catRef)}</c:f>${strCache(d.cats)}</c:strRef>` : `<c:strLit><c:ptCount val="${d.cats.length}"/>${d.cats.map((v, k) => `<c:pt idx="${k}"><c:v>${xe(v)}</c:v></c:pt>`).join('')}</c:strLit>`}</c:cat>`;
      const valXml = `<c:${t === 'scatter' ? 'yVal' : 'val'}><c:numRef><c:f>${xe(s.valRef)}</c:f>${numCache(s.vals)}</c:numRef></c:${t === 'scatter' ? 'yVal' : 'val'}>`;
      let pts = ''; if (t === 'pie' || t === 'doughnut') pts = d.cats.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${col(k)}"/></a:solidFill></c:spPr></c:dPt>`).join('');
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${t === 'pie' || t === 'doughnut' ? '' : fill}${t === 'line' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>' : ''}${pts}${catXml}${valXml}${t === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    }).join('');
    const ax = '<c:axId val="111"/><c:axId val="222"/>';
    let plot;
    if (t === 'col' || t === 'bar') plot = `<c:barChart><c:barDir val="${t === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${ch.stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/>${ch.stacked ? '<c:overlap val="100"/>' : ''}${ax}</c:barChart>`;
    else if (t === 'line') plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}<c:marker val="1"/>${ax}</c:lineChart>`;
    else if (t === 'area') plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${ser}${ax}</c:areaChart>`;
    else if (t === 'scatter') plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ser}${ax}</c:scatterChart>`;
    else if (t === 'doughnut') plot = `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
    else plot = `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>`;
    const axes = (t === 'pie' || t === 'doughnut') ? '' : (t === 'scatter' ? `<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/></c:valAx>` : `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${t === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="0"/><c:tickLblPos val="nextTo"/><c:crossAx val="222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`) + `<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${t === 'bar' ? 'b' : 'l'}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="111"/><c:crosses val="autoZero"/></c:valAx>`;
    const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr lang="en-US" sz="1400" b="0"/><a:t>${xe(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}"><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea>${ch.legend === false ? '' : '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }

  XL.toXlsx = async function (wb) {
    wb.ensureCalc(); const files = []; const strings = [], sIdx = new Map(); const si = s => { let i = sIdx.get(s); if (i === undefined) { i = strings.length; strings.push(s); sIdx.set(s, i); } return i; };
    /* styles */
    const fonts = [], fontMap = new Map(), fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillMap = new Map(), borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], borderMap = new Map(), nfs = [], nfMap = new Map();
    const bid = o => { const k = JSON.stringify(o); let i = borderMap.get(k); if (i === undefined) { const side = (n, b) => b ? `<${n} style="${BSTYLE[b.s] || 'thin'}"><color rgb="${argb(b.c)}"/></${n}>` : `<${n}/>`; borders.push(`<border>${side('left', o.bL)}${side('right', o.bR)}${side('top', o.bT)}${side('bottom', o.bB)}<diagonal/></border>`); i = borders.length - 1; borderMap.set(k, i); } return i; };
    const fid = st => { const x = `<font>${st.b ? '<b/>' : ''}${st.i ? '<i/>' : ''}${st.st ? '<strike/>' : ''}${st.u ? '<u/>' : ''}<sz val="${st.sz || 11}"/><color rgb="${argb(st.color || '#000000')}"/><name val="${xe(st.font || 'Calibri')}"/><family val="2"/></font>`; let i = fontMap.get(x); if (i === undefined) { fonts.push(x); i = fonts.length - 1; fontMap.set(x, i); } return i; };
    const flid = st => { if (!st.bg) return 0; const x = `<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.bg)}"/><bgColor indexed="64"/></patternFill></fill>`; let i = fillMap.get(x); if (i === undefined) { fills.push(x); i = fills.length - 1; fillMap.set(x, i); } return i; };
    const rev = {}; Object.keys(XL.BUILTIN_FMT).forEach(k => rev[XL.BUILTIN_FMT[k]] = +k); rev['mm-dd-yy'] = 14;
    const nid = f => { if (!f || f === 'General') return 0; if (rev[f] !== undefined) return rev[f]; let i = nfMap.get(f); if (i === undefined) { i = 164 + nfs.length; nfs.push(`<numFmt numFmtId="${i}" formatCode="${xe(f)}"/>`); nfMap.set(f, i); } return i; };
    const dxfs = []; const dxf = st => { const x = `<dxf>${st.color || st.b || st.i ? `<font>${st.b ? '<b/>' : ''}${st.i ? '<i/>' : ''}${st.color ? `<color rgb="${argb(st.color)}"/>` : ''}</font>` : ''}${st.bg ? `<fill><patternFill><bgColor rgb="${argb(st.bg)}"/></patternFill></fill>` : ''}</dxf>`; dxfs.push(x); return dxfs.length - 1; };
    fonts.push('<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>'); fontMap.set(fonts[0], 0);
    const xfs = wb.styles.map(st => { const al = (st.ha || st.va || st.wrap || st.ind) ? `<alignment${st.ha ? ` horizontal="${st.ha}"` : ''}${st.va ? ` vertical="${st.va === 'middle' ? 'center' : st.va}"` : ''}${st.wrap ? ' wrapText="1"' : ''}${st.ind ? ` indent="${st.ind}"` : ''}/>` : ''; const b = bid(st); const bi = (st.bL || st.bR || st.bT || st.bB) ? b : 0; return `<xf numFmtId="${nid(st.nf)}" fontId="${fid(st)}" fillId="${flid(st)}" borderId="${bi}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`; });
    /* sheets */
    const sheetsXml = [], sheetRels = [], drawings = [], charts = []; let chartN = 0;
    wb.sheets.forEach((sh, idx) => {
      const n = idx + 1; const rels = []; let relId = 0;
      const rows = new Map(); const addCell = (r, c, x) => { if (!rows.has(r)) rows.set(r, []); rows.get(r).push([c, x]); };
      sh.cells.forEach((cell, k) => {
        const r = Math.floor(k / KC), c = k % KC; const a = XL.addr(r, c); const s = cell.s ? ` s="${cell.s}"` : '';
        if (cell.f != null) {
          const v = wb.value(sh, r, c); let f = fixF(cell.f); const isArr = cell.spillRect;
          const fx = isArr ? `<f t="array" ref="${XL.rangeStr(cell.spillRect)}">${xe(f)}</f>` : (cell.ast && cell.ast.t === 'fn' && ARR.has(cell.ast.name) ? `<f t="array" ref="${a}">${xe(f)}</f>` : `<f>${xe(f)}</f>`);
          let t = '', vv = '';
          if (typeof v === 'number') vv = `<v>${v}</v>`; else if (typeof v === 'string') { t = ' t="str"'; vv = `<v>${xe(v)}</v>`; } else if (typeof v === 'boolean') { t = ' t="b"'; vv = `<v>${v ? 1 : 0}</v>`; } else if (XL.isErr(v)) { t = ' t="e"'; vv = `<v>${xe(v.err)}</v>`; }
          addCell(r, c, `<c r="${a}"${s}${t}>${fx}${vv}</c>`);
        } else if (cell.v !== undefined && cell.v !== null && cell.v !== '') {
          const v = cell.v;
          if (typeof v === 'number') addCell(r, c, `<c r="${a}"${s}><v>${v}</v></c>`); else if (typeof v === 'boolean') addCell(r, c, `<c r="${a}"${s} t="b"><v>${v ? 1 : 0}</v></c>`); else if (XL.isErr(v)) addCell(r, c, `<c r="${a}"${s} t="e"><v>${xe(v.err)}</v></c>`); else addCell(r, c, `<c r="${a}"${s} t="s"><v>${si(String(v))}</v></c>`);
        } else if (cell.s) addCell(r, c, `<c r="${a}"${s}/>`);
      });
      sh.spill.forEach((v, k) => { const r = Math.floor(k / KC), c = k % KC; const a = XL.addr(r, c); if (typeof v === 'number') addCell(r, c, `<c r="${a}"><v>${v}</v></c>`); else if (typeof v === 'string') addCell(r, c, `<c r="${a}" t="str"><v>${xe(v)}</v></c>`); else if (typeof v === 'boolean') addCell(r, c, `<c r="${a}" t="b"><v>${v ? 1 : 0}</v></c>`); });
      Object.keys(sh.rowH).concat(Array.from(sh.hiddenR).map(String)).forEach(r => { if (!rows.has(+r)) rows.set(+r, []); });
      const rowStr = Array.from(rows.keys()).sort((a, b) => a - b).map(r => { const cs = rows.get(r).sort((a, b) => a[0] - b[0]).map(x => x[1]).join(''); const ht = sh.rowH[r] !== undefined ? ` ht="${(sh.rowH[r] * 0.75).toFixed(2)}" customHeight="1"` : ''; const hd = sh.hiddenR.has(r) ? ' hidden="1"' : ''; return `<row r="${r + 1}"${ht}${hd}>${cs}</row>`; }).join('');
      const ur = wb.usedRange(sh); const dim = ur ? XL.rangeStr(ur) : 'A1';
      const colsSet = new Set(Object.keys(sh.colW).map(Number).concat(Array.from(sh.hiddenC))); const cols = Array.from(colsSet).sort((a, b) => a - b).map(c => `<col min="${c + 1}" max="${c + 1}" width="${((sh.colW[c] !== undefined ? sh.colW[c] : sh.defaultW) - 5) / 7 + 0.71}" customWidth="1"${sh.hiddenC.has(c) ? ' hidden="1"' : ''}/>`).join('');
      let pane = ''; if (sh.freeze.r || sh.freeze.c) { const tl = XL.addr(sh.freeze.r, sh.freeze.c); pane = `<pane${sh.freeze.c ? ` xSplit="${sh.freeze.c}"` : ''}${sh.freeze.r ? ` ySplit="${sh.freeze.r}"` : ''} topLeftCell="${tl}" activePane="${sh.freeze.r && sh.freeze.c ? 'bottomRight' : sh.freeze.r ? 'bottomLeft' : 'topRight'}" state="frozen"/>`; }
      let extra = '';
      if (sh.merges.length) extra += `<mergeCells count="${sh.merges.length}">${sh.merges.map(g => `<mergeCell ref="${XL.rangeStr(g)}"/>`).join('')}</mergeCells>`;
      // conditional formatting
      let pr = 1; sh.cf.forEach(rule => {
        const ref = XL.rangeStr(rule.range); let inner = '';
        const st = rule.style || {}; const tl = XL.addr(rule.range.r1, rule.range.c1);
        const op = { '>': 'greaterThan', '<': 'lessThan', '>=': 'greaterThanOrEqual', '<=': 'lessThanOrEqual', '=': 'equal', '<>': 'notEqual', between: 'between', notbetween: 'notBetween' };
        if (rule.type === 'cell') inner = `<cfRule type="cellIs" dxfId="${dxf(st)}" priority="${pr++}" operator="${op[rule.op || '>']}"><formula>${xe(rule.v1)}</formula>${rule.v2 !== undefined && (rule.op === 'between' || rule.op === 'notbetween') ? `<formula>${xe(rule.v2)}</formula>` : ''}</cfRule>`;
        else if (rule.type === 'text') { const t = rule.op || 'contains'; const f = t === 'begins' ? `LEFT(${tl},${String(rule.v1).length})="${rule.v1}"` : t === 'ends' ? `RIGHT(${tl},${String(rule.v1).length})="${rule.v1}"` : t === 'notcontains' ? `ISERROR(SEARCH("${rule.v1}",${tl}))` : `NOT(ISERROR(SEARCH("${rule.v1}",${tl})))`; inner = `<cfRule type="expression" dxfId="${dxf(st)}" priority="${pr++}"><formula>${xe(f)}</formula></cfRule>`; }
        else if (rule.type === 'dup') inner = `<cfRule type="${rule.op === 'unique' ? 'uniqueValues' : 'duplicateValues'}" dxfId="${dxf(st)}" priority="${pr++}"/>`;
        else if (rule.type === 'top' || rule.type === 'bottom') inner = `<cfRule type="top10" dxfId="${dxf(st)}" priority="${pr++}"${rule.type === 'bottom' ? ' bottom="1"' : ''}${rule.percent ? ' percent="1"' : ''} rank="${rule.v1 || 10}"/>`;
        else if (rule.type === 'avg') inner = `<cfRule type="aboveAverage" dxfId="${dxf(st)}" priority="${pr++}"${rule.op === 'below' ? ' aboveAverage="0"' : ''}/>`;
        else if (rule.type === 'blanks') inner = `<cfRule type="containsBlanks" dxfId="${dxf(st)}" priority="${pr++}"><formula>LEN(TRIM(${tl}))=0</formula></cfRule>`;
        else if (rule.type === 'errors') inner = `<cfRule type="containsErrors" dxfId="${dxf(st)}" priority="${pr++}"><formula>ISERROR(${tl})</formula></cfRule>`;
        else if (rule.type === 'formula') inner = `<cfRule type="expression" dxfId="${dxf(st)}" priority="${pr++}"><formula>${xe(String(rule.v1).replace(/^=/, ''))}</formula></cfRule>`;
        else if (rule.type === 'databar') inner = `<cfRule type="dataBar" priority="${pr++}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(rule.color || '#638ec6')}"/></dataBar></cfRule>`;
        else if (rule.type === 'scale') { const cs = rule.colors || ['#f8696b', '#ffeb84', '#63be7b']; inner = `<cfRule type="colorScale" priority="${pr++}"><colorScale>${cs.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>'}${cs.map(c => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`; }
        else if (rule.type === 'icons') { const n3 = rule.set === '5arrows' ? 5 : 3; const nm = { '3arrows': '3Arrows', '3traffic': '3TrafficLights1', '3flags': '3Flags', '5arrows': '5Arrows' }[rule.set || '3arrows'] || '3Arrows'; inner = `<cfRule type="iconSet" priority="${pr++}"><iconSet iconSet="${nm}">${Array.from({ length: n3 }, (_, i) => `<cfvo type="percent" val="${Math.round(i * 100 / n3)}"/>`).join('')}</iconSet></cfRule>`; }
        if (inner) extra += `<conditionalFormatting sqref="${ref}">${inner}</conditionalFormatting>`;
      });
      sh._dxfUsed = true;
      if (sh.dv.length) extra += `<dataValidations count="${sh.dv.length}">${sh.dv.map(d => { const t = { list: 'list', whole: 'whole', decimal: 'decimal', date: 'date', textLength: 'textLength', custom: 'custom' }[d.type]; const f1 = d.type === 'list' ? (Array.isArray(d.list) ? '"' + d.list.join(',') + '"' : (String(d.list).startsWith('=') ? String(d.list).slice(1) : '"' + d.list + '"')) : d.v1; return `<dataValidation type="${t}"${d.type !== 'list' && d.type !== 'custom' ? ` operator="${d.op || 'between'}"` : ''} allowBlank="${d.allowBlank === false ? 0 : 1}" showInputMessage="1" showErrorMessage="1" sqref="${XL.rangeStr(d.range)}"${d.errMsg ? ` error="${xe(d.errMsg)}"` : ''}><formula1>${xe(f1)}</formula1>${d.v2 !== undefined && d.v2 !== '' ? `<formula2>${xe(d.v2)}</formula2>` : ''}</dataValidation>`; }).join('')}</dataValidations>`;
      let hl = ''; if (sh.links.size) { hl = '<hyperlinks>' + Array.from(sh.links).map(([k, u]) => { const { r, c } = XL.unkey(k); const id = 'rIdL' + (++relId); rels.push(`<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${xe(u)}" TargetMode="External"/>`); return `<hyperlink ref="${XL.addr(r, c)}" r:id="${id}"/>`; }).join('') + '</hyperlinks>'; }
      const af = sh.filter ? `<autoFilter ref="${XL.rangeStr(sh.filter.range)}"/>` : '';
      // charts
      let drawingTag = ''; if (sh.charts.length) {
        let dx = ''; const drels = [];
        sh.charts.forEach((ch, i) => { chartN++; const p = posToCell(sh, wb, ch.x, ch.y); const cid = 'rId' + (i + 1); drels.push(`<Relationship Id="${cid}" Type="${REL}/chart" Target="../charts/chart${chartN}.xml"/>`); charts.push([`xl/charts/chart${chartN}.xml`, chartXml(wb, ch)]);
          dx += `<xdr:oneCellAnchor><xdr:from><xdr:col>${p.c}</xdr:col><xdr:colOff>${p.co * 9525}</xdr:colOff><xdr:row>${p.r}</xdr:row><xdr:rowOff>${p.ro * 9525}</xdr:rowOff></xdr:from><xdr:ext cx="${Math.round(ch.w * 9525)}" cy="${Math.round(ch.h * 9525)}"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${cid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`; });
        drawings.push([`xl/drawings/drawing${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${REL}">${dx}</xdr:wsDr>`], [`xl/drawings/_rels/drawing${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drels.join('')}</Relationships>`]);
        const id = 'rIdD' + (++relId); rels.push(`<Relationship Id="${id}" Type="${REL}/drawing" Target="../drawings/drawing${n}.xml"/>`); drawingTag = `<drawing r:id="${id}"/>`;
      }
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet ${NS}><sheetPr>${sh.tabColor ? `<tabColor rgb="${argb(sh.tabColor)}"/>` : ''}</sheetPr><dimension ref="${dim}"/><sheetViews><sheetView${idx === wb.activeIndex ? ' tabSelected="1"' : ''}${sh.showGrid ? '' : ' showGridLines="0"'}${sh.zoom && sh.zoom !== 1 ? ` zoomScale="${Math.round(sh.zoom * 100)}"` : ''} workbookViewId="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="${(sh.defaultH * 0.75).toFixed(2)}"${sh.defaultH !== 20 ? ' customHeight="1"' : ''}/>${cols ? `<cols>${cols}</cols>` : ''}<sheetData>${rowStr}</sheetData>${af}${extra.includes('<mergeCells') ? extra.match(/<mergeCells[\s\S]*?<\/mergeCells>/)[0] : ''}${extra.replace(/<mergeCells[\s\S]*?<\/mergeCells>/, '')}${hl}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>${drawingTag}</worksheet>`;
      sheetsXml.push(xml); sheetRels.push(rels);
    });
    // dxfs were created while writing CF; rebuild styles after
    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet ${NS}>${nfs.length ? `<numFmts count="${nfs.length}">${nfs.join('')}</numFmts>` : ''}<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs></styleSheet>`;
    // fix dxfId order: dxfs were pushed in order of use so ids match
    const names = Object.values(wb.names).map(nm => `<definedName name="${xe(nm.name)}">${xe(nm.ref)}</definedName>`).join('');
    const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook ${NS}><bookViews><workbookView activeTab="${wb.activeIndex || 0}"/></bookViews><sheets>${wb.sheets.map((s, i) => `<sheet name="${xe(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${names ? `<definedNames>${names}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${wb.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>${drawings.filter(d => d[0].endsWith('.xml') && !d[0].includes('_rels')).map(d => `<Override PartName="/${d[0]}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join('')}${charts.map(c => `<Override PartName="/${c[0]}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`).join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const all = [['[Content_Types].xml', ct], ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`],
      ['xl/workbook.xml', wbXml], ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wb.sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rIdSt" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rIdSs" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`],
      ['xl/styles.xml', stylesXml], ['xl/sharedStrings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst ${NS} count="${strings.length}" uniqueCount="${strings.length}">${strings.map(s => `<si><t xml:space="preserve">${xe(s)}</t></si>`).join('')}</sst>`],
      ['docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Office Web</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`],
      ['docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Office Web</Application></Properties>`]];
    sheetsXml.forEach((x, i) => { all.push([`xl/worksheets/sheet${i + 1}.xml`, x]); if (sheetRels[i].length) all.push([`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRels[i].join('')}</Relationships>`]); });
    drawings.forEach(d => all.push(d)); charts.forEach(c => all.push(c));
    return OA.zip(all);
  };

  /* ---------------------------------------------------------------- READ */
  const IDX = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];
  const tag = (n, name) => n ? Array.from(n.getElementsByTagNameNS('*', name)) : [];
  const kid = (n, name) => { if (!n) return null; for (const c of n.children) if (c.localName === name) return c; return null; };
  const kids = (n, name) => n ? Array.from(n.children).filter(c => c.localName === name) : [];
  function colorOf(el) {
    if (!el) return null; const rgb = el.getAttribute('rgb'); let hex = null;
    if (rgb) hex = rgb.length === 8 ? rgb.slice(2) : rgb; else if (el.hasAttribute('theme')) hex = THEME[+el.getAttribute('theme')] || '000000'; else if (el.hasAttribute('indexed')) { const i = +el.getAttribute('indexed'); hex = i === 64 ? '000000' : i === 65 ? 'FFFFFF' : IDX[i] || '000000'; } else if (el.getAttribute('auto')) return null;
    if (!hex) return null; const tint = parseFloat(el.getAttribute('tint') || 0); if (tint) hex = OA.mixColor(hex, tint); return '#' + hex.toUpperCase();
  }
  function resolvePath(base, target) { if (target[0] === '/') return target.slice(1); const parts = base.split('/'); parts.pop(); target.split('/').forEach(p => { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); }); return parts.join('/'); }
  function relsOf(z, path) { const rp = path.replace(/([^/]+)$/, '_rels/$1.rels'); const out = {}; if (!z[rp]) return out; tag(OA.xmlDoc(z[rp]), 'Relationship').forEach(r => { out[r.getAttribute('Id')] = { type: r.getAttribute('Type').split('/').pop(), target: r.getAttribute('TargetMode') === 'External' ? r.getAttribute('Target') : resolvePath(path, r.getAttribute('Target')), ext: r.getAttribute('TargetMode') === 'External' }; }); return out; }
  const stripX = f => f.replace(/_xlfn\.|_xlws\./g, '');

  XL.fromXlsx = async function (buf) {
    const z = await OA.unzip(buf); if (!z['xl/workbook.xml']) throw new Error('This is not a valid .xlsx file');
    const wb = new XL.Workbook(); wb.sheets = []; wb.styles = [{}]; wb._styleMap = { '{}': 0 };
    const wbd = OA.xmlDoc(z['xl/workbook.xml']); const wrels = relsOf(z, 'xl/workbook.xml');
    // shared strings
    const sst = []; if (z['xl/sharedStrings.xml']) tag(OA.xmlDoc(z['xl/sharedStrings.xml']), 'si').forEach(si => { sst.push(tag(si, 't').filter(t => !t.parentNode || t.parentNode.localName !== 'rPh').map(t => t.textContent).join('')); });
    // styles
    const numFmts = {}; let fonts = [], fills = [], borders = [], xfs = [], dxfs = [];
    if (z['xl/styles.xml']) {
      const sd = OA.xmlDoc(z['xl/styles.xml']); tag(sd, 'numFmt').forEach(n => numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode'));
      const fe = kid(sd.documentElement, 'fonts'); fonts = kids(fe, 'font').map(f => ({ b: !!kid(f, 'b') && kid(f, 'b').getAttribute('val') !== '0', i: !!kid(f, 'i') && kid(f, 'i').getAttribute('val') !== '0', u: !!kid(f, 'u') && kid(f, 'u').getAttribute('val') !== 'none', st: !!kid(f, 'strike'), sz: kid(f, 'sz') ? +kid(f, 'sz').getAttribute('val') : 11, font: kid(f, 'name') ? kid(f, 'name').getAttribute('val') : 'Calibri', color: colorOf(kid(f, 'color')) }));
      fills = kids(kid(sd.documentElement, 'fills'), 'fill').map(f => { const p = kid(f, 'patternFill'); if (!p || p.getAttribute('patternType') !== 'solid') return null; return colorOf(kid(p, 'fgColor')); });
      const bs = n => { const b = kid(n, 'x') || n; return null; };
      borders = kids(kid(sd.documentElement, 'borders'), 'border').map(b => { const o = {}; [['left', 'bL'], ['right', 'bR'], ['top', 'bT'], ['bottom', 'bB']].forEach(([n, k]) => { const e = kid(b, n); if (e && e.getAttribute('style')) o[k] = { s: e.getAttribute('style').replace(/^(mediumDashed|dashDot|dashDotDot|mediumDashDot|slantDashDot|mediumDashDotDot)$/, 'dashed'), c: colorOf(kid(e, 'color')) || '#000000' }; }); return o; });
      xfs = kids(kid(sd.documentElement, 'cellXfs'), 'xf').map(x => {
        const st = {}; const f = fonts[+x.getAttribute('fontId') || 0]; if (f) { if (f.b) st.b = true; if (f.i) st.i = true; if (f.u) st.u = true; if (f.st) st.st = true; if (f.sz && f.sz !== 11) st.sz = f.sz; if (f.font && f.font !== 'Calibri') st.font = f.font; if (f.color && f.color !== '#000000') st.color = f.color; }
        const fl = fills[+x.getAttribute('fillId') || 0]; if (fl) st.bg = fl; const b = borders[+x.getAttribute('borderId') || 0]; if (b) Object.assign(st, b);
        const nid = x.getAttribute('numFmtId'); const nf = numFmts[nid] || XL.BUILTIN_FMT[nid]; if (nf && nf !== 'General') st.nf = nf === 'mm-dd-yy' ? 'm/d/yyyy' : nf;
        const al = kid(x, 'alignment'); if (al) { if (al.getAttribute('horizontal') && al.getAttribute('horizontal') !== 'general') st.ha = al.getAttribute('horizontal'); const v = al.getAttribute('vertical'); if (v && v !== 'bottom') st.va = v === 'center' ? 'middle' : v; if (al.getAttribute('wrapText') === '1') st.wrap = true; if (al.getAttribute('indent')) st.ind = +al.getAttribute('indent'); }
        return wb.styleIdx(st);
      });
      dxfs = tag(sd, 'dxf').filter(d => d.parentNode.localName === 'dxfs').map(d => { const st = {}; const f = kid(d, 'font'); if (f) { if (kid(f, 'b')) st.b = true; if (kid(f, 'i')) st.i = true; const c = colorOf(kid(f, 'color')); if (c) st.color = c; } const fl = kid(d, 'fill'); if (fl) { const p = kid(fl, 'patternFill'); const c = colorOf(kid(p, 'bgColor')) || colorOf(kid(p, 'fgColor')); if (c) st.bg = c; } return st; });
    }
    // defined names
    tag(wbd, 'definedName').forEach(n => { const nm = n.getAttribute('name'); if (/^_xlnm\./.test(nm) || n.getAttribute('hidden') === '1') return; wb.names[nm.toLowerCase()] = { name: nm, ref: n.textContent.replace(/^=/, '') }; });
    const sheetEls = tag(wbd, 'sheet'); let activeTab = 0; const bv = tag(wbd, 'workbookView')[0]; if (bv && bv.getAttribute('activeTab')) activeTab = +bv.getAttribute('activeTab');
    for (const se of sheetEls) {
      const rid = se.getAttribute('r:id') || se.getAttributeNS(REL, 'id') || Array.from(se.attributes).find(a => a.localName === 'id').value; const rel = wrels[rid]; if (!rel || !z[rel.target]) continue;
      const sh = new XL.Sheet(se.getAttribute('name')); wb.sheets.push(sh); const d = OA.xmlDoc(z[rel.target]); const srels = relsOf(z, rel.target); const root = d.documentElement;
      const sp = kid(root, 'sheetPr'); if (sp && kid(sp, 'tabColor')) sh.tabColor = colorOf(kid(sp, 'tabColor'));
      const sv = tag(root, 'sheetView')[0]; if (sv) { if (sv.getAttribute('showGridLines') === '0') sh.showGrid = false; if (sv.getAttribute('zoomScale')) sh.zoom = +sv.getAttribute('zoomScale') / 100; const pane = kid(sv, 'pane'); if (pane && pane.getAttribute('state') === 'frozen') sh.freeze = { r: +(pane.getAttribute('ySplit') || 0), c: +(pane.getAttribute('xSplit') || 0) }; }
      const sf = kid(root, 'sheetFormatPr'); if (sf && sf.getAttribute('defaultRowHeight')) sh.defaultH = Math.round(parseFloat(sf.getAttribute('defaultRowHeight')) * 4 / 3); if (sh.defaultH < 15) sh.defaultH = 20;
      kids(kid(root, 'cols'), 'col').forEach(c => { const min = +c.getAttribute('min') - 1, max = Math.min(+c.getAttribute('max') - 1, min + 300); const w = Math.round(parseFloat(c.getAttribute('width') || 8.43) * 7 + 5); for (let i = min; i <= max; i++) { if (c.getAttribute('hidden') === '1') sh.hiddenC.add(i); else if (c.getAttribute('customWidth') === '1' || w !== 64) sh.colW[i] = w; } if (c.getAttribute('style') && +c.getAttribute('style')) for (let i = min; i <= max; i++) sh.colStyle[i] = xfs[+c.getAttribute('style')] || 0; });
      const shared = {}; const arrays = [];
      kids(kid(root, 'sheetData'), 'row').forEach(row => {
        const r = +row.getAttribute('r') - 1; if (row.getAttribute('hidden') === '1') sh.hiddenR.add(r); if (row.getAttribute('customHeight') === '1' && row.getAttribute('ht')) sh.rowH[r] = Math.round(parseFloat(row.getAttribute('ht')) * 4 / 3);
        kids(row, 'c').forEach(c => {
          const ref = XL.parseAddr(c.getAttribute('r')); if (!ref) return; const t = c.getAttribute('t'); const s = xfs[+c.getAttribute('s') || 0] || 0; const fe = kid(c, 'f'), ve = kid(c, 'v'); let v;
          if (t === 's') v = sst[+ve.textContent]; else if (t === 'str') v = ve ? ve.textContent : ''; else if (t === 'inlineStr') v = tag(kid(c, 'is'), 't').map(x => x.textContent).join(''); else if (t === 'b') v = ve && ve.textContent === '1'; else if (t === 'e') v = XL.ERR(ve ? ve.textContent : '#VALUE!'); else if (ve && ve.textContent !== '') v = parseFloat(ve.textContent); else v = undefined;
          let cell = null;
          if (fe) {
            let src = fe.textContent; const ft = fe.getAttribute('t');
            if (ft === 'shared') { const si = fe.getAttribute('si'); if (src) shared[si] = { src: stripX(src), r: ref.r, c: ref.c }; else if (shared[si]) { try { const ast = XL.shiftAst(XL.parse(shared[si].src), ref.r - shared[si].r, ref.c - shared[si].c); src = XL.astToStr(ast); cell = { f: src, ast, s }; } catch (e) { } } if (src && !cell) { try { const ast = XL.parse(stripX(src)); cell = { f: XL.astToStr(ast), ast, s }; } catch (e) { } } }
            else if (ft === 'array') { try { const ast = XL.parse(stripX(src)); cell = { f: XL.astToStr(ast), ast, s }; arrays.push({ ref: fe.getAttribute('ref'), r: ref.r, c: ref.c }); } catch (e) { } }
            else if (src) { try { const ast = XL.parse(stripX(src)); cell = { f: XL.astToStr(ast), ast, s }; } catch (e) { cell = null; } }
            if (!cell && v !== undefined) cell = { v, s };
          } else if (v !== undefined) cell = { v, s }; else if (s) cell = { s };
          if (cell) { if (!cell.s) delete cell.s; sh.cells.set(key(ref.r, ref.c), cell); }
        });
      });
      // remove cached values inside array spill regions
      arrays.forEach(a => { const g = XL.parseRange(a.ref); if (!g || (g.r1 === g.r2 && g.c1 === g.c2)) return; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if (r !== a.r || c !== a.c) { const k = key(r, c); const x = sh.cells.get(k); if (x && x.f == null) { if (x.s) sh.cells.set(k, { s: x.s }); else sh.cells.delete(k); } } });
      kids(kid(root, 'mergeCells'), 'mergeCell').forEach(m => { const g = XL.parseRange(m.getAttribute('ref')); if (g) sh.merges.push(g); });
      tag(root, 'hyperlink').forEach(h => { const a = XL.parseAddr(h.getAttribute('ref').split(':')[0]); const rid = h.getAttribute('r:id') || (Array.from(h.attributes).find(x => x.localName === 'id') || {}).value; const rl = rid && srels[rid]; if (a && rl) sh.links.set(key(a.r, a.c), rl.target); else if (a && h.getAttribute('location')) sh.links.set(key(a.r, a.c), '#' + h.getAttribute('location')); });
      const af = kid(root, 'autoFilter'); if (af && af.getAttribute('ref')) { const g = XL.parseRange(af.getAttribute('ref')); if (g) sh.filter = { range: g, crit: {} }; }
      // conditional formatting
      kids(root, 'conditionalFormatting').forEach(cf => { const g = XL.parseRange((cf.getAttribute('sqref') || '').split(' ')[0]); if (!g) return; kids(cf, 'cfRule').forEach(r => {
        const t = r.getAttribute('type'); const style = dxfs[+r.getAttribute('dxfId')] || {}; const fs = kids(r, 'formula').map(f => f.textContent); const rule = { range: g, style };
        if (t === 'cellIs') { const om = { greaterThan: '>', lessThan: '<', greaterThanOrEqual: '>=', lessThanOrEqual: '<=', equal: '=', notEqual: '<>', between: 'between', notBetween: 'notbetween' }; Object.assign(rule, { type: 'cell', op: om[r.getAttribute('operator')] || '>', v1: fs[0], v2: fs[1] }); }
        else if (t === 'containsText') Object.assign(rule, { type: 'text', op: 'contains', v1: r.getAttribute('text') || '' }); else if (t === 'beginsWith') Object.assign(rule, { type: 'text', op: 'begins', v1: r.getAttribute('text') || '' }); else if (t === 'endsWith') Object.assign(rule, { type: 'text', op: 'ends', v1: r.getAttribute('text') || '' });
        else if (t === 'duplicateValues' || t === 'uniqueValues') Object.assign(rule, { type: 'dup', op: t === 'uniqueValues' ? 'unique' : 'dup' }); else if (t === 'top10') Object.assign(rule, { type: r.getAttribute('bottom') === '1' ? 'bottom' : 'top', v1: +r.getAttribute('rank') || 10, percent: r.getAttribute('percent') === '1' });
        else if (t === 'aboveAverage') Object.assign(rule, { type: 'avg', op: r.getAttribute('aboveAverage') === '0' ? 'below' : 'above' }); else if (t === 'containsBlanks') Object.assign(rule, { type: 'blanks' }); else if (t === 'containsErrors') Object.assign(rule, { type: 'errors' }); else if (t === 'expression') Object.assign(rule, { type: 'formula', v1: fs[0] });
        else if (t === 'dataBar') Object.assign(rule, { type: 'databar', color: colorOf(tag(r, 'color')[0]) || '#638ec6' }); else if (t === 'colorScale') Object.assign(rule, { type: 'scale', colors: tag(r, 'color').map(c => colorOf(c) || '#ffffff') }); else if (t === 'iconSet') { const nm = (tag(r, 'iconSet')[0] || r).getAttribute('iconSet') || '3Arrows'; Object.assign(rule, { type: 'icons', set: /5/.test(nm) ? '5arrows' : /Traffic/.test(nm) ? '3traffic' : /Flags/.test(nm) ? '3flags' : '3arrows' }); } else return;
        rule.id = XL.uid(); sh.cf.push(rule); }); });
      tag(root, 'dataValidation').forEach(dv => { const g = XL.parseRange((dv.getAttribute('sqref') || '').split(' ')[0]); if (!g) return; const type = dv.getAttribute('type'); if (!type || type === 'none') return; const f1 = tag(dv, 'formula1')[0], f2 = tag(dv, 'formula2')[0]; const rule = { range: g, type, op: dv.getAttribute('operator') || 'between', allowBlank: dv.getAttribute('allowBlank') !== '0', errMsg: dv.getAttribute('error') || '' };
        if (type === 'list') { const t = f1 ? f1.textContent : ''; rule.list = /^".*"$/.test(t) ? t.slice(1, -1).split(',') : '=' + t; } else { rule.v1 = f1 ? f1.textContent : ''; rule.v2 = f2 ? f2.textContent : ''; } sh.dv.push(rule); });
      // comments (notes)
      Object.values(srels).filter(r => r.type === 'comments').forEach(r => { if (!z[r.target]) return; tag(OA.xmlDoc(z[r.target]), 'comment').forEach(cm => { const a = XL.parseAddr(cm.getAttribute('ref')); if (a) sh.notes.set(key(a.r, a.c), tag(cm, 't').map(t => t.textContent).join('')); }); });
      // charts
      Object.values(srels).filter(r => r.type === 'drawing').forEach(dr => {
        if (!z[dr.target]) return; const dd = OA.xmlDoc(z[dr.target]); const drel = relsOf(z, dr.target);
        Array.from(dd.documentElement.children).forEach(anc => {
          const gf = tag(anc, 'chart')[0]; if (!gf) return; const rid = gf.getAttribute('r:id') || (Array.from(gf.attributes).find(a => a.localName === 'id') || {}).value; const cr = drel[rid]; if (!cr || !z[cr.target]) return;
          const cd = OA.xmlDoc(z[cr.target]); const plot = tag(cd, 'plotArea')[0]; if (!plot) return; let type = 'col';
          const kidsP = Array.from(plot.children).map(c => c.localName);
          if (kidsP.includes('barChart')) type = (tag(plot, 'barDir')[0] || { getAttribute: () => 'col' }).getAttribute('val') === 'bar' ? 'bar' : 'col'; else if (kidsP.includes('lineChart')) type = 'line'; else if (kidsP.includes('pieChart')) type = 'pie'; else if (kidsP.includes('doughnutChart')) type = 'doughnut'; else if (kidsP.includes('areaChart')) type = 'area'; else if (kidsP.includes('scatterChart')) type = 'scatter'; else return;
          const refs = tag(cd, 'f').map(f => f.textContent); const rects = refs.map(r => { const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(r); return m ? { sh: m[1] ? m[1].replace(/''/g, "'") : m[2], g: XL.parseRange(m[3]) } : null; }).filter(x => x && x.g); if (!rects.length) return;
          let r1 = 1e9, c1 = 1e9, r2 = -1, c2 = -1; rects.forEach(x => { r1 = Math.min(r1, x.g.r1); c1 = Math.min(c1, x.g.c1); r2 = Math.max(r2, x.g.r2); c2 = Math.max(c2, x.g.c2); });
          const valRect = rects.find(x => x.g.r1 !== x.g.r2 || x.g.c1 === x.g.c2); const rowsAsSeries = tag(cd, 'val')[0] ? (() => { const f = tag(tag(cd, 'val')[0], 'f')[0]; const m = f && /!(.+)$/.exec(f.textContent); const g = m && XL.parseRange(m[1]); return g ? g.r1 === g.r2 && g.c1 !== g.c2 : false; })() : false;
          const title = tag(tag(cd, 'title')[0], 't').map(t => t.textContent).join('') || '';
          const fr = tag(anc, 'from')[0]; let x = 0, y = 0; if (fr) { const col = +tag(fr, 'col')[0].textContent, row = +tag(fr, 'row')[0].textContent; for (let i = 0; i < col; i++) x += wb.colWidth(sh, i); for (let i = 0; i < row; i++) y += wb.rowHeight(sh, i); x += (+tag(fr, 'colOff')[0].textContent) / 9525; y += (+tag(fr, 'rowOff')[0].textContent) / 9525; }
          let w = 480, h = 288; const to = tag(anc, 'to')[0], ext = tag(anc, 'ext')[0]; if (ext && ext.getAttribute('cx')) { w = +ext.getAttribute('cx') / 9525; h = +ext.getAttribute('cy') / 9525; } else if (to && fr) { let x2 = 0, y2 = 0; const col = +tag(to, 'col')[0].textContent, row = +tag(to, 'row')[0].textContent; for (let i = 0; i < col; i++) x2 += wb.colWidth(sh, i); for (let i = 0; i < row; i++) y2 += wb.rowHeight(sh, i); w = Math.max(120, x2 - x); h = Math.max(100, y2 - y); }
          sh.charts.push({ id: XL.uid(), type, title, ref: XL.quoteSheet(rects[0].sh) + '!' + XL.rangeStr({ r1, c1, r2, c2 }), rowsAsSeries, x, y, w, h, legend: !!tag(cd, 'legend')[0], stacked: !!tag(cd, 'grouping').find(g => g.getAttribute('val') === 'stacked') });
        });
      });
    }
    if (!wb.sheets.length) wb.addSheet('Sheet1');
    wb.activeIndex = Math.min(activeTab, wb.sheets.length - 1); wb.dirty(); return wb;
  };
  XL.toHTMLTable = function (wb, sh, g) { return wb.rangeHTML(sh, g); };
})(window);
