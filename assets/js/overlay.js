/* =====================================================================
   Ассистент — предпросмотр документа, наложение печати 40 мм, выгрузка «изображение в PDF».
   Печати берутся из кэша (Stamp Studio → «Добавить в файл», до 4 штук), ставятся мышью, можно несколько раз на одном файле.
   В DOCX печать вставляется плавающей картинкой, привязанной к абзацу, в XLSX — картинкой, привязанной к ячейке:
   так она остаётся рядом с нужным текстом, даже если разбивка на страницы в Word чуть отличается от предпросмотра.
   ===================================================================== */
(function () {
  'use strict';
  const E = window.AssistantEngine, I = E.internals, { parseXml, serialize, kids, firstKid, byNS, NS_S, NS_W } = I;
  const PX_MM = 96 / 25.4, STAMP_MM = 40, STAMP_PX = STAMP_MM * PX_MM, EMU_MM = 36000, EMU_PX = 9525, STAMP_EMU = STAMP_MM * EMU_MM;
  const NS = {
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing', a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture', r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships', ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
    xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  };
  const extOf = n => (String(n).split('.').pop() || '').toLowerCase();
  const dataUrlToBytes = u => { const b = atob(u.split(',')[1]); const a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  /* ---------------------------------------------------------------
     ПРЕДПРОСМОТР: DOCX (docx-preview) и XLSX (собственная отрисовка)
     --------------------------------------------------------------- */
  async function renderDocx(blob, host) {
    if (!window.docx) throw new Error('Не загрузилась библиотека предпросмотра Word');
    host.innerHTML = '';
    await window.docx.renderAsync(blob, host, null, { className: 'docx', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: false, useBase64URL: true });
    const pages = Array.from(host.querySelectorAll('section.docx'));
    pages.forEach((p, i) => { p.classList.add('ov-page'); p.dataset.page = i; });
    return { kind: 'docx', pages };
  }

  const NUMFMT_BUILTIN = { 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 37: '#,##0', 38: '#,##0', 39: '#,##0.00', 40: '#,##0.00' };
  const colPx = w => Math.max(0, Math.round(w * 7 + 5));
  function excelDate(n) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); return String(d.getUTCDate()).padStart(2, '0') + '.' + String(d.getUTCMonth() + 1).padStart(2, '0') + '.' + d.getUTCFullYear(); }
  function fmtNumber(v, code) {
    const n = parseFloat(v); if (!isFinite(n)) return v;
    if (/[dmy]/i.test(code.replace(/\[[^\]]*\]|"[^"]*"/g, '')) && !/0\.0|#/.test(code)) return excelDate(n);
    const pct = /%/.test(code), dm = /\.([0#]+)/.exec(code), dec = dm ? dm[1].length : 0, grp = /#,#|,0/.test(code);
    let x = pct ? n * 100 : n, s = Math.abs(x).toFixed(dec);
    if (grp) s = s.replace(/^(\d+)/, m => m.replace(/\B(?=(\d{3})+(?!\d))/g, ' '));
    return (x < 0 ? '-' : '') + s.replace('.', ',') + (pct ? '%' : '');
  }
  async function loadXlsxModel(zip) {
    const sst = await I.readShared(zip), sf = zip.file('xl/styles.xml'), st = sf ? parseXml(await sf.async('string')) : null;
    const fonts = [], fills = [], borders = [], xfs = [], numFmts = {};
    if (st) {
      byNS(st, NS_S, 'numFmt').forEach(n => { numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode'); });
      const f = byNS(st, NS_S, 'fonts')[0]; if (f) kids(f, 'font').forEach(x => { const sz = firstKid(x, 'sz'), col = firstKid(x, 'color'); fonts.push({ b: !!firstKid(x, 'b'), i: !!firstKid(x, 'i'), sz: sz ? parseFloat(sz.getAttribute('val')) : 11, color: col && col.getAttribute('rgb') ? '#' + col.getAttribute('rgb').slice(-6) : '' }); });
      const fl = byNS(st, NS_S, 'fills')[0]; if (fl) kids(fl, 'fill').forEach(x => { const pf = firstKid(x, 'patternFill'), fg = pf && firstKid(pf, 'fgColor'); fills.push(pf && pf.getAttribute('patternType') === 'solid' && fg && fg.getAttribute('rgb') ? '#' + fg.getAttribute('rgb').slice(-6) : ''); });
      const bd = byNS(st, NS_S, 'borders')[0]; if (bd) kids(bd, 'border').forEach(x => { const o = {}; ['left', 'right', 'top', 'bottom'].forEach(s => { const e = firstKid(x, s); o[s] = !!(e && e.getAttribute('style')); }); borders.push(o); });
      const cx = byNS(st, NS_S, 'cellXfs')[0]; if (cx) kids(cx, 'xf').forEach(x => { const al = firstKid(x, 'alignment'); xfs.push({ font: +x.getAttribute('fontId') || 0, fill: +x.getAttribute('fillId') || 0, border: +x.getAttribute('borderId') || 0, fmt: x.getAttribute('numFmtId') || '0', h: al ? al.getAttribute('horizontal') : '', v: al ? al.getAttribute('vertical') : '', wrap: al ? al.getAttribute('wrapText') === '1' : false }); });
    }
    return { sst, fonts, fills, borders, xfs, numFmts };
  }
  async function renderXlsx(blob, host) {
    const zip = await I.JSZip.loadAsync(await blob.arrayBuffer()), M = await loadXlsxModel(zip);
    const wb = zip.file('xl/workbook.xml'), names = wb ? byNS(parseXml(await wb.async('string')), NS_S, 'sheet').map(s => s.getAttribute('name')) : [];
    const files = Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort((a, b) => parseInt(a.replace(/\D/g, ''), 10) - parseInt(b.replace(/\D/g, ''), 10));
    host.innerHTML = ''; const pages = [];
    for (let si = 0; si < files.length; si++) {
      const doc = parseXml(await zip.file(files[si]).async('string'));
      const fmt = byNS(doc, NS_S, 'sheetFormatPr')[0], defW = fmt && fmt.getAttribute('defaultColWidth') ? parseFloat(fmt.getAttribute('defaultColWidth')) : 8.43, defH = fmt && fmt.getAttribute('defaultRowHeight') ? parseFloat(fmt.getAttribute('defaultRowHeight')) : 15;
      const rowEls = byNS(doc, NS_S, 'row'); let maxR = 0, maxC = 1;
      const cells = {}; // r -> c -> {el}
      rowEls.forEach(r => { const rn = +r.getAttribute('r'); maxR = Math.max(maxR, rn); cells[rn] = {}; kids(r, 'c').forEach(c => { const ref = I.splitRef(c.getAttribute('r') || ''); if (ref) { cells[rn][ref.col] = c; maxC = Math.max(maxC, ref.col); } }); });
      const merges = byNS(doc, NS_S, 'mergeCell').map(m => { const [a, b] = m.getAttribute('ref').split(':'); const A = I.splitRef(a), B = I.splitRef(b || a); return { r1: A.row, c1: A.col, r2: B.row, c2: B.col }; });
      merges.forEach(m => { maxR = Math.max(maxR, m.r2); maxC = Math.max(maxC, m.c2); });
      maxR = Math.min(maxR, 700); maxC = Math.min(maxC, 40);
      const colW = new Array(maxC + 1).fill(colPx(defW)); colW[0] = 0;
      byNS(doc, NS_S, 'col').forEach(c => { const w = c.getAttribute('hidden') === '1' ? 0 : colPx(parseFloat(c.getAttribute('width') || defW)); for (let k = +c.getAttribute('min'); k <= Math.min(+c.getAttribute('max'), maxC); k++) colW[k] = w; });
      const rowH = new Array(maxR + 1).fill(Math.round(defH * 96 / 72)); rowH[0] = 0;
      rowEls.forEach(r => { const rn = +r.getAttribute('r'); if (rn > maxR) return; if (r.getAttribute('hidden') === '1') rowH[rn] = 0; else if (r.getAttribute('ht')) rowH[rn] = Math.round(parseFloat(r.getAttribute('ht')) * 96 / 72); });
      const mAt = {}, covered = new Set();
      merges.forEach(m => { mAt[m.r1 + ':' + m.c1] = m; for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(r + ':' + c); });
      const wrap = document.createElement('div'); wrap.className = 'ov-page xl-sheet'; wrap.dataset.page = pages.length; wrap.dataset.sheet = si;
      const title = document.createElement('div'); title.className = 'xl-title'; title.textContent = 'Лист: ' + (names[si] || (si + 1)); wrap.appendChild(title);
      const holder = document.createElement('div'); holder.className = 'xl-holder'; wrap.appendChild(holder);
      const tbl = document.createElement('table'); tbl.className = 'xl-table';
      tbl.style.width = colW.reduce((a, b) => a + b, 0) + 'px';
      const cg = document.createElement('colgroup'); for (let c = 1; c <= maxC; c++) { const e = document.createElement('col'); e.style.width = colW[c] + 'px'; cg.appendChild(e); } tbl.appendChild(cg);
      for (let r = 1; r <= maxR; r++) {
        if (!rowH[r]) continue;
        const tr = document.createElement('tr'); tr.style.height = rowH[r] + 'px'; const skip = new Set();
        for (let c = 1; c <= maxC; c++) {
          if (!colW[c] || covered.has(r + ':' + c) || skip.has(c)) continue;
          const el = cells[r] && cells[r][c], td = document.createElement('td'); const m = mAt[r + ':' + c];
          if (m) { td.rowSpan = m.r2 - m.r1 + 1; td.colSpan = m.c2 - m.c1 + 1; }
          if (el) {
            const xf = M.xfs[+el.getAttribute('s') || 0] || {}, font = M.fonts[xf.font] || {}, t = el.getAttribute('t'), v = firstKid(el, 'v');
            let text = I.xlCellText(el, M.sst);
            if (!t || t === 'n') { if (text !== '') { const code = M.numFmts[xf.fmt] || NUMFMT_BUILTIN[xf.fmt] || ((+xf.fmt >= 14 && +xf.fmt <= 22) ? 'dd.mm.yyyy' : ''); if (code) text = fmtNumber(text, code); } }
            td.textContent = text;
            if (font.b) td.style.fontWeight = '700'; if (font.i) td.style.fontStyle = 'italic'; if (font.sz) td.style.fontSize = Math.round(font.sz * 96 / 72) + 'px'; if (font.color && font.color !== '#000000') td.style.color = font.color;
            const fill = M.fills[xf.fill]; if (fill) td.style.background = fill;
            const bd = M.borders[xf.border]; if (bd) ['left', 'right', 'top', 'bottom'].forEach(s => { if (bd[s]) td.style['border' + s[0].toUpperCase() + s.slice(1)] = '1px solid #000'; });
            td.style.textAlign = xf.h === 'center' || xf.h === 'centerContinuous' ? 'center' : xf.h === 'right' ? 'right' : (!xf.h && (!t || t === 'n') && text !== '' ? 'right' : 'left');
            td.style.verticalAlign = xf.v === 'center' ? 'middle' : xf.v === 'top' ? 'top' : 'bottom';
            if (xf.wrap) td.style.whiteSpace = 'normal';
            // как в Excel: текст без переноса, прижатый влево, вытекает в пустые соседние ячейки (иначе в предпросмотре «половина надписи не видна»)
            if (!m && text !== '' && !xf.wrap && (xf.h === 'left' || (!xf.h && (t === 's' || t === 'inlineStr' || t === 'str')))) {
              let k = c + 1;
              while (k <= maxC && colW[k] && !covered.has(r + ':' + k) && !mAt[r + ':' + k]) {
                const ne = cells[r] && cells[r][k]; if (ne && I.xlCellText(ne, M.sst) !== '') break;
                const nx = ne ? M.xfs[+ne.getAttribute('s') || 0] : null, nb = nx && M.borders[nx.border]; if (nb && (nb.left || nb.right || nb.top || nb.bottom)) break;
                td.colSpan = (td.colSpan || 1) + 1; skip.add(k); k++;
              }
            }
          }
          tr.appendChild(td);
        }
        tbl.appendChild(tr);
      }
      holder.appendChild(tbl); host.appendChild(wrap);
      const cumC = [0], cumR = [0]; for (let c = 1; c <= maxC; c++) cumC[c] = cumC[c - 1] + colW[c]; for (let r = 1; r <= maxR; r++) cumR[r] = cumR[r - 1] + rowH[r];
      wrap._geom = { cumC, cumR, maxC, maxR, file: files[si], holder }; holder._wrap = wrap;
      pages.push(wrap);
    }
    return { kind: 'xlsx', pages };
  }
  async function renderPreview(blob, name, host) {
    const ext = extOf(name);
    if (ext === 'docx' || ext === 'dotx') return renderDocx(blob, host);
    if (ext === 'xlsx' || ext === 'xlsm') return renderXlsx(blob, host);
    if (ext === 'xls') { const cv = await I.xlsToXlsx(await blob.arrayBuffer()); return renderXlsx(new Blob([cv.data]), host); }
    throw new Error('Предпросмотр: формат .' + ext + ' не поддерживается. Нужен .docx, .xlsx или .xls');
  }

  /* ---------------------------------------------------------------
     ВЫГРУЗКА «ИЗОБРАЖЕНИЕ В PDF»: страницы → картинки → PDF
     --------------------------------------------------------------- */
  async function pagesToPdf(pages, kind) {
    if (!window.html2canvas || !window.PDFLib) throw new Error('Не загрузились библиотеки для PDF');
    // html2canvas неверно считает размеры при CSS-масштабе страницы: на время съёмки возвращаем 100 %
    const de = document.documentElement, prevZoom = de.style.zoom, sx = window.scrollX, sy = window.scrollY;
    de.style.zoom = '1';
    try { return await pagesToPdfRaw(pages, kind); }
    finally { de.style.zoom = prevZoom; window.scrollTo(sx, sy); }
  }
  async function pagesToPdfRaw(pages, kind) {
    const pdf = await window.PDFLib.PDFDocument.create(), A4 = 1.4142;
    for (const page of pages) {
      const target = kind === 'xlsx' ? (page._geom.holder) : page;
      const prevMax = target.style.maxHeight, prevOv = target.style.overflow; target.style.maxHeight = 'none'; target.style.overflow = 'visible';
      const canvas = await window.html2canvas(target, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false });
      target.style.maxHeight = prevMax; target.style.overflow = prevOv;
      const slices = [];
      if (kind === 'xlsx') { const h = Math.round(canvas.width * A4); for (let y = 0; y < canvas.height; y += h) slices.push([y, Math.min(h, canvas.height - y)]); }
      else slices.push([0, canvas.height]);
      for (const [y, h] of slices) {
        const c = document.createElement('canvas'); c.width = canvas.width; c.height = kind === 'xlsx' ? Math.round(canvas.width * A4) : h;
        const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); g.drawImage(canvas, 0, y, canvas.width, h, 0, 0, canvas.width, h);
        const jpg = await pdf.embedJpg(dataUrlToBytes(c.toDataURL('image/jpeg', 0.92)));
        const wPt = c.width * 0.375, hPt = c.height * 0.375;          // scale 2 → 1 CSS px = 0.75 pt, значит 0.375 pt на пиксель холста
        const pg = pdf.addPage([wPt, hPt]); pg.drawImage(jpg, { x: 0, y: 0, width: wPt, height: hPt });
      }
    }
    return pdf.save();
  }
  async function blobToPdf(blob, name) {         // без предпросмотра: рисуем документ за экраном
    const host = document.createElement('div'); host.className = 'ov-offscreen'; document.body.appendChild(host);
    try { const r = await renderPreview(blob, name, host); await new Promise(res => setTimeout(res, 150)); return await pagesToPdf(r.pages, r.kind); }
    finally { host.remove(); }
  }

  /* ---------------------------------------------------------------
     ВСТРАИВАНИЕ ПЕЧАТИ В ФАЙЛ
     --------------------------------------------------------------- */
  async function addContentTypePng(zip) {
    const f = zip.file('[Content_Types].xml'); const doc = parseXml(await f.async('string')), root = doc.documentElement;
    if (!kids(root, 'Default').some(d => (d.getAttribute('Extension') || '').toLowerCase() === 'png')) { const d = doc.createElementNS(NS.ct, 'Default'); d.setAttribute('Extension', 'png'); d.setAttribute('ContentType', 'image/png'); root.insertBefore(d, root.firstChild); zip.file('[Content_Types].xml', serialize(doc)); }
  }
  async function addOverride(zip, partName, type) {
    const f = zip.file('[Content_Types].xml'); const doc = parseXml(await f.async('string')), root = doc.documentElement;
    if (kids(root, 'Override').some(d => d.getAttribute('PartName') === partName)) return;
    const o = doc.createElementNS(NS.ct, 'Override'); o.setAttribute('PartName', partName); o.setAttribute('ContentType', type); root.appendChild(o); zip.file('[Content_Types].xml', serialize(doc));
  }
  async function addRel(zip, relsPath, type, target) {
    let doc, root; const f = zip.file(relsPath);
    if (f) { doc = parseXml(await f.async('string')); root = doc.documentElement; }
    else { doc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="' + NS.rel + '"/>'); root = doc.documentElement; }
    const ids = kids(root, 'Relationship').map(r => r.getAttribute('Id')); let n = 1; while (ids.includes('rId' + n)) n++;
    const id = 'rId' + n, rel = doc.createElementNS(NS.rel, 'Relationship'); rel.setAttribute('Id', id); rel.setAttribute('Type', type); rel.setAttribute('Target', target); root.appendChild(rel);
    zip.file(relsPath, serialize(doc)); return id;
  }
  const REL_IMAGE = NS.r + '/image', REL_DRAWING = NS.r + '/drawing';

  /** Позиция печати в DOCX: ближайший абзац предпросмотра → абзац исходного файла */
  function docxAnchorFor(st, pageEl, srcParas) {
    const rendered = Array.from(pageEl.closest('.docx-wrapper').querySelectorAll('section.docx article p'));
    const norm = s => (s || '').replace(/\s+/g, ' ').trim();
    const rt = rendered.map(p => norm(p.textContent)), srcT = srcParas.map(x => norm(I.mapParagraph(x.p).text));
    const map = new Array(rendered.length).fill(-1);
    if (rendered.length === srcParas.length) rendered.forEach((_, i) => { map[i] = i; });
    else { let j = 0; rendered.forEach((_, i) => { if (!rt[i]) return; for (let k = j; k < Math.min(srcT.length, j + 60); k++) if (srcT[k] === rt[i]) { map[i] = k; j = k + 1; break; } }); }
    const pr = pageEl.getBoundingClientRect(), zk = pageEl.offsetWidth ? pr.width / pageEl.offsetWidth : 1;   // zk — масштаб сайта (0,8)
    const cy = pr.top + (st.y + STAMP_PX / 2) * zk;
    let best = -1, bd = 1e9;
    rendered.forEach((p, i) => { if (map[i] < 0 || !pageEl.contains(p)) return; const r = p.getBoundingClientRect(); const d = cy < r.top ? r.top - cy : cy > r.bottom ? cy - r.bottom : 0; if (d < bd) { bd = d; best = i; } });
    if (best < 0) return null;
    const p = rendered[best], r = p.getBoundingClientRect(), td = p.closest('td');
    const stampLeft = pr.left + st.x * zk, stampTop = pr.top + st.y * zk;
    let hFrom = 'page', hOff = st.x;
    if (td) { const tr = td.getBoundingClientRect(), pad = parseFloat(getComputedStyle(td).paddingLeft) || 0; hFrom = 'column'; hOff = (stampLeft - (tr.left + pad * zk)) / zk; }
    return { index: map[best], hFrom, hEmu: Math.round(hOff * EMU_PX), vEmu: Math.round((stampTop - r.top) / zk * EMU_PX) };
  }
  /** placements: [{ key, anchor }], images: { key: Uint8Array } — у каждой печати свой рисунок */
  async function stampDocx(buf, placements, images) {
    const zip = await I.JSZip.loadAsync(buf), docXml = zip.file('word/document.xml');
    if (!docXml) throw new Error('Не найден word/document.xml');
    const doc = parseXml(await docXml.async('string'));
    const T = window.AssistantTech, paras = [];
    (function walk(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; const nm = I.lname(c); if (nm === 'p') paras.push({ p: c }); else if (['tbl', 'tr', 'tc', 'sdt', 'sdtContent'].includes(nm)) walk(c); } })(byNS(doc, NS_W, 'body')[0]);
    const rids = {}, names = {}; await addContentTypePng(zip);
    for (const key of Array.from(new Set(placements.map(p => p.key)))) {
      let n = 1; while (zip.file('word/media/stamp' + n + '.png')) n++;
      zip.file('word/media/stamp' + n + '.png', images[key]); names[key] = 'stamp' + n + '.png';
      rids[key] = await addRel(zip, 'word/_rels/document.xml.rels', REL_IMAGE, 'media/stamp' + n + '.png');
    }
    let docId = 9000; byNS(doc, NS.wp, 'docPr').forEach(d => { docId = Math.max(docId, (+d.getAttribute('id') || 0) + 1); });
    let placed = 0, skipped = 0;
    for (const st of placements) {
      const a = st.anchor; if (!a || !paras[a.index]) { skipped++; continue; }
      const xml = `<w:r xmlns:w="${NS_W}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:r="${NS.r}"><w:drawing>` +
        `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${251658240 + placed}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
        `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${a.hFrom}"><wp:posOffset>${a.hEmu}</wp:posOffset></wp:positionH>` +
        `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${a.vEmu}</wp:posOffset></wp:positionV>` +
        `<wp:extent cx="${STAMP_EMU}" cy="${STAMP_EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
        `<wp:docPr id="${docId}" name="Печать ${docId}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${NS.pic}"><pic:pic>` +
        `<pic:nvPicPr><pic:cNvPr id="0" name="${names[st.key]}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rids[st.key]}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${STAMP_EMU}" cy="${STAMP_EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
      const node = doc.importNode(parseXml(xml).documentElement, true);
      paras[a.index].p.appendChild(node); docId++; placed++;
    }
    zip.file('word/document.xml', serialize(doc));
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), placed, skipped };
  }
  async function stampXlsx(buf, placements, images) {
    const zip = await I.JSZip.loadAsync(buf), media = {}; await addContentTypePng(zip);
    for (const key of Array.from(new Set(placements.map(p => p.key)))) {
      let n = 1; while (zip.file('xl/media/stamp' + n + '.png')) n++;
      zip.file('xl/media/stamp' + n + '.png', images[key]); media[key] = 'stamp' + n + '.png';
    }
    const bySheet = {}; placements.forEach(p => { (bySheet[p.sheetFile] = bySheet[p.sheetFile] || []).push(p); });
    let placed = 0;
    for (const file of Object.keys(bySheet)) {
      const sheetName = file.split('/').pop(), relsPath = 'xl/worksheets/_rels/' + sheetName + '.rels';
      const doc = parseXml(await zip.file(file).async('string')), root = doc.documentElement;
      let drawingPath = null;
      const rf = zip.file(relsPath);
      if (rf) { const rd = parseXml(await rf.async('string')); const r = kids(rd.documentElement, 'Relationship').find(x => /\/drawing$/.test(x.getAttribute('Type') || '')); if (r) drawingPath = 'xl/' + r.getAttribute('Target').replace(/^\.\.\//, '').replace(/^\/xl\//, ''); }
      const hasDrawing = drawingPath && zip.file(drawingPath);
      if (!hasDrawing) {
        let k = 1; while (zip.file('xl/drawings/drawing' + k + '.xml')) k++;
        drawingPath = 'xl/drawings/drawing' + k + '.xml';
        zip.file(drawingPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}"></xdr:wsDr>`);
        await addOverride(zip, '/' + drawingPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
        const rid = await addRel(zip, relsPath, REL_DRAWING, '../drawings/drawing' + k + '.xml');
        if (!root.getAttribute('xmlns:r')) root.setAttribute('xmlns:r', NS.r);
        const el = doc.createElementNS(NS_S, 'drawing'); el.setAttributeNS(NS.r, 'r:id', rid);
        const after = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
        const ref = Array.from(root.childNodes).find(c => c.nodeType === 1 && after.includes(I.lname(c)));
        if (ref) root.insertBefore(el, ref); else root.appendChild(el);
        zip.file(file, serialize(doc));
      }
      const drelsPath = drawingPath.replace(/drawings\/([^/]+)$/, 'drawings/_rels/$1.rels');
      const imgRid = {};
      for (const key of Array.from(new Set(bySheet[file].map(p => p.key)))) imgRid[key] = await addRel(zip, drelsPath, REL_IMAGE, '../media/' + media[key]);
      let xml = await zip.file(drawingPath).async('string');
      const m = /<(\w+:)?wsDr[^>]*>/.exec(xml), pre = m && m[1] ? m[1] : '';
      const ids = Array.from(xml.matchAll(/cNvPr[^>]*\sid="(\d+)"/g)).map(x => +x[1]); let nid = Math.max.apply(null, ids.concat([1])) + 1;
      let add = '';
      for (const st of bySheet[file]) {
        add += `<${pre}oneCellAnchor><${pre}from><${pre}col>${st.col}</${pre}col><${pre}colOff>${st.colOff}</${pre}colOff><${pre}row>${st.row}</${pre}row><${pre}rowOff>${st.rowOff}</${pre}rowOff></${pre}from>` +
          `<${pre}ext cx="${STAMP_EMU}" cy="${STAMP_EMU}"/><${pre}pic xmlns:a="${NS.a}" xmlns:r="${NS.r}"><${pre}nvPicPr><${pre}cNvPr id="${nid}" name="Печать ${nid}"/><${pre}cNvPicPr><a:picLocks noChangeAspect="1"/></${pre}cNvPicPr></${pre}nvPicPr>` +
          `<${pre}blipFill><a:blip r:embed="${imgRid[st.key]}"/><a:stretch><a:fillRect/></a:stretch></${pre}blipFill><${pre}spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${STAMP_EMU}" cy="${STAMP_EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></${pre}spPr></${pre}pic><${pre}clientData/></${pre}oneCellAnchor>`;
        nid++; placed++;
      }
      xml = /\/>\s*$/.test(xml.trim()) && !/<\/(\w+:)?wsDr>/.test(xml) ? xml.replace(/\/>\s*$/, '>' + add + `</${pre}wsDr>`) : xml.replace(new RegExp('</' + pre + 'wsDr>\\s*$'), add + `</${pre}wsDr>`);
      zip.file(drawingPath, xml);
    }
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), placed, skipped: placements.length - placed };
  }
  const colIndexAt = (cum, px) => { for (let c = 1; c < cum.length; c++) if (px < cum[c]) return c; return cum.length - 1; };

  /* ---------------------------------------------------------------
     КОМПОНЕНТ: блок «Наложение печати» (в конце каждого блока программы)
     --------------------------------------------------------------- */
  function mount(root, opts) {
    root.classList.add('ov-box');
    root.innerHTML = `
      <div class="ov-head"><b>${opts.title || 'Наложение печати'}</b><span class="ov-cache" data-tip="Печати берутся из памяти браузера. Подготовьте их в Stamp Studio и нажмите «Добавить в файл»: в памяти помещается до четырёх печатей, каждая хранится 30 минут."></span></div>
      <div class="ov-row">
        <select class="ov-file" data-tip="Выберите документ, на который нужно поставить печать. Это могут быть готовые документы из этого блока."></select>
        <label class="btn small ghost ov-up" data-tip="Или загрузите свой готовый документ (.docx, .xlsx или .xls).">${opts.hero ? 'Выбрать файл…' : 'Свой файл…'}<input type="file" accept=".docx,.xlsx,.xls" hidden></label>
        <button class="btn small ov-show" type="button" data-tip="Показать документ на экране, чтобы поставить на него печать мышью.">Показать документ</button>
        <button class="btn small ghost ov-forget" type="button" data-tip="Стереть из памяти браузера все сохранённые печати прямо сейчас.">Забыть все печати</button>
        <span class="ov-thumbs" hidden></span>
      </div>
      <div class="ov-msg hint"></div>
      <div class="ov-stage" hidden>
        <div class="ov-tools">
          <span class="hint">Выберите печать кнопкой-миниатюрой («Печать 1…4») и нажмите на документ — в этом месте встанет печать 40 мм. Печати можно ставить друг на друга: нажмите на уже стоящую печать — новая ляжет сверху. Передвинуть печать — потянуть за значок ✥, убрать — крестик ×. После скачивания использованные печати стираются из памяти.</span>
          <button class="btn small ghost ov-clear" type="button" data-tip="Убрать все поставленные печати с этого документа.">Убрать все печати</button>
        </div>
        <div class="ov-view"></div>
        <div class="ov-actions">
          <button class="btn primary ov-dl" type="button" data-tip="Скачать документ в исходном формате (Word или Excel) с наложенными печатями.">Скачать с печатью</button>
          <button class="btn ov-pdf" type="button" data-tip="Скачать документ как картинку в формате PDF. Печати сохранятся на странице.">Скачать изображение в PDF</button>
        </div>
      </div>`;
    const SC = window.StampCache, zoomK = () => window.UIZoom.k();
    const kOf = el => (el.offsetWidth ? el.getBoundingClientRect().width / el.offsetWidth : 1) || 1;   // масштаб сайта вокруг элемента
    const $q = s => root.querySelector(s);
    const sel = $q('.ov-file'), msg = $q('.ov-msg'), stage = $q('.ov-stage'), view = $q('.ov-view'), cacheEl = $q('.ov-cache'), thumbs = $q('.ov-thumbs');
    let uploaded = null, cur = null, stamps = [], armed = true, ghost = null, current = null, activeId = null, thumbSig = '';
    const say = (t, cls) => { msg.textContent = t || ''; msg.className = 'ov-msg hint' + (cls ? ' ' + cls : ''); };

    function files() { const l = (opts.getFiles ? opts.getFiles() : []).filter(f => /\.(docx|xlsx)$/i.test(f.name)); if (uploaded) l.push(uploaded); return l; }
    function refreshList() {
      const keep = sel.value, l = files(); sel.innerHTML = '';
      if (!l.length) { const o = document.createElement('option'); o.textContent = 'Нет готовых документов — выполните заполнение выше или загрузите свой файл'; o.value = ''; sel.append(o); }
      l.forEach((f, i) => { const o = document.createElement('option'); o.value = i; o.textContent = f.name; sel.append(o); });
      if (keep && l[keep]) sel.value = keep;
    }
    /** Миниатюры сохранённых печатей: «Печать 1…4» — рядом с кнопкой «Забыть все печати», появляются после «Показать документ» */
    function refreshThumbs() {
      const l = SC.list(); if (!l.some(r => r.id === activeId)) activeId = l[0] ? l[0].id : null;
      thumbs.hidden = stage.hidden || !l.length;
      const sig = l.map(r => r.id).join(',') + '|' + activeId; if (sig === thumbSig) return; thumbSig = sig;
      thumbs.innerHTML = '';
      l.forEach(r => {
        const b = document.createElement('button'); b.type = 'button'; b.className = 'ov-thumb' + (r.id === activeId ? ' on' : ''); b.setAttribute('aria-pressed', String(r.id === activeId));
        b.setAttribute('data-tip', 'Печать ' + r.slot + ': выберите её, затем нажмите на документ — печать встанет в это место. После скачивания файла использованная печать стирается из памяти.');
        const im = document.createElement('img'); im.alt = ''; im.src = r.dataUrl; const t = document.createElement('span'); t.textContent = 'Печать ' + r.slot; b.append(im, t);
        b.addEventListener('click', () => { activeId = r.id; thumbSig = ''; refreshThumbs(); ghost = null; say('Выбрана «Печать ' + r.slot + '». Нажмите на документ — она встанет в это место.'); });
        thumbs.append(b);
      });
    }
    function refreshCache() {
      const l = SC.list();
      if (l.length) { const m = Math.max(1, Math.ceil((Math.min.apply(null, l.map(r => r.expires)) - Date.now()) / 60000)); cacheEl.textContent = 'Печатей в памяти: ' + l.length + ' из ' + SC.MAX + ' · ближайшая сотрётся через ≈ ' + m + ' мин'; cacheEl.className = 'ov-cache on'; }
      else { cacheEl.innerHTML = 'Печатей в памяти нет. <a href="stamp.html">Подготовьте их в Stamp Studio</a> и нажмите «Добавить в файл»'; cacheEl.className = 'ov-cache off'; }
      refreshThumbs();
    }
    SC.onChange(refreshCache); setInterval(refreshCache, 20000); refreshCache(); refreshList();
    root._refresh = refreshList;
    $q('.ov-forget').addEventListener('click', () => { SC.clear(); say('Все печати стёрты из памяти.'); });
    async function takeFile(f) {
      if (!f) return;
      if (!/\.(docx|xlsx|xls)$/i.test(f.name)) { say('Подходят файлы Word (.docx) и Excel (.xlsx, .xls).', 'err'); return; }
      try {
        if (/\.xls$/i.test(f.name)) { say('Перевожу .xls в .xlsx…'); const cv = await I.xlsToXlsx(await f.arrayBuffer()); f = { name: f.name.replace(/\.xls$/i, '') + '.xlsx', blob: new Blob([cv.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) }; }
        else f = { name: f.name, blob: f };
      } catch (err) { say('Не удалось прочитать файл: ' + (err.message || err), 'err'); return; }
      uploaded = f; refreshList(); sel.value = String(files().length - 1); say('Файл добавлен: ' + f.name);
      if (opts.hero) $q('.ov-show').click();      // в блоке «Свой файл» документ сразу показываем
    }
    $q('.ov-up input').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; takeFile(f); });
    if (opts.hero) {
      root.classList.add('ov-hero'); sel.hidden = true;
      ['dragenter', 'dragover'].forEach(ev => root.addEventListener(ev, e => { e.preventDefault(); root.classList.add('over'); }));
      ['dragleave', 'drop'].forEach(ev => root.addEventListener(ev, e => { e.preventDefault(); root.classList.remove('over'); }));
      root.addEventListener('drop', e => takeFile(e.dataTransfer.files[0]));
    }
    sel.addEventListener('mousedown', refreshList);

    const hostOf = page => (page._geom ? page._geom.holder : page);
    function addStamp(page, x, y) {
      const c = SC.get(activeId); if (!c) { say('Печатей в памяти нет (или прошло 30 минут). Подготовьте их заново в Stamp Studio.', 'err'); return; }
      const el = document.createElement('div'); el.className = 'stamp-mark'; el.style.width = el.style.height = STAMP_PX + 'px';
      el.innerHTML = '<img alt="печать" draggable="false"><span class="mv" title="Передвинуть печать (потяните)">✥</span><button type="button" class="x" title="Убрать печать">×</button>'; el.firstChild.src = c.dataUrl;
      hostOf(page).appendChild(el); const st = { el, page, x, y, data: c.dataUrl, sid: c.id, slot: c.slot }; place(st); stamps.push(st);
      el.querySelector('.x').addEventListener('click', ev => { ev.stopPropagation(); remove(st); });
      // Сама печать «прозрачна» для мыши, поэтому щелчок по ней ставит новую печать поверх (печати могут перекрывать друг друга).
      // Двигают печать за значок ✥ в левом верхнем углу, убирают крестиком в правом верхнем.
      el.querySelector('.mv').addEventListener('pointerdown', ev => {
        ev.preventDefault(); ev.stopPropagation();
        const sx = ev.clientX, sy = ev.clientY, ox = st.x, oy = st.y, k = kOf(hostOf(page));
        const mv = e2 => { st.x = ox + (e2.clientX - sx) / k; st.y = oy + (e2.clientY - sy) / k; place(st); };
        const up = () => { document.removeEventListener('pointermove', mv); document.removeEventListener('pointerup', up); document.removeEventListener('pointercancel', up); };
        document.addEventListener('pointermove', mv); document.addEventListener('pointerup', up); document.addEventListener('pointercancel', up);
      });
      el.querySelector('.mv').addEventListener('click', ev => ev.stopPropagation());
      say('Печатей на документе: ' + stamps.length + '.');
    }
    const place = st => { st.el.style.left = st.x + 'px'; st.el.style.top = st.y + 'px'; };
    function remove(st) { st.el.remove(); stamps = stamps.filter(s => s !== st); say('Печатей на документе: ' + stamps.length + '.'); }
    $q('.ov-clear').addEventListener('click', () => { stamps.slice().forEach(remove); });

    view.addEventListener('mousemove', e => {
      const c = SC.get(activeId), page = e.target.closest && e.target.closest('.ov-page'); if (!armed || !page || !c) { if (ghost) ghost.style.display = 'none'; return; }
      if (!ghost) { ghost = document.createElement('img'); ghost.className = 'stamp-ghost'; ghost.draggable = false; }
      if (ghost.dataset.sid !== c.id) { ghost.src = c.dataUrl; ghost.dataset.sid = c.id; }
      ghost.style.width = ghost.style.height = STAMP_PX + 'px'; const host = hostOf(page); if (ghost.parentNode !== host) host.appendChild(ghost);
      const r = host.getBoundingClientRect(), k = kOf(host); ghost.style.display = 'block'; ghost.style.left = ((e.clientX - r.left) / k - STAMP_PX / 2) + 'px'; ghost.style.top = ((e.clientY - r.top) / k - STAMP_PX / 2) + 'px';
    });
    view.addEventListener('mouseleave', () => { if (ghost) ghost.style.display = 'none'; });
    view.addEventListener('click', e => {
      const page = e.target.closest && e.target.closest('.ov-page'); if (!page || e.target.closest('.mv, .x') || !armed) return;
      const host = hostOf(page), r = host.getBoundingClientRect(), k = kOf(host); addStamp(page, (e.clientX - r.left) / k - STAMP_PX / 2, (e.clientY - r.top) / k - STAMP_PX / 2);
    });

    $q('.ov-show').addEventListener('click', async () => {
      const l = files(), f = l[+sel.value]; if (!f) { say('Сначала выберите документ.', 'err'); return; }
      say('Готовлю предпросмотр…'); stamps = []; ghost = null;
      try { cur = await renderPreview(f.blob, f.name, view); current = f; stage.hidden = false; thumbSig = ''; refreshThumbs(); const n = SC.count(); say(n ? 'Документ открыт. Выберите печать (миниатюры справа от «Забыть все печати») и нажмите на нужное место.' : 'Документ открыт, но печатей в памяти нет — подготовьте их в Stamp Studio.', n ? '' : 'err'); }
      catch (err) { say('Не удалось показать документ: ' + (err.message || err), 'err'); }
    });

    const baseName = n => n.replace(/\.[^.]+$/, '');
    function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); }
    /** Печати, которые попали в скачанный файл, считаются использованными — стираем их из памяти */
    async function spendStamps() {
      const ids = Array.from(new Set(stamps.map(s => s.sid))); if (!ids.length) return '';
      await SC.removeMany(ids); return ' Использованные печати (' + ids.length + ') удалены из памяти.';
    }
    $q('.ov-dl').addEventListener('click', async () => {
      if (!current || !cur) return; if (!stamps.length) { say('Поставьте хотя бы одну печать — нажмите на документ.', 'err'); return; }
      try {
        const buf = await current.blob.arrayBuffer(), images = {}; stamps.forEach(s => { if (!images[s.sid]) images[s.sid] = dataUrlToBytes(s.data); });
        let res;
        if (cur.kind === 'docx') {
          const zip = await I.JSZip.loadAsync(buf), doc = parseXml(await zip.file('word/document.xml').async('string')), paras = [];
          (function walk(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; const nm = I.lname(c); if (nm === 'p') paras.push({ p: c }); else if (['tbl', 'tr', 'tc', 'sdt', 'sdtContent'].includes(nm)) walk(c); } })(byNS(doc, NS_W, 'body')[0]);
          const pl = stamps.map(s => ({ key: s.sid, anchor: docxAnchorFor(s, s.page, paras) }));
          res = await stampDocx(buf, pl, images);
        } else {
          const pl = stamps.map(s => {
            const g = s.page._geom, px = s.x, py = s.y;
            const col = colIndexAt(g.cumC, Math.max(0, px)), row = colIndexAt(g.cumR, Math.max(0, py));
            return { key: s.sid, sheetFile: g.file, col: col - 1, row: row - 1, colOff: Math.max(0, Math.round((px - g.cumC[col - 1]) * EMU_PX)), rowOff: Math.max(0, Math.round((py - g.cumR[row - 1]) * EMU_PX)) };
          });
          res = await stampXlsx(buf, pl, images);
        }
        const ext = extOf(current.name);
        download(new Blob([res.data], { type: ext === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), baseName(current.name) + ' (с печатью).' + ext);
        const spent = res.placed ? await spendStamps() : '';
        say('Готово: печатей вставлено — ' + res.placed + (res.skipped ? ', не удалось привязать — ' + res.skipped : '') + '.' + spent, res.skipped ? 'err' : 'ok');
      } catch (err) { say('Не удалось вставить печать: ' + (err.message || err), 'err'); }
    });
    $q('.ov-pdf').addEventListener('click', async () => {
      if (!current || !cur) return; say('Собираю PDF…');
      root.classList.add('capturing'); if (ghost) ghost.style.display = 'none';
      try { const bytes = await pagesToPdf(cur.pages, cur.kind); download(new Blob([bytes], { type: 'application/pdf' }), baseName(current.name) + ' (изображение).pdf'); say('PDF скачан: каждая страница — картинка.' + await spendStamps(), 'ok'); }
      catch (err) { say('Не удалось собрать PDF: ' + (err.message || err), 'err'); }
      finally { root.classList.remove('capturing'); }
    });
  }

  window.Overlay = { mount, renderPreview, pagesToPdf, blobToPdf, stampDocx, stampXlsx, STAMP_PX };
})();
