/* =====================================================================
   Ассистент — чтение старого формата Excel (.xls, BIFF8) и перевод в .xlsx
   Значения, объединённые ячейки, ширины и высоты, шрифты, заливки, границы,
   выравнивание и числовые форматы сохраняются. Формулы заменяются своими
   последними значениями, картинки и макросы не переносятся.
   Контейнер OLE читает библиотека cfb; если файл нестандартный (BIFF5, HTML-таблица,
   XML 2003), подключается запасной вариант — SheetJS (только значения).
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('jszip'), require('cfb'));
  else root.AssistantXls = factory(root.JSZip, root.CFB);
})(typeof self !== 'undefined' ? self : this, function (JSZip, CFB) {
  'use strict';

  const PALETTE = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];
  const BORDER = ['', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
  const HALIGN = ['', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
  const VALIGN = ['top', 'center', 'bottom', 'justify', 'distributed'];
  const PATTERN = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis', 'gray125', 'gray0625'];

  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  const colName = n => { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const fnum = v => { const s = String(v); return s; };

  class Reader {
    constructor(u8) { this.u = u8; this.dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength); }
  }

  /* строка из записи: cch + флаг + символы; len16 — длина 2 байта (иначе 1 байт). Возвращает { s, next } */
  function readXlStr(u, dv, off, len16) {
    let cch, p = off;
    if (len16) { cch = dv.getUint16(p, true); p += 2; } else { cch = u[p]; p += 1; }
    const flags = u[p]; p += 1;
    let runs = 0, ext = 0;
    if (flags & 8) { runs = dv.getUint16(p, true); p += 2; }
    if (flags & 4) { ext = dv.getUint32(p, true); p += 4; }
    let s = '';
    if (flags & 1) { for (let i = 0; i < cch; i++) s += String.fromCharCode(dv.getUint16(p + i * 2, true)); p += cch * 2; }
    else { for (let i = 0; i < cch; i++) s += String.fromCharCode(u[p + i]); p += cch; }
    p += runs * 4 + ext;
    return { s, next: p };
  }

  /* таблица общих строк с учётом записей CONTINUE: на границе записи у продолжения строки есть свой байт флага */
  function readSst(chunks) {
    const first = chunks[0];
    const dv0 = new DataView(first.buffer, first.byteOffset, first.byteLength);
    const uniq = dv0.getUint32(4, true);
    let ci = 0, u = chunks[0], dv = dv0, p = 8;
    const out = [];
    const nextChunk = () => { ci++; u = chunks[ci]; dv = new DataView(u.buffer, u.byteOffset, u.byteLength); p = 0; };
    const need = () => { while (p >= u.length && ci < chunks.length - 1) nextChunk(); };
    for (let k = 0; k < uniq; k++) {
      need(); if (p >= u.length) break;
      if (p + 3 > u.length) { // заголовок строки не помещается — крайне редкий случай
        if (ci >= chunks.length - 1) break; nextChunk();
      }
      const cch = dv.getUint16(p, true), flags = u[p + 2]; p += 3;
      let runs = 0, ext = 0;
      if (flags & 8) { runs = dv.getUint16(p, true); p += 2; }
      if (flags & 4) { ext = dv.getUint32(p, true); p += 4; }
      let hi = !!(flags & 1), s = '', left = cch;
      while (left > 0) {
        need();
        const bytes = hi ? 2 : 1, avail = Math.floor((u.length - p) / bytes);
        const take = Math.min(left, avail);
        if (hi) for (let i = 0; i < take; i++) s += String.fromCharCode(dv.getUint16(p + i * 2, true));
        else for (let i = 0; i < take; i++) s += String.fromCharCode(u[p + i]);
        p += take * bytes; left -= take;
        if (left > 0) { if (ci >= chunks.length - 1) break; nextChunk(); hi = !!(u[p] & 1); p += 1; }   // байт флага в начале продолжения
      }
      let skip = runs * 4 + ext;
      while (skip > 0) { need(); const t = Math.min(skip, u.length - p); p += t; skip -= t; if (skip > 0 && ci >= chunks.length - 1) break; }
      out.push(s);
    }
    return out;
  }

  function rkToNum(rk) {
    let v;
    if (rk & 2) v = rk >> 2;                                   // целое со знаком
    else { const b = new DataView(new ArrayBuffer(8)); b.setUint32(4, (rk & 0xFFFFFFFC) >>> 0, true); v = b.getFloat64(0, true); }
    if (rk & 1) v /= 100;
    return v;
  }

  function parseWorkbook(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    const recs = [];                                          // склеиваем CONTINUE с предыдущей записью
    for (let p = 0; p + 4 <= u8.length;) {
      const id = dv.getUint16(p, true), len = dv.getUint16(p + 2, true);
      if (p + 4 + len > u8.length) break;
      const data = u8.subarray(p + 4, p + 4 + len);
      if (id === 0x003C && recs.length) recs[recs.length - 1].parts.push(data);
      else recs.push({ id, off: p, parts: [data] });
      p += 4 + len;
    }
    const wb = { fonts: [], xfs: [], formats: {}, palette: null, sst: [], sheets: [], date1904: false, bofs: [] };
    let i = 0, level = 0, cur = null;
    const sheetAt = {};
    // глобальная часть
    const sheetsByOffset = {};
    for (; i < recs.length; i++) {
      const r = recs[i], d = r.parts[0], rdv = new DataView(d.buffer, d.byteOffset, d.byteLength);
      if (r.id === 0x0809) {
        const ver = rdv.getUint16(0, true), dt = rdv.getUint16(2, true);
        if (i === 0) { if (ver !== 0x0600) throw new Error('BIFF' + (ver === 0x0500 ? '5' : '?')); }
        if (dt === 0x0010 || dt === 0x0020 || dt === 0x0040) { break; }
      } else if (r.id === 0x002F) throw new Error('Файл защищён паролем');
      else if (r.id === 0x0022) wb.date1904 = rdv.getUint16(0, true) === 1;
      else if (r.id === 0x0031) {
        const h = rdv.getUint16(0, true), gr = rdv.getUint16(2, true), ic = rdv.getUint16(4, true), bls = rdv.getUint16(6, true), uls = d[10], cch = d[14], fl = d[15];
        let name = ''; if (fl & 1) for (let k = 0; k < cch; k++) name += String.fromCharCode(rdv.getUint16(16 + k * 2, true)); else for (let k = 0; k < cch; k++) name += String.fromCharCode(d[16 + k]);
        wb.fonts.push({ h, i: !!(gr & 2), strike: !!(gr & 8), color: ic, bold: bls >= 700, u: uls, name });
        if (wb.fonts.length === 4) wb.fonts.push(null);       // шрифта №4 в BIFF нет
      } else if (r.id === 0x041E) { const id = rdv.getUint16(0, true); wb.formats[id] = readXlStr(d, rdv, 2, true).s; }
      else if (r.id === 0x00E0) {
        const al = d[6], b0 = rdv.getUint16(10, true), c0 = rdv.getUint16(12, true), c1 = rdv.getUint32(14, true), c2 = rdv.getUint16(18, true);
        wb.xfs.push({
          font: rdv.getUint16(0, true), fmt: rdv.getUint16(2, true), style: !!(rdv.getUint16(4, true) & 4),
          h: al & 7, wrap: !!(al & 8), v: (al >> 4) & 7, rot: d[7], indent: d[8] & 15, shrink: !!(d[8] & 16),
          bl: b0 & 15, br: (b0 >> 4) & 15, bt: (b0 >> 8) & 15, bb: (b0 >> 12) & 15,
          cl: c0 & 127, cr: (c0 >> 7) & 127, ct: c1 & 127, cb: (c1 >> 7) & 127, pat: (c1 >>> 26) & 63,
          fg: c2 & 127, bg: (c2 >> 7) & 127,
        });
      } else if (r.id === 0x0092) { const n = rdv.getUint16(0, true); wb.palette = []; for (let k = 0; k < n; k++) wb.palette.push(((d[2 + k * 4] << 16) | (d[3 + k * 4] << 8) | d[4 + k * 4]).toString(16).padStart(6, '0').toUpperCase()); }
      else if (r.id === 0x0085) { const off = rdv.getUint32(0, true), st = d[4], dt = d[5], nm = readXlStr(d, rdv, 6, false).s; wb.sheets.push({ off, hidden: st !== 0, type: dt, name: nm, cells: {}, rows: {}, cols: [], merges: [], defH: null, defW: null }); }
      else if (r.id === 0x00FC) wb.sst = readSst(r.parts);
    }
    // листы
    for (const sh of wb.sheets) sheetsByOffset[sh.off] = sh;
    for (; i < recs.length; i++) {
      const r = recs[i], d = r.parts[0], rdv = new DataView(d.buffer, d.byteOffset, d.byteLength);
      if (r.id === 0x0809) {
        const dt = rdv.getUint16(2, true);
        if (level === 0) cur = sheetsByOffset[r.off] || null;
        level++;
        if (cur && level === 1 && !(dt === 0x0010)) cur = null;
        continue;
      }
      if (r.id === 0x000A) { level = Math.max(0, level - 1); if (level === 0) cur = null; continue; }
      if (!cur || level !== 1) continue;
      const S = cur;
      const put = (row, col, cell) => { if (row > 1048575 || col > 16383) return; (S.cells[row] = S.cells[row] || {})[col] = cell; };
      switch (r.id) {
        case 0x0208: { const rw = rdv.getUint16(0, true), hh = rdv.getUint16(6, true), gr = rdv.getUint16(12, true), ix = rdv.getUint16(14, true);
          S.rows[rw] = { ht: (hh & 0x7FFF) / 20, custom: !!(gr & 0x40), hidden: !!(gr & 0x20), xf: (ix & 0x1000) || (gr & 0x80) ? (ix & 0xFFF) : null }; break; }
        case 0x007D: { const a = rdv.getUint16(0, true), b = rdv.getUint16(2, true), w = rdv.getUint16(4, true), xf = rdv.getUint16(6, true), gr = rdv.getUint16(8, true); S.cols.push({ a, b, w: w / 256, xf, hidden: !!(gr & 1) }); break; }
        case 0x0055: S.defW = rdv.getUint16(0, true); break;
        case 0x0225: S.defH = rdv.getUint16(2, true) / 20; break;
        case 0x00E5: { const n = rdv.getUint16(0, true); for (let k = 0; k < n; k++) { const o = 2 + k * 8; S.merges.push([rdv.getUint16(o, true), rdv.getUint16(o + 2, true), rdv.getUint16(o + 4, true), rdv.getUint16(o + 6, true)]); } break; }
        case 0x00FD: put(rdv.getUint16(0, true), rdv.getUint16(2, true), { xf: rdv.getUint16(4, true), t: 's', v: wb.sst[rdv.getUint32(6, true)] || '' }); break;
        case 0x0204: put(rdv.getUint16(0, true), rdv.getUint16(2, true), { xf: rdv.getUint16(4, true), t: 's', v: readXlStr(d, rdv, 6, true).s }); break;
        case 0x0203: put(rdv.getUint16(0, true), rdv.getUint16(2, true), { xf: rdv.getUint16(4, true), t: 'n', v: rdv.getFloat64(6, true) }); break;
        case 0x027E: put(rdv.getUint16(0, true), rdv.getUint16(2, true), { xf: rdv.getUint16(4, true), t: 'n', v: rkToNum(rdv.getUint32(6, true)) }); break;
        case 0x00BD: { const rw = rdv.getUint16(0, true), c0 = rdv.getUint16(2, true), n = (d.length - 6) / 6; for (let k = 0; k < n; k++) put(rw, c0 + k, { xf: rdv.getUint16(4 + k * 6, true), t: 'n', v: rkToNum(rdv.getUint32(6 + k * 6, true)) }); break; }
        case 0x0201: put(rdv.getUint16(0, true), rdv.getUint16(2, true), { xf: rdv.getUint16(4, true), t: 'b0' }); break;
        case 0x00BE: { const rw = rdv.getUint16(0, true), c0 = rdv.getUint16(2, true), n = (d.length - 6) / 2; for (let k = 0; k < n; k++) put(rw, c0 + k, { xf: rdv.getUint16(4 + k * 2, true), t: 'b0' }); break; }
        case 0x0205: { const v = d[6], isErr = d[7]; put(rdv.getUint16(0, true), rdv.getUint16(2, true), { xf: rdv.getUint16(4, true), t: isErr ? 'e' : 'b', v: isErr ? ({ 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A' }[v] || '#N/A') : v }); break; }
        case 0x0006: {
          const rw = rdv.getUint16(0, true), col = rdv.getUint16(2, true), xf = rdv.getUint16(4, true);
          if (d[12] === 0xFF && d[13] === 0xFF) {
            const k = d[6];
            if (k === 0) { cur._pendStr = { rw, col, xf }; put(rw, col, { xf, t: 's', v: '' }); }
            else if (k === 1) put(rw, col, { xf, t: 'b', v: d[8] });
            else if (k === 2) put(rw, col, { xf, t: 'e', v: '#N/A' });
            else put(rw, col, { xf, t: 's', v: '' });
          } else put(rw, col, { xf, t: 'n', v: rdv.getFloat64(6, true) });
          break;
        }
        case 0x0207: { if (cur._pendStr) { const q = cur._pendStr; cur._pendStr = null; put(q.rw, q.col, { xf: q.xf, t: 's', v: readXlStr(d, rdv, 0, true).s }); } break; }
        case 0x0081: S.fitPage = !!(rdv.getUint16(0, true) & 0x100); break;
        case 0x00A1: { const gr = rdv.getUint16(10, true); S.page = { paper: rdv.getUint16(0, true), scale: rdv.getUint16(2, true), fitW: rdv.getUint16(6, true), fitH: rdv.getUint16(8, true), portrait: !!(gr & 2), noPls: !!(gr & 4) }; break; }
        case 0x0026: S.mL = rdv.getFloat64(0, true); break;
        case 0x0027: S.mR = rdv.getFloat64(0, true); break;
        case 0x0028: S.mT = rdv.getFloat64(0, true); break;
        case 0x0029: S.mB = rdv.getFloat64(0, true); break;
        default: break;
      }
    }
    return wb;
  }

  /* ---------- сборка xlsx ---------- */
  function buildStyles(wb) {
    const pal = i => { if (i >= 8 && i < 8 + 56) return (wb.palette && wb.palette[i - 8]) || PALETTE[i]; if (i < 8) return PALETTE[i]; if (i === 64 || i === 0x7FFF) return null; if (i === 65) return 'FFFFFF'; return null; };
    const colorXml = (tag, idx) => { const c = pal(idx); return c ? '<' + tag + ' rgb="FF' + c + '"/>' : '<' + tag + ' auto="1"/>'; };
    // шрифты
    const fonts = wb.fonts.map(f => f || wb.fonts[0] || { h: 200, name: 'Arial', color: 64 });
    if (!fonts.length) fonts.push({ h: 200, name: 'Arial', color: 64, bold: false, i: false, u: 0 });
    const fontXml = fonts.map(f => '<font>' + (f.bold ? '<b/>' : '') + (f.i ? '<i/>' : '') + (f.strike ? '<strike/>' : '') + (f.u ? '<u' + (f.u === 2 || f.u === 0x22 ? ' val="double"' : '') + '/>' : '') + '<sz val="' + (f.h / 20) + '"/>' + (f.color === 0x7FFF || f.color === 64 ? '' : colorXml('color', f.color)) + '<name val="' + esc(f.name || 'Arial') + '"/></font>').join('');
    // заливки и границы: уникальные
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillIdx = {};
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], borderIdx = {};
    const fillOf = x => {
      if (!x.pat) return 0;
      const p = PATTERN[x.pat] || 'solid';
      const key = p + ':' + x.fg + ':' + x.bg; if (fillIdx[key] != null) return fillIdx[key];
      const fgc = pal(x.fg), bgc = pal(x.bg);
      let xml;
      if (p === 'solid') xml = '<fill><patternFill patternType="solid"><fgColor rgb="FF' + (fgc || '000000') + '"/><bgColor indexed="64"/></patternFill></fill>';
      else xml = '<fill><patternFill patternType="' + p + '">' + (fgc ? '<fgColor rgb="FF' + fgc + '"/>' : '') + (bgc ? '<bgColor rgb="FF' + bgc + '"/>' : '') + '</patternFill></fill>';
      fills.push(xml); return (fillIdx[key] = fills.length - 1);
    };
    const borderOf = x => {
      const key = [x.bl, x.br, x.bt, x.bb, x.cl, x.cr, x.ct, x.cb].join(',');
      if (borderIdx[key] != null) return borderIdx[key];
      const side = (tag, st, c) => st && BORDER[st] ? '<' + tag + ' style="' + BORDER[st] + '">' + colorXml('color', c) + '</' + tag + '>' : '<' + tag + '/>';
      borders.push('<border>' + side('left', x.bl, x.cl) + side('right', x.br, x.cr) + side('top', x.bt, x.ct) + side('bottom', x.bb, x.cb) + '<diagonal/></border>');
      return (borderIdx[key] = borders.length - 1);
    };
    const xfs = wb.xfs.length ? wb.xfs : [{ font: 0, fmt: 0, h: 0, wrap: false, v: 2, rot: 0, indent: 0, shrink: false, bl: 0, br: 0, bt: 0, bb: 0, cl: 0, cr: 0, ct: 0, cb: 0, pat: 0, fg: 64, bg: 65 }];
    const xfXml = xfs.map(x => {
      const fid = fillOf(x), bid = borderOf(x), fo = x.font < fonts.length ? x.font : 0;
      let al = '';
      const attrs = [];
      if (x.h && HALIGN[x.h]) attrs.push('horizontal="' + HALIGN[x.h] + '"');
      if (x.v !== 2) attrs.push('vertical="' + (VALIGN[x.v] || 'bottom') + '"');
      if (x.wrap) attrs.push('wrapText="1"');
      if (x.rot) attrs.push('textRotation="' + (x.rot === 255 ? 255 : x.rot > 90 ? 90 - x.rot : x.rot) + '"');
      if (x.indent) attrs.push('indent="' + x.indent + '"');
      if (x.shrink) attrs.push('shrinkToFit="1"');
      if (attrs.length) al = '<alignment ' + attrs.join(' ') + '/>';
      return '<xf numFmtId="' + x.fmt + '" fontId="' + fo + '" fillId="' + fid + '" borderId="' + bid + '" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"' + (al ? '>' + al + '</xf>' : '/>');
    }).join('');
    const custom = Object.keys(wb.formats).map(id => '<numFmt numFmtId="' + id + '" formatCode="' + esc(wb.formats[id]) + '"/>').join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (custom ? '<numFmts count="' + Object.keys(wb.formats).length + '">' + custom + '</numFmts>' : '') +
      '<fonts count="' + fonts.length + '">' + fontXml + '</fonts><fills count="' + fills.length + '">' + fills.join('') + '</fills><borders count="' + borders.length + '">' + borders.join('') + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="' + xfs.length + '">' + xfXml + '</cellXfs>' +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  }

  function sheetXml(wb, S, sstIndex) {
    const rowsN = Object.keys(S.cells).map(Number), rowKeys = new Set(rowsN.concat(Object.keys(S.rows).map(Number)));
    let maxR = -1, maxC = -1;
    for (const r of rowsN) { maxR = Math.max(maxR, r); for (const c of Object.keys(S.cells[r])) maxC = Math.max(maxC, +c); }
    for (const m of S.merges) { maxR = Math.max(maxR, m[1]); maxC = Math.max(maxC, m[3]); }
    const dim = maxR < 0 ? 'A1' : 'A1:' + colName(maxC) + (maxR + 1);
    const defW = S.defW != null ? S.defW : 8;
    let x = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
    if (S.fitPage) x += '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>';
    x += '<dimension ref="' + dim + '"/><sheetViews><sheetView workbookViewId="0"/></sheetViews>';
    x += '<sheetFormatPr defaultRowHeight="' + (S.defH || 12.75) + '" defaultColWidth="' + (defW + 0.71).toFixed(2) + '"/>';
    const cols = S.cols.filter(c => c.a <= 16383).sort((a, b) => a.a - b.a);
    if (cols.length) x += '<cols>' + cols.map(c => '<col min="' + (c.a + 1) + '" max="' + (Math.min(c.b, 16383) + 1) + '" width="' + (c.w > 0 ? c.w.toFixed(2) : '0') + '" customWidth="1"' + (c.hidden ? ' hidden="1"' : '') + (c.xf && wb.xfs[c.xf] ? ' style="' + c.xf + '"' : '') + '/>').join('') + '</cols>';
    x += '<sheetData>';
    for (const r of Array.from(rowKeys).sort((a, b) => a - b)) {
      const ri = S.rows[r] || {}, cells = S.cells[r] || {};
      let attrs = 'r="' + (r + 1) + '"';
      if (ri.ht != null && ri.ht > 0) attrs += ' ht="' + ri.ht + '"' + (ri.custom ? ' customHeight="1"' : '');
      if (ri.hidden) attrs += ' hidden="1"';
      let cx = '';
      for (const c of Object.keys(cells).map(Number).sort((a, b) => a - b)) {
        const cell = cells[c], ref = colName(c) + (r + 1), s = cell.xf < wb.xfs.length ? ' s="' + cell.xf + '"' : '';
        if (cell.t === 's') { if (cell.v === '') cx += '<c r="' + ref + '"' + s + '/>'; else cx += '<c r="' + ref + '"' + s + ' t="s"><v>' + sstIndex(cell.v) + '</v></c>'; }
        else if (cell.t === 'n') cx += '<c r="' + ref + '"' + s + '><v>' + fnum(cell.v) + '</v></c>';
        else if (cell.t === 'b') cx += '<c r="' + ref + '"' + s + ' t="b"><v>' + (cell.v ? 1 : 0) + '</v></c>';
        else if (cell.t === 'e') cx += '<c r="' + ref + '"' + s + ' t="e"><v>' + esc(cell.v) + '</v></c>';
        else cx += '<c r="' + ref + '"' + s + '/>';
      }
      x += '<row ' + attrs + '>' + cx + '</row>';
    }
    x += '</sheetData>';
    if (S.merges.length) x += '<mergeCells count="' + S.merges.length + '">' + S.merges.filter(m => m[0] !== m[1] || m[2] !== m[3]).map(m => '<mergeCell ref="' + colName(m[2]) + (m[0] + 1) + ':' + colName(m[3]) + (m[1] + 1) + '"/>').join('') + '</mergeCells>';
    const ml = S.mL != null ? S.mL : 0.7, mr = S.mR != null ? S.mR : 0.7, mt = S.mT != null ? S.mT : 0.75, mb = S.mB != null ? S.mB : 0.75;
    x += '<pageMargins left="' + ml + '" right="' + mr + '" top="' + mt + '" bottom="' + mb + '" header="0.3" footer="0.3"/>';
    if (S.page) {
      const p = S.page;
      x += '<pageSetup' + (p.paper ? ' paperSize="' + p.paper + '"' : '') + (!S.fitPage && p.scale && p.scale !== 100 ? ' scale="' + p.scale + '"' : '') + (S.fitPage ? ' fitToWidth="' + p.fitW + '" fitToHeight="' + p.fitH + '"' : '') + ' orientation="' + (p.portrait ? 'portrait' : 'landscape') + '"/>';
    }
    x += '</worksheet>';
    // fitToWidth/Height = 0 означает «сколько страниц нужно»
    return x;
  }

  function wbToXlsx(wb) {
    const sheets = wb.sheets.filter(s => s.type === 0);
    if (!sheets.length) throw new Error('В файле нет листов с данными');
    const sstList = [], sstMap = new Map();
    const sstIndex = v => { let k = sstMap.get(v); if (k == null) { k = sstList.length; sstList.push(v); sstMap.set(v, k); } return k; };
    const zip = new JSZip();
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map((_, i) => '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>').join('') +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + (wb.date1904 ? '<workbookPr date1904="1"/>' : '') + '<bookViews><workbookView/></bookViews><sheets>' +
      sheets.map((s, i) => '<sheet name="' + esc(s.name.replace(/[\\/?*[\]:]/g, '_').slice(0, 31) || 'Лист' + (i + 1)) + '" sheetId="' + (i + 1) + '"' + (s.hidden ? ' state="hidden"' : '') + ' r:id="rId' + (i + 1) + '"/>').join('') + '</sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>').join('') +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId' + (sheets.length + 2) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');
    zip.file('xl/styles.xml', buildStyles(wb));
    sheets.forEach((s, i) => zip.file('xl/worksheets/sheet' + (i + 1) + '.xml', sheetXml(wb, s, sstIndex)));
    zip.file('xl/sharedStrings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sstList.length + '" uniqueCount="' + sstList.length + '">' + sstList.map(t => '<si><t xml:space="preserve">' + esc(t) + '</t></si>').join('') + '</sst>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  /* ---------- запасной путь: SheetJS (только значения) ---------- */
  let sheetjsPromise = null;
  function loadSheetJs() {
    if (typeof XLSX !== 'undefined' && root().XLSX) return Promise.resolve(root().XLSX);
    if (root().XLSX) return Promise.resolve(root().XLSX);
    if (!sheetjsPromise) sheetjsPromise = new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = (window.XLS_VENDOR_BASE || 'assets/vendor/') + 'xlsx.full.min.js';
      s.onload = () => res(window.XLSX); s.onerror = () => { sheetjsPromise = null; rej(new Error('Не удалось загрузить модуль чтения .xls')); };
      document.head.appendChild(s);
    });
    return sheetjsPromise;
  }
  const root = () => (typeof self !== 'undefined' ? self : globalThis);
  async function viaSheetJs(u8) {
    const X = await loadSheetJs();
    const wb = X.read(u8, { type: 'array', cellStyles: true, cellDates: false });
    const out = X.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true });
    return new Uint8Array(out);
  }

  /** .xls → .xlsx. Возвращает { data: Uint8Array, mode: 'full' | 'values' } */
  async function toXlsx(buf) {
    const u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    if (u8.length > 3 && u8[0] === 0x50 && u8[1] === 0x4B) return { data: u8, mode: 'xlsx' };       // уже xlsx
    const isOle = u8.length > 8 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0;
    if (isOle && CFB) {
      try {
        const cf = CFB.read(u8, { type: 'array' });
        const ent = CFB.find(cf, '/Workbook') || CFB.find(cf, '/Book');
        if (!ent || !ent.content) throw new Error('нет потока Workbook');
        const wb = parseWorkbook(ent.content instanceof Uint8Array ? ent.content : new Uint8Array(ent.content));
        return { data: await wbToXlsx(wb), mode: 'full' };
      } catch (e) {
        if (/парол/.test(e.message)) throw e;
        if (typeof document === 'undefined') throw e;
      }
    }
    if (typeof document === 'undefined') throw new Error('Файл .xls не удалось прочитать');
    try { return { data: await viaSheetJs(u8), mode: 'values' }; }
    catch (e) { throw new Error('Не удалось прочитать файл .xls: ' + (e.message || e)); }
  }

  return { toXlsx, parseWorkbook, wbToXlsx };
});
