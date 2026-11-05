/* Ассистент — интерфейс */
(function () {
  'use strict';
  const E = window.AssistantEngine;
  const $ = id => document.getElementById(id);

  const CARD_FIELDS = [
    ['name_full', 'Полное наименование', 'wide'],
    ['name_short', 'Сокращённое наименование', 'wide'],
    ['inn', 'ИНН'], ['kpp', 'КПП'], ['ogrn', 'ОГРН / ОГРНИП'], ['okpo', 'ОКПО'],
    ['addr_legal', 'Юридический адрес', 'wide'],
    ['addr_post', 'Почтовый адрес', 'wide'],
    ['addr_fact', 'Фактический адрес', 'wide'],
    ['phone', 'Телефон'], ['fax', 'Факс'], ['email', 'E-mail'], ['site', 'Сайт'],
    ['bank_name', 'Банк', 'wide'], ['bik', 'БИК'], ['rs', 'Расчётный счёт'], ['ks', 'Корр. счёт'],
    ['director_pos', 'Должность подписанта'], ['director_fio', 'ФИО подписанта'],
    ['director_pos_gen', 'Должность в родительном падеже («в лице …»)'], ['director_fio_gen', 'ФИО в родительном падеже («в лице …»)'],
    ['basis', 'Действует на основании'], ['tax', 'Налогообложение / НДС'],
  ];

  let card = {};
  let forms = [];     // {id, file, mode}
  let results = [];   // {name, blob, stats, items, error}
  let seq = 0;

  /* ---------- карта партнёра ---------- */
  function saveCard() { try { localStorage.setItem('assistant.card', JSON.stringify(card)); } catch (e) {} }
  function loadCard() { try { const s = localStorage.getItem('assistant.card'); if (s) card = JSON.parse(s) || {}; } catch (e) {} }

  function renderCardForm() {
    const grid = $('cardGrid');
    grid.innerHTML = '';
    for (const [key, label, cls] of CARD_FIELDS) {
      const l = document.createElement('label');
      if (cls) l.className = cls;
      l.append(label);
      const inp = document.createElement('input');
      inp.type = 'text'; inp.value = card[key] || '';
      inp.addEventListener('input', () => { card[key] = inp.value.trim(); if (key === 'director_pos' && !card.director_pos_gen_manual) { /* подсказка не перезаписывает ручной ввод */ } saveCard(); });
      l.append(inp);
      grid.append(l);
    }
    $('cardForm').hidden = false;
    updateRunButton();
  }

  function setStatus(el, text, cls) { el.textContent = text; el.className = 'status' + (cls ? ' ' + cls : ''); }

  async function pdfToText(buf) {
    const pdfjs = window.pdfjsLib;
    if (!pdfjs) throw new Error('Не загрузилась библиотека PDF');
    pdfjs.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    const pdf = await pdfjs.getDocument({ data: buf }).promise;
    const lines = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent();
      const rows = [];
      for (const it of content.items) {
        const y = Math.round(it.transform[5] / 3), x = it.transform[4];
        let row = rows.find(r => r.y === y);
        if (!row) rows.push(row = { y, parts: [] });
        row.parts.push({ x, w: it.width, s: it.str });
      }
      rows.sort((a, b) => b.y - a.y).forEach(r => {
        r.parts.sort((a, b) => a.x - b.x);
        let out = '', prevEnd = null;
        for (const p of r.parts) {
          if (!p.s.trim()) continue;
          out += (prevEnd !== null ? (p.x - prevEnd > 12 ? ' | ' : ' ') : '') + p.s;
          prevEnd = p.x + p.w;
        }
        if (out.trim()) lines.push(out.trim());
      });
    }
    return lines.join('\n');
  }

  async function readCardFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    const buf = await file.arrayBuffer();
    if (ext === 'docx') return E.docxToText(buf);
    if (ext === 'xlsx') return E.xlsxToText(buf);
    if (ext === 'pdf') return pdfToText(buf);
    if (ext === 'txt') return new TextDecoder('utf-8').decode(buf);
    throw new Error('Формат .' + ext + ' не поддерживается. Сохраните карту как .docx, .xlsx, .pdf или .txt');
  }

  function applyParsed(text, sourceName) {
    const parsed = E.parseCardText(text);
    const n = Object.keys(parsed).length;
    if (!n) { setStatus($('cardStatus'), 'Не удалось найти реквизиты. Проверьте файл или вставьте текст вручную.', 'err'); return; }
    card = parsed; saveCard(); renderCardForm();
    setStatus($('cardStatus'), `${sourceName}: распознано полей — ${n}. Проверьте их ниже.`, 'ok');
  }

  async function handleCardFile(file) {
    if (!file) return;
    setStatus($('cardStatus'), 'Читаю ' + file.name + '…');
    try { applyParsed(await readCardFile(file), file.name); }
    catch (e) { setStatus($('cardStatus'), 'Ошибка: ' + e.message, 'err'); }
  }

  /* ---------- бланки ---------- */
  function renderForms() {
    const ul = $('formList');
    ul.innerHTML = '';
    for (const f of forms) {
      const li = document.createElement('li');
      const nm = document.createElement('span'); nm.className = 'name'; nm.textContent = f.file.name;
      const sel = document.createElement('select');
      [['default', 'Режим: как в настройках'], ['full', 'Заполнить всё'], ['min', 'Только наименование и подписант']].forEach(([v, t]) => {
        const o = document.createElement('option'); o.value = v; o.textContent = t; if (f.mode === v) o.selected = true; sel.append(o);
      });
      sel.addEventListener('change', () => { f.mode = sel.value; });
      const x = document.createElement('button'); x.className = 'x'; x.type = 'button'; x.title = 'Убрать'; x.textContent = '×';
      x.addEventListener('click', () => { forms = forms.filter(q => q.id !== f.id); renderForms(); });
      li.append(nm, sel, x);
      ul.append(li);
    }
    updateRunButton();
  }
  function addForms(files) {
    for (const file of files) forms.push({ id: ++seq, file, mode: 'default' });
    renderForms();
  }
  function updateRunButton() { $('btnRun').disabled = !(forms.length && Object.keys(card).length); }

  /* ---------- запуск ---------- */
  function uniqueName(name, used) {
    const dot = name.lastIndexOf('.');
    const base = dot > 0 ? name.slice(0, dot) : name, ext = dot > 0 ? name.slice(dot) : '';
    let cand = base + ' (заполнено)' + ext, i = 2;
    while (used.has(cand)) cand = base + ' (заполнено ' + (i++) + ')' + ext;
    used.add(cand);
    return cand;
  }

  const MIME = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  };

  async function run() {
    const btn = $('btnRun');
    btn.disabled = true; btn.textContent = 'Заполняю…';
    results = [];
    const used = new Set();
    const g = {
      highlight: $('optHighlight').checked,
      fallbackAddr: $('optFallback').checked,
      mode: $('optMode').value,
      docDate: $('optDate').value,
    };
    for (const f of forms) {
      const name = uniqueName(f.file.name, used);
      try {
        const buf = await f.file.arrayBuffer();
        const out = await E.fillFile(f.file.name, buf, card, Object.assign({}, g, { mode: f.mode === 'default' ? g.mode : f.mode }));
        const ext = f.file.name.split('.').pop().toLowerCase();
        results.push({ name, blob: new Blob([out.data], { type: MIME[ext] || 'application/octet-stream' }), stats: out.stats, items: out.items });
      } catch (e) {
        // файл всё равно возвращаем — в исходном виде, чтобы количество файлов совпало
        results.push({ name: f.file.name, blob: f.file, error: e.message || String(e) });
      }
    }
    renderResults();
    btn.disabled = false; btn.textContent = 'Заполнить';
  }

  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  function renderResults() {
    const ul = $('resultList');
    ul.innerHTML = '';
    for (const r of results) {
      const li = document.createElement('li'); li.className = 'res';
      const line = document.createElement('div'); line.className = 'line';
      const nm = document.createElement('span'); nm.className = 'name'; nm.textContent = r.name; nm.style.flex = '1 1 220px';
      line.append(nm);
      if (r.error) {
        const b = document.createElement('span'); b.className = 'badge err'; b.textContent = 'не обработан: ' + r.error; line.append(b);
      } else {
        const a = document.createElement('span'); a.className = 'badge ok'; a.textContent = 'заполнено: ' + r.stats.filled; line.append(a);
        if (r.stats.highlighted) { const b = document.createElement('span'); b.className = 'badge hl'; b.textContent = 'выделено жёлтым: ' + r.stats.highlighted; line.append(b); }
      }
      const dl = document.createElement('button'); dl.className = 'btn small'; dl.type = 'button';
      dl.textContent = r.error ? 'Скачать без изменений' : 'Скачать';
      dl.addEventListener('click', () => download(r.blob, r.name));
      line.append(dl);
      li.append(line);
      if (r.items && r.items.length) {
        const d = document.createElement('details');
        const s = document.createElement('summary'); s.textContent = 'Что именно сделано'; d.append(s);
        const u = document.createElement('ul');
        for (const it of r.items) {
          const l = document.createElement('li'); l.style.cssText = 'border:0;padding:0;display:list-item';
          if (it.kind === 'fill') l.textContent = `${it.key}: ${it.val}`;
          else if (it.kind === 'hl') { const m = document.createElement('mark'); m.textContent = it.key; l.append('Не заполнено — ', m); }
          else l.textContent = 'Ошибка: ' + it.key;
          u.append(l);
        }
        d.append(u); li.append(d);
      }
      ul.append(li);
    }
    $('resultCard').hidden = false;
    $('resultCard').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  async function downloadZip() {
    const zip = new JSZip();
    for (const r of results) zip.file(r.name, r.blob);
    download(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), 'Заполненные_документы.zip');
  }

  /* ---------- drag & drop ---------- */
  function bindDrop(zone, input, handler) {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } });
    ['dragenter', 'dragover'].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => zone.addEventListener(ev, e => { e.preventDefault(); zone.classList.remove('over'); }));
    zone.addEventListener('drop', e => handler(Array.from(e.dataTransfer.files)));
    input.addEventListener('change', () => { handler(Array.from(input.files)); input.value = ''; });
  }

  /* ---------- старт ---------- */
  bindDrop($('dropCard'), $('fileCard'), files => handleCardFile(files[0]));
  bindDrop($('dropForms'), $('fileForms'), addForms);
  $('btnParsePaste').addEventListener('click', () => applyParsed($('pasteCard').value, 'Вставленный текст'));
  $('btnClearCard').addEventListener('click', () => { card = {}; saveCard(); $('cardForm').hidden = true; setStatus($('cardStatus'), ''); updateRunButton(); });
  $('btnRun').addEventListener('click', run);
  $('btnZip').addEventListener('click', downloadZip);

  loadCard();
  if (Object.keys(card).length) { renderCardForm(); setStatus($('cardStatus'), 'Загружена карта, сохранённая в этом браузере.', 'ok'); }
})();
