/* Ассистент — автозаполнение: интерфейс (5 блоков) */
(function () {
  'use strict';
  const E = window.AssistantEngine, TECH = window.AssistantTech, LIB = window.AssistantLibrary;
  const $ = (root, sel) => root.querySelector(sel);

  /* ---------- поля ---------- */
  const ORG_GROUPS = [
    ['Организация', [['name_full', 'Полное наименование', 'wide'], ['name_short', 'Сокращённое наименование', 'wide'], ['inn', 'ИНН'], ['kpp', 'КПП'], ['ogrn', 'ОГРН / ОГРНИП'], ['okpo', 'ОКПО'], ['reg_date', 'Дата регистрации'], ['tax_reg_date', 'Дата постановки на налоговый учёт']]],
    ['Адреса и связь', [['addr_legal', 'Юридический адрес', 'wide'], ['addr_post', 'Почтовый адрес', 'wide'], ['addr_fact', 'Фактический адрес', 'wide'], ['phone', 'Телефон'], ['email', 'E-mail'], ['site', 'Сайт']]],
    ['Банк', [['bank_name', 'Банк', 'wide'], ['bik', 'БИК'], ['rs', 'Расчётный счёт'], ['ks', 'Корр. счёт'], ['bank_inn', 'ИНН банка'], ['bank_addr', 'Адрес банка']]],
    ['Руководитель и подписант', [['director_pos', 'Должность'], ['director_fio', 'ФИО'], ['director_pos_gen', 'Должность, род. падеж («в лице …»)', 0, 'gen'], ['director_fio_gen', 'ФИО, род. падеж («в лице …»)', 0, 'gen'], ['basis', 'Действует на основании', 'wide'], ['director_inn', 'ИНН руководителя'], ['accountant_fio', 'Главный бухгалтер']]],
    ['Контактное лицо и налогообложение', [['contact_fio', 'Контактное лицо (ФИО)'], ['contact_pos', 'Должность контактного лица'], ['contact_phone', 'Телефон контактного лица'], ['contact_email', 'E-mail контактного лица'], ['tax_system', 'Система налогообложения', 'sel'], ['branches', 'Филиалы'], ['holding', 'Холдинг / группа компаний', 'wide']]],
  ];
  const PERSON_GROUPS = [
    ['Работник', [['name_full', 'ФИО полностью', 'wide'], ['director_pos', 'Должность по договору'], ['department', 'Структурное подразделение'], ['salary', 'Оклад, руб.'], ['start_date', 'Дата начала работы'], ['end_date', 'Дата окончания (срочный договор)'], ['term_reason', 'Основание срочного договора', 'wide'], ['probation', 'Испытательный срок'], ['tab_no', 'Табельный номер']]],
    ['Личные данные', [['birth_date', 'Дата рождения'], ['birth_place', 'Место рождения'], ['citizenship', 'Гражданство'], ['passport', 'Паспорт (серия и номер)'], ['passport_date', 'Дата выдачи паспорта'], ['passport_issued', 'Кем выдан паспорт', 'wide'], ['passport_code', 'Код подразделения'], ['snils', 'СНИЛС'], ['inn', 'ИНН']]],
    ['Адрес и связь', [['addr_legal', 'Адрес регистрации', 'wide'], ['addr_fact', 'Адрес проживания', 'wide'], ['phone', 'Телефон'], ['email', 'E-mail']]],
  ];
  const TAX = ['', 'ОСН (с НДС)', 'УСН (без НДС)', 'УСН (с НДС)', 'ПСН', 'НПД'];
  const DEMO_ORG = {
    a: `Полное наименование: Общество с ограниченной ответственностью «Ромашка»\nИНН: 7700000001\nКПП: 770001001\nОГРН: 1027700000001\nЮридический адрес: 123112, г. Москва, Тестовая ул., д. 1\nБанк: АО «Тестбанк»\nБИК: 044525001\nРасчётный счёт: 40702810900000000001\nКорр. счёт: 30101810400000000001\nТелефон: +7 (495) 000-00-01\nE-mail: info@romashka.test\nГенеральный директор: Петров Пётр Петрович`,
    b: `Полное наименование: Общество с ограниченной ответственностью «Василёк»\nИНН: 7800000002\nКПП: 780001001\nОГРН: 1027800000002\nОКПО: 12345678\nЮридический адрес: 190000, г. Санкт-Петербург, Образцовая ул., д. 2\nБанк: АО «Образецбанк»\nБИК: 044030002\nРасчётный счёт: 40702810100000000002\nКорр. счёт: 30101810500000000002\nТелефон: +7 (812) 000-00-02\nE-mail: office@vasilek.test\nГенеральный директор: Сидорова Анастасия Юрьевна`,
  };
  const DEMO_PERSON = `ФИО: Иванов Иван Иванович\nДолжность: Менеджер по продажам\nДата рождения: 01.02.1990\nПаспорт: 4510 123456\nПаспорт выдан: ОУФМС России по г. Москве\nДата выдачи: 05.03.2015\nКод подразделения: 770-001\nСНИЛС: 123-456-789 01\nИНН: 770100000012\nАдрес регистрации: г. Москва, ул. Лесная, д. 5, кв. 7\nТелефон: +7 900 111-22-33`;

  const MIME = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
  const setStatus = (el, t, cls) => { el.textContent = t; el.classList.remove('ok', 'err'); el.classList.add('status'); if (cls) el.classList.add(cls); };
  const tip = t => ' data-tip="' + String(t).replace(/"/g, '&quot;') + '"';
  const extOf = n => (n.split('.').pop() || '').toLowerCase();
  function uniqueName(name, used, suffix) { const dot = name.lastIndexOf('.'), base = dot > 0 ? name.slice(0, dot) : name, ext = dot > 0 ? name.slice(dot) : ''; let cand = base + ' (' + suffix + ')' + ext, i = 2; while (used.has(cand)) cand = base + ' (' + suffix + ' ' + (i++) + ')' + ext; used.add(cand); return cand; }
  function download(blob, name) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); }

  /* ---------- чтение файла с реквизитами ---------- */
  async function pdfToText(buf) {
    const pdfjs = window.pdfjsLib; if (!pdfjs) throw new Error('Не загрузилась библиотека PDF (нужен интернет)');
    pdfjs.GlobalWorkerOptions.workerSrc = 'assets/vendor/pdfjs/pdf.worker.min.js';
    const pdf = await pdfjs.getDocument({ data: buf }).promise, lines = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent(), rows = [];
      for (const it of content.items) { const y = Math.round(it.transform[5] / 3), x = it.transform[4]; let row = rows.find(r => r.y === y); if (!row) rows.push(row = { y, parts: [] }); row.parts.push({ x, w: it.width, s: it.str }); }
      rows.sort((a, b) => b.y - a.y).forEach(r => {
        r.parts.sort((a, b) => a.x - b.x); let out = '', prevEnd = null;
        for (const p of r.parts) { if (!p.s.trim()) continue; out += (prevEnd !== null ? (p.x - prevEnd > 12 ? ' | ' : ' ') : '') + p.s; prevEnd = p.x + p.w; }
        if (out.trim()) lines.push(out.trim());
      });
    }
    return lines.join('\n');
  }
  async function readCardFile(file) {
    const ext = extOf(file.name), buf = await file.arrayBuffer();
    if (ext === 'docx') return E.docxToText(buf); if (ext === 'xlsx') return E.xlsxToText(buf); if (ext === 'xls') return E.xlsxToText((await E.xlsToXlsx(buf)).data); if (ext === 'pdf') return pdfToText(buf);
    if (ext === 'txt') return new TextDecoder('utf-8').decode(buf);
    throw new Error('Формат .' + ext + ' не поддерживается. Сохраните как .docx, .xlsx, .xls, .pdf или .txt');
  }

  /* ---------- общий блок: чекбокс жёлтого, результаты, печать ---------- */
  function hlBar(id) {
    return `<label class="hl-bar"${tip('Если галочка стоит, всё, что программа не смогла заполнить, будет выделено жёлтым цветом — так легко увидеть, что дописать вручную.')}><input type="checkbox" id="${id}" checked> <span>Выделять жёлтым цветом не заполненные поля</span></label>`;
  }
  function makeResults(root, block) {
    const list = $(root, '.res-list'), card = $(root, '.res-card'), memEl = $(root, '.res-mem'), state = { items: [] };
    const DC = window.DocCache, blk = root.closest('.blk'), pill = document.createElement('span'); pill.className = 'blk-mem'; pill.hidden = true;
    if (blk && blk.querySelector('.blk-t')) blk.querySelector('.blk-t').append(pill);
    /** Подпись о том, что документы лежат в памяти браузера, и сколько им осталось */
    function memLabel() {
      const n = state.items.filter(r => r.cid).length, left = DC.remaining(block.key), m = Math.max(1, Math.ceil(left / 60000));
      pill.hidden = !n; pill.textContent = 'в памяти: ' + n + ' · ≈ ' + m + ' мин';
      memEl.textContent = n ? 'Готовые документы сохранены в памяти браузера ещё ≈ ' + m + ' мин: можно перейти в другую программу и вернуться — они останутся.' : '';
    }
    function render(quiet) {
      list.innerHTML = '';
      for (const r of state.items) {
        const li = document.createElement('li'); li.className = 'res';
        const line = document.createElement('div'); line.className = 'line';
        const nm = document.createElement('span'); nm.className = 'name'; nm.textContent = r.name; nm.style.flex = '1 1 220px'; line.append(nm);
        const pill = (t, c) => { const s = document.createElement('span'); s.className = 'pill ' + c; s.textContent = t; line.append(s); };
        if (r.error) pill('не обработан: ' + r.error, 'err');
        else { pill(r.stats.label || 'заполнено', 'ok'); pill(String(r.stats.filled), 'ok'); if (r.stats.highlighted) pill('жёлтым выделено: ' + r.stats.highlighted, 'hl'); }
        const dl = document.createElement('button'); dl.className = 'btn small'; dl.type = 'button'; dl.textContent = r.error ? 'Скачать без изменений' : 'Скачать';
        dl.setAttribute('data-tip', 'Скачать документ в исходном формате (Word или Excel).'); dl.addEventListener('click', () => download(r.blob, r.name)); line.append(dl);
        if (!r.error && /\.(docx|xlsx)$/i.test(r.name)) {
          const pd = document.createElement('button'); pd.className = 'btn small'; pd.type = 'button'; pd.textContent = 'Скачать изображение в PDF';
          pd.setAttribute('data-tip', 'Документ превратится в картинку и будет сохранён как PDF — удобно отправить или распечатать. Страницы — только изображения, текст в PDF не выделяется.');
          pd.addEventListener('click', async () => { pd.disabled = true; const old = pd.textContent; pd.textContent = 'Готовлю PDF…'; try { download(new Blob([await window.Overlay.blobToPdf(r.blob, r.name)], { type: 'application/pdf' }), r.name.replace(/\.[^.]+$/, '') + ' (изображение).pdf'); } catch (e) { alert('Не удалось собрать PDF: ' + (e.message || e)); } pd.disabled = false; pd.textContent = old; });
          line.append(pd);
        }
        li.append(line);
        if (r.items && r.items.length) {
          const d = document.createElement('details'), s = document.createElement('summary'); s.textContent = 'Что именно сделано'; d.append(s);
          const u = document.createElement('ul');
          for (const it of r.items) {
            const l = document.createElement('li');
            if (it.role) { const t = document.createElement('span'); t.className = 'tag ' + (it.role === 'customer' ? 'c' : 'e'); t.textContent = (block.roleNames || {})[it.role] || (it.role === 'customer' ? 'Заказчик' : 'Участник'); l.append(t); }
            if (it.kind === 'fill') l.append(it.key + ': ' + String(it.val).slice(0, 90));
            else if (it.kind === 'hl') { const m = document.createElement('mark'); m.textContent = it.key; l.append('Заполните вручную — ', m); }
            else if (it.kind === 'rep') l.append(it.name + ': «' + it.from + '» → «' + it.to + '» · ' + it.rule);
            else l.append('Ошибка: ' + it.key);
            u.append(l);
          }
          d.append(u); li.append(d);
        }
        list.append(li);
      }
      card.hidden = !state.items.length;
      if (state.items.length && !quiet) card.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      if (root._ovRefresh) root._ovRefresh();
      memLabel();
    }
    $(root, '.zip-btn').addEventListener('click', async () => { const zip = new JSZip(); for (const r of state.items) zip.file(r.name, r.blob); download(await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' }), 'Готовые_документы.zip'); });
    const ovHost = $(root, '.ov-host');
    window.Overlay.mount(ovHost, { getFiles: () => state.items.filter(r => !r.error).map(r => ({ name: r.name, blob: r.blob })) });
    root._ovRefresh = ovHost._refresh;
    // документы живут в памяти 30 минут: возвращаем их после перехода между программами; по истечении срока убираем
    DC.onChange(() => {
      const alive = new Set(DC.list(block.key).map(x => x.cid));
      if (state.items.some(r => r.cid && !alive.has(r.cid))) { state.items = state.items.filter(r => !r.cid || alive.has(r.cid)); render(true); }
      else if (!state.items.length && alive.size) { state.items = DC.list(block.key); render(true); }
      else memLabel();
    });
    DC.ready.then(() => { if (!state.items.length) { const l = DC.list(block.key); if (l.length) { state.items = l; render(true); } } memLabel(); });
    setInterval(memLabel, 20000);
    return { set(items) { state.items = items; render(); DC.save(block.key, items).then(ok => { if (!ok) memEl.textContent = 'Не удалось сохранить документы в память браузера (мало места) — скачайте их сразу.'; }); }, state };
  }
  const resultsHtml = zipTip => `
    <section class="res-card" hidden><h3 class="sec">Результат</h3><div class="hint res-mem"></div><ul class="files res-list"></ul>
      <div class="row"><button class="btn zip-btn" type="button"${tip(zipTip || 'Скачать все готовые документы одним архивом ZIP.')}>Скачать всё одним архивом (ZIP)</button></div></section>
    <section class="ov-host"></section>`;

  /* ---------- окно реквизитов стороны ---------- */
  function buildParty(host, cfg, store) {
    const role = cfg.role, u = { inputs: {} }, manual = new Set();
    host.innerHTML = `<h2><span class="n">${cfg.n}</span> ${cfg.title}<span class="badge-role">${cfg.badge}</span></h2><div class="sub">${cfg.sub}</div>
      <div class="lib-row"><select class="lib-sel"${tip('Список ваших сохранённых данных. Выберите запись — поля заполнятся сразу. Списки хранятся только в этом браузере; все они собраны в «Моей библиотеке» вверху страницы.')}></select><button class="btn small" type="button" data-act="libsave"${tip('Запомнить то, что сейчас введено, в список «' + cfg.lib[0].title + '». Запись с тем же названием обновится.')}>Сохранить в «${cfg.lib[0].title}»</button><button class="btn small ghost" type="button" data-act="libdel"${tip('Убрать выбранную запись из списка. Введённые поля не изменятся.')} disabled>Удалить из списка</button></div>
      <div class="drop" tabindex="0"${tip('Перетащите сюда файл с реквизитами или нажмите и выберите его на компьютере: Word, Excel, PDF с текстом или TXT.')}><b>Перетащите карту сюда</b> или нажмите, чтобы выбрать<small>.docx, .xlsx, .xls, .pdf (с текстом), .txt</small><input type="file" accept=".docx,.xlsx,.xls,.pdf,.txt" hidden></div>
      <details class="paste"><summary${tip('Если файла нет, скопируйте реквизиты и вставьте их сюда как обычный текст.')}>Или вставить текст реквизитов</summary><textarea rows="6" placeholder="${cfg.person ? 'ФИО: Иванов Иван Иванович&#10;Паспорт: 4510 123456&#10;Дата рождения: 01.02.1990' : 'Наименование: ООО «Ромашка»&#10;ИНН: 7701234567&#10;КПП: 770101001&#10;Генеральный директор: Иванов Иван Иванович'}"></textarea><div class="row"><button class="btn small" type="button" data-act="parse"${tip('Программа найдёт в тексте реквизиты и разложит их по полям ниже.')}>Распознать</button></div></details>
      <div class="row"><button class="btn small ghost" type="button" data-act="demo"${tip('Заполнить поля вымышленными данными — чтобы потренироваться.')}>Подставить вымышленный пример</button><button class="btn small ghost" type="button" data-act="clear"${tip('Стереть все поля этого окна.')}>Очистить</button></div>
      ${cfg.labor ? `<details class="take"><summary${tip('Если у вас уже есть другой трудовой договор, программа найдёт в нём ' + (cfg.person ? 'данные работника' : 'данные работодателя') + ' и перенесёт в поля ниже.')}>Взять ${cfg.person ? 'данные работника' : 'данные работодателя'} из другого трудового договора</summary><div class="take-body">Загрузите готовый договор (.docx) — программа прочитает вводную часть и реквизиты сторон. Вы сможете проверить и поправить данные в полях ниже.<div class="row"><label class="btn small take-up">Выбрать договор…<input type="file" accept=".docx" hidden></label></div><div class="take-list"></div></div></details>` : ''}
      <div class="status"></div>
      <details class="req-fields"><summary${tip('Поля с реквизитами свёрнуты, чтобы не мешать. Раскройте их, чтобы проверить или поправить данные вручную.')}><span class="t">Поля реквизитов</span><span class="cnt"></span></summary><div class="grid"></div></details>`;
    const drop = $(host, '.drop'), file = $(host, 'input[type=file]'), ta = $(host, 'textarea'); u.status = $(host, '.status'); u.grid = $(host, '.grid');
    const card = () => store.cards[role];
    drop.addEventListener('click', () => file.click());
    drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); file.click(); } });
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    /* --- мои списки --- */
    const sel = $(host, '.lib-sel'), bSave = $(host, '[data-act=libsave]'), bDel = $(host, '[data-act=libdel]');
    function fillSel(keep) {
      sel.innerHTML = ''; const o0 = document.createElement('option'); o0.value = ''; o0.textContent = '— Выбрать из списка: ' + cfg.lib.map(l => l.title).join(' / ') + ' —'; sel.append(o0);
      for (const l of cfg.lib) { const arr = LIB.list(l.key); const g = document.createElement('optgroup'); g.label = l.title + ' (' + arr.length + ')'; if (!arr.length) { const o = document.createElement('option'); o.disabled = true; o.textContent = 'пока пусто'; g.append(o); }
        for (const r of arr) { const o = document.createElement('option'); o.value = l.key + '|' + r.id; o.textContent = LIB.nameOf(r.card) + (r.card.inn ? ' · ИНН ' + r.card.inn : ''); g.append(o); } sel.append(g); }
      if (keep && [...sel.options].some(o => o.value === keep)) sel.value = keep; bDel.disabled = !sel.value;
    }
    fillSel(); LIB.onChange(() => fillSel(sel.value));
    sel.addEventListener('change', () => {
      bDel.disabled = !sel.value; if (!sel.value) return; const [k, id] = sel.value.split('|'), c = LIB.get(k, id); if (!c) return;
      store.cards[role] = c; manual.clear(); store.save(); render(); $(host, 'details.req-fields').open = false;
      setStatus(u.status, 'Загружено из списка: ' + LIB.nameOf(c) + '. Проверьте данные ниже.', 'ok');
    });
    bSave.addEventListener('click', () => {
      const r = LIB.add(cfg.lib[0].key, card());
      setStatus(u.status, r.ok ? (r.updated ? 'Запись в списке «' + cfg.lib[0].title + '» обновлена.' : 'Сохранено в список «' + cfg.lib[0].title + '».') : r.why === 'empty' ? 'Сначала введите хотя бы наименование или ФИО.' : 'Не удалось сохранить: в браузере не хватает места.', r.ok ? 'ok' : 'err');
      if (r.ok) fillSel(cfg.lib[0].key + '|' + r.id);
    });
    bDel.addEventListener('click', () => { const [k, id] = sel.value.split('|'); if (!id) return; if (!confirm('Убрать эту запись из списка?')) return; LIB.remove(k, id); sel.value = ''; setStatus(u.status, 'Запись убрана из списка.'); });
    function apply(text, source) { applyCard(E.parseCardText(text), source); }
    function applyCard(parsed, source) {
      const n = Object.keys(parsed).length;
      if (!n) { setStatus(u.status, 'Не удалось найти реквизиты. Проверьте файл или вставьте текст вручную.', 'err'); return; }
      if (cfg.person) { parsed.name_full = parsed.name_full || parsed.director_fio || ''; parsed.director_fio = parsed.director_fio || parsed.name_full; parsed.name_short = E.shortFio(parsed.name_full); ['basis', 'director_pos_gen', 'director_fio_gen'].forEach(k => delete parsed[k]); }
      const keep = {}; for (const k of ['tax_system', 'contact_fio', 'contact_pos', 'contact_phone', 'contact_email']) if (card()[k]) keep[k] = card()[k];
      store.cards[role] = Object.assign(parsed, keep); manual.clear(); store.save(); render(); if (n) $(host, 'details.req-fields').open = false;
      const need = cfg.person ? ['name_full', 'passport', 'birth_date'] : ['name_full', 'inn', 'addr_legal', 'director_fio'];
      const miss = need.filter(k => !card()[k]).map(k => E.FIELD_TITLES[k]);
      setStatus(u.status, source + ': распознано полей — ' + n + '.' + (miss.length ? ' Не найдено: ' + miss.join(', ') + '.' : '') + ' Проверьте данные ниже.', miss.length ? '' : 'ok');
    }
    async function handle(f) { if (!f) return; setStatus(u.status, 'Читаю ' + f.name + '…'); try { apply(await readCardFile(f), f.name); } catch (e) { setStatus(u.status, 'Ошибка: ' + e.message, 'err'); } }
    drop.addEventListener('drop', e => handle(e.dataTransfer.files[0]));
    file.addEventListener('change', () => { handle(file.files[0]); file.value = ''; });
    $(host, '[data-act=parse]').addEventListener('click', () => apply(ta.value, 'Вставленный текст'));
    $(host, '[data-act=demo]').addEventListener('click', () => apply(cfg.demo, 'Вымышленный пример'));
    $(host, '[data-act=clear]').addEventListener('click', () => { store.cards[role] = {}; manual.clear(); store.save(); render(); setStatus(u.status, ''); });
    function render() {
      u.grid.innerHTML = ''; u.inputs = {};
      for (const [gname, fields] of cfg.groups) {
        const g = document.createElement('div'); g.className = 'grp'; g.textContent = gname; u.grid.append(g);
        for (const [key, label, cls, kind] of fields) {
          const l = document.createElement('label'); if (cls === 'wide') l.className = 'wide'; l.append(label); let inp;
          if (cls === 'sel') { inp = document.createElement('select'); TAX.forEach(t => { const o = document.createElement('option'); o.value = t; o.textContent = t || '— не указано —'; inp.append(o); }); inp.value = card()[key] || ''; }
          else { inp = document.createElement('input'); inp.type = 'text'; inp.value = card()[key] || ''; if (kind === 'gen') inp.classList.add('auto'); }
          inp.addEventListener('input', () => edit(key, inp)); u.inputs[key] = inp; l.append(inp); u.grid.append(l);
        }
      }
      count(); store.changed();
    }
    function count() {
      const all = cfg.groups.reduce((a, g) => a + g[1].length, 0), got = cfg.groups.reduce((a, g) => a + g[1].filter(f => card()[f[0]]).length, 0);
      const nm = card().name_full || card().name_short || '', cnt = $(host, 'details.req-fields .cnt');
      cnt.textContent = (nm ? nm.slice(0, 48) + (nm.length > 48 ? '…' : '') + ' · ' : '') + 'заполнено ' + got + ' из ' + all; cnt.classList.toggle('ok', got > 0);
    }
    function edit(key, inp) {
      const c = card(); c[key] = inp.value.trim();
      if (cfg.person && key === 'name_full') { c.director_fio = c[key]; c.name_short = E.shortFio(c[key]); }
      if (key === 'director_pos_gen' || key === 'director_fio_gen') manual.add(key);
      if (!cfg.person) {
        if (key === 'director_fio' && !manual.has('director_fio_gen')) { c.director_fio_gen = E.fioGen(c.director_fio); u.inputs.director_fio_gen.value = c.director_fio_gen; }
        if (key === 'director_pos' && !manual.has('director_pos_gen')) { c.director_pos_gen = E.posGen(c.director_pos); u.inputs.director_pos_gen.value = c.director_pos_gen; }
      }
      store.save(); count(); store.changed();
    }
    /** ИП: после ввода («change», а не на каждую букву) подтягиваем полное наименование и руководителя по правилам */
    function ipNormalize() {
      if (cfg.person) return; const c = card(), before = JSON.stringify(c), n = E.applyIpRules(Object.assign({}, c));
      if (JSON.stringify(n) === before) return;
      for (const k of Object.keys(n)) if (n[k] !== c[k]) { c[k] = n[k]; if (u.inputs[k]) u.inputs[k].value = n[k]; }
      if (c.kpp === undefined && u.inputs.kpp) u.inputs.kpp.value = '';
      store.save(); count(); store.changed();
    }
    u.grid.addEventListener('change', e => { if (e.target && /^(name_full|name_short|ogrn|director_fio)$/.test(Object.keys(u.inputs).find(k => u.inputs[k] === e.target) || '')) ipNormalize(); });
    /* --- данные из другого трудового договора --- */
    if (cfg.labor) {
      const take = $(host, 'details.take'), tl = $(take, '.take-list'), tup = $(take, '.take-up input');
      async function fromBuf(buf, name) {
        setStatus(u.status, 'Читаю ' + name + '…');
        try {
          const r = await window.AssistantLabor.extractParties(buf), c = r[cfg.person ? 'worker' : 'employer'];
          if (!Object.keys(c).length) { setStatus(u.status, 'В договоре «' + name + '» не нашлось данных ' + (cfg.person ? 'работника' : 'работодателя') + '. Проверьте, что это трудовой договор с вводной частью и реквизитами.', 'err'); return; }
          applyCard(c, 'Из договора «' + name + '»'); take.open = false;
        } catch (e) { setStatus(u.status, 'Ошибка: ' + (e.message || e), 'err'); }
      }
      tup.addEventListener('change', async () => { const f = tup.files[0]; tup.value = ''; if (f) fromBuf(await f.arrayBuffer(), f.name); });
      take.addEventListener('toggle', () => {
        tl.innerHTML = ''; if (!take.open) return;
        const cands = (store.files || []).filter(f => /\.docx$/i.test(f.file.name) && !f.builtin).map(f => ({ name: f.file.name, get: () => f.file.arrayBuffer() }));
        for (const r of (store.results ? store.results() : [])) if (/\.docx$/i.test(r.name)) cands.push({ name: r.name + ' (готовый)', get: () => r.blob.arrayBuffer() });
        cands.forEach(cd => { const b = document.createElement('button'); b.type = 'button'; b.className = 'btn small ghost'; b.textContent = 'Из бланка в списке: ' + cd.name; b.addEventListener('click', async () => fromBuf(await cd.get(), cd.name)); tl.append(b); });
      });
    }
    render();
  }

  /* ---------- блок «заполнение по реквизитам» (1, 2 и 4) ---------- */
  function fillBlock(el, cfg) {
    const body = $(el, '.blk-body'), key = 'assistant.cards.' + cfg.key;
    const store = { cards: { customer: {}, executor: {} }, files: [], seq: 0, results: null };
    try { const s = JSON.parse(localStorage.getItem(cfg.key === 'forms' ? 'assistant.cards.v2' : key) || 'null'); if (s) { store.cards.customer = s.customer || {}; store.cards.executor = s.executor || {}; } } catch (e) {}
    store.save = () => { try { localStorage.setItem(cfg.key === 'forms' ? 'assistant.cards.v2' : key, JSON.stringify(store.cards)); if (cfg.docInfo) localStorage.setItem('assistant.docinfo.' + cfg.key, JSON.stringify(cfg.docInfo.read(body))); } catch (e) {} };
    store.changed = () => updateRun();
    body.innerHTML = `
      ${hlBar('hl-' + cfg.key)}
      <div class="guide"><div><b>1. Реквизиты</b>${cfg.guide[0]}</div><div><b>2. Бланки</b>${cfg.guide[1]}</div><div><b>3. Проверка</b>Скачайте результат и просмотрите места, выделенные жёлтым: правила эвристические, итог проверяйте глазами.</div></div>
      <div class="cols"><section class="card party customer"></section><section class="card party executor"></section></div>
      <section class="card" style="margin-top:20px"><h2><span class="n">3</span> ${cfg.filesTitle}</h2><div class="sub">${cfg.filesSub}</div>
        <div class="drop files-drop" tabindex="0"${tip(cfg.filesTip)}><b>Перетащите файлы сюда</b> или нажмите, чтобы выбрать<small>${cfg.accept.replace(/\./g, '.')}</small><input type="file" accept="${cfg.accept}" multiple hidden></div>
        ${cfg.templates && cfg.templates.length ? `<div class="tpl-head">Или выберите готовый бланк</div><div class="tpl-sub">${cfg.tplSub || 'Бланки встроены в программу — загружать файл не нужно. «Добавить» кладёт бланк в список ниже, «Скачать пустой» отдаёт бланк без данных.'}</div><div class="tpl-grid"></div>` : ''}
        <div class="tpl-head">Мои бланки</div><div class="tpl-sub">Ваши сохранённые бланки этого блока. Чтобы пополнить список, добавьте файл и нажмите «В мою библиотеку» рядом с ним.</div><div class="tpl-grid my-tpl"></div>
        <ul class="files form-list"></ul></section>
      ${cfg.settingsHtml || ''}
      <div class="row" style="margin-top:16px"><button class="btn primary run-btn" type="button" disabled${tip('Программа впишет реквизиты в загруженные документы. Исходные файлы не меняются — получатся новые копии.')}>${cfg.runLabel}</button><span class="hint run-hint">Добавьте документы и реквизиты.</span></div>
      <div class="res-wrap">${resultsHtml()}</div>`;
    const runBtn = $(body, '.run-btn');
    const cols = body.querySelectorAll('.party'); buildParty(cols[0], Object.assign({ role: 'customer', n: 1, labor: cfg.labor }, cfg.parties.customer), store); buildParty(cols[1], Object.assign({ role: 'executor', n: 2, labor: cfg.labor }, cfg.parties.executor), store);
    const res = makeResults($(body, '.res-wrap'), { key: cfg.key, roleNames: cfg.roleNames });
    store.results = () => res.state.items;
    const drop = $(body, '.files-drop'), fin = $(drop, 'input');
    drop.addEventListener('click', () => fin.click());
    drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fin.click(); } });
    ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', e => add(Array.from(e.dataTransfer.files))); fin.addEventListener('change', () => { add(Array.from(fin.files)); fin.value = ''; });
    function add(files, extra) { for (const file of files) store.files.push(Object.assign({ id: ++store.seq, file, mode: 'default', type: 'default' }, extra || {})); renderFiles(); }
    /* --- встроенные бланки --- */
    const b64bytes = b => { const bin = atob(b), a = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i); return a; };
    const tplGrid = $(body, '.tpl-grid:not(.my-tpl)');
    if (tplGrid) for (const t of cfg.templates) {
      const card = document.createElement('div'); card.className = 'tpl';
      card.innerHTML = '<b></b><span></span><div class="row"><button type="button" class="btn small" data-a="add">Добавить</button><button type="button" class="btn small ghost" data-a="empty">Скачать пустой</button></div>';
      card.querySelector('b').textContent = t.title; card.querySelector('span').textContent = t.desc;
      card.querySelector('[data-a=add]').setAttribute('data-tip', 'Положить этот бланк в список документов на заполнение.');
      card.querySelector('[data-a=empty]').setAttribute('data-tip', 'Скачать этот бланк без данных: на местах реквизитов останутся строки «________».');
      card._t = t; tplGrid.append(card);
      card.querySelector('[data-a=add]').addEventListener('click', () => { if (store.files.some(f => f.tplId === t.id)) return; add([new File([b64bytes(t.b64)], t.name, { type: MIME.docx })], { builtin: true, tplId: t.id }); });
      card.querySelector('[data-a=empty]').addEventListener('click', async () => {
        try { const r = await E.fillFile(t.name, b64bytes(t.b64), { customer: {}, executor: {} }, {}, { highlight: false, docType: 'contract' }); download(new Blob([r.data], { type: MIME.docx }), t.name.replace(/\.docx$/, '') + ' (пустой).docx'); } catch (e) { alert('Не удалось подготовить бланк: ' + (e.message || e)); }
      });
    }
    /* --- мои бланки (библиотека) --- */
    const myGrid = $(body, '.my-tpl'), blockKey = cfg.key;
    const bytesToB64 = u8 => { let bin = ''; for (let i = 0; i < u8.length; i += 0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(bin); };
    function renderMy() {
      myGrid.innerHTML = ''; const arr = LIB.templates(blockKey);
      if (!arr.length) { const e = document.createElement('div'); e.className = 'tpl-sub'; e.textContent = 'Пока пусто.'; myGrid.append(e); return; }
      for (const t of arr) {
        const c = document.createElement('div'); c.className = 'tpl'; c.innerHTML = '<b></b><span></span><div class="row"><button type="button" class="btn small" data-a="add">Добавить</button><button type="button" class="btn small ghost" data-a="del">Удалить</button></div>';
        c.querySelector('b').textContent = t.name; c.querySelector('span').textContent = 'Сохранён ' + new Date(t.ts).toLocaleDateString('ru-RU');
        c.querySelector('[data-a=add]').addEventListener('click', () => { if (store.files.some(f => f.libId === t.id)) return; const ext = /\.xlsx$/i.test(t.name) ? 'xlsx' : 'docx'; add([new File([b64bytes(t.b64)], t.name, { type: MIME[ext] })], { libId: t.id }); });
        c.querySelector('[data-a=del]').addEventListener('click', () => { if (confirm('Удалить бланк «' + t.name + '» из библиотеки?')) LIB.removeTemplate(t.id); });
        myGrid.append(c);
      }
    }
    renderMy(); LIB.onChange(renderMy);
    function markTpl() { if (!tplGrid) return; for (const c of tplGrid.children) c.classList.toggle('in', store.files.some(f => f.tplId === c._t.id)); }
    function renderFiles() {
      const ul = $(body, '.form-list'); ul.innerHTML = '';
      for (const f of store.files) {
        const li = document.createElement('li'), nm = document.createElement('span'); nm.className = 'name'; nm.textContent = f.file.name; nm.title = f.file.name; li.append(nm);
        if (f.builtin) { const tg = document.createElement('span'); tg.className = 'tag b'; tg.textContent = 'встроенный бланк'; li.append(tg); }
        else if (f.libId) { const tg = document.createElement('span'); tg.className = 'tag b'; tg.textContent = 'из библиотеки'; li.append(tg); }
        else if (/\.(docx|xlsx)$/i.test(f.file.name)) { const sv = document.createElement('button'); sv.type = 'button'; sv.className = 'btn small ghost'; sv.textContent = 'В мою библиотеку'; sv.setAttribute('data-tip', 'Сохранить этот бланк в браузере: он появится в «Моих бланках» этого блока, и загружать его заново не придётся.');
          sv.addEventListener('click', async () => { const r = LIB.addTemplate(blockKey, f.file.name, bytesToB64(new Uint8Array(await f.file.arrayBuffer()))); sv.textContent = r.ok ? 'Сохранено ✓' : 'Не хватило места'; sv.disabled = r.ok; }); li.append(sv); }
        if (cfg.perFile) {
          const s1 = document.createElement('select'); s1.setAttribute('data-tip', 'Подсказать программе, что это за документ. Обычно хватает «как в настройках».');
          [['default', 'Тип: как в настройках'], ['form', 'Форма участника'], ['contract', 'Договор / спецификация']].forEach(([v, t]) => { const o = document.createElement('option'); o.value = v; o.textContent = t; if (f.type === v) o.selected = true; s1.append(o); }); s1.addEventListener('change', () => { f.type = s1.value; });
          const s2 = document.createElement('select'); s2.setAttribute('data-tip', 'Сколько полей заполнять в этом документе.');
          [['default', 'Объём: как в настройках'], ['full', 'Заполнить всё'], ['min', 'Только наименование и подписант']].forEach(([v, t]) => { const o = document.createElement('option'); o.value = v; o.textContent = t; if (f.mode === v) o.selected = true; s2.append(o); }); s2.addEventListener('change', () => { f.mode = s2.value; });
          li.append(s1, s2);
        }
        const x = document.createElement('button'); x.className = 'x'; x.type = 'button'; x.title = ''; x.setAttribute('data-tip', 'Убрать файл из списка.'); x.textContent = '×'; x.addEventListener('click', () => { store.files = store.files.filter(q => q.id !== f.id); renderFiles(); }); li.append(x); ul.append(li);
      }
      markTpl(); updateRun();
    }
    function updateRun() {
      const hasCard = Object.keys(store.cards.executor).length || Object.keys(store.cards.customer).length;
      runBtn.disabled = !(store.files.length && hasCard);
      $(body, '.run-hint').textContent = !store.files.length ? 'Добавьте документы.' : !hasCard ? 'Заполните реквизиты хотя бы одной стороны.' : 'Готово к заполнению.';
    }
    if (cfg.docInfo) { try { const d = JSON.parse(localStorage.getItem('assistant.docinfo.' + cfg.key) || 'null'); if (d) cfg.docInfo.write(body, d); } catch (e) {} body.querySelectorAll('.settings input, .settings select').forEach(i => i.addEventListener('input', store.save)); }
    runBtn.addEventListener('click', async () => {
      runBtn.disabled = true; runBtn.textContent = 'Заполняю…'; const used = new Set(), out = [];
      const g = cfg.readOpts(body), info = cfg.docInfo ? cfg.docInfo.read(body) : {};
      for (const f of store.files) {
        const isXls = /\.xls$/i.test(f.file.name), name = uniqueName(isXls ? f.file.name.replace(/\.xls$/i, '.xlsx') : f.file.name, used, cfg.suffix);
        try {
          let buf = await f.file.arrayBuffer(); const type = f.type === 'default' ? g.type : f.type; let rep = null;
          // трудовой договор: в уже заполненном бланке сначала заменяем данные прежних сторон на новые
          if (cfg.labor && $(body, '.o-replace').checked && !f.builtin && /\.docx$/i.test(f.file.name) && (await E.docxToText(buf)).indexOf('{{') < 0) {
            const old = await window.AssistantLabor.extractParties(buf);
            if (old.found.employer + old.found.worker > 0) { rep = await window.AssistantLabor.replaceOldData(buf, old, store.cards, { highlight: $(body, '#hl-' + cfg.key).checked }); buf = rep.data; }
          }
          const r = await E.fillFile(f.file.name, buf, store.cards, info, Object.assign({}, g.engine, { highlight: $(body, '#hl-' + cfg.key).checked, mode: f.mode === 'default' ? g.engine.mode : f.mode, docType: type === 'form' ? 'form' : (type || g.engine.docType) }));
          if (r.convertedFrom === 'xls') r.items.unshift({ kind: 'fill', key: 'Файл .xls переведён в .xlsx', val: r.convertMode === 'values' ? 'запасным способом: оформление могло измениться' : 'оформление сохранено, формулы заменены значениями' });
          const items = (rep ? rep.items : []).concat(r.items);
          out.push({ name, blob: new Blob([r.data], { type: MIME[isXls ? 'xlsx' : extOf(f.file.name)] || 'application/octet-stream' }), stats: { filled: r.stats.filled + (rep ? rep.replaced : 0), highlighted: r.stats.highlighted + (rep ? rep.left : 0), label: rep ? 'заполнено и заменено:' : 'заполнено:' }, items });
        } catch (e) { out.push({ name: f.file.name, blob: f.file, error: e.message || String(e) }); }
      }
      res.set(out); runBtn.textContent = cfg.runLabel; updateRun();
    });
    renderFiles();
    return { store, add };
  }

  /* ---------- блок 3: техническое предложение ---------- */
  function techBlock(el) {
    const body = $(el, '.blk-body');
    const st = { tz: null, entries: null, tps: [], seq: 0 };
    body.innerHTML = `
      ${hlBar('hl-tech')}
      <div class="guide"><div><b>1. Техническое задание</b>Загрузите ТЗ заказчика (.docx или .xlsx): программа найдёт все показатели — «не менее», «не более», «не шире», «не выше», «от … до» и другие.</div><div><b>2. Техническое предложение</b>Загрузите пустой бланк предложения (.docx или .xlsx) — в нём показатели будут заменены на конкретные значения.</div><div><b>3. Проверка</b>В списке «Что именно сделано» видно каждую замену и правило. Значения — граничные: сверьте их с реальным товаром.</div></div>
      <div class="cols">
        <section class="card party customer"><h2><span class="n">1</span> Техническое задание (ТЗ)<span class="badge-role">Заказчик</span></h2><div class="sub">Документ заказчика с требованиями к товару, работам или услугам.</div>
          <div class="drop tz-drop" tabindex="0"${tip('Перетащите сюда ТЗ или нажмите и выберите файл Word или Excel.')}><b>Перетащите ТЗ сюда</b> или нажмите, чтобы выбрать<small>.docx, .xlsx</small><input type="file" accept=".docx,.xlsx" hidden></div>
          <div class="status tz-status"></div><ul class="files tz-list"></ul></section>
        <section class="card party executor"><h2><span class="n">2</span> Техническое предложение<span class="badge-role">Участник</span></h2><div class="sub">Бланк, который нужно заполнить. Структура, форма и стиль сохраняются.</div>
          <div class="drop tp-drop" tabindex="0"${tip('Перетащите сюда бланк технического предложения (можно несколько файлов) или нажмите и выберите.')}><b>Перетащите бланк сюда</b> или нажмите, чтобы выбрать<small>.docx, .xlsx</small><input type="file" accept=".docx,.xlsx" multiple hidden></div>
          <ul class="files tp-list"></ul></section>
      </div>
      <section class="card" style="margin-top:20px"><h2><span class="n">3</span> Правила замены</h2><div class="sub">Как поступать с показателями, заданными диапазоном.</div>
        <div class="settings">
          <label${tip('«От 5 до 20 мм» не является значением. Выберите, что подставлять: нижнюю границу, середину или верхнюю. Показатели вида «диапазон измерения» остаются диапазоном всегда, если выбрано «Оставить диапазон» — для всех.')}>Показатель «от … до …»
            <select class="o-range"><option value="min" selected>Подставить нижнюю границу</option><option value="mid">Подставить середину</option><option value="max">Подставить верхнюю границу</option><option value="keep">Оставить диапазон как есть</option></select></label>
          <label class="check"${tip('Если выключить, заменяться будут только таблицы. Обычный текст документа останется как есть.')}><input type="checkbox" class="o-text" checked> Заменять и в обычном тексте (вне таблиц)</label>
        </div>
        <div class="rules"><b>Как это работает.</b> «Не менее / не ниже / не уже / не тоньше / не короче / минимум / ≥ / от N» → <i>N</i>. «Не более / не выше / не шире / не толще / не длиннее / максимум / ≤ / до N» → <i>N</i>. «Более / выше / шире / толще / свыше» → число <i>чуть больше</i> (500 → 501; 2,5 → 2,6). «Менее / ниже / уже / тоньше» → число <i>чуть меньше</i>. «N ± d», «около N» → <i>N</i>. После замены слово согласуется с числом («не менее 24 месяцев» → «24 месяца»).</div>
      </section>
      <div class="row" style="margin-top:16px"><button class="btn primary run-btn" type="button" disabled${tip('Программа найдёт показатели в ТЗ и вставит конкретные значения в техническое предложение. Исходные файлы не меняются.')}>Заполнить техническое предложение</button><span class="hint run-hint">Загрузите ТЗ и бланк предложения.</span></div>
      <div class="res-wrap">${resultsHtml('Скачать все готовые предложения одним архивом ZIP.')}</div>`;
    const res = makeResults($(body, '.res-wrap'), { key: 'tech' });
    const wire = (dropSel, multi, onFiles) => {
      const d = $(body, dropSel), fi = $(d, 'input');
      d.addEventListener('click', () => fi.click()); d.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
      ['dragenter', 'dragover'].forEach(ev => d.addEventListener(ev, e => { e.preventDefault(); d.classList.add('over'); })); ['dragleave', 'drop'].forEach(ev => d.addEventListener(ev, e => { e.preventDefault(); d.classList.remove('over'); }));
      d.addEventListener('drop', e => onFiles(Array.from(e.dataTransfer.files))); fi.addEventListener('change', () => { onFiles(Array.from(fi.files)); fi.value = ''; });
    };
    const runBtn = $(body, '.run-btn');
    const upd = () => { runBtn.disabled = !(st.entries && st.tps.length); $(body, '.run-hint').textContent = !st.tz ? 'Загрузите ТЗ.' : !st.tps.length ? 'Загрузите бланк предложения.' : 'Готово к заполнению.'; };
    wire('.tz-drop', false, async files => {
      const f = files[0]; if (!f) return; const stEl = $(body, '.tz-status'); setStatus(stEl, 'Читаю ' + f.name + '…');
      try { st.entries = await TECH.readTz(f.name, await f.arrayBuffer()); st.tz = f; $(body, '.tz-list').innerHTML = ''; const li = document.createElement('li'); li.innerHTML = '<span class="name"></span>'; li.firstChild.textContent = f.name; $(body, '.tz-list').append(li);
        const n = st.entries.length, q = st.entries.filter(e => TECH.hasQualifier(e.req)).length;
        setStatus(stEl, 'Найдено показателей: ' + n + ', из них с формулировками «не менее / не более / …»: ' + q + '.' + (n ? '' : ' Показатели не распознаны — проверьте, что ТЗ оформлено таблицей или строками «Показатель: значение».'), n ? 'ok' : 'err'); }
      catch (e) { st.entries = null; setStatus(stEl, 'Ошибка: ' + e.message, 'err'); } upd();
    });
    wire('.tp-drop', true, files => { files.forEach(f => st.tps.push({ id: ++st.seq, file: f })); renderTp(); });
    function renderTp() {
      const ul = $(body, '.tp-list'); ul.innerHTML = '';
      st.tps.forEach(f => { const li = document.createElement('li'), nm = document.createElement('span'); nm.className = 'name'; nm.textContent = f.file.name; const x = document.createElement('button'); x.className = 'x'; x.type = 'button'; x.textContent = '×'; x.setAttribute('data-tip', 'Убрать файл из списка.'); x.addEventListener('click', () => { st.tps = st.tps.filter(q => q.id !== f.id); renderTp(); }); li.append(nm, x); ul.append(li); }); upd();
    }
    runBtn.addEventListener('click', async () => {
      runBtn.disabled = true; runBtn.textContent = 'Заполняю…'; const used = new Set(), out = [];
      const opts = { range: $(body, '.o-range').value, replaceInText: $(body, '.o-text').checked, highlight: $(body, '#hl-tech').checked };
      for (const f of st.tps) {
        const name = uniqueName(f.file.name, used, 'предложение');
        try { const r = await TECH.fillProposal(f.file.name, await f.file.arrayBuffer(), st.entries, opts);
          out.push({ name, blob: new Blob([r.data], { type: MIME[extOf(f.file.name)] }), stats: { filled: r.stats.replaced, highlighted: r.stats.highlighted, label: 'заменено:' }, items: r.report.map(x => x.kind === 'hl' ? { kind: 'hl', key: x.name + (x.why ? ' (' + x.why + ')' : '') } : { kind: 'rep', name: x.name, from: x.from, to: x.to, rule: x.rule }) });
        } catch (e) { out.push({ name: f.file.name, blob: f.file, error: e.message || String(e) }); }
      }
      res.set(out); runBtn.textContent = 'Заполнить техническое предложение'; upd();
    });
    upd();
  }

  /* ---------- настройки блоков ---------- */
  const docSettings = (idp, withMoney) => `
    <section class="card" style="margin-top:20px"><h2><span class="n">4</span> Данные документа и настройки</h2><div class="sub">Номер, дата и город. Пустые поля останутся жёлтыми.</div>
      <div class="settings">
        <h3>Документ</h3>
        <label>Номер документа<input type="text" class="d-number" placeholder="например, 15/11-26"></label>
        <label>Дата документа<input type="date" class="d-date"></label>
        <label>Город<input type="text" class="d-city" placeholder="Москва"></label>
        ${withMoney ? `<label>Номер спецификации<input type="text" class="d-spec" placeholder="1"></label>
        <label>Сумма, руб.<input type="text" class="d-amount" placeholder="1 234 567,89"></label>
        <label>Ставка НДС<select class="d-vat"><option value="22" selected>22 %</option><option value="10">10 %</option><option value="0">без НДС</option></select></label>
        <label>Цены в спецификации<select class="d-vatmode"><option value="incl" selected>включают НДС</option><option value="excl">указаны без НДС</option></select></label>
        <label>Тип документов по умолчанию<select class="o-type"><option value="auto" selected>Определять автоматически</option><option value="form">Форма участника (всё — за Участника)</option><option value="contract">Договор (Заказчик и Исполнитель)</option></select></label>` : ''}
        <h3>Параметры заполнения</h3>
        ${idp === 'labor' ? `<label class="check"${tip('Если загруженный договор уже заполнен (другой работодатель, другой работник), программа найдёт в нём прежние данные и заменит их на введённые выше. Если выключить — заполнятся только пустые места.')}><input type="checkbox" class="o-replace" checked> Заменять старые данные в уже заполненном бланке на новые</label><div class="rep-note" style="grid-column:1/-1">После замены проверьте вводную часть договора: окончания слов («именуемый / именуемая», «действующего / действующей») и должность руководителя в тексте программа подгоняет не всегда.</div>` : ''}
        <label class="check"><input type="checkbox" class="o-fallback" checked> Нет почтового/фактического адреса — подставлять юридический</label>
        <label class="check"><input type="checkbox" class="o-contact" checked> Нет контактного лица — подставлять руководителя</label>
        ${withMoney ? `<label>Режим по умолчанию<select class="o-mode"><option value="full">Заполнить все поля</option><option value="min">Только наименование и подписант</option></select></label>` : ''}
      </div></section>`;
  const docInfo = withMoney => ({
    read(b) { const v = s => { const e = $(b, s); return e ? e.value.trim() : ''; }; return { number: v('.d-number'), city: v('.d-city'), specNo: v('.d-spec'), amount: v('.d-amount'), vat: v('.d-vat') || '22', vatMode: v('.d-vatmode') || 'incl', date: v('.d-date') }; },
    write(b, d) { const s = (sel, val) => { const e = $(b, sel); if (e && val != null) e.value = val; }; s('.d-number', d.number); s('.d-city', d.city); s('.d-spec', d.specNo); s('.d-amount', d.amount); s('.d-vat', d.vat); s('.d-vatmode', d.vatMode); s('.d-date', d.date); },
  });
  const readBase = (b, extra) => { const c = s => { const e = $(b, s); return e ? e.checked : true; }, v = (s, d) => { const e = $(b, s); return e ? e.value : d; }; return { type: v('.o-type', 'auto'), engine: Object.assign({ fallbackAddr: c('.o-fallback'), fallbackContact: c('.o-contact'), mode: v('.o-mode', 'full'), docType: v('.o-type', 'auto') }, extra || {}) }; };

  /* ---------- запуск ---------- */
  const blocks = {};
  const el = id => document.getElementById('blk-' + id);
  blocks.labor = fillBlock(el('labor'), {
    key: 'labor', labor: true, templates: (window.AssistantTemplates || {}).labor, suffix: 'заполнено', runLabel: 'Заполнить трудовой договор', roleNames: { customer: 'Работодатель', executor: 'Работник' }, perFile: false,
    accept: '.docx,.dotx', filesTitle: 'Бланк трудового договора', filesSub: 'Выберите один из четырёх готовых бланков ниже или загрузите свой договор в Word (.docx) — пустой или уже заполненный: старые данные можно заменить на новые.', filesTip: 'Перетащите сюда бланк трудового договора в формате Word или нажмите и выберите файл.',
    guide: ['Введите данные работодателя и работника вручную либо загрузите карточки из файлов Word, PDF, Excel или TXT.', 'Выберите готовый бланк из встроенных или загрузите свой договор (.docx). Программа впишет данные в пустые места, а в уже заполненном договоре заменит старые данные на новые.'],
    parties: { customer: { lib: [{ key: 'employers', title: 'Мои работодатели' }], title: 'Работодатель', badge: 'Работодатель', sub: 'Организация или ИП, которые принимают человека на работу.', groups: ORG_GROUPS, demo: DEMO_ORG.a }, executor: { lib: [{ key: 'workers', title: 'Мои работники' }], title: 'Работник', badge: 'Работник', sub: 'Человек, которого принимают на работу: ФИО, паспорт, адрес, должность и оклад.', groups: PERSON_GROUPS, person: true, demo: DEMO_PERSON } },
    settingsHtml: docSettings('labor', false), docInfo: docInfo(false), readOpts: b => readBase(b, { docType: 'contract' }),
  });
  blocks.forms = fillBlock(el('forms'), {
    key: 'forms', templates: (window.AssistantTemplates || {}).forms, tplSub: 'Готовые формы договоров встроены в программу. При выборе данные сторон подставляются из реквизитов Заказчика и Участника / Исполнителя.', suffix: 'заполнено', runLabel: 'Заполнить документы', roleNames: { customer: 'Заказчик', executor: 'Участник' }, perFile: true,
    accept: '.docx,.dotx,.xlsx', filesTitle: 'Бланки и документы', filesSub: 'Заявки, анкеты, договоры, спецификации. Для каждого файла можно уточнить тип.', filesTip: 'Перетащите сюда бланки (Word или Excel) или нажмите и выберите файлы. Можно сразу много.',
    guide: ['Загрузите карты обеих сторон или введите данные вручную. Для заявки участника достаточно блока «Участник / Исполнитель».', 'Добавьте формы заявки, анкеты, договоры, спецификации (.docx, .xlsx) — сколько угодно файлов сразу.'],
    parties: { customer: { lib: [{ key: 'customers', title: 'Мои заказчики' }], title: 'Реквизиты Заказчика', badge: 'Заказчик', sub: 'Сторона, которая заказывает работы, покупает, арендует. Нужна для договоров и спецификаций.', groups: ORG_GROUPS, demo: DEMO_ORG.a }, executor: { lib: [{ key: 'participants', title: 'Мои участники' }, { key: 'contractors', title: 'Мои подрядчики / исполнители' }], title: 'Реквизиты Участника / Исполнителя', badge: 'Участник · Исполнитель', sub: 'Ваша сторона: участник закупки, подрядчик, поставщик. По этим данным заполняются заявки, анкеты и формы.', groups: ORG_GROUPS, demo: DEMO_ORG.b } },
    settingsHtml: docSettings('forms', true), docInfo: docInfo(true), readOpts: b => readBase(b),
  });
  blocks.ks = fillBlock(el('ks'), {
    key: 'ks', suffix: 'КС заполнено', runLabel: 'Заполнить КС-2 / КС-3', roleNames: { customer: 'Заказчик (Генподрядчик)', executor: 'Подрядчик (Субподрядчик)' }, perFile: false,
    accept: '.xlsx,.xlsm,.xls', filesTitle: 'КС-2 и КС-3 (Excel)', filesSub: 'Загрузите акт КС-2 и справку КС-3 в формате Excel. Можно несколько файлов сразу.', filesTip: 'Перетащите сюда КС-2 и КС-3 в формате Excel (.xlsx или .xls) или нажмите и выберите файлы.',
    guide: ['Окно «Заказчик (Генподрядчик)» и окно «Подрядчик (Субподрядчик)»: введите реквизиты и подписантов или загрузите из файлов.', 'Добавьте КС-2 и КС-3 (.xlsx или старый .xls). Программа вставит реквизиты сторон, а под «Сдал» и «Принял» — должности и подписантов.'],
    parties: { customer: { lib: [{ key: 'customers', title: 'Мои заказчики' }], title: 'Заказчик (Генподрядчик)', badge: 'Заказчик · Принял', sub: 'Сторона, которая принимает работы. Подписант этой стороны попадёт в строку «Принял».', groups: ORG_GROUPS, demo: DEMO_ORG.a }, executor: { lib: [{ key: 'contractors', title: 'Мои подрядчики / исполнители' }], title: 'Подрядчик (Субподрядчик)', badge: 'Подрядчик · Сдал', sub: 'Сторона, которая сдаёт работы. Подписант этой стороны попадёт в строку «Сдал».', groups: ORG_GROUPS, demo: DEMO_ORG.b } },
    settingsHtml: `<section class="card" style="margin-top:20px"><h2><span class="n">4</span> Параметры КС-2 / КС-3</h2><div class="sub">Как вписывать реквизиты в акт и справку.</div><div class="settings">
      <label class="check"${tip('Если в форме уже стоят другие реквизиты (старый заказчик, старый подрядчик, старые подписанты) — они будут заменены на новые. Если выключить, заполняться будут только пустые места.')}><input type="checkbox" class="o-overwrite" checked> Заменять уже заполненные реквизиты и подписантов</label>
      <label class="check"${tip('В строке «Заказчик» и «Подрядчик» будет написано: наименование, ИНН, КПП, адрес, телефон — как принято в КС-2. Если выключить, впишется только наименование.')}><input type="checkbox" class="o-line" checked> В строках сторон писать наименование, ИНН, КПП, адрес и телефон</label>
      <label class="check"><input type="checkbox" class="o-contact" checked> Нет контактного лица — подставлять руководителя</label></div></section>`,
    readOpts: b => ({ type: 'form', engine: { mode: 'full', docType: 'form', defaultRole: 'executor', fallbackAddr: true, fallbackContact: $(b, '.o-contact').checked, overwrite: $(b, '.o-overwrite').checked, partyLine: $(b, '.o-line').checked, signers: true, ks: true } }),
  });
  techBlock(el('tech'));

  /* ---------- блок 5: свой файл ---------- */
  (function ownBlock(box) {
    const body = $(box, '.blk-body');
    body.innerHTML = `
      <div class="guide"><div><b>1. Файл</b>Перетащите в рамку ниже свой документ Word (.docx) или Excel (.xlsx, .xls) — любой, не обязательно заполненный в этой программе.</div><div><b>2. Печати</b>Подготовьте печати в <a href="stamp.html">Stamp Studio</a> и нажмите «Добавить в файл». Здесь выберите печать и нажмите на нужное место документа.</div><div><b>3. Результат</b>Скачайте документ в исходном формате или как изображение в PDF. Файл обрабатывается в вашем браузере и никуда не отправляется.</div></div>
      <div class="ov-host"></div>`;
    window.Overlay.mount($(body, '.ov-host'), { hero: true, title: 'Файл и печати', getFiles: () => [] });
  })(el('own'));

  /* ---------- «Моя библиотека» ---------- */
  (function libraryPanel() {
    const box = document.getElementById('libPanel'), head = $(box, '.lib-toggle'), body = $(box, '.lib-body'), sum = document.getElementById('libSummary');
    function render() {
      const total = Object.keys(LIB.LISTS).reduce((a, k) => a + LIB.count(k), 0), nt = LIB.templates().length;
      sum.textContent = total || nt ? 'Записей: ' + total + ' · бланков: ' + nt : 'Сохранённые стороны и бланки — всегда под рукой';
      if (body.hidden) return; body.innerHTML = '';
      const note = document.createElement('div'); note.className = 'hint'; note.textContent = 'Данные хранятся только в этом браузере на этом компьютере и никуда не отправляются. Записи появляются, когда вы нажимаете «Сохранить в …» в окнах реквизитов. Паспортные данные работников сохраняйте только если это допустимо у вас. Чтобы перенести библиотеку на другой компьютер или коллеге — сохраните копию в файл и загрузите её там.'; body.append(note);
      const mk = (t, c, ...kids) => { const e = document.createElement(t); if (c) e.className = c; kids.forEach(k => e.append(k)); return e; };
      const grid = mk('div', 'lib-grid');
      for (const [k, title] of Object.entries(LIB.LISTS)) {
        const card = mk('div', 'lib-card'), h = mk('h3'); h.textContent = title + ' (' + LIB.count(k) + ')'; card.append(h);
        const arr = LIB.list(k); if (!arr.length) card.append(mk('div', 'tpl-sub', document.createTextNode('Пока пусто.')));
        for (const r of arr) { const row = mk('div', 'lib-item'), nm = mk('span', 'name'); nm.textContent = LIB.nameOf(r.card) + (r.card.inn ? ' · ИНН ' + r.card.inn : ''); nm.title = nm.textContent; const x = mk('button', 'x'); x.type = 'button'; x.textContent = '×'; x.setAttribute('data-tip', 'Убрать из списка.'); x.addEventListener('click', () => { if (confirm('Убрать «' + LIB.nameOf(r.card) + '» из списка «' + title + '»?')) LIB.remove(k, r.id); }); row.append(nm, x); card.append(row); }
        grid.append(card);
      }
      const tc = mk('div', 'lib-card'), th = mk('h3'); th.textContent = 'Мои бланки (' + nt + ')'; tc.append(th);
      const BN = { labor: 'трудовые договоры', forms: 'формы и договоры', ks: 'КС-2, КС-3' };
      if (!nt) tc.append(mk('div', 'tpl-sub', document.createTextNode('Пока пусто. Добавьте бланк в блок и нажмите «В мою библиотеку».')));
      for (const t of LIB.templates()) { const row = mk('div', 'lib-item'), nm = mk('span', 'name'); nm.textContent = t.name + ' — ' + (BN[t.block] || t.block); nm.title = nm.textContent; const x = mk('button', 'x'); x.type = 'button'; x.textContent = '×'; x.addEventListener('click', () => { if (confirm('Удалить бланк «' + t.name + '»?')) LIB.removeTemplate(t.id); }); row.append(nm, x); tc.append(row); }
      grid.append(tc); body.append(grid);
      const act = mk('div', 'row'), b1 = mk('button', 'btn small'), b2 = mk('label', 'btn small ghost'), b3 = mk('button', 'btn small ghost'), st = mk('span', 'hint');
      b1.type = 'button'; b1.textContent = 'Сохранить копию библиотеки (.json)'; b1.addEventListener('click', () => download(new Blob([LIB.exportJson()], { type: 'application/json' }), 'Моя_библиотека_ассистент.json'));
      b2.textContent = 'Загрузить копию…'; const fi = document.createElement('input'); fi.type = 'file'; fi.accept = '.json,application/json'; fi.hidden = true; b2.append(fi);
      fi.addEventListener('change', async () => { const f = fi.files[0]; fi.value = ''; if (!f) return; try { const n = LIB.importJson(await f.text()); st.textContent = 'Добавлено или обновлено записей: ' + n + '.'; } catch (e) { st.textContent = 'Ошибка: ' + (e.message || e); } });
      b3.type = 'button'; b3.textContent = 'Очистить библиотеку'; b3.addEventListener('click', () => { if (confirm('Удалить из браузера все списки и бланки библиотеки? Это нельзя отменить.')) LIB.clearAll(); });
      act.append(b1, b2, b3, st); body.append(act);
    }
    head.addEventListener('click', () => { const open = body.hidden; body.hidden = !open; head.setAttribute('aria-expanded', String(open)); box.classList.toggle('open', open); render(); });
    LIB.onChange(render); render();
  })();

  /* ---------- раскрытие блоков ---------- */
  document.querySelectorAll('.blk').forEach(b => {
    const head = $(b, '.blk-head'), body = $(b, '.blk-body');
    head.addEventListener('click', () => { const open = body.hidden; body.hidden = !open; head.setAttribute('aria-expanded', String(open)); b.classList.toggle('open', open); });
  });
  if (location.hash) { const t = document.querySelector(location.hash.replace('#', '#blk-')); if (t) $(t, '.blk-head').click(); }
  window.AssistantDocs = { blocks };
})();
