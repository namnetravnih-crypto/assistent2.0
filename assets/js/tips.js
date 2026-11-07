/* Ассистент — всплывающие памятки. Любому элементу достаточно атрибута data-tip="текст". */
(function () {
  'use strict';
  const box = document.createElement('div'); box.className = 'tip-pop'; box.setAttribute('role', 'tooltip'); box.hidden = true;
  document.addEventListener('DOMContentLoaded', () => document.body.appendChild(box));
  if (document.body) document.body.appendChild(box);
  let cur = null, timer = null;
  function place(el) {
    const r = el.getBoundingClientRect(); box.style.maxWidth = Math.min(320, window.innerWidth - 24) + 'px';
    box.hidden = false; box.style.left = '0px'; box.style.top = '0px';
    const b = box.getBoundingClientRect();
    let x = r.left + r.width / 2 - b.width / 2, y = r.bottom + 10;
    if (y + b.height > window.innerHeight - 8) y = Math.max(8, r.top - b.height - 10);
    x = Math.max(8, Math.min(x, window.innerWidth - b.width - 8));
    const k = window.UIZoom ? window.UIZoom.k() : 1;          // сайт показан в масштабе 80 %: переводим экранные пиксели в пиксели вёрстки
    box.style.left = (x / k) + 'px'; box.style.top = (y / k) + 'px';
  }
  function show(el) { const t = el.getAttribute('data-tip'); if (!t) return; cur = el; box.textContent = t; place(el); }
  function hide() { cur = null; box.hidden = true; clearTimeout(timer); }
  document.addEventListener('mouseover', e => { const el = e.target.closest && e.target.closest('[data-tip]'); if (!el || el === cur) return; clearTimeout(timer); timer = setTimeout(() => show(el), 220); });
  document.addEventListener('mouseout', e => { const el = e.target.closest && e.target.closest('[data-tip]'); if (el && !el.contains(e.relatedTarget)) hide(); });
  document.addEventListener('focusin', e => { const el = e.target.closest && e.target.closest('[data-tip]'); if (el) show(el); });
  document.addEventListener('focusout', hide);
  document.addEventListener('scroll', hide, true); document.addEventListener('mousedown', hide);

  const T = (sel, text) => document.querySelectorAll(sel).forEach(el => { if (!el.hasAttribute('data-tip')) { el.setAttribute('data-tip', text); el.removeAttribute('title'); } });
  function stampTips() {
    if (!document.getElementById('btnDownload')) return;
    T('#uploadLabel', 'Выберите фото, скан или PDF с оттиском. Файл остаётся только в вашем браузере.');
    T('#dropZone', 'Перетащите сюда файл или нажмите на поле — откроется выбор файла.');
        T('#btnAuto', 'Программа заново найдёт оттиск и уберёт фон, блики и тени.');
    T('#btnMakeRound', 'Нажмите, если оттиск на фото получился овальным — программа выровняет его в круг.');
    T('#manualBox > summary', 'Тонкая настройка для сложных случаев. Новичку обычно не нужна.');
    T('#btnPick', 'Нажмите, затем щёлкните по цвету на картинке — этот цвет исчезнет.');
    T('#presetChips .chip', 'Быстро убрать этот цвет с оттиска.');
    T('#btnMagicCursor', 'Уточнение края: внутри круга остаётся только плотный оттиск, фон убирается.');
    T('#tbEdgeSens', 'Чувствительность «Уточнения края». Правее — остаются и тонкие, бледные линии; левее — только самый плотный оттиск. Действует на новые мазки.');
    T('#tbRemoveText', 'Для цветного оттиска: чёрный и серый текст вокруг исчезнет. Для чёрного или серого оттиска не включайте.');
    T('#tbTextPower', 'Сила лишнего текста: чем больше, тем смелее убирается чёрный и серый текст. Работает при включённой галочке выше.');
    T('#btnTextMask', 'Маска: нажмите на результат — в этом месте вернётся чёрный и серый текст, убранный галочкой. Ширину круга меняет ползунок «Размер» и колесо мыши.');
    T('#btnEraser', 'Ластик: стирает всё, что попало под круг.');
    T('#btnRestore', 'Вернуть: возвращает то, что стёрла очистка.');
    T('#brushSize', 'Размер круга для «Уточнения края», «Маски», ластика и «Вернуть». Его же меняет колесо мыши над результатом.');
    T('#btnUndo', 'Отменить последнее действие (Ctrl+Z).');
    T('#bgSwitch button', 'Фон под результатом — только для просмотра, на файл не влияет.');
    T('#btnDownload', 'Скачать готовую печать в PNG. Диаметр печати — ровно 40 мм.');
    T('#btnAddToFile', 'Печать сохранится в памяти браузера (до четырёх печатей одновременно) на 30 минут, потом сотрётся сама — либо сразу после скачивания файла, на который её поставили. Наложить её можно в программе «Автозаполнение».');
    T('#btnReset', 'Очистить всё и начать сначала.');
    T('#memClear', 'Стереть из памяти браузера все сохранённые печати.');
    document.querySelectorAll('.step').forEach(st => { const n = st.querySelector('.step-title'); if (n) st.setAttribute('data-tip', 'Шаг «' + n.textContent + '».'); });
  }
  function portalTips() {
    T('.tile[href="stamp.html"]', 'Программа I: подготовка печати — очистка фона и подписи, PNG 40 мм.');
    T('.tile[href="docs.html"]', 'Программа II: заполнение трудовых договоров, форм и договоров, техпредложений, КС-2/КС-3 и наложение печати на любой свой файл.');
    T('.crumbs a[href="index.html"]', 'Вернуться на главную страницу.');
    T('.crumbs a[href="stamp.html"]', 'Перейти к программе подготовки печати.');
    T('.crumbs a[href="docs.html"]', 'Перейти к программе автозаполнения документов.');
  }
  document.addEventListener('DOMContentLoaded', () => { stampTips(); portalTips(); });
  window.Tips = { T };
})();
