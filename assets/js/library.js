/* Ассистент — «Моя библиотека»: списки сторон (заказчики, работодатели, работники, подрядчики/исполнители, участники) и свои бланки.
   Хранится только в этом браузере (localStorage). Можно выгрузить копию в файл .json и загрузить на другом компьютере. */
(function () {
  'use strict';
  const KEY = 'assistant.library.v1';
  const LISTS = {
    customers: 'Мои заказчики', employers: 'Мои работодатели', workers: 'Мои работники', contractors: 'Мои подрядчики / исполнители', participants: 'Мои участники',
  };
  let data = { lists: {}, templates: [] }; const subs = [];
  try { const s = JSON.parse(localStorage.getItem(KEY) || 'null'); if (s && typeof s === 'object') data = { lists: s.lists || {}, templates: s.templates || [] }; } catch (e) {}
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  function save() { try { localStorage.setItem(KEY, JSON.stringify(data)); } catch (e) { return false; } subs.forEach(f => { try { f(); } catch (e) {} }); return true; }
  const norm = s => String(s || '').toLowerCase().replace(/[«»"'\s.,]+/g, ' ').trim();
  const nameOf = c => c.name_full || c.name_short || c.director_fio || '';

  const L = {
    LISTS,
    onChange(f) { subs.push(f); },
    list(k) { return (data.lists[k] || []).slice().sort((a, b) => nameOf(a.card).localeCompare(nameOf(b.card), 'ru')); },
    count(k) { return (data.lists[k] || []).length; },
    /** Сохраняет карточку; запись с тем же ИНН (или тем же названием) обновляется. Возвращает { id, updated, ok } */
    add(k, card) {
      const c = JSON.parse(JSON.stringify(card || {})), nm = nameOf(c); if (!nm) return { ok: false, why: 'empty' };
      const arr = data.lists[k] = data.lists[k] || [];
      const ex = arr.find(r => (c.inn && r.card.inn === c.inn && norm(nameOf(r.card)).split(' ')[0] === norm(nm).split(' ')[0]) || norm(nameOf(r.card)) === norm(nm));
      if (ex) { ex.card = c; ex.ts = Date.now(); return { ok: save(), id: ex.id, updated: true }; }
      const rec = { id: uid(), card: c, ts: Date.now() }; arr.push(rec); return { ok: save(), id: rec.id, updated: false };
    },
    get(k, id) { const r = (data.lists[k] || []).find(x => x.id === id); return r ? JSON.parse(JSON.stringify(r.card)) : null; },
    remove(k, id) { data.lists[k] = (data.lists[k] || []).filter(r => r.id !== id); save(); },
    nameOf,
    /* бланки */
    templates(block) { return data.templates.filter(t => !block || t.block === block); },
    addTemplate(block, name, b64) { const ex = data.templates.find(t => t.block === block && t.name === name); if (ex) { ex.b64 = b64; ex.ts = Date.now(); return { ok: save(), updated: true }; } data.templates.push({ id: uid(), block, name, b64, ts: Date.now() }); return { ok: save() }; },
    removeTemplate(id) { data.templates = data.templates.filter(t => t.id !== id); save(); },
    /* копия */
    exportJson() { return JSON.stringify({ app: 'assistant-library', v: 1, data }, null, 1); },
    importJson(text) {
      const j = JSON.parse(text); if (!j || j.app !== 'assistant-library' || !j.data) throw new Error('Это не файл библиотеки АССИСТЕНТА');
      let n = 0;
      for (const k of Object.keys(j.data.lists || {})) for (const r of j.data.lists[k]) { if (!LISTS[k] || !r.card) continue; const x = L.add(k, r.card); if (x.ok) n++; }
      for (const t of j.data.templates || []) if (t.name && t.b64) { L.addTemplate(t.block || 'forms', t.name, t.b64); n++; }
      return n;
    },
    clearAll() { data = { lists: {}, templates: [] }; save(); },
  };
  window.AssistantLibrary = L;
})();
