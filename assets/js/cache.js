/* =====================================================================
   Ассистент — общая кэш-память браузера (IndexedDB), одна на обе программы.
   • Печати: до 4 штук одновременно (Печать 1…4). Каждая стирается через 30 минут
     либо сразу после того, как файл, на который её поставили, скачан.
   • Готовые документы автозаполнения: хранятся 30 минут отдельно по блокам.
   Данные лежат в браузере, переживают переход между программами и обновление страницы.
   ===================================================================== */
(function () {
  'use strict';
  const TTL = 30 * 60 * 1000, MAX = 4, DB_NAME = 'assistant-cache', OLD_KEY = 'assistant.stamp.cache.v1';
  let db = null, bc = null, timer = null;
  const mem = { stamps: [], docs: [] };      // зеркало базы в памяти: чтение для интерфейса всегда мгновенное
  const listeners = [];
  const emit = () => listeners.forEach(f => { try { f(); } catch (e) {} });
  const rid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

  /* ---------- IndexedDB ---------- */
  function openDb() {
    return new Promise(res => {
      try {
        const rq = indexedDB.open(DB_NAME, 1);
        rq.onupgradeneeded = () => { const d = rq.result; ['stamps', 'docs'].forEach(s => { if (!d.objectStoreNames.contains(s)) d.createObjectStore(s, { keyPath: 'id' }); }); };
        rq.onsuccess = () => res(rq.result); rq.onerror = () => res(null); rq.onblocked = () => res(null);
      } catch (e) { res(null); }
    });
  }
  const getAll = store => new Promise(res => {
    if (!db) return res([]);
    try { const rq = db.transaction(store, 'readonly').objectStore(store).getAll(); rq.onsuccess = () => res(rq.result || []); rq.onerror = () => res([]); } catch (e) { res([]); }
  });
  const put = (store, rec) => new Promise(res => {
    if (!db) return res(true);                       // базы нет (закрытый режим) — работаем только в памяти страницы
    try { const t = db.transaction(store, 'readwrite'); t.objectStore(store).put(rec); t.oncomplete = () => res(true); t.onerror = t.onabort = () => res(false); } catch (e) { res(false); }
  });
  const del = (store, ids) => new Promise(res => {
    if (!db || !ids.length) return res(true);
    try { const t = db.transaction(store, 'readwrite'), os = t.objectStore(store); ids.forEach(i => os.delete(i)); t.oncomplete = () => res(true); t.onerror = t.onabort = () => res(false); } catch (e) { res(false); }
  });
  const notify = () => { try { if (bc) bc.postMessage('changed'); } catch (e) {} };

  /* ---------- срок жизни ---------- */
  function purge() {
    const now = Date.now(), ds = mem.stamps.filter(r => r.expires <= now), dd = mem.docs.filter(r => r.expires <= now);
    if (ds.length) { mem.stamps = mem.stamps.filter(r => r.expires > now); del('stamps', ds.map(r => r.id)); }
    if (dd.length) { mem.docs = mem.docs.filter(r => r.expires > now); del('docs', dd.map(r => r.id)); }
    return ds.length + dd.length;
  }
  function schedule() {
    if (timer) { clearTimeout(timer); timer = null; }
    const all = mem.stamps.concat(mem.docs); if (!all.length) return;
    const next = Math.min.apply(null, all.map(r => r.expires));
    timer = setTimeout(() => { if (purge()) { notify(); emit(); } schedule(); }, Math.min(2147000000, Math.max(0, next - Date.now()) + 60));
  }
  const changed = () => { schedule(); notify(); emit(); };

  /* ---------- загрузка ---------- */
  const valid = r => r && r.id && r.expires > Date.now();
  async function reload() {
    const [s, d] = await Promise.all([getAll('stamps'), getAll('docs')]);
    mem.stamps = s.filter(r => valid(r) && r.dataUrl).sort((a, b) => a.slot - b.slot);
    mem.docs = d.filter(valid).sort((a, b) => a.order - b.order);
    const removed = purge(); schedule(); emit(); if (removed) notify();
  }
  const ready = (async () => {
    db = await openDb();
    try { if ('BroadcastChannel' in window) { bc = new BroadcastChannel('assistant-cache'); bc.onmessage = () => { reload(); }; } } catch (e) {}
    await reload();
    // перенос печати из прежней версии кэша (одна печать в localStorage)
    try {
      const old = JSON.parse(localStorage.getItem(OLD_KEY) || 'null');
      if (old) { localStorage.removeItem(OLD_KEY); if (old.dataUrl && old.expires > Date.now() && !mem.stamps.length) { const rec = { id: 's' + rid(), slot: 1, dataUrl: old.dataUrl, w: old.w, h: old.h, saved: old.saved || Date.now(), expires: old.expires }; mem.stamps.push(rec); await put('stamps', rec); changed(); } }
    } catch (e) {}
  })();
  setInterval(() => { if (purge()) { notify(); emit(); } }, 30000);

  /* ---------- печати ---------- */
  const StampCache = {
    TTL, MAX, ready,
    list() { const n = Date.now(); return mem.stamps.filter(r => r.expires > n).slice(); },
    get(id) { const l = this.list(); return id ? (l.find(r => r.id === id) || null) : (l[0] || null); },
    count() { return this.list().length; },
    remaining(rec) { const r = rec || this.get(); return r ? Math.max(0, r.expires - Date.now()) : 0; },
    async save(dataUrl, w, h) {
      await ready; purge();
      if (mem.stamps.length >= MAX) return { ok: false, reason: 'full' };
      const used = new Set(mem.stamps.map(r => r.slot)); let slot = 1; while (used.has(slot)) slot++;
      const rec = { id: 's' + rid(), slot, dataUrl, w, h, saved: Date.now(), expires: Date.now() + TTL };
      mem.stamps.push(rec); mem.stamps.sort((a, b) => a.slot - b.slot);
      if (!(await put('stamps', rec))) { mem.stamps = mem.stamps.filter(r => r !== rec); emit(); return { ok: false, reason: 'storage' }; }
      changed(); return { ok: true, slot, id: rec.id };
    },
    async removeMany(ids) {
      ids = (ids || []).filter(Boolean); if (!ids.length) return;
      mem.stamps = mem.stamps.filter(r => !ids.includes(r.id)); changed(); await del('stamps', ids);
    },
    remove(id) { return this.removeMany([id]); },
    clear() { return this.removeMany(mem.stamps.map(r => r.id)); },
    onChange(f) { listeners.push(f); },
    bytesToDataUrl(bytes) { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return 'data:image/png;base64,' + btoa(s); },
  };

  /* ---------- готовые документы автозаполнения ---------- */
  const DocCache = {
    TTL, ready,
    /** Сохраняет результат работы блока (прежние документы этого блока заменяются). Номера cid назначаются сразу. */
    save(block, items) {
      const old = mem.docs.filter(r => r.block === block).map(r => r.id);
      mem.docs = mem.docs.filter(r => r.block !== block);
      const now = Date.now(), recs = items.map((it, i) => {
        const rec = { id: 'd' + rid(), block, order: i, name: it.name, blob: it.blob, stats: it.stats || null, items: it.items || null, error: it.error || null, saved: now, expires: now + TTL };
        it.cid = rec.id; it.expires = rec.expires; return rec;
      });
      mem.docs.push.apply(mem.docs, recs); changed();
      return ready.then(async () => { await del('docs', old); let ok = true; for (const r of recs) if (!(await put('docs', r))) ok = false; return ok; });
    },
    /** Документы блока в виде элементов списка результатов */
    list(block) {
      const n = Date.now();
      return mem.docs.filter(r => r.block === block && r.expires > n).sort((a, b) => a.order - b.order)
        .map(r => { const it = { name: r.name, blob: r.blob, cid: r.id, expires: r.expires }; if (r.stats) it.stats = r.stats; if (r.items) it.items = r.items; if (r.error) it.error = r.error; return it; });
    },
    count(block) { return this.list(block).length; },
    remaining(block) { const l = this.list(block); return l.length ? Math.max(0, Math.min.apply(null, l.map(x => x.expires)) - Date.now()) : 0; },
    clearBlock(block) { const ids = mem.docs.filter(r => r.block === block).map(r => r.id); mem.docs = mem.docs.filter(r => r.block !== block); changed(); return del('docs', ids); },
    onChange(f) { listeners.push(f); },
  };

  window.StampCache = StampCache; window.DocCache = DocCache;
})();
