/* Ассистент — трудовые договоры: данные сторон из чужого договора и замена данных в уже заполненном бланке.
   Работает поверх engine.js (AssistantEngine.internals); без обращения к сети. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./engine.js'));
  else root.AssistantLabor = factory(root.AssistantEngine);
})(typeof self !== 'undefined' ? self : this, function (E) {
  'use strict';
  const I = E.internals, JSZip = I.JSZip;
  const ucFirst = s => s ? s[0].toUpperCase() + s.slice(1) : s;
  const same = (a, b) => String(a || '').replace(/\s+/g, ' ').trim().toLowerCase() === String(b || '').replace(/\s+/g, ' ').trim().toLowerCase();

  /* ---------- обратное склонение: родительный → именительный ---------- */
  const NOM_SPECIAL = { 'петра': 'Пётр', 'пётра': 'Пётр', 'павла': 'Павел', 'льва': 'Лев', 'ильи': 'Илья', 'кузьмы': 'Кузьма', 'любови': 'Любовь', 'фомы': 'Фома', 'саввы': 'Савва', 'никиты': 'Никита', 'ильичa': 'Ильич' };
  function nomWord(w, kind) {   // kind: 'sur' | 'name' | 'pat'; sex: из контекста
    const low = w.toLowerCase(); if (NOM_SPECIAL[low]) return NOM_SPECIAL[low];
    const cap = ucFirst, rep = (re, to) => cap(low.replace(re, to));
    if (kind === 'pat') { if (/(вича|ьича|ича)$/.test(low)) return rep(/а$/, ''); if (/(вны|чны)$/.test(low)) return rep(/ы$/, 'а'); return cap(low); }
    // женские формы
    if (/(овой|евой|ёвой|иной|ыной)$/.test(low)) return rep(/ой$/, 'а');
    if (/(ской|цкой)$/.test(low)) return rep(/ой$/, 'ая');
    if (kind === 'name') {
      if (/(ии)$/.test(low)) return rep(/и$/, 'я');
      if (/[жшщчцкгхбвдзлмнпрстфй]ы$/.test(low) || /[аяеоуыиэю]ы$/.test(low)) return rep(/ы$/, 'а');
      if (/[бвгджзклмнпрстфхцчшщ]и$/.test(low)) return rep(/и$/, 'я');
      if (/ея$/.test(low)) return rep(/я$/, 'й');
      if (/ия$/.test(low)) return rep(/я$/, 'й');
      if (/[рлн]я$/.test(low)) return rep(/я$/, 'ь');
      if (/а$/.test(low) && !/[аеёиоуыэюя]а$/.test(low)) return rep(/а$/, '');
      return cap(low);
    }
    // фамилия (мужская)
    if (/ского$/.test(low)) return rep(/ого$/, 'ий');
    if (/(ого)$/.test(low)) return rep(/ого$/, 'ый');
    if (/(ова|ева|ёва|ина|ына)$/.test(low)) return rep(/а$/, '');
    if (/[бвгджзклмнпрстфхцчшщй]а$/.test(low) && low.length > 4) return rep(/а$/, '');
    if (/я$/.test(low) && low.length > 4) return rep(/я$/, 'ь');
    return cap(low);
  }
  function fioNom(gen) {
    const w = String(gen || '').trim().split(/\s+/).filter(Boolean); if (!w.length) return '';
    const out = w.map((x, i) => i === 0 ? nomWord(x, 'sur') : i === 1 ? nomWord(x, 'name') : nomWord(x, 'pat'));
    return out.join(' ');
  }
  function posNom(gen) {
    return String(gen || '').trim().split(/\s+/).map((x, i) => {
      const low = x.toLowerCase(); let r = low;
      if (/ского$/.test(low)) r = low.replace(/ого$/, 'ий'); else if (/ого$/.test(low)) r = low.replace(/ого$/, 'ый'); else if (/ей$/.test(low) && i === 0 && false) r = low;
      else if (/(теля|еля)$/.test(low)) r = low.replace(/я$/, 'ь'); else if (/а$/.test(low) && !/(ства|сва)$/.test(low)) r = low.replace(/а$/, ''); else if (/я$/.test(low)) r = low.replace(/я$/, 'ь');
      return i === 0 ? ucFirst(r) : r;
    }).join(' ');
  }

  /* ---------- строки документа: абзацы с учётом ячеек ---------- */
  async function loadLines(buf) {
    const zip = await JSZip.loadAsync(buf), f = zip.file('word/document.xml'); if (!f) throw new Error('Это не DOCX-файл');
    const doc = I.parseXml(await f.async('string')), lines = [];
    (function walk(node, cell) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue; const n = I.lname(c);
        if (n === 'p') { const t = I.clean(I.mapParagraph(c).text).trim(); if (t) lines.push({ t, cell }); }
        else if (n === 'tc') walk(c, c);
        else if (n === 'tbl' || n === 'tr' || n === 'sdt' || n === 'sdtContent') walk(c, cell);
      }
    })(I.byNS(doc, I.NS_W, 'body')[0] || doc.documentElement, null);
    return lines;
  }
  const ROLE_WORD = /^(работодател[а-яё]*|работник[а-яё]*)(?=$|[\s:.\-–—«"])/i;
  const roleOfLine = t => { const m = ROLE_WORD.exec(t.trim()); return m ? (/^работодател/i.test(m[1]) ? 'employer' : 'worker') : null; };

  /* ---------- вводная часть: «Организация, в лице …, именуемое «Работодатель», … и ФИО, именуемый «Работник»» ---------- */
  function parsePreamble(lines) {
    const head = lines.slice(0, 14).map(l => l.t).join(' ');
    const m = /^([\s\S]*?)с\s+одной\s+стороны[,\s]+и\s+([\s\S]*?)с\s+другой\s+стороны/i.exec(head); if (!m) return {};
    const out = {};
    for (const part of [m[1], m[2]]) {
      const isEmp = /[«"]\s*работодател/i.test(part) || /именуем[а-яё]*\s+(?:в\s+дальнейшем\s+)?[«"]?работодател/i.test(part);
      const isWrk = /[«"]\s*работник/i.test(part) || /именуем[а-яё]*\s+(?:в\s+дальнейшем\s+)?[«"]?работник/i.test(part);
      if (!isEmp && !isWrk) continue;
      const c = {};
      let name = part.split(/,?\s*(?:в\s+лице|именуем[а-яё]*)/i)[0].replace(/^[\s,]+|[\s,]+$/g, '');
      if (isWrk) name = name.replace(/^(?:гражданин(?:ка)?\s+(?:российской\s+федерации|рф|р\.ф\.)?|гр\.\s*)/i, '').trim();
      if (name) c.name_full = name;
      const lic = /в\s+лице\s+([^,]+?)\s+([А-ЯЁ][а-яё]+(?:\s+[А-ЯЁ][а-яё]+){1,2})\s*,/i.exec(part);
      if (lic && isEmp) { c.director_pos = posNom(lic[1]); c.director_fio = fioNom(lic[2]); c.director_pos_gen = lic[1].trim(); c.director_fio_gen = lic[2].trim(); }
      const bas = /действующ[а-яё]*\s+на\s+основании\s+([^,;]+?)(?:\s*,|\s+с\s+(?:одной|другой)|$)/i.exec(part); if (bas && isEmp) c.basis = bas[1].trim();
      out[isEmp ? 'employer' : 'worker'] = c;
    }
    return out;
  }

  /** Достаёт данные работодателя и работника из готового трудового договора. */
  async function extractParties(buf) {
    const lines = await loadLines(buf), pre = parsePreamble(lines);
    // раздел реквизитов — с последнего заголовка «Адреса/Реквизиты/Подписи», иначе последняя треть
    let start = -1; lines.forEach((l, i) => { if (l.t.length < 90 && /реквизит|адреса\s+и|подписи\s+сторон|подписи$/i.test(l.t) && !/^\d+\.\d+/.test(l.t)) start = i; });
    if (start < 0) start = Math.floor(lines.length * 0.62);
    const bucket = { employer: [], worker: [] }; let role = null, lastCell = null;
    for (let i = start; i < lines.length; i++) {
      const l = lines[i]; if (l.cell !== lastCell) { lastCell = l.cell; role = null; }
      let t = l.t; const r = roleOfLine(t);
      if (r) { role = r; t = t.replace(ROLE_WORD, '').replace(/^[\s:.\-–—]+/, '').trim(); if (!t) continue; }
      if (!role || /_{3,}|^м\.?\s*п\.?$/i.test(t) || /^второй\s+экземпляр/i.test(t)) continue;
      bucket[role].push(t);
    }
    const res = { employer: {}, worker: {} };
    for (const r of ['employer', 'worker']) { if (r === 'employer' && !bucket.employer.some(t => /^(?:банк|р\/?с)/i.test(t))) { /* банк без подписи */ } try { res[r] = bucket[r].length ? E.parseCardText(bucket[r].join('\n')) : {}; } catch (e) { res[r] = {}; } }
    if (!res.employer.bank_name) { const b = bucket.employer.find(t => /банк/i.test(t) && !/[:：]/.test(t) && t.length < 90); if (b) res.employer.bank_name = b; }
    // вводная часть дополняет то, чего нет в реквизитах
    const pe = pre.employer || {}, pw = pre.worker || {};
    for (const k of Object.keys(pe)) if (!res.employer[k]) res.employer[k] = pe[k];
    if (!res.employer.name_full && res.employer.name_short) res.employer.name_full = res.employer.name_short;
    if (pw.name_full && !res.worker.name_full) res.worker.name_full = pw.name_full;
    if (res.worker.name_full) { res.worker.director_fio = res.worker.director_fio || res.worker.name_full; res.worker.name_short = E.shortFio(res.worker.name_full); }
    // должность / оклад / срок из текста договора (не из реквизитов)
    const all = lines.map(l => l.t).join('\n');
    const pos = /(?:в\s+должности|на\s+должность)\s+[«"]?([^»",.;\n]{3,80}?)[»"]?(?=\s+в\s+структурн|[,.;]|\s*$|\s+(?:с|на)\s)/i.exec(all); if (pos && !res.worker.director_pos) res.worker.director_pos = pos[1].trim();
    const dep = /структурн\w+\s+подразделени\w+\s+[«"]([^»"]{2,80})[»"]/i.exec(all); if (dep) res.worker.department = dep[1].trim();
    const sal = /оклад[^\d]{0,60}?([\d][\d\s ]*(?:[.,]\d{1,2})?)\s*(?:\(|руб)/i.exec(all); if (sal && !res.worker.salary) res.worker.salary = sal[1].replace(/[\s ]+/g, '').replace(',', '.').replace(/\.00$/, '');
    const sd = /дата\s+начала\s+работы\s*:?\s*(\d{2}\.\d{2}\.\d{4})/i.exec(all); if (sd) res.worker.start_date = sd[1];
    const ed = /(?:на\s+срок\s+до|по)\s+(\d{2}\.\d{2}\.\d{4})/i.exec(all); if (ed) res.worker.end_date = ed[1];
    const pr = /испытательн\w+\s+срок[^.:]{0,30}:\s*([^.;\n]{2,40})/i.exec(all); if (pr && !/_{3,}/.test(pr[1])) res.worker.probation = pr[1].trim();
    for (const k of ['employer', 'worker']) for (const kk of Object.keys(res[k])) if (!res[k][kk] || /_{3,}/.test(res[k][kk])) delete res[k][kk];
    return { employer: res.employer, worker: res.worker, found: { employer: Object.keys(res.employer).length, worker: Object.keys(res.worker).length } };
  }

  /* ---------- замена старых данных на новые в заполненном бланке ---------- */
  const NORM_MAP = { ' ': ' ', ' ': ' ', '“': '"', '”': '"', '«': '"', '»': '"', '„': '"', '‟': '"', '–': '-', '—': '-', 'ё': 'е', 'Ё': 'Е' };
  const normText = s => s.replace(/[  “”«»„‟–—ёЁ]/g, c => NORM_MAP[c]);
  const esc = s => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  function textRe(from) { const parts = normText(String(from).trim()).split(/\s+/).map(esc); return new RegExp('(?<![A-Za-zА-Яа-я0-9])' + parts.join('[\\s]+') + '(?![A-Za-zА-Яа-я0-9])', 'gi'); }
  function numRe(from) { const d = String(from).replace(/\D/g, ''); if (d.length < 4) return null; return new RegExp('(?<!\\d)\\+?' + d.split('').join('[\\s\\-().]{0,2}') + '(?!\\d)', 'g'); }
  const NUMERIC = new Set(['inn', 'kpp', 'ogrn', 'bik', 'rs', 'ks', 'passport', 'passport_code', 'snils', 'phone', 'bank_inn', 'okpo']);
  const DATEY = new Set(['birth_date', 'passport_date', 'start_date', 'end_date']);
  const initials = fio => { const w = String(fio || '').trim().split(/\s+/); return w.length >= 2 ? w.slice(1).map(x => x[0].toUpperCase() + '.').join('') : ''; };
  const initialsFirst = fio => { const sh = E.shortFio(fio); const m = /^(\S+)\s+(\S+)$/.exec(sh || ''); return m ? m[2] + ' ' + m[1] : ''; };

  /** Составляет список замен: [{from, to, label, num?}] */
  function buildMap(oldC, newC, person) {
    const list = [], add = (from, to, label, kind) => { from = String(from || '').trim(); if (from.length < 2 || same(from, to)) return; list.push({ from, to: String(to || '').trim(), label, kind }); };
    const keys = new Set([...Object.keys(oldC), ...Object.keys(newC)]);
    const SOFT = new Set(['director_pos', 'department', 'tab_no', 'start_date', 'end_date', 'probation']);   // условия работы: если в новых данных пусто — оставляем как было
    const FIELDS = person
      ? ['birth_date', 'birth_place', 'citizenship', 'passport', 'passport_date', 'passport_issued', 'passport_code', 'snils', 'inn', 'addr_legal', 'addr_fact', 'phone', 'email', 'director_pos', 'department', 'tab_no', 'start_date', 'end_date', 'probation']
      : ['name_full', 'name_short', 'inn', 'kpp', 'ogrn', 'okpo', 'addr_legal', 'addr_post', 'addr_fact', 'bank_name', 'bik', 'rs', 'ks', 'phone', 'email', 'site'];
    for (const k of FIELDS) if (keys.has(k) && oldC[k] && !(SOFT.has(k) && !newC[k])) add(oldC[k], newC[k], E.FIELD_TITLES[k] || k, NUMERIC.has(k) ? 'num' : DATEY.has(k) ? 'date' : 'text');
    // имена: именительный, родительный, инициалы
    const fioOld = person ? oldC.name_full : oldC.director_fio, fioNew = person ? newC.name_full : newC.director_fio;
    const ttl = person ? 'ФИО работника' : 'ФИО руководителя';
    if (fioOld) {
      add(fioOld, fioNew, ttl, 'text');
      const go = oldC[person ? 'name_full_gen' : 'director_fio_gen'] || E.fioGen(fioOld), gn = fioNew ? (newC[person ? 'name_full_gen' : 'director_fio_gen'] || E.fioGen(fioNew)) : '';
      add(go, gn, ttl + ' (род. падеж)', 'text');
      add(E.shortFio(fioOld), fioNew ? E.shortFio(fioNew) : '', ttl + ' (инициалы)', 'text'); add(initialsFirst(fioOld), fioNew ? initialsFirst(fioNew) : '', ttl + ' (инициалы)', 'text');
    }
    if (!person) {
      if (oldC.director_pos) { add(oldC.director_pos, newC.director_pos, 'Должность руководителя', 'text'); add(oldC.director_pos_gen || E.posGen(oldC.director_pos), newC.director_pos ? (newC.director_pos_gen || E.posGen(newC.director_pos)) : '', 'Должность (род. падеж)', 'text'); }
      const core = s => { const m = /[«"]([^»"]{2,})[»"]/.exec(s || ''); return m ? m[1].trim() : ''; };
      const co = core(oldC.name_full) || core(oldC.name_short), cn = core(newC.name_full) || core(newC.name_short);
      if (co) add('«' + co + '»', cn ? '«' + cn + '»' : '', 'Название организации', 'text');
    }
    return list.sort((a, b) => b.from.length - a.from.length);
  }

  /**
   * Заменяет данные старых сторон (oldParties = {employer, worker}) на новые карточки (cards = {customer, executor}).
   * Возвращает { data, replaced, left, items }.
   */
  async function replaceOldData(buf, oldParties, cards, opts) {
    opts = Object.assign({ highlight: true }, opts || {});
    const zip = await JSZip.loadAsync(buf), items = []; let replaced = 0, left = 0;
    const newEmp = E.applyIpRules(Object.assign({}, cards.customer || {})), newWrk = Object.assign({}, cards.executor || {});
    const hasC = Object.keys(cards.customer || {}).length > 0, hasW = Object.keys(cards.executor || {}).length > 0;   // сторона без новых данных остаётся как есть
    const maps = (hasC ? buildMap(oldParties.employer || {}, newEmp, false) : []).concat(hasW ? buildMap(oldParties.worker || {}, newWrk, true) : []);
    const parts = ['word/document.xml'].concat(Object.keys(zip.files).filter(n => /^word\/(header|footer)\d*\.xml$/.test(n)));
    for (const part of parts) {
      const doc = I.parseXml(await zip.file(part).async('string'));
      for (const p of I.byNS(doc, I.NS_W, 'p')) {
        const text = I.mapParagraph(p).text; if (text.length < 2) continue;
        const nt = normText(text), hits = [];
        for (const m of maps) {
          const re = m.kind === 'num' ? numRe(m.from) : textRe(m.from); if (!re) continue; re.lastIndex = 0; let x;
          while ((x = re.exec(nt))) { hits.push({ s: x.index, e: x.index + x[0].length, m }); if (x[0].length === 0) re.lastIndex++; }
        }
        if (!hits.length) continue;
        hits.sort((a, b) => (b.e - b.s) - (a.e - a.s) || a.s - b.s);
        const take = []; for (const h of hits) if (!take.some(t => h.s < t.e && h.e > t.s)) take.push(h);
        take.sort((a, b) => b.s - a.s);
        for (const h of take) {
          const old = text.slice(h.s, h.e);
          if (h.m.to) { I.fillRange(p, h.s, h.e, h.m.to); replaced++; items.push({ kind: 'rep', name: h.m.label, from: old, to: h.m.to, rule: 'замена данных' }); }
          else { I.fillRange(p, h.s, h.e, '________'); if (opts.highlight) I.highlightRange(p, h.s, h.s + 8); left++; items.push({ kind: 'hl', key: h.m.label + ' — в новых реквизитах нет, было «' + old + '»' }); }
        }
      }
      zip.file(part, I.serialize(doc));
    }
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), replaced, left, items };
  }

  return { extractParties, replaceOldData, fioNom, posNom, buildMap };
});
