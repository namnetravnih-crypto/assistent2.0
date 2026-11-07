/* =====================================================================
   Ассистент — движок автозаполнения (DOCX / XLSX), версия 2
   Две стороны: Заказчик и Участник/Исполнитель. Всё в браузере.
   Структуру, шрифты и таблицы документа не меняет: правится только текст.
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('jszip'), require('@xmldom/xmldom'));
  else root.AssistantEngine = factory(root.JSZip, { DOMParser: root.DOMParser, XMLSerializer: root.XMLSerializer });
})(typeof self !== 'undefined' ? self : this, function (JSZip, X) {
  'use strict';
  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_XML = 'http://www.w3.org/XML/1998/namespace';

  /* ---------- XML ---------- */
  const parseXml = s => new X.DOMParser().parseFromString(s, 'application/xml');
  const serialize = d => { let s = new X.XMLSerializer().serializeToString(d); if (!/^<\?xml/.test(s)) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; return s; };
  const lname = n => n.localName || String(n.nodeName).replace(/^.*:/, '');
  const kids = (n, name) => { const o = []; for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && lname(c) === name) o.push(c); return o; };
  const firstKid = (n, name) => kids(n, name)[0] || null;
  const byNS = (n, ns, name) => Array.from(n.getElementsByTagNameNS(ns, name));

  /* ---------- Поля сторон ---------- */
  const FIELD_TITLES = {
    name_full: 'Полное наименование', name_short: 'Сокращённое наименование', inn: 'ИНН', kpp: 'КПП', ogrn: 'ОГРН', okpo: 'ОКПО', okved: 'ОКВЭД', oktmo: 'ОКТМО',
    reg_date: 'Дата регистрации', tax_reg_date: 'Дата постановки на налоговый учёт', addr_legal: 'Юридический адрес', addr_post: 'Почтовый адрес', addr_fact: 'Фактический адрес',
    phone: 'Телефон', fax: 'Факс', email: 'E-mail', site: 'Сайт', bank_name: 'Банк', bank_addr: 'Адрес банка', bank_inn: 'ИНН банка', bank_phone: 'Телефон банка',
    bik: 'БИК', rs: 'Расчётный счёт', ks: 'Корр. счёт', director_pos: 'Должность руководителя', director_fio: 'ФИО руководителя', director_inn: 'ИНН руководителя',
    director_pos_gen: 'Должность (род. падеж)', director_fio_gen: 'ФИО (род. падеж)', basis: 'Действует на основании', tax_system: 'Система налогообложения',
    accountant_fio: 'Главный бухгалтер', contact_fio: 'Контактное лицо', contact_pos: 'Должность контактного лица', contact_phone: 'Телефон контактного лица', contact_email: 'E-mail контактного лица',
    // данные работника (трудовой договор)
    birth_date: 'Дата рождения', birth_place: 'Место рождения', citizenship: 'Гражданство', passport: 'Паспорт (серия и номер)', passport_issued: 'Паспорт выдан (кем)',
    passport_date: 'Дата выдачи паспорта', passport_code: 'Код подразделения', snils: 'СНИЛС', salary: 'Оклад / заработная плата', start_date: 'Дата начала работы',
    department: 'Структурное подразделение', probation: 'Испытательный срок', tab_no: 'Табельный номер',
    end_date: 'Дата окончания срочного договора', term_reason: 'Основание срочного договора',
  };
  const T = k => FIELD_TITLES[k] || k;

  /* ---------- Разбор слов ---------- */
  const lc = s => String(s || '').toLowerCase().replace(/ё/g, 'е');
  const clean = s => String(s || '').replace(/[   ]/g, ' ').replace(/\s+/g, ' ').trim();
  const stripParens = s => { let o = s, p; do { p = o; o = o.replace(/\([^()]*\)/g, ' '); } while (o !== p); return o.replace(/\[[^\]]*\]/g, ' '); };

  // Работодатель — «сторона 1» (customer), Работник — «сторона 2» (executor): так трудовой договор ложится на те же два окна
  const CUSTOMER_RE = /заказчик[а-яёa-z0-9_]*|покупател[а-яёa-z0-9_]*|арендатор[а-яёa-z0-9_]*|генподрядчик[а-яёa-z0-9_]*|работодател[а-яёa-z0-9_]*|наймодател[а-яёa-z0-9_]*(?![а-яёa-z0-9_]*не)/;
  const EXEC_RE = /исполнител[а-яёa-z0-9_]*|подрядчик[а-яёa-z0-9_]*|поставщик[а-яёa-z0-9_]*|продавец[а-яёa-z0-9_]*|продавц[а-яёa-z0-9_]*|арендодател[а-яёa-z0-9_]*|субподрядчик[а-яёa-z0-9_]*|участник[а-яёa-z0-9_]*|претендент[а-яёa-z0-9_]*|заявител[а-яёa-z0-9_]*|работник[а-яёa-z0-9_]*/;
  function roleOfWord(w) {
    const t = lc(w);
    if (/^(?:суб)?подрядчик|^исполнител|^поставщик|^продав|^арендодател|^участник|^претендент|^заявител|^работник/.test(t)) return 'executor';
    if (/^заказчик|^покупател|^арендатор|^генподрядчик|^работодател/.test(t)) return 'customer';
    return null;
  }
  function roleInText(text) {           // последняя упомянутая роль в тексте
    const t = lc(text); let best = null, pos = -1;
    for (const m of t.matchAll(new RegExp('(?<![а-я])(?:' + CUSTOMER_RE.source + '|' + EXEC_RE.source + ')', 'g'))) if (m.index >= pos) { pos = m.index; best = roleOfWord(m[0]); }
    return best;
  }
  function headingRole(text) {          // «ЗАКАЗЧИК:», «От Исполнителя», «Сторона: Подрядчик»
    const t = lc(clean(text)).replace(/[«»"':;.,\-–—_()]/g, ' ').replace(/\s+/g, ' ').trim();
    const m = /^(?:(?:от|за)\s+)?(?:сторона\s+)?(заказчик[а-яёa-z0-9_]*|покупател[а-яёa-z0-9_]*|арендатор[а-яёa-z0-9_]*|генподрядчик[а-яёa-z0-9_]*|исполнител[а-яёa-z0-9_]*|подрядчик[а-яёa-z0-9_]*|поставщик[а-яёa-z0-9_]*|продав[а-яёa-z0-9_]*|арендодател[а-яёa-z0-9_]*|субподрядчик[а-яёa-z0-9_]*|работодател[а-яёa-z0-9_]*|работник[а-яёa-z0-9_]*)(?:\s+\d)?$/.exec(t);
    return m ? roleOfWord(m[1]) : null;
  }

  /* ---------- Деньги и числа ---------- */
  function parseNum(s) { const t = String(s == null ? '' : s).replace(/[\s ]/g, '').replace(',', '.'); if (!/^-?\d+(\.\d+)?$/.test(t)) return null; return parseFloat(t); }
  function fmtMoney(n) { const [a, b] = (Math.round(n * 100) / 100).toFixed(2).split('.'); return a.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + b; }
  const UN = [['', 'один', 'два', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять'], ['', 'одна', 'две', 'три', 'четыре', 'пять', 'шесть', 'семь', 'восемь', 'девять']];
  const TEN = ['десять', 'одиннадцать', 'двенадцать', 'тринадцать', 'четырнадцать', 'пятнадцать', 'шестнадцать', 'семнадцать', 'восемнадцать', 'девятнадцать'];
  const TENS = ['', '', 'двадцать', 'тридцать', 'сорок', 'пятьдесят', 'шестьдесят', 'семьдесят', 'восемьдесят', 'девяносто'];
  const HUN = ['', 'сто', 'двести', 'триста', 'четыреста', 'пятьсот', 'шестьсот', 'семьсот', 'восемьсот', 'девятьсот'];
  const plural = (n, f) => { const m = n % 100, d = n % 10; return (m > 10 && m < 20) ? f[2] : d === 1 ? f[0] : (d >= 2 && d <= 4) ? f[1] : f[2]; };
  function triad(n, fem) { const o = []; o.push(HUN[(n / 100) | 0]); const r = n % 100; if (r >= 10 && r < 20) o.push(TEN[r - 10]); else { o.push(TENS[(r / 10) | 0]); o.push(UN[fem ? 1 : 0][r % 10]); } return o.filter(Boolean).join(' '); }
  function numWords(n) {
    if (n === 0) return 'ноль';
    const parts = [[1e9, ['миллиард', 'миллиарда', 'миллиардов'], 0], [1e6, ['миллион', 'миллиона', 'миллионов'], 0], [1e3, ['тысяча', 'тысячи', 'тысяч'], 1]], out = [];
    for (const [d, f, fem] of parts) { const q = Math.floor(n / d) % 1000; if (q) out.push(triad(q, fem) + ' ' + plural(q, f)); }
    const r = n % 1000; if (r) out.push(triad(r, 0));
    return out.join(' ');
  }
  function rubWords(amount) {
    const rub = Math.floor(amount + 1e-9), kop = Math.round((amount - rub) * 100), w = numWords(rub);
    return w.charAt(0).toUpperCase() + w.slice(1) + ' ' + plural(rub, ['рубль', 'рубля', 'рублей']) + ' ' + String(kop).padStart(2, '0') + ' ' + plural(kop, ['копейка', 'копейки', 'копеек']);
  }

  /* ---------- Склонение ---------- */
  const POS_GEN = { 'генеральный директор': 'Генерального директора', 'исполнительный директор': 'Исполнительного директора', 'коммерческий директор': 'Коммерческого директора', 'директор': 'Директора', 'президент': 'Президента', 'руководитель': 'Руководителя', 'председатель': 'Председателя', 'управляющий': 'Управляющего', 'индивидуальный предприниматель': 'Индивидуального предпринимателя', 'и.о. директора': 'и.о. директора', 'главный врач': 'Главного врача', 'ректор': 'Ректора', 'начальник': 'Начальника' };
  function posGen(pos) {
    const t = clean(pos); if (!t) return '';
    const k = lc(t); if (POS_GEN[k]) return POS_GEN[k];
    let stop = false; return t.split(' ').map((w, i) => { const l = w.toLowerCase(); let r = l; if (/^(по|в|на|при|и|для|от|с|со|к)$/.test(l)) stop = true; if (stop) return i === 0 ? w : l; if (/(ый|ой)$/.test(l)) r = l.replace(/(ый|ой)$/, 'ого'); else if (/ий$/.test(l)) r = l.replace(/ий$/, 'ого'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'я'); else if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) r = l + 'а'; return i === 0 ? r.charAt(0).toUpperCase() + r.slice(1) : r; }).join(' ');
  }
  function declWord(w, kind, fem) {  // родительный падеж одного слова ФИО
    const l = w.toLowerCase(); let r = null;
    const keepCase = x => w[0] === w[0].toUpperCase() ? x.charAt(0).toUpperCase() + x.slice(1) : x;
    if (kind === 'patr') { if (/ич$/.test(l)) r = l + 'а'; else if (/на$/.test(l)) r = l.replace(/а$/, 'ы'); else return w; return keepCase(r); }
    if (kind === 'sur') {
      if (fem) { if (/(ова|ева|ёва|ина|ына)$/.test(l)) r = l.replace(/а$/, 'ой'); else if (/ая$/.test(l)) r = l.replace(/ая$/, 'ой'); else return w; return keepCase(r); }
      if (/(ов|ев|ёв|ин|ын)$/.test(l)) r = l + 'а'; else if (/(ский|цкий|ской|ный|ной|вой|тый)$/.test(l)) r = l.replace(/(ий|ой|ый)$/, 'ого'); else if (/ий$/.test(l)) r = l.replace(/ий$/, 'ия'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'я'); else if (/й$/.test(l)) r = l.replace(/й$/, 'я'); else if (/(ко|о|е|у|ю|и|ых|их|а|я)$/.test(l) && !/(ка|ша)$/.test(l)) return w; else if (/(ка|га|ха)$/.test(l)) r = l.replace(/а$/, 'и'); else if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) r = l + 'а'; else return w;
      return keepCase(r);
    }
    if (fem) { if (/ия$/.test(l)) r = l.replace(/я$/, 'и'); else if (/[гкх]а$/.test(l) || /ша$/.test(l) || /жа$/.test(l) || /ча$/.test(l) || /ща$/.test(l)) r = l.replace(/а$/, 'и'); else if (/а$/.test(l)) r = l.replace(/а$/, 'ы'); else if (/я$/.test(l)) r = l.replace(/я$/, 'и'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'и'); else return w; return keepCase(r); }
    if (/ей$/.test(l) && l.length > 4) r = l.replace(/ей$/, 'ея'); else if (/ий$/.test(l)) r = l.replace(/ий$/, 'ия'); else if (/й$/.test(l)) r = l.replace(/й$/, 'я'); else if (/ь$/.test(l)) r = l.replace(/ь$/, 'я'); else if (/(а|я)$/.test(l)) r = /[гкх]а$/.test(l) || /[шжчщ]а$/.test(l) ? l.replace(/а$/, 'и') : l.replace(/а$/, 'ы').replace(/я$/, 'и'); else if (/(ев|ёв|лев)$/.test(l) && false) r = l; else if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) r = (l === 'лев' ? 'льва' : l === 'павел' ? 'павла' : l + 'а'); else return w;
    return keepCase(r);
  }
  function fioGen(fio) {
    const p = clean(fio).split(' '); if (p.length < 2) return clean(fio);
    const patr = p[2] || '', fem = /(на|зы)$/i.test(patr) || (!patr && /(а|я)$/i.test(p[1]) && /(ва|на|ая)$/i.test(p[0]));
    return [declWord(p[0], 'sur', fem), declWord(p[1], 'name', fem), patr ? declWord(patr, 'patr', fem) : ''].filter(Boolean).concat(p.slice(3)).join(' ');
  }
  function shortFio(fio) {
    const f = clean(fio); if (!f) return ''; if (/\./.test(f)) return f;
    const p = f.split(' '); if (p.length < 2) return f;
    return p[0] + ' ' + p.slice(1, 3).map(x => x[0].toUpperCase() + '.').join('');
  }
  const MONTHS = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const parseIso = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? { y: +m[1], mo: +m[2], d: +m[3] } : null; };
  const dateLong = iso => { const d = parseIso(iso); return d ? `«${String(d.d).padStart(2, '0')}» ${MONTHS[d.mo - 1]} ${d.y} г.` : ''; };
  const dateShort = iso => { const d = parseIso(iso); return d ? `${String(d.d).padStart(2, '0')}.${String(d.mo).padStart(2, '0')}.${d.y}` : ''; };
  const isIp = c => /^(ип|индивидуальный предприниматель)\s/i.test(c.name_full || c.name_short || '');

  /* ---------- Индивидуальный предприниматель (ИП) ----------
     Если загружены реквизиты ИП:
       • полное наименование — «Индивидуальный предприниматель Фамилия Имя Отчество»;
       • руководитель (подписант) — сам предприниматель, если в карточке нет другого руководителя;
       • КПП у ИП нет. */
  const IP_NAME_RE = /^(?:ип|индивидуальн[а-яё]*\s+предпринимател[а-яё]*)\s+(.+)$/i;
  const FIO_RE = /^[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё.-]+(?:\s+[А-ЯЁ][а-яё.-]+)?$/;
  const looksFioLoose = v => FIO_RE.test(clean(v)) && !/\d/.test(v);
  function detectIp(card, allText) {
    const nm = clean(card.name_full || card.name_short || '');
    if (IP_NAME_RE.test(nm)) return true;
    if (/^\d{15}$/.test(String(card.ogrn || '').replace(/\s/g, ''))) return true;                  // ОГРНИП — 15 цифр
    const t = String(allText || '');
    if (/огрнип/i.test(t)) return true;
    if (/индивидуальн[а-яё]*\s+предпринимател/i.test(t) && !/(?:^|\s)(?:ООО|АО|ПАО|НАО|ЗАО|ОАО)(?:\s|«|")|общество|акционерн/i.test(nm)) return true;
    return false;
  }
  /** Приводит карточку ИП к правилам. Дописывает только недостающее, но полное наименование всегда пишет по формуле. */
  function applyIpRules(card, allText) {
    if (!card || !detectIp(card, allText)) return card;
    const nameSrc = clean(card.name_full || card.name_short || ''), m = IP_NAME_RE.exec(nameSrc);
    let fio = '';
    if (m && looksFioLoose(m[1])) fio = clean(m[1]);
    else if (card.director_fio && looksFioLoose(card.director_fio)) fio = clean(card.director_fio);
    else if (looksFioLoose(nameSrc)) fio = nameSrc;
    else { const mm = String(allText || '').match(/(?:ИП|Индивидуальный\s+предприниматель)\s+([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2})/); if (mm) fio = clean(mm[1]); }
    if (!fio) return card;
    // если в наименовании только инициалы, а у руководителя полное ФИО — берём полное
    const dFio = clean(card.director_fio || '');
    if (/\./.test(fio) && dFio && !/\./.test(dFio) && dFio.split(' ')[0].toLowerCase() === fio.split(' ')[0].toLowerCase()) fio = dFio;
    if (!card.director_fio) card.director_fio = fio;                       // руководитель — сам предприниматель
    if (!card.director_pos) card.director_pos = 'Индивидуальный предприниматель';
    card.name_full = 'Индивидуальный предприниматель ' + fio;
    if (!card.name_short || !/^ип\s/i.test(clean(card.name_short))) card.name_short = 'ИП ' + fio;
    delete card.kpp;
    if (!card.director_pos_gen) card.director_pos_gen = posGen(card.director_pos);
    if (!card.director_fio_gen) card.director_fio_gen = fioGen(card.director_fio);
    if (!card.basis || /^устава$/i.test(clean(card.basis))) card.basis = 'свидетельства о государственной регистрации' + (card.ogrn ? ' (ОГРНИП ' + card.ogrn + ')' : '');
    return card;
  }

  /* пол по ФИО: женский — отчество на «-вна / -чна», мужской — «-вич / -ич / -оглы» */
  function genderOf(fio) {
    const p = clean(fio).split(' '), patr = (p[2] || '').toLowerCase(), name = (p[1] || '').toLowerCase(), sur = (p[0] || '').toLowerCase();
    if (/(вна|чна|кызы|гызы)$/.test(patr)) return 'f';
    if (/(вич|ич|оглы|улы)$/.test(patr)) return 'm';
    if (/(ова|ева|ёва|ина|ына|ская|цкая|ная)$/.test(sur)) return 'f';
    if (/(ов|ев|ёв|ин|ын|ский|цкий|ной|ый)$/.test(sur)) return 'm';
    if (/[ая]$/.test(name) && !/^(никита|илья|кузьма|лука|фома|савва|данила|гаврила)$/.test(name)) return 'f';
    return 'm';
  }
  const ORG_NEUTER = /^(?:общество|акционерное|публичное|непубличное|товарищество|предприятие|учреждение|объединение|ооо|ао|пао|нао|зао|оао|муп|гуп|фгуп|мбу|гбу|ано)(?![а-яё])/i;
  const ORG_MASC = /^(?:фонд|союз|центр|институт|университет|комитет|банк|кооператив|холдинг|колледж|лицей|техникум|комбинат)(?![а-яё])/i;
  const ORG_FEM = /^(?:компания|корпорация|ассоциация|организация|фирма|академия|школа|больница|клиника)(?![а-яё])/i;
  /** вид стороны: ИП, физическое лицо (работник), организация */
  function kindOf(c) {
    const nm = clean(c.name_full || c.name_short || '');
    if (!nm) return 'org';
    if (IP_NAME_RE.test(nm)) return 'ip';
    if (ORG_NEUTER.test(nm) || ORG_MASC.test(nm) || ORG_FEM.test(nm) || /[«"“]/.test(nm)) return 'org';
    if (looksFioLoose(nm)) return 'person';
    return 'org';
  }

  /* =====================================================================
     Распознавание подписи поля
     ===================================================================== */
  const NLB = '(?<![а-яa-z0-9])', NRB = '(?![а-яa-z])';
  const ORGW = '(?:организации|юридического\\s+лица|участника(?:\\s+закупки)?|контрагента|претендента|заявителя|поставщика|подрядчика|исполнителя|заказчика|покупателя|компании|предприятия|партнера|субъекта|стороны|арендатора|арендодателя)';
  const RULES = [
    // данные работника — идут первыми, чтобы «Паспорт выдан», «Дата рождения», «Должностной оклад» не приняли за другие поля
    ['passport_issued', /паспорт[а-яёa-z0-9_]*\s*(?:рф\s*)?(?:выдан[а-яёa-z0-9_]*|кем)|кем\s+выдан[а-яёa-z0-9_]*|выдан[а-яёa-z0-9_]*\s+кем|орган[а-яёa-z0-9_]*,?\s+выдав[а-яёa-z0-9_]*/g],
    ['passport_date', /дата\s+выдачи(?:\s+паспорта)?|когда\s+выдан[а-яёa-z0-9_]*/g],
    ['passport_code', /код\s+подразделения/g],
    ['birth_place', /место\s+рождения/g],
    ['birth_date', /дата\s+рождения|родил[а-яё]+(?:\s+в)?|год\s+рождения/g],
    ['snils', /снилс|страхов[а-яёa-z0-9_]*\s+свидетельств[а-яёa-z0-9_]*(?:\s+пенсионного)?/g],
    ['passport', /серия\s+и\s+номер|серия[,\s]+номер|паспорт[а-яёa-z0-9_]*(?:\s+(?:рф|гражданина(?:\s+рф)?))?/g],
    ['salary', /(?:должностн[а-яёa-z0-9_]*\s+)?оклад[а-яёa-z0-9_]*|заработн[а-яёa-z0-9_]*\s+плат[а-яёa-z0-9_]*|тарифн[а-яёa-z0-9_]*\s+ставк[а-яёa-z0-9_]*|размер\s+оплаты/g],
    ['start_date', /дата\s+(?:начала\s+работы|приема|приёма|трудоустройства|вступления)|приступ[а-яёa-z0-9_]*\s+к\s+работе/g],
    ['probation', /испытательн[а-яёa-z0-9_]*\s+срок|срок\s+испытания/g],
    ['tab_no', /табельн[а-яёa-z0-9_]*\s+номер/g],
    ['citizenship', /гражданств[а-яёa-z0-9_]*/g],
    ['department', /структурн[а-яёa-z0-9_]*\s+подразделени[а-яёa-z0-9_]*|подразделение(?![а-я])|(?<![а-я])отдел(?![а-я])/g],
    ['inn_kpp', new RegExp(NLB + 'инн\\s*[/,\\\\и]\\s*кпп' + NRB, 'g')],
    ['name_short', /сокращ[а-яёa-z0-9_]*\s+(?:фирменн[а-яёa-z0-9_]*\s+)?(?:наименован[а-яёa-z0-9_]*|назван[а-яёa-z0-9_]*)/g],
    ['name_full', new RegExp('(?:полн[а-яёa-z0-9_]*\\s+)?(?:фирменн[а-яёa-z0-9_]*\\s+)?наименован[а-яёa-z0-9_]*\\s+' + ORGW + '|полн[а-яёa-z0-9_]*\\s+(?:фирменн[а-яёa-z0-9_]*\\s+)?наименован[а-яёa-z0-9_]*|фирменн[а-яёa-z0-9_]*\\s+наименован[а-яёa-z0-9_]*|название\\s+(?:организации|компании)|участник\\s+(?:процедуры\\s+)?закупки|участник\\s+конкурса|лидер\\s+коллективного\\s+участника', 'g')],
    ['inn', new RegExp(NLB + 'инн' + NRB, 'g')],
    ['kpp', new RegExp(NLB + 'кпп' + NRB, 'g')],
    ['ogrn', new RegExp(NLB + 'огрн(?:ип)?' + NRB + '|основной\\s+государственный\\s+регистрационный\\s+номер', 'g')],
    ['okpo', new RegExp(NLB + 'окпо' + NRB, 'g')],
    ['oktmo', new RegExp(NLB + 'октмо' + NRB, 'g')],
    ['okved', new RegExp(NLB + 'оквэд[а-яёa-z0-9_]*', 'g')],
    ['tax_reg_date', /дата\s+постановки\s+на\s+(?:налоговый\s+)?учет/g],
    ['reg_date', /дата\s+(?:государственной\s+)?(?:регистрации|создания|внесения|присвоения)/g],
    ['addr_post', /почтов[а-яёa-z0-9_]*\s+адрес|(?<!электронн[а-яёa-z0-9_]*\s)адрес\s+для\s+(?:корреспонденции|направления|почтовой)/g],
    ['addr_fact', /фактическ[а-яёa-z0-9_]*\s+(?:адрес|местонахождение|место)|адрес\s+фактическ[а-яёa-z0-9_]*|место\s+осуществления\s+деятельности|складск[а-яёa-z0-9_]*\s+помещени/g],
    ['addr_legal', /юридическ[а-яёa-z0-9_]*\s+адрес|адрес\s*(?:места\s+нахождения|местонахождения|регистрации|юридическ[а-яёa-z0-9_]*)|место\s*нахождени[а-яёa-z0-9_]*|местонахождени[а-яёa-z0-9_]*|место\s+жительств[а-яёa-z0-9_]*|место\s+регистрац[а-яёa-z0-9_]*/g],
    ['email', new RegExp('e-?mail|' + NLB + 'эл[а-яёa-z0-9_]*\\.?\\s*почт[а-яёa-z0-9_]*|электронн[а-яёa-z0-9_]+\\s+(?:почт[а-яёa-z0-9_]*|адрес)|адрес[а-яёa-z0-9_]*\\s+электронной\\s+почты', 'g')],
    ['site', /сайт[а-яёa-z0-9_]*|веб-?[а-яёa-z0-9_]*|интернет-?[а-яёa-z0-9_]*\s*сайт/g],
    ['addr_any', new RegExp(NLB + 'адрес(?:а|у|ом)?' + NRB, 'g')],
    ['phone', new RegExp('телефон\\s*[/,]\\s*факс|телефон[а-яёa-z0-9_]*|' + NLB + 'тел\\.?' + NRB + '|контактн[а-яёa-z0-9_]*\\s+номер', 'g')],
    ['fax', /факс/g],
    ['bank_name', new RegExp('наименование\\s+банка|обслуживающ[а-яёa-z0-9_]+\\s+банк|' + NLB + 'банк' + NRB, 'g')],
    ['bik', new RegExp(NLB + 'бик' + NRB, 'g')],
    ['ks', new RegExp(NLB + 'к\\s*[/\\\\]\\s*с' + NRB + '|кор[а-яёa-z0-9_]*\\.?\\s*сч[её]т[а-яёa-z0-9_]*|корреспондентск[а-яёa-z0-9_]+', 'g')],
    ['rs', new RegExp(NLB + 'р\\s*[/\\\\]\\s*с' + NRB + '|р\\.\\s*с\\.|расч[её]тн[а-яёa-z0-9_]*\\s+сч[её]т|рас\\.?\\s*сч[её]т|номер\\s+сч[её]та', 'g')],
    ['tax_system', /систем[а-яёa-z0-9_]+\s+налогообложени[а-яёa-z0-9_]+|налогообложени[а-яёa-z0-9_]+/g],
    ['vat_payer', /плательщик[а-яёa-z0-9_]*\s+ндс|является\s+плательщиком/g],
    ['vat_text', new RegExp(NLB + 'ндс' + NRB, 'g')],
    ['signer_pos_fio', /должност[а-яёa-z0-9_]*\s*[,/и]\s*ф\.?\s*и\.?\s*о\.?|ф\.?\s*и\.?\s*о\.?\s*[,/и]\s*должност[а-яёa-z0-9_]*/g],
    ['signer_fio', new RegExp('ф\\.\\s*и\\.\\s*о\\.?|' + NLB + 'фио' + NRB + '|фамилия[, ]+имя[, ]+отчество', 'g')],
    ['signer_pos', new RegExp(NLB + 'должност[а-яёa-z0-9_]*', 'g')],
    ['inperson', /в\s+лице/g],
    ['basis', /действующ[а-яёa-z0-9_]+\s+на\s+основании|на\s+основании\s+(?:устава|доверенности)/g],
    ['doc_date', /дата\s+(?:документа|составления|заявки|подписания|договора)|(?<![а-я])дата(?![а-я])/g],
    ['branches', /филиал[а-яёa-z0-9_]*/g],
    ['holding', /холдинг[а-яёa-z0-9_]*|концерн[а-яёa-z0-9_]*|финансово-промышленн[а-яёa-z0-9_]*/g],
  ];
  const BARE_NAME = /^(?:наименование|организация|участник|исполнитель|подрядчик|субподрядчик|генподрядчик|поставщик|претендент|заявитель|контрагент|компания|предприятие|юридическое\s+лицо|фирма|заказчик|покупатель|работодатель|работник)[\s:,.*№\-–—]*$/;
  const BARE_BOSS = /^(?:руководитель(?:\s+организации)?|подписант|уполномоченное\s+лицо|представитель|генеральный\s+директор|директор|президент|управляющий)[\s:,.*\-–—]*$/;
  const CODE_KEYS = ['inn', 'ogrn', 'kpp', 'okpo', 'okved', 'oktmo'];

  function scanOrder(t, keys) {          // ключи, упомянутые в строке, в порядке появления
    const found = [];
    for (const [k, re] of RULES) if (keys.includes(k)) { re.lastIndex = 0; const m = re.exec(t); if (m) found.push([m.index, k]); }
    return found.sort((a, b) => a[0] - b[0]).map(x => x[1]);
  }
  const uniq = a => a.filter((x, i) => a.indexOf(x) === i);

  /* classify(label) → { items:[ключи], role, bare } | null.  opt.tail — подпись взята из хвоста длинного абзаца */
  function classify(raw, opt) {
    opt = opt || {};
    const original = lc(clean(raw)), t = clean(lc(stripParens(clean(raw)))).replace(/[:：*\s]+$/, '');
    if (!t && !original) return null;
    if (/подпис|(?<![а-я])м\.?\s?п\.?(?![а-я])/.test(t) && !/[а-я]{4,}\s+подпис/.test(t) && t.length < 40) return { items: [], skip: true };
    const role = roleInText(t.slice(-60));
    const out = items => ({ items, role, bare: false });

    if (!opt.tail) {
      // банковские реквизиты одной строкой
      if (/банковск[а-яёa-z0-9_]*\s+реквизит|реквизит[а-яёa-z0-9_]*\s+банка|платежн[а-яёa-z0-9_]*\s+реквизит/.test(t)) {
        const o = original, it = [];
        if (/наименовани[а-яёa-z0-9_]*(?:\s+и\s+адрес)?\s+банка|(?<![а-я])банк(?![а-я])/.test(o)) it.push('bank_name');
        if (/адрес[а-яёa-z0-9_]*\s+банка/.test(o)) it.push('bank_addr');
        if (/р\s*\/\s*с|расч[а-яёa-z0-9_]*\s*сч|номер\s+сч/.test(o)) it.push('rs');
        if (/к\s*\/\s*с|кор[а-яёa-z0-9_]*\.?\s*сч/.test(o)) it.push('ks');
        if (/(?<![а-я])бик(?![а-я])/.test(o)) it.push('bik');
        if (/инн\s+банка/.test(o)) it.push('bank_inn');
        if (/телефон[а-яёa-z0-9_]*\s+банка/.test(o)) it.push('bank_phone');
        const order = ['bank_name', 'bank_addr', 'rs', 'ks', 'bik', 'bank_inn', 'bank_phone'];
        return out(it.length ? order.filter(k => it.includes(k)) : ['bank_name', 'bik', 'rs', 'ks']);
      }
      const QUESTION = /наличи|судим|задолженн|решени|сведени[а-яёa-z0-9_]*\s+о|отсутств|не\s+находит|подтвержда|предусмотрен/.test(t);
      if (QUESTION) { /* вопрос анкеты — не поле реквизитов */ }
      else if (/бухгалтер/.test(t)) return out(['accountant_fio']);
      if (!QUESTION && /контактн[а-яёa-z0-9_]*\s+лиц|ответственн[а-яёa-z0-9_]*\s+лиц/.test(t)) {
        const it = [/должност/.test(original) ? 'contact_fio_pos' : 'contact_fio'];
        if (/телефон/.test(original)) it.push('contact_phone');
        if (/почт|e-?mail/.test(original)) it.push('contact_email');
        return out(it);
      }
      const DIR = /руковод|директор|единолич|подписант|уполномоченн[а-яёa-z0-9_]*\s+лиц|представител|первое\s+лицо/.test(t);
      if (DIR && !QUESTION && t.length <= 150 && !/наименовани|адрес|реквизит|телефон\s*$/.test(t)) {
        if (/инн/.test(t)) return out(['director_inn']);
        const FIOW = /ф\.?\s?и\.?\s?о|фамили|имя|отчеств/.test(t);
        const it = [(/должност/.test(original) || !FIOW) ? 'director_pos_fio' : 'director_fio'];
        if (/телефон/.test(original)) it.push('director_phone');
        return out(it);
      }
      // составные строки: «ИНН; ОГРН; КПП …», «Фактический; Почтовый адрес», «сайт, e-mail»
      const codes = scanOrder(t, CODE_KEYS);
      if (codes.length >= 2 && !/инн\s*[/,\\и]\s*кпп/.test(t)) return out(codes);
      const adr = [];
      if (/юридическ/.test(t)) adr.push(['addr_legal', t.search(/юридическ/)]);
      if (/фактическ/.test(t)) adr.push(['addr_fact', t.search(/фактическ/)]);
      if (/почтов/.test(t)) adr.push(['addr_post', t.search(/почтов/)]);
      if (adr.length >= 2) return out(adr.sort((a, b) => a[1] - b[1]).map(x => x[0]));
      const con = scanOrder(t, ['phone', 'email', 'site', 'fax']);
      if (con.length >= 2 && !/адрес\s+электронной\s+почты/.test(t.replace(/(?:сайт[а-яёa-z0-9_]*|веб[а-яёa-z0-9_]*)[^;]*/, ''))) return out(con);
      if (con.length >= 2) return out(con);
      if (!/^(?:дата|номер|№)/.test(t)) {
        if (BARE_NAME.test(t) && !opt.noBare) return { items: ['name_full'], role: roleInText(t), bare: true };
        if (BARE_BOSS.test(t) && !opt.noBare) return { items: ['director_pos_fio'], role: null, bare: true, boss: true };
      }
    }
    // одиночное поле: в обычном режиме — первое по приоритету, в «хвостовом» — самое близкое к концу
    let best = null;
    for (const [key, re] of RULES) {
      re.lastIndex = 0;
      for (const m of t.matchAll(re)) {
        const end = m.index + m[0].length, len = m[0].length;
        if (opt.tail) { if (!best || end > best.end || (end === best.end && len > best.len)) best = { key, end, len }; }
        else if (!best) best = { key, end, len };
      }
      if (best && !opt.tail) break;
    }
    if (!best) return null;
    if (!opt.tail && best.key === 'signer_pos' && /наименован|параметр|документ|продукц/.test(t)) return null;
    return out([best.key]);
  }

  /* =====================================================================
     Разбор карты стороны (текст → поля)
     ===================================================================== */
  const OPF = [[/Общество с ограниченной ответственностью/gi, 'ООО'], [/Публичное акционерное общество/gi, 'ПАО'], [/Акционерное общество/gi, 'АО'], [/Индивидуальный предприниматель/gi, 'ИП']];
  const POS_RE = /^(генеральный\s+директор|исполнительный\s+директор|коммерческий\s+директор|директор|президент|руководитель|председатель|управляющий|индивидуальный\s+предприниматель|и\.\s*о\.\s*директора|главный\s+врач|ректор|начальник)\s+(.+)$/i;
  const splitPosFio = v => { const m = POS_RE.exec(clean(v)); return m ? { pos: m[1].replace(/^./, c => c.toUpperCase()), fio: m[2].trim() } : null; };
  const looksFio = v => /^[А-ЯЁ][а-яё-]+\s+[А-ЯЁ][а-яё.-]+(\s+[А-ЯЁ][а-яё.-]+)?$/.test(clean(v));

  function parseCardText(text) {
    const card = {};
    const set = (k, v) => { v = clean(v); if (v && !card[k]) card[k] = v; };
    const lines = String(text || '').split(/\r?\n/).map(s => clean(s)).filter(Boolean);
    const pairs = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      if (line.includes(' | ')) {
        const c = line.split(' | ').map(s => s.trim());
        if (c.length === 2) pairs.push([c[0], c[1]]);
        else if (c.length === 4) { pairs.push([c[0], c[1]]); pairs.push([c[2], c[3]]); }
        else if (c.length > 2) { pairs.push([c[0], c.slice(1).filter(Boolean).join(' ')]); pairs.push([c[c.length - 2], c[c.length - 1]]); }
        continue;
      }
      const ci = line.indexOf(':');
      if (ci > 0 && ci <= 80) {
        const lab = line.slice(0, ci), val = line.slice(ci + 1).trim();
        if (val) { pairs.push([lab, val]); continue; }
        if (lines[i + 1] && !classify(lines[i + 1], { noBare: true })) pairs.push([lab, lines[i + 1]]);
        continue;
      }
      if (line.length <= 60 && classify(line) && lines[i + 1] && !lines[i + 1].includes(' | ') && !classify(lines[i + 1])) pairs.push([line, lines[i + 1]]);
    }
    const digits = (v, re) => { const m = re.exec(String(v).replace(/[\s-]/g, '')); return m ? m[0] : ''; };
    let bankSeen = false;                       // после строки «Банк» ИНН и КПП относятся к банку, а не к организации
    for (const [label, value] of pairs) {
      if (!value) continue;
      const r = classify(label); if (!r || r.skip || !r.items.length) continue;
      if (r.items.includes('bank_name')) bankSeen = true;
      if (bankSeen && r.items.length === 1 && ['inn', 'kpp', 'inn_kpp', 'ogrn', 'okpo'].includes(r.items[0])) {
        if (r.items[0] === 'inn') set('bank_inn', digits(value, /\d{10,12}/));
        else if (r.items[0] === 'inn_kpp') set('bank_inn', (String(value).match(/\d{10,12}/) || [''])[0]);
        continue;
      }
      if (r.items.includes('bank_name') && r.items.length > 1) { set('bank_name', String(value).split(/,\s*(?:р\/с|к\/с|бик|расч|корр|р\.\s*с)/i)[0].replace(/^банк[:\s]+/i, '')); continue; }
      if (r.items.length > 1) {          // составная строка карты: разложим числа по длине
        if (r.items.every(k => CODE_KEYS.includes(k))) for (const n of (String(value).match(/\d{8,15}/g) || [])) {
          if (n.length === 9) set('kpp', n); else if (n.length === 10 || n.length === 12) { if (!card.inn) set('inn', n); else set('okpo', n); } else if (n.length === 13 || n.length === 15) set('ogrn', n); else set('okpo', n);
        }
        continue;
      }
      const k = r.items[0];
      switch (k) {
        case 'inn_kpp': { const m = String(value).match(/\d{9,12}/g) || []; set('inn', m.find(x => x.length >= 10)); set('kpp', m.find(x => x.length === 9)); break; }
        case 'inn': set('inn', digits(value, /\d{10,12}/)); break;
        case 'kpp': set('kpp', digits(value, /\d{9}/)); break;
        case 'ogrn': set('ogrn', digits(value, /\d{13,15}/)); break;
        case 'okpo': set('okpo', digits(value, /\d{8,10}/)); break;
        case 'bik': set('bik', digits(value, /\d{9}/)); break;
        case 'rs': set('rs', digits(value, /\d{20}/)); break;
        case 'ks': set('ks', digits(value, /\d{20}/)); break;
        case 'director_inn': set('director_inn', digits(value, /\d{12}/)); break;
        case 'signer_fio': set('director_fio', value); break;
        case 'signer_pos': set('director_pos', value); break;
        case 'director_fio': set('director_fio', value); break;
        case 'director_pos_fio': case 'signer_pos_fio': {
          const sp = splitPosFio(value);
          if (sp) { set('director_pos', sp.pos); set('director_fio', sp.fio); }
          else if (looksFio(value)) { if (BARE_BOSS.test(lc(clean(label))) && !/^(руководитель|подписант|уполномоч|представитель)/i.test(label)) set('director_pos', label.replace(/[:\s]+$/, '').replace(/^./, c => c.toUpperCase())); set('director_fio', value); }
          break;
        }
        case 'addr_any': set('addr_legal', value); break;
        case 'contact_fio_pos': case 'contact_fio': { const sp = splitPosFio(value); if (sp) { set('contact_pos', sp.pos); set('contact_fio', sp.fio); } else set('contact_fio', value); break; }
        case 'inperson': case 'doc_date': case 'vat_text': case 'vat_payer': break;
        default: if (FIELD_TITLES[k] || k === 'branches' || k === 'holding') set(k, value);
      }
    }
    const all = lines.join('\n');
    const fb = (k, re, g) => { if (!card[k]) { const m = all.match(re); if (m) set(k, m[g || 1]); } };
    const bankAt = all.search(/(?:^|\n)\s*(?:банковские реквизиты|банк\b)/i), head = bankAt > 0 ? all.slice(0, bankAt) : all;   // ИНН и КПП организации ищем выше банковского блока
    const fbh = (k, re) => { if (!card[k]) { const m = head.match(re); if (m) set(k, m[1]); } };
    fbh('inn', /ИНН[^\d]{0,15}(?<!\d)(\d{12}|\d{10})(?!\d)/i);
    fbh('kpp', /КПП[^\d]{0,15}(?<!\d)(\d{9})(?!\d)/i);
    fb('ogrn', /ОГРН(?:ИП)?[^\d]{0,15}(?<!\d)(\d{15}|\d{13})(?!\d)/i);
    fb('bik', /БИК[^\d]{0,15}(?<!\d)(\d{9})(?!\d)/i);
    for (const m of all.matchAll(/(?<![\d])(\d{20})(?!\d)/g)) { if (m[1].startsWith('301')) set('ks', m[1]); else set('rs', m[1]); }
    fb('email', /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/, 0);
    if (!card.phone) {          // телефон — только отдельное число, а не кусок длинного номера (счёта, ИНН, ОГРН)
      const m = all.match(/(?<![\d])(?:\+7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}(?![\d])/);
      if (m) set('phone', m[0]);
    }
    if (!card.name_full) {
      const m = all.match(/(?:Общество с ограниченной ответственностью|Акционерное общество|Публичное акционерное общество)\s*[«"“][^»"”\n]+[»"”]|(?:ООО|АО|ПАО|ЗАО|ОАО|НАО|АНО|МУП|ГУП|ФГУП)\s*[«"“][^»"”\n]+[»"”]|Индивидуальный предприниматель\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2}|ИП\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2}/);
      if (m) set('name_full', m[0]);
    }
    if (!card.director_fio) {
      const m = all.match(/(Генеральный директор|Исполнительный директор|Директор|Руководитель|Президент)[:\s–—-]+([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2})/);
      if (m) { set('director_pos', m[1]); set('director_fio', m[2]); }
    }
    return finishCard(card, all);
  }

  function finishCard(card, allText) {
    const set = (k, v) => { v = clean(v); if (v && !card[k]) card[k] = v; };
    if (!card.name_short && card.name_full) { let s = card.name_full; for (const [re, to] of OPF) s = s.replace(re, to); set('name_short', s); }
    if (!card.name_full && card.name_short) set('name_full', card.name_short);
    if (card.director_pos && !card.director_pos_gen) set('director_pos_gen', posGen(card.director_pos));
    if (card.director_fio && !card.director_fio_gen) set('director_fio_gen', fioGen(card.director_fio));
    if (!card.basis && card.director_fio) set('basis', 'Устава');
    // ИП: наименование «Индивидуальный предприниматель ФИО», руководитель — сам предприниматель, КПП нет
    if (detectIp(card, allText)) { if (card.basis === 'Устава') delete card.basis; applyIpRules(card, allText); }
    return card;
  }

  /* =====================================================================
     Контекст заполнения: стороны, документ, журнал
     ===================================================================== */
  const ALLOWED_MIN = new Set(['name_full', 'name_short', 'signer_fio', 'signer_fio_short', 'signer_pos', 'signer_pos_fio', 'director_pos_fio', 'director_fio', 'inperson']);
  const SHORT = { inn: 'ИНН', kpp: 'КПП', okpo: 'ОКПО', okved: 'ОКВЭД2', oktmo: 'ОКТМО', addr_fact: 'Фактический', addr_post: 'Почтовый', addr_legal: 'Юридический', phone: 'тел.', email: 'e-mail', site: 'сайт', fax: 'факс', rs: 'р/с', ks: 'к/с', bik: 'БИК', bank_inn: 'ИНН банка', bank_addr: 'адрес банка', bank_phone: 'тел. банка', contact_phone: 'тел.', contact_email: 'e-mail', director_phone: 'тел.' };

  function makeCtx(parties, doc, opts) {
    opts = Object.assign({ mode: 'full', highlight: true, fallbackAddr: true, fallbackContact: true, docType: 'auto', defaultRole: 'executor' }, opts || {});
    doc = Object.assign({ date: '', number: '', specNo: '', city: '', amount: '', vat: '22', vatMode: 'incl' }, doc || {});
    // реквизиты ИП приводим к правилам (полное наименование, руководитель) — работает и для карточек, введённых вручную
    const normParty = c => applyIpRules(Object.assign({}, c || {}));
    parties = { customer: normParty(parties && parties.customer), executor: normParty(parties && parties.executor) };
    const ctx = {
      parties, doc, opts, handled: new Set(), skipP: new WeakSet(), skipT: new WeakSet(), items: [], stats: { filled: 0, highlighted: 0 }, role: null, contract: false,
      allowed(key) { return opts.mode !== 'min' || ALLOWED_MIN.has(key); },
      amountNum() { return parseNum(doc.amount); },
      vatRate() { return parseFloat(doc.vat) || 0; },
      single(key, role, arg) {
        const c = parties[role] || {}; const d = doc;
        const dirFio = c.director_fio || '';
        switch (key) {
          case 'name_full': return c.name_full || c.name_short || '';
          case 'name_short': return c.name_short || c.name_full || '';
          case 'name_addr': return [c.name_full || c.name_short, c.addr_legal].filter(Boolean).join(', ');
          // строка стороны в КС-2 / КС-3: «наименование, ИНН, КПП, адрес, тел.»
          case 'party_line': return c.name_full || c.name_short ? [c.name_full || c.name_short, c.inn ? 'ИНН ' + c.inn : '', c.kpp ? 'КПП ' + c.kpp : '', c.addr_legal, c.phone ? 'тел. ' + c.phone : ''].filter(Boolean).join(', ') : '';
          case 'inn_kpp': return c.inn ? (c.kpp ? c.inn + '/' + c.kpp : c.inn) : '';
          case 'ogrn': return c.ogrn || '';
          case 'addr_any': case 'addr_legal': return c.addr_legal || '';
          case 'addr_post': return c.addr_post || (opts.fallbackAddr ? c.addr_legal : '') || '';
          case 'addr_fact': return c.addr_fact || (opts.fallbackAddr ? c.addr_legal : '') || '';
          case 'signer_fio': case 'director_fio': return dirFio;
          case 'signer_fio_short': return shortFio(dirFio);
          case 'signer_pos': case 'director_pos': return c.director_pos || '';
          case 'signer_pos_fio': case 'director_pos_fio': case 'sig_title': return dirFio ? [c.director_pos, dirFio].filter(Boolean).join(' ') : '';
          case 'director_inn': return c.director_inn || (isIp(c) ? 'Отсутствует' : '');
          case 'director_phone': return c.phone || '';
          case 'inperson': return (c.director_pos_gen && c.director_fio_gen) ? c.director_pos_gen + ' ' + c.director_fio_gen : '';
          case 'contact_fio_pos': { const f = c.contact_fio || (opts.fallbackContact ? dirFio : ''); const p = c.contact_fio ? c.contact_pos : (opts.fallbackContact ? c.director_pos : ''); return f ? [p, f].filter(Boolean).join(' ') : ''; }
          case 'contact_fio': return c.contact_fio || (opts.fallbackContact ? dirFio : '');
          case 'contact_phone': return c.contact_phone || (opts.fallbackContact ? c.phone : '') || '';
          case 'contact_email': return c.contact_email || (opts.fallbackContact ? c.email : '') || '';
          case 'vat_payer': return c.tax_system ? (/осн|с\s*ндс/i.test(c.tax_system) ? 'Да' : 'Нет') : '';
          case 'vat_text': return c.tax_system ? (/осн|с\s*ндс/i.test(c.tax_system) ? 'Да, с НДС по ставке ' + (d.vat || 22) + ' %' : 'Без НДС') : '';
          case 'doc_date': return dateShort(d.date);
          case 'doc_no': return d.number || '';
          case 'spec_no': return d.specNo || '';
          case 'city': return d.city || '';
          case 'amount_num': { const n = parseNum(d.amount); return n == null ? '' : fmtMoney(n); }
          case 'amount_words': { const n = parseNum(d.amount); return n == null ? '' : rubWords(n); }
          case 'vat_amount': { const n = parseNum(d.amount), r = parseFloat(d.vat) || 0; if (n == null || !r) return ''; return fmtMoney(d.vatMode === 'excl' ? n * r / 100 : n * r / (100 + r)); }
          case 'amount': { const n = parseNum(d.amount); return n == null ? '' : fmtMoney(n) + ' руб. (' + rubWords(n) + ')'; }
          // ---- составные значения для готовых бланков ----
          case 'named': { const k = kindOf(c); if (k === 'org') return ORG_MASC.test(clean(c.name_full || '')) ? 'именуемый' : ORG_FEM.test(clean(c.name_full || '')) ? 'именуемая' : 'именуемое'; return genderOf(c.director_fio || c.name_full || '') === 'f' ? 'именуемая' : 'именуемый'; }
          case 'acting': return genderOf(dirFio) === 'f' ? 'действующей' : 'действующего';
          case 'fio_gen': { const src = kindOf(c) === 'person' ? (c.name_full || c.director_fio) : (c.director_fio || c.name_full); return src ? fioGen(src) : ''; }
          case 'fio_short': return shortFio(kindOf(c) === 'person' ? (c.name_full || c.director_fio || '') : (c.director_fio || ''));
          case 'ogrn_line': return c.ogrn ? (isIp(c) ? 'ОГРНИП ' : 'ОГРН ') + c.ogrn : '';
          case 'reg_line': return [c.ogrn ? (isIp(c) ? 'ОГРНИП ' : 'ОГРН ') + c.ogrn : '', c.inn ? 'ИНН ' + c.inn : '', c.kpp ? 'КПП ' + c.kpp : ''].filter(Boolean).join(', ');
          case 'bank_line': return [c.rs ? 'р/с ' + c.rs : '', c.bank_name ? 'в ' + c.bank_name : '', c.ks ? 'к/с ' + c.ks : '', c.bik ? 'БИК ' + c.bik : ''].filter(Boolean).join(', ');
          case 'salary_num': { const n = parseNum(String(c.salary || '').replace(/[^\d,.\s]/g, '')); return n == null ? '' : fmtMoney(n); }
          case 'salary_words': { const n = parseNum(String(c.salary || '').replace(/[^\d,.\s]/g, '')); return n == null ? '' : rubWords(n); }
          case 'salary_text': { if (!c.salary) return ''; const n = parseNum(String(c.salary).replace(/[^\d,.\s]/g, '')); return n == null ? clean(c.salary) : fmtMoney(n) + ' руб. (' + rubWords(n).toLowerCase() + ')'; }
          case 'intro': {            // «ООО «Ромашка», именуемое в дальнейшем «Работодатель», в лице Директора Иванова И.И., действующего на основании Устава»
            const nm = c.name_full || c.name_short || ''; if (!nm) return '';
            const label = arg || 'Сторона', k = kindOf(c);
            let out = nm + ', ' + ctx.single('named', role) + ' в дальнейшем «' + label + '»';
            if (k === 'org' && c.director_pos_gen && c.director_fio_gen) out += ', в лице ' + c.director_pos_gen + ' ' + c.director_fio_gen + ', ' + ctx.single('acting', role) + ' на основании ' + (c.basis || 'Устава');
            else if (k === 'ip' && c.basis) out += ', ' + (genderOf(dirFio) === 'f' ? 'действующая' : 'действующий') + ' на основании ' + c.basis;
            return out;
          }
          default: return c[key] || '';
        }
      },
      // значение «токена» готового бланка: {{employer.name_full}}, {{worker.passport}}, {{doc.number}} …
      token(who, key, arg) {
        const role = TOKEN_ROLE[who]; const d = doc;
        if (who === 'doc') {
          switch (key) {
            case 'number': return d.number || '';
            case 'date': case 'date_long': return dateLong(d.date);
            case 'date_short': return dateShort(d.date);
            case 'city': return d.city || '';
            case 'spec_no': return d.specNo || '';
            case 'amount': { const n = parseNum(d.amount); if (n == null) return ''; const w = rubWords(n); return fmtMoney(n) + ' руб. (' + w.charAt(0).toLowerCase() + w.slice(1) + ')'; }
            case 'amount_num': case 'amount_words': case 'vat_amount': return this.single(key, null);
            case 'vat_rate': { const r = parseFloat(d.vat); return isFinite(r) && r > 0 ? String(r) : ''; }
            case 'vat_phrase': { const r = parseFloat(d.vat) || 0, n = parseNum(d.amount); if (!r) return d.amount ? 'НДС не облагается' : ''; const v = this.single('vat_amount', null); return 'в том числе НДС ' + r + ' %' + (v ? ' (' + v + ' руб.)' : ''); }
            default: return '';
          }
        }
        let k = key; if (!role) return '';
        if (who === 'worker' || who === 'employee') { if (key === 'fio') k = 'name_full'; if (key === 'position') k = 'director_pos'; }
        return this.single(k, role, arg || DEFAULT_LABEL[who]);
      },
      // значение для набора ключей: один ключ — как есть, несколько — «Подпись значение»
      render(items, role) {
        const vals = items.map(k => [k, this.single(k, role)]);
        if (items.length === 1) return { text: vals[0][1], miss: vals[0][1] ? [] : [items[0]] };
        const parts = [], miss = [];
        for (const [k, v] of vals) { if (!v) { miss.push(k); continue; } const sh = k === 'ogrn' ? (isIp(parties[role] || {}) ? 'ОГРНИП' : 'ОГРН') : SHORT[k]; parts.push((sh ? sh + (/^addr_/.test(k) ? ': ' : ' ') : '') + v); }
        const bankish = items.some(k => /^(bank_|rs$|ks$|bik$|contact_|director_)/.test(k));
        return { text: parts.join(bankish ? ', ' : '; '), miss };
      },
      dateLong() { return dateLong(doc.date); },
      filled(key, val, role) { ctx.stats.filled++; ctx.items.push({ kind: 'fill', key, val, role }); },
      hl(what, role) { ctx.stats.highlighted++; ctx.items.push({ kind: 'hl', key: what, role }); },
    };
    return ctx;
  }
  const TOKEN_ROLE = { customer: 'customer', employer: 'customer', executor: 'executor', worker: 'executor', employee: 'executor', doc: null };
  const DEFAULT_LABEL = { employer: 'Работодатель', worker: 'Работник', employee: 'Работник', customer: 'Заказчик', executor: 'Исполнитель' };
  const titleOf = k => (FIELD_TITLES[k] || ({ party_line: 'Реквизиты стороны', director_pos_fio: 'Должность и ФИО руководителя', signer_fio: 'ФИО подписанта', signer_fio_short: 'ФИО (инициалы)', signer_pos: 'Должность подписанта', signer_pos_fio: 'Должность и ФИО', sig_title: 'Руководитель', inperson: 'В лице (род. падеж)', inn_kpp: 'ИНН/КПП', addr_any: 'Адрес', contact_fio_pos: 'Контактное лицо', vat_payer: 'Плательщик НДС', vat_text: 'НДС', doc_date: 'Дата документа', doc_no: '№ договора', spec_no: '№ спецификации', city: 'Город', amount: 'Сумма договора', branches: 'Филиалы', holding: 'Холдинг', director_phone: 'Телефон руководителя', name_addr: 'Наименование и адрес', amount_num: 'Сумма договора', amount_words: 'Сумма прописью', vat_amount: 'Сумма НДС', amount: 'Сумма договора', intro: 'Вводная часть стороны', named: 'именуемый / именуемое', acting: 'действующий на основании', fio_gen: 'ФИО (род. падеж)', fio_short: 'ФИО (инициалы)', ogrn_line: 'ОГРН / ОГРНИП', reg_line: 'ОГРН, ИНН, КПП', bank_line: 'Банковские реквизиты', salary_text: 'Оклад', salary_num: 'Оклад', salary_words: 'Оклад прописью', date_long: 'Дата документа', date: 'Дата документа', number: '№ документа', vat_rate: 'Ставка НДС', vat_phrase: 'НДС' })[k] || k);

  /* =====================================================================
     Токены готовых бланков: {{employer.name_full}}, {{worker.passport}}, {{doc.number}}, {{customer.intro|Заказчик}}
     Стороны: employer / worker (трудовой договор), customer / executor (договоры и формы), doc — данные документа.
     В отличие от «пустых мест» (____) токен точно говорит, что и чьё вписать.
     ===================================================================== */
  const TOKEN_RE = /\{\{\s*([A-Za-z]+)\s*\.\s*([A-Za-z_0-9]+)\s*(?:\|\s*([^{}]{0,40}?)\s*)?\}\}/g;

  /* =====================================================================
     DOCX: операции с прогонами (w:r)
     ===================================================================== */
  const mkW = (doc, name) => doc.createElementNS(NS_W, 'w:' + name);
  function runText(r) { let s = ''; for (let c = r.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; const n = lname(c); if (n === 't') s += c.textContent; else if (n === 'tab') s += '\t'; } return s; }
  function isSimpleRun(r) { for (let c = r.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; if (!['rPr', 't', 'tab'].includes(lname(c))) return false; } return true; }
  function setRunText(r, text) {
    const doc = r.ownerDocument;
    for (const c of Array.from(r.childNodes)) if (c.nodeType === 1 && ['t', 'tab'].includes(lname(c))) r.removeChild(c);
    text.split('\t').forEach((part, i) => {
      if (i > 0) r.appendChild(mkW(doc, 'tab'));
      if (part !== '') { const t = mkW(doc, 't'); t.setAttributeNS(NS_XML, 'xml:space', 'preserve'); t.appendChild(doc.createTextNode(part)); r.appendChild(t); }
    });
  }
  const RECURSE = new Set(['hyperlink', 'ins', 'smartTag', 'fldSimple', 'sdt', 'sdtContent', 'customXml']);
  function paraRuns(p) { const out = []; (function walk(n) { for (let c = n.firstChild; c; c = c.nextSibling) { if (c.nodeType !== 1) continue; const nm = lname(c); if (nm === 'r') out.push(c); else if (RECURSE.has(nm)) walk(c); } })(p); return out; }
  function mapParagraph(p) { let pos = 0; const items = paraRuns(p).map(r => { const t = runText(r); const o = { r, t, s: pos, e: pos + t.length }; pos += t.length; return o; }); return { items, text: items.map(i => i.t).join('') }; }
  function splitAt(p, k) {
    for (const it of mapParagraph(p).items) if (it.s < k && k < it.e) {
      if (!isSimpleRun(it.r)) return;
      const cl = it.r.cloneNode(true); setRunText(it.r, it.t.slice(0, k - it.s)); setRunText(cl, it.t.slice(k - it.s)); it.r.parentNode.insertBefore(cl, it.r.nextSibling); return;
    }
  }
  function rangeRuns(p, s, e) { splitAt(p, s); splitAt(p, e); return mapParagraph(p).items.filter(it => it.t.length > 0 && it.s < e && it.e > s).map(it => it.r); }
  const HL_BEFORE = ['u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  function addHighlight(run) {
    const doc = run.ownerDocument; let rPr = firstKid(run, 'rPr');
    if (!rPr) { rPr = mkW(doc, 'rPr'); run.insertBefore(rPr, run.firstChild); }
    for (const h of kids(rPr, 'highlight')) rPr.removeChild(h);
    const hl = mkW(doc, 'highlight'); hl.setAttributeNS(NS_W, 'w:val', 'yellow');
    const before = Array.from(rPr.childNodes).find(c => c.nodeType === 1 && HL_BEFORE.includes(lname(c)));
    if (before) rPr.insertBefore(hl, before); else rPr.appendChild(hl);
  }
  const stripFmt = (rPr, names) => { for (const c of Array.from(rPr.childNodes)) if (c.nodeType === 1 && names.includes(lname(c))) rPr.removeChild(c); };
  function cleanRPr(rPr, extra) { if (!rPr) return null; const c = rPr.cloneNode(true); stripFmt(c, ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'highlight'].concat(extra || [])); return c; }
  function makeRun(doc, text, rPr) { const r = mkW(doc, 'r'); if (rPr) r.appendChild(rPr); setRunText(r, text); return r; }
  function fillRange(p, s, e, val) {
    const runs = rangeRuns(p, s, e); if (!runs.length) return false;
    setRunText(runs[0], val); const rp = firstKid(runs[0], 'rPr'); if (rp) stripFmt(rp, ['highlight', 'u']);
    for (let i = 1; i < runs.length; i++) runs[i].parentNode.removeChild(runs[i]);
    return true;
  }
  function highlightRange(p, s, e) { const runs = rangeRuns(p, s, e); runs.forEach(addHighlight); return runs.length > 0; }
  const NBSP6 = '      ';
  function lastTextRunRPr(p) { const runs = paraRuns(p).filter(r => runText(r).trim()); const r = runs[runs.length - 1]; return r ? firstKid(r, 'rPr') : null; }
  function appendToParagraph(doc, p, text, highlight) { const run = makeRun(doc, text, cleanRPr(lastTextRunRPr(p), ['b', 'bCs'])); if (highlight) addHighlight(run); p.appendChild(run); }

  /* ---------- пустые места в абзаце ---------- */
  const BLANK_RE = /_{3,}|\.{5,}|…{2,}/g;
  const PH_RE = /\[([^\]\n]{2,160})\]|\{\{([^}\n]{2,60})\}\}|(?<![\w/])<([^<>\n]{2,60})>/g;
  const DATE_RE = /[«"“]\s*_*\s*[»"”]\s*_{2,}\s*(?:20\d?_{1,}|20\d\d)\s*(?:г\.?|года)?/;
  function findBlanks(map) {
    const ranges = [];
    for (const m of map.text.matchAll(BLANK_RE)) ranges.push({ s: m.index, e: m.index + m[0].length });
    for (const it of map.items) if (it.t.length >= 3 && !it.t.trim()) {
      const rPr = firstKid(it.r, 'rPr'), u = rPr && firstKid(rPr, 'u');
      if (u && (u.getAttributeNS(NS_W, 'val') || u.getAttribute('w:val')) !== 'none') ranges.push({ s: it.s, e: it.e });
    }
    for (const m of map.text.matchAll(PH_RE)) ranges.push({ s: m.index, e: m.index + m[0].length, ph: true, inner: m[1] || m[2] || m[3] });
    ranges.sort((a, b) => a.s - b.s || b.e - a.e);
    const out = [];
    for (const r of ranges) { const last = out[out.length - 1]; if (last && r.s < last.e) continue; out.push(r); }
    return out;
  }
  const trimEnd = s => s.replace(/[\s ]+$/, ''), trimStart = s => s.replace(/^[\s ]+/, '');

  /* что стоит перед пустым местом → какие поля вписать */
  function resolveContext(before, after, prevText, nextText, ctx) {
    let b = trimEnd(before), a = trimStart(after);
    if (/^[\s.,;:_\-–—]*$/.test(a)) a = '';
    const lineIsBlankOnly = !b && !a;
    if (!b && prevText && prevText.length < 160 && !/именуем/.test(a)) b = trimEnd(prevText);
    if (lineIsBlankOnly && nextText) a = trimStart(nextText);
    const K = (items, role) => ({ act: 'key', items, role: role || null });
    if (/^[\s)\]/(]*(?:подпис|м\.\s*п)/i.test(a)) return { act: 'skip' };
    if (/(?:м\.\s*п\.?|подпись\s*[:\-–]?)[)\s]*$/i.test(b.slice(-30))) return { act: 'skip' };
    // название стороны договора: «_____, именуемое в дальнейшем «Заказчик»»
    if (/^[,\s]*(?:именуем|далее\s+именуем|в\s+дальнейшем)/i.test(a)) return K(['name_full']);
    const nm = /^[,\s]*(?:\(\s*)?(полное\s+)?(?:фирменное\s+)?наименование(\s+и\s+адрес)?/i.exec(a);
    if (nm && lineIsBlankOnly || (nm && !b)) return K([nm[2] ? 'name_addr' : 'name_full'], roleInText(a.slice(0, 80)));
    const hm = a.match(/^[)\]/]*\s*\(\s*([^)]{0,70})\)/);
    if (hm) {
      const h = lc(hm[1]);
      if (/должност/.test(h) && /ф\.?\s*и\.?\s*о|фамили/.test(h)) return K([/в\s+лице\s*$/i.test(b) ? 'inperson' : 'director_pos_fio'], roleInText(b.slice(-80)));
      if (/ф\.?\s*и\.?\s*о|фамили|расшифровк|инициал/.test(h)) return K([/расшифровк|инициал/.test(h) ? 'signer_fio_short' : 'signer_fio'], roleInText(b.slice(-80)));
      if (/должност/.test(h)) return K(['signer_pos'], roleInText(b.slice(-80)));
      if (/наименовани|организаци/.test(h)) return K(['name_full'], roleInText(b.slice(-80)));
      if (/подпис/.test(h)) return { act: 'skip' };
    }
    if (/\/\s*$/.test(b)) return K(['signer_fio_short'], roleInText(b.slice(-120)));
    if (/^\//.test(a)) return { act: 'skip' };
    const tail = b.slice(-120), lt = lc(tail);
    if (/(?:^|[\s«»"])г\.?\s*$|город\s*$/.test(lt) || /^г\.?$/.test(lt)) return K(['city']);
    if (/(?:№|n)\s*$/i.test(lt)) {
      const wide = lc(b);
      if (/спецификаци|приложени/.test(wide) && !/к\s+договор[а-яёa-z0-9_]*\s*$/.test(wide)) return K(['spec_no']);
      if (/договор|контракт/.test(wide) || ctx.contract) return K(['doc_no']);
      return { act: 'unknown', label: tail.slice(-40) };
    }
    if (/(?:в\s+том\s+числе|включая|из\s+них)\s+ндс[^.]{0,40}$|сумма\s+ндс[^.]{0,30}$/.test(lt) || /ндс[^.]{0,30}(?:составляет|—|-|:)\s*$/.test(lt)) return K(['vat_amount']);
    if (/(?:сумм[а-яёa-z0-9_]*|цен[а-яёa-z0-9_]*|стоимост[а-яёa-z0-9_]*)[^:.]{0,80}$/.test(lt) && !/ндс/.test(lt.slice(-30))) return K([/\(\s*$/.test(tail) ? 'amount_words' : 'amount_num']);
    if (BARE_NAME.test(lt.replace(/[:\s]+$/, ''))) return K(['name_full'], roleInText(lt));
    const r = classify(tail, { tail: true });
    if (!r) return { act: 'unknown', label: tail.slice(-40) };
    if (r.items[0] === 'director_pos_fio' && r.boss) return { act: 'skip' };
    if (r.items[0] === 'signer_pos_fio' && /:\s*$/.test(tail) === false && !a) return { act: 'skip' };
    return K(r.items, r.role);
  }
  const siblingParaText = (p, dir) => { let n = dir < 0 ? p.previousSibling : p.nextSibling; while (n && n.nodeType !== 1) n = dir < 0 ? n.previousSibling : n.nextSibling; return n && lname(n) === 'p' ? mapParagraph(n).text : ''; };

  function pickRole(ctx, explicit, extra) {
    return explicit || extra || ctx.role || (ctx.contract ? null : ctx.opts.defaultRole);
  }
  function applyResolution(p, b, r, ctx, role) {
    if (r.act === 'skip') return;
    if (r.act === 'unknown') { if (ctx.opts.mode !== 'min' && ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl(r.label || 'unknown', null); } return; }
    const items = r.items.filter(k => ctx.allowed(k)); if (!items.length) return;
    const needsRole = items.some(k => !['doc_date', 'doc_no', 'spec_no', 'city', 'amount_num', 'amount_words', 'vat_amount'].includes(k));
    if (needsRole && !role) { if (ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl('не определена сторона: ' + items.map(titleOf).join(', '), null); } return; }
    const out = ctx.render(items, role);
    if (out.text) { fillRange(p, b.s, b.e, out.text); ctx.filled(items.map(titleOf).join(' + '), out.text, needsRole ? role : null); }
    else if (ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl(items.map(titleOf).join(' + '), needsRole ? role : null); }
  }

  function processParagraph(doc, p, ctx) {
    let map = mapParagraph(p);
    const text0 = map.text; if (!text0.trim()) return;
    const hr = headingRole(text0);
    if (hr) { ctx.role = hr; return; }
    // метки сторон в абзаце: «именуемое в дальнейшем «Заказчик»»
    const markers = []; for (const m of text0.matchAll(/именуем[а-яёa-z0-9_]*(?:\s*\(\s*[-–]?[а-яё]+\s*\))?(?:\s+в\s+дальнейшем)?\s*[,(]?\s*[«"“]?\s*([А-Яа-яЁё]+)/gi)) { const rl = roleOfWord(m[1]); if (rl) markers.push({ idx: m.index, rl }); }
    let dateRange = null;
    const dm = map.text.match(DATE_RE);
    if (dm) {
      dateRange = { s: dm.index, e: dm.index + dm[0].length };
      const dv = ctx.dateLong();
      if (dv && ctx.allowed('doc_date')) { fillRange(p, dateRange.s, dateRange.e, dv); ctx.filled('Дата документа', dv, null); dateRange = null; }
      else if (ctx.opts.mode !== 'min' && ctx.opts.highlight) { highlightRange(p, dateRange.s, dateRange.e); ctx.hl('Дата документа', null); }
      map = mapParagraph(p);
    }
    let blanks = findBlanks(map);
    if (dateRange) blanks = blanks.filter(b => !(b.s >= dateRange.s && b.e <= dateRange.e));
    const prevText = siblingParaText(p, -1), nextText = siblingParaText(p, +1);
    const textNow = map.text;
    if (blanks.length) {
      // роли считаем по исходным позициям, правим с конца
      const plan = blanks.map((b, i) => {
        let r;
        if (b.ph) { const c = classify(b.inner, { noBare: true }); r = c && c.items.length ? { act: 'key', items: c.items, role: c.role } : { act: 'skip' }; }
        else r = resolveContext(textNow.slice(0, b.s), textNow.slice(b.e), prevText, nextText, ctx);
        let mr = null;
        const lastB = markers.filter(m => m.idx < b.s).pop(), nextA = markers.find(m => m.idx >= b.e);
        if (r.act === 'key' && r.items.includes('name_full') && nextA && !(blanks[i + 1] && blanks[i + 1].s < nextA.idx) && nextA.idx - b.e < 220) mr = nextA.rl;
        else if (lastB) mr = lastB.rl;
        return { b, r, role: pickRole(ctx, r.role, mr) };
      });
      for (let i = 0; i < plan.length - 1; i++) {      // «_____ [наименование участника]» — оставляем одно значение
        const a1 = plan[i], a2 = plan[i + 1];
        if (!a1.b.ph && a2.b.ph && !textNow.slice(a1.b.e, a2.b.s).trim()) {
          if (a1.r.act !== 'key' && a2.r.act === 'key') { a1.r = a2.r; a1.role = pickRole(ctx, a2.r.role, a1.role); }
          a2.drop = true;
        }
      }
      for (let i = plan.length - 1; i >= 0; i--) { if (plan[i].drop) fillRange(p, plan[i].b.s, plan[i].b.e, ''); else applyResolution(p, plan[i].b, plan[i].r, ctx, plan[i].role); }
    } else {
      const t = trimEnd(map.text);
      if (t.length <= 140 && /[:：]$/.test(t)) {
        const c = classify(t.slice(0, -1), { noBare: false });
        if (c && !c.skip && c.items.length && !c.boss) {
          const items = c.items.filter(k => ctx.allowed(k));
          const role = pickRole(ctx, c.role);
          if (items.length && role) {
            const out = ctx.render(items, role);
            if (out.text) { appendToParagraph(doc, p, ' ' + out.text, false); ctx.filled(items.map(titleOf).join(' + '), out.text, role); }
            else if (ctx.opts.highlight) { appendToParagraph(doc, p, ' ' + NBSP6, true); ctx.hl(items.map(titleOf).join(' + '), role); }
          }
        }
      }
    }
    if (markers.length) ctx.role = markers[markers.length - 1].rl;
  }

  /* =====================================================================
     DOCX: таблицы
     ===================================================================== */
  const cellText = tc => byNS(tc, NS_W, 't').map(t => t.textContent).join('');
  const isBlankText = s => /^[\s _.…]*$/.test(s);
  const hasObjects = tc => byNS(tc, NS_W, 'drawing').length > 0 || byNS(tc, NS_W, 'tbl').length > 0;
  const looksLikeField = l => /[:：]\s*$/.test(l) || (l.split(/\s+/).length >= 2 && !/\d{3,}/.test(l));
  const AMOUNT_LABEL = /(?:сумма|цена|стоимость)\s+(?:договора|контракта|заявки|предложения)|общая\s+(?:сумма|стоимость)\s+(?:договора|контракта)?/;

  function writeCell(doc, tc, val, o) {
    o = o || {};
    const p = kids(tc, 'p')[0]; if (!p) return;
    const runs = paraRuns(p); let rPr = null;
    if (runs.length) rPr = cleanRPr(firstKid(runs[0], 'rPr'));
    if (!rPr) { const pPr = firstKid(p, 'pPr'); const mark = pPr && firstKid(pPr, 'rPr'); if (mark) rPr = cleanRPr(mark); }
    if (!rPr && o.labelTc) { const lp = kids(o.labelTc, 'p')[0]; const lr = lp && paraRuns(lp).find(r => runText(r).trim()); const rp = lr && firstKid(lr, 'rPr'); if (rp) rPr = cleanRPr(rp, ['b', 'bCs', 'u']); }
    for (const r of runs) r.parentNode.removeChild(r);
    const run = makeRun(doc, o.highlight ? NBSP6 : val, rPr);
    if (o.highlight) addHighlight(run);
    p.appendChild(run);
  }

  function handleTarget(doc, ctx, tcLabel, tcTarget, label, noBare, cellRole) {
    const lab = lc(label);
    let items, role = null, explicit = null;
    if (AMOUNT_LABEL.test(lab) && !/ндс/.test(lab.replace(/с\s+учет[а-яёa-z0-9_]*[^,]*ндс/, ''))) { items = ['amount']; }
    else {
      const r = classify(label, { noBare });
      if (r && r.skip) return;
      if (r && r.items.length) { items = r.items; explicit = r.role; }
    }
    if (items) {
      items = items.filter(k => ctx.allowed(k)); if (!items.length) return;
      const needsRole = items.some(k => !['doc_date', 'doc_no', 'spec_no', 'city', 'amount', 'amount_num', 'amount_words', 'vat_amount'].includes(k));
      role = pickRole(ctx, explicit, cellRole);
      if (needsRole && !role) { if (ctx.opts.highlight) { writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true }); ctx.hl('не определена сторона: ' + items.map(titleOf).join(', '), null); } return; }
      const out = ctx.render(items, role);
      if (out.text) { writeCell(doc, tcTarget, out.text, { labelTc: tcLabel }); ctx.filled(items.map(titleOf).join(' + '), out.text, needsRole ? role : null); }
      else if (ctx.opts.highlight) { writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true }); ctx.hl(items.map(titleOf).join(' + '), needsRole ? role : null); }
    } else if (ctx.opts.mode !== 'min' && ctx.opts.highlight && looksLikeField(label)) {
      writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true });
      ctx.hl(clean(label).slice(0, 50), null);
    }
  }

  /* таблица-спецификация: суммы строк, ИТОГО, НДС */
  const gridCols = tr => { let c = 0; return kids(tr, 'tc').map(tc => { const pr = firstKid(tc, 'tcPr'), gs = pr && firstKid(pr, 'gridSpan'); const start = c; c += gs ? (parseInt(gs.getAttributeNS(NS_W, 'val') || gs.getAttribute('w:val'), 10) || 1) : 1; return { tc, start, span: c - start }; }); };
  function processSpec(doc, tbl, ctx) {
    const rows = kids(tbl, 'tr'); if (rows.length < 2) return false;
    let hi = -1, qty = -1, price = -1, sum = -1, priceNoVat = false;
    for (let i = 0; i < Math.min(rows.length, 3) && hi < 0; i++) {
      const g = gridCols(rows[i]); let q = -1, p = -1, s = -1, pn = false;
      g.forEach(c => { const t = lc(cellText(c.tc)); if (/кол-?\s?во|количеств/.test(t) && q < 0) q = c.start; else if (/цена/.test(t) && p < 0) { p = c.start; pn = /без\s+ндс/.test(t); } else if (/сумма|стоимост|всего/.test(t) && !/цена/.test(t) && s < 0) s = c.start; });
      if (q >= 0 && p >= 0 && s >= 0) { hi = i; qty = q; price = p; sum = s; priceNoVat = pn; }
    }
    if (hi < 0) return false;
    const cellAt = (tr, col) => gridCols(tr).find(c => c.start === col) || null;
    let total = 0, cnt = 0, missing = 0, totalRowIdx = -1, vatRowIdx = -1;
    const vatExcl = priceNoVat || ctx.doc.vatMode === 'excl';
    rows.slice(hi + 1).forEach((tr, k) => {
      const idx = hi + 1 + k, first = lc(cellText(kids(tr, 'tc')[0] || tr));
      if (/^\s*итого/.test(first)) { totalRowIdx = idx; return; }
      if (/ндс/.test(first) && /(?:в\s+том\s+числе|включая|сумма|ндс)/.test(first) && idx > hi + 1) { vatRowIdx = idx; return; }
      const cq = cellAt(tr, qty), cp = cellAt(tr, price), cs = cellAt(tr, sum); if (!cq || !cp || !cs) return;
      const q = parseNum(cellText(cq.tc)), pr = parseNum(cellText(cp.tc)), sv = parseNum(cellText(cs.tc));
      if (isBlankText(cellText(cp.tc)) && q != null) { if (ctx.opts.highlight && ctx.opts.mode !== 'min') { writeCell(doc, cp.tc, '', { highlight: true }); ctx.hl('Цена за единицу', null); } }
      if (q != null && pr != null) { const v = Math.round(q * pr * 100) / 100; if (isBlankText(cellText(cs.tc))) { writeCell(doc, cs.tc, fmtMoney(v), { labelTc: cp.tc }); ctx.filled('Сумма по позиции', fmtMoney(v), null); } total += v; cnt++; }
      else if (sv != null) { total += sv; cnt++; } else missing++;
    });
    const rate = ctx.vatRate();
    const fillTotal = (idx, val, name) => {
      if (idx < 0 || val == null) return; const c = cellAt(rows[idx], sum) || gridCols(rows[idx]).slice(-1)[0];
      if (c && isBlankText(cellText(c.tc))) { writeCell(doc, c.tc, fmtMoney(val), {}); ctx.filled(name, fmtMoney(val), null); }
    };
    let grand = null;
    if (cnt && !missing) grand = vatExcl && rate ? total * (1 + rate / 100) : total;
    if (grand == null) { const a = parseNum(ctx.doc.amount); if (a != null && cnt === 0) grand = a; }
    if (grand != null) {
      fillTotal(totalRowIdx, vatExcl && rate && cnt && !missing ? total : grand, 'ИТОГО');
      const vat = rate ? (vatExcl ? grand - grand / (1 + rate / 100) : grand * rate / (100 + rate)) : null;
      if (vat != null) fillTotal(vatRowIdx, vat, 'В том числе НДС');
    } else if (ctx.opts.highlight && ctx.opts.mode !== 'min') {
      for (const idx of [totalRowIdx, vatRowIdx]) if (idx >= 0) { const c = cellAt(rows[idx], sum) || gridCols(rows[idx]).slice(-1)[0]; if (c && isBlankText(cellText(c.tc))) { writeCell(doc, c.tc, '', { highlight: true }); ctx.hl(idx === totalRowIdx ? 'ИТОГО' : 'НДС', null); } }
    }
    return true;
  }

  /** Подставляет токены во всех абзацах части документа. Абзацы и таблицы с токенами дальше обычными правилами не трогаются. */
  function tokenPass(doc, ctx) {
    let total = 0;
    for (const p of byNS(doc, NS_W, 'p')) {
      const map = mapParagraph(p); if (map.text.indexOf('{{') < 0) continue;
      const hits = [];
      for (const m of map.text.matchAll(TOKEN_RE)) { const who = m[1].toLowerCase(); if (who in TOKEN_ROLE) hits.push({ s: m.index, e: m.index + m[0].length, who, key: m[2].toLowerCase(), arg: m[3] }); }
      if (!hits.length) continue;
      for (let i = hits.length - 1; i >= 0; i--) {
        const h = hits[i], val = ctx.token(h.who, h.key, h.arg), role = TOKEN_ROLE[h.who], what = titleOf(h.key) + (role ? '' : '');
        if (val) { fillRange(p, h.s, h.e, val); ctx.filled(what, val, role); }
        else { fillRange(p, h.s, h.e, '________'); if (ctx.opts.highlight) highlightRange(p, h.s, h.s + 8); ctx.hl(what, role); }
        total++;
      }
      ctx.skipP.add(p);
      for (let a = p.parentNode; a; a = a.parentNode) if (a.nodeType === 1 && lname(a) === 'tbl') ctx.skipT.add(a);
    }
    return total;
  }

  function walkBlock(doc, node, ctx) {
    for (let c = node.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const n = lname(c);
      if (n === 'p') { if (!insideHandled(c, ctx) && !ctx.skipP.has(c)) { try { processParagraph(doc, c, ctx); } catch (e) { ctx.items.push({ kind: 'err', key: String(e && e.message || e) }); } } }
      else if (n === 'tbl') { if (!ctx.skipT.has(c)) processTable(doc, c, ctx); }
      else if (n === 'sdt' || n === 'sdtContent') walkBlock(doc, c, ctx);
    }
  }
  function insideHandled(p, ctx) { for (let n = p.parentNode; n; n = n.parentNode) if (ctx.handled.has(n)) return true; return false; }

  function processTable(doc, tbl, ctx) {
    const rows = kids(tbl, 'tr');
    const saveRole = ctx.role;
    if (processSpec(doc, tbl, ctx)) { ctx.role = saveRole; return; }
    // роли колонок: «ЗАКАЗЧИК | ИСПОЛНИТЕЛЬ» в первых строках
    const colRole = {};
    rows.slice(0, 2).forEach(tr => gridCols(tr).forEach((c, i) => { const hr = headingRole(cellText(c.tc)); if (hr) colRole[i] = hr; }));
    const twoRole = Object.keys(colRole).length >= 2;
    for (let ri = 0; ri < rows.length; ri++) {
      const tr = rows[ri], cells = kids(tr, 'tc');
      for (let i = 0; i < cells.length - 1; i++) {
        const label = cellText(cells[i]).trim();
        if (!label || isBlankText(label) || label.length > 700 || /_{3,}|\.{5,}|…{2,}/.test(label)) continue;
        if (twoRole || headingRole(label)) continue;
        const next = cells[i + 1];
        if (isBlankText(cellText(next)) && !hasObjects(next)) { handleTarget(doc, ctx, cells[i], next, label, false, colRole[i + 1] || null); ctx.handled.add(cells[i]); ctx.handled.add(next); i++; }
        else if (classify(label) || /[:：]\s*$/.test(label)) ctx.handled.add(cells[i]);
      }
      const nextRow = rows[ri + 1];
      if (nextRow) {
        const below = kids(nextRow, 'tc');
        const allHeaders = cells.length >= 1 && cells.every(c => !isBlankText(cellText(c)));
        const belowEmpty = below.length === cells.length && below.every(c => { const t = cellText(c).trim(); return isBlankText(t) || /^\d+[.)]?$/.test(t); }) && below.some(c => isBlankText(cellText(c)));
        if (allHeaders && belowEmpty && !cells.some(c => ctx.handled.has(c)) && !twoRole) {
          cells.forEach((hc, j) => { const tcb = below[j]; if (!isBlankText(cellText(tcb)) || ctx.handled.has(tcb)) return; handleTarget(doc, ctx, hc, tcb, cellText(hc).trim(), true, null); ctx.handled.add(tcb); });
          cells.forEach(c => ctx.handled.add(c));
        }
      }
      // остальное содержимое ячеек — абзацы и вложенные таблицы, с ролью колонки
      cells.forEach((tc, i) => {
        if (ctx.handled.has(tc)) return;
        const fp = kids(tc, 'p')[0], hr = fp ? headingRole(mapParagraph(fp).text) : null;
        const prev = ctx.role; if (hr || colRole[i]) ctx.role = hr || colRole[i];
        walkBlock(doc, tc, ctx);
        ctx.role = prev;
      });
    }
    ctx.role = saveRole;
  }

  async function fillDocx(buf, parties, docInfo, opts) {
    const zip = await JSZip.loadAsync(buf);
    if (!zip.file('word/document.xml')) throw new Error('Не найден word/document.xml — это не DOCX');
    const ctx = makeCtx(parties, docInfo, opts);
    const parts = ['word/document.xml'].concat(Object.keys(zip.files).filter(n => /^word\/(header|footer)\d*\.xml$/.test(n)));
    for (const part of parts) {
      const doc = parseXml(await zip.file(part).async('string'));
      if (part === 'word/document.xml') {
        const all = lc(byNS(doc, NS_W, 't').map(t => t.textContent).join(' '));
        const t = ctx.opts.docType;
        ctx.contract = t === 'contract' || (t === 'auto' && (/именуем[а-яёa-z0-9_]*\s+(?:в\s+дальнейшем\s+)?[«"]?(?:заказчик|исполнитель|подрядчик|поставщик|покупатель)/.test(all) || /спецификаци/.test(all.slice(0, 600)) && /договор/.test(all.slice(0, 900))));
      }
      ctx.role = null;
      if (tokenPass(doc, ctx)) ctx.tplMode = true;   // бланк с токенами: остальной текст не трогаем
      if (!ctx.tplMode) walkBlock(doc, byNS(doc, NS_W, 'body')[0] || doc.documentElement, ctx);
      zip.file(part, serialize(doc));
    }
    const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { data, stats: ctx.stats, contract: ctx.contract, items: ctx.items };
  }

  /* =====================================================================
     XLSX
     ===================================================================== */
  const colToNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const numToCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const splitRef = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? { col: colToNum(m[1]), row: +m[2] } : null; };
  async function readShared(zip) { const f = zip.file('xl/sharedStrings.xml'); if (!f) return []; const doc = parseXml(await f.async('string')); return byNS(doc, NS_S, 'si').map(si => byNS(si, NS_S, 't').map(t => t.textContent).join('')); }
  function xlCellText(c, sst) { const t = c.getAttribute('t'), v = firstKid(c, 'v'); if (t === 's') return v ? (sst[+v.textContent] || '') : ''; if (t === 'inlineStr') return byNS(c, NS_S, 't').map(x => x.textContent).join(''); return v ? v.textContent : ''; }
  function setXlText(doc, c, text) {
    for (const ch of Array.from(c.childNodes)) c.removeChild(ch);
    c.setAttribute('t', 'inlineStr');
    const is = doc.createElementNS(NS_S, 'is'), t = doc.createElementNS(NS_S, 't');
    t.setAttributeNS(NS_XML, 'xml:space', 'preserve'); t.appendChild(doc.createTextNode(text)); is.appendChild(t); c.appendChild(is);
  }
  function makeHighlightStyler(stylesDoc) {
    if (!stylesDoc) return null;
    const fills = byNS(stylesDoc, NS_S, 'fills')[0], xfsEl = byNS(stylesDoc, NS_S, 'cellXfs')[0]; if (!fills || !xfsEl) return null;
    let fillId = null; const cache = {};
    return function (oldS) {
      oldS = oldS || 0; if (cache[oldS] != null) return cache[oldS];
      if (fillId == null) {
        const fill = stylesDoc.createElementNS(NS_S, 'fill'), pf = stylesDoc.createElementNS(NS_S, 'patternFill'); pf.setAttribute('patternType', 'solid');
        const fg = stylesDoc.createElementNS(NS_S, 'fgColor'); fg.setAttribute('rgb', 'FFFFFF00'); const bg = stylesDoc.createElementNS(NS_S, 'bgColor'); bg.setAttribute('indexed', '64');
        pf.appendChild(fg); pf.appendChild(bg); fill.appendChild(pf); fills.appendChild(fill); fillId = kids(fills, 'fill').length - 1; fills.setAttribute('count', String(fillId + 1));
      }
      const xfs = kids(xfsEl, 'xf'), base = xfs[oldS] || xfs[0], xf = base ? base.cloneNode(true) : stylesDoc.createElementNS(NS_S, 'xf');
      xf.setAttribute('fillId', String(fillId)); xf.setAttribute('applyFill', '1'); xfsEl.appendChild(xf);
      const idx = kids(xfsEl, 'xf').length - 1; xfsEl.setAttribute('count', String(idx + 1)); return (cache[oldS] = idx);
    };
  }
  function processSheet(doc, sst, ctx, styler) {
    const merges = {};
    for (const m of byNS(doc, NS_S, 'mergeCell')) { const [a, b] = (m.getAttribute('ref') || '').split(':'); const A = splitRef(a), B = b && splitRef(b); if (A && B) merges[a] = B.col; }
    const colRole = {}, usedT = new Set();
    for (const row of byNS(doc, NS_S, 'row')) {
      const rn = +row.getAttribute('r'), cellEls = kids(row, 'c'), byCol = {};
      cellEls.forEach(c => { const ref = splitRef(c.getAttribute('r') || ''); if (ref) byCol[ref.col] = c; });
      for (const c of cellEls) {
        const ref = splitRef(c.getAttribute('r') || ''); if (!ref) continue;
        const label = xlCellText(c, sst).trim();
        if (!label || !/[A-Za-zА-Яа-яЁё]/.test(label) || label.length > 700) continue;
        if (/^\(.*\)$/.test(label)) continue;          // пояснение под линией «(должность)», «(организация, адрес…)» — не поле
        if (ctx.opts.overwrite && (usedT.has(c.getAttribute('r')) || /\d{8,}/.test(label.replace(/[\s-]/g, '')))) continue;   // в режиме замены старые значения не считаем подписями
        const hr = headingRole(label); if (hr) { colRole[ref.col] = hr; continue; }
        const tcol = (merges[c.getAttribute('r')] || ref.col) + 1;
        const cellRole = colRole[tcol] || colRole[ref.col] || null;
        const hlCell = (cell, what, role) => { if (ctx.opts.highlight && styler) { cell.setAttribute('s', String(styler(+cell.getAttribute('s')))); ctx.hl(what, role); } };
        const bm = /_{3,}\s*$/.exec(label);
        if (bm) {
          const r = classify(label.slice(0, bm.index), { tail: true });
          if (r && !r.skip && r.items.length && !r.boss) {
            const items = r.items.filter(k => ctx.allowed(k)), role = pickRole(ctx, r.role, cellRole); if (!items.length || !role) continue;
            const out = ctx.render(items, role);
            if (out.text) { setXlText(doc, c, label.slice(0, bm.index).replace(/\s*$/, ' ') + out.text); ctx.filled(items.map(titleOf).join(' + '), out.text, role); } else hlCell(c, items.map(titleOf).join(' + '), role);
          }
          continue;
        }
        let tgt = byCol[tcol];
        if (tgt) usedT.add(tgt.getAttribute('r'));
        const tgtText = tgt ? xlCellText(tgt, sst).trim() : '';
        const filledTgt = !!tgt && !isBlankText(tgtText);
        if (filledTgt) {
          // режим замены (КС-2 / КС-3): перезаписываем реквизиты, но не трогаем соседние подписи и числа-коды форм
          if (!ctx.opts.overwrite) continue;
          if (tgtText.length < 45 && /^[А-Яа-яЁёA-Za-z\s()\/.,:\-]+$/.test(tgtText) && classify(tgtText)) continue;
        }
        const lab = lc(label);
        let items = null, explicit = null, bare = false;
        if (AMOUNT_LABEL.test(lab)) items = ['amount'];
        else { const r = classify(label); if (r && !r.skip && r.items.length) { items = r.items; explicit = r.role; bare = !!r.bare; } }
        if (items && ctx.opts.partyLine && bare && items.length === 1 && items[0] === 'name_full') items = ['party_line'];
        const ensure = () => {
          if (!tgt) { tgt = doc.createElementNS(NS_S, 'c'); tgt.setAttribute('r', numToCol(tcol) + rn); const after = cellEls.find(x => { const q = splitRef(x.getAttribute('r') || ''); return q && q.col > tcol; }); if (after) row.insertBefore(tgt, after); else row.appendChild(tgt); byCol[tcol] = tgt; }
          return tgt;
        };
        if (items) {
          items = items.filter(k => ctx.allowed(k)); if (!items.length) continue;
          const role = pickRole(ctx, explicit, cellRole);
          const needsRole = items.some(k => !['doc_date', 'doc_no', 'spec_no', 'city', 'amount', 'amount_num', 'amount_words', 'vat_amount'].includes(k));
          if (needsRole && !role) continue;
          const out = ctx.render(items, role);
          if (out.text) { setXlText(doc, ensure(), out.text); ctx.filled(items.map(titleOf).join(' + '), out.text, needsRole ? role : null); }
          else if (tgt && !filledTgt) hlCell(tgt, items.map(titleOf).join(' + '), needsRole ? role : null);
        } else if (ctx.opts.mode !== 'min' && tgt && !filledTgt && looksLikeField(label)) hlCell(tgt, clean(label).slice(0, 50), null);
      }
    }
  }
  /* КС-2 / КС-3: подписанты. «Сдал» — Подрядчик (Субподрядчик), «Принял» — Заказчик (Генподрядчик).
     Ищем строку подписей под «Сдал / Принял»: над пояснениями «(должность)» и «(расшифровка подписи)» вписываем
     должность и фамилию с инициалами. Если пояснений нет — берём первую и последнюю свободные ячейки строки. */
  function processSigners(doc, sst, ctx, styler) {
    const grid = {}, rowEls = {};
    for (const row of byNS(doc, NS_S, 'row')) {
      const rn = +row.getAttribute('r'); rowEls[rn] = row; grid[rn] = {};
      for (const c of kids(row, 'c')) { const ref = splitRef(c.getAttribute('r') || ''); if (ref) grid[rn][ref.col] = c; }
    }
    const txt = (r, c) => { const el = grid[r] && grid[r][c]; return el ? xlCellText(el, sst).trim() : ''; };
    const anchors = [];
    for (const rn of Object.keys(grid).map(Number)) for (const cn of Object.keys(grid[rn]).map(Number)) {
      const m = /^(сдал|принял)[а-я]*(?:\s*\([^)]*\))?\s*[:.]?$/.exec(lc(clean(txt(rn, cn))));
      if (m) anchors.push({ r: rn, c: cn, role: m[1] === 'сдал' ? 'executor' : 'customer' });
    }
    if (!anchors.length) return 0;
    const getCell = (rn, col) => {
      if (grid[rn] && grid[rn][col]) return grid[rn][col];
      const row = rowEls[rn]; if (!row) return null;
      const c = doc.createElementNS(NS_S, 'c'); c.setAttribute('r', numToCol(col) + rn);
      const after = kids(row, 'c').find(x => { const q = splitRef(x.getAttribute('r') || ''); return q && q.col > col; });
      if (after) row.insertBefore(c, after); else row.appendChild(c);
      grid[rn][col] = c; return c;
    };
    let done = 0;
    for (const a of anchors) {
      const sameRow = anchors.filter(x => x.r === a.r && x.c > a.c).map(x => x.c), sameCol = anchors.filter(x => x.c === a.c && x.r > a.r).map(x => x.r);
      const colStart = a.c + 1, colEnd = sameRow.length ? Math.min.apply(null, sameRow) - 1 : 1e9, rowEnd = sameCol.length ? Math.min.apply(null, sameCol) - 1 : a.r + 6;
      let posCol = null, fioCol = null, fioFull = false, tr = a.r;
      for (let rr = a.r; rr <= rowEnd && posCol == null && fioCol == null; rr++) {
        for (const cn of Object.keys(grid[rr] || {}).map(Number).sort((x, y) => x - y)) {
          if (cn < colStart || cn > colEnd) continue;
          const t = lc(txt(rr, cn));
          if (/должност/.test(t) && posCol == null) posCol = cn;
          else if (/расшифр|ф\.?\s?и\.?\s?о|фамил/.test(t) && fioCol == null) { fioCol = cn; fioFull = !/расшифр|инициал/.test(t); }
        }
        if (posCol != null || fioCol != null) tr = Math.max(a.r, rr - 1);
      }
      if (posCol == null && fioCol == null) {          // пояснений нет — берём свободные ячейки в строке «Сдал / Принял»
        const free = Object.keys(grid[a.r] || {}).map(Number).filter(cn => cn >= colStart && cn <= colEnd && isBlankText(txt(a.r, cn))).sort((x, y) => x - y);
        if (free.length >= 2) { posCol = free[0]; fioCol = free[free.length - 1]; }
        else if (free.length === 1) fioCol = free[0];
      }
      const c = ctx.parties[a.role] || {};
      const put = (col, text, what) => {
        if (col == null) return;
        const cell = getCell(tr, col); if (!cell) return;
        const cur = xlCellText(cell, sst).trim();
        if (!isBlankText(cur) && !ctx.opts.overwrite) return;
        if (text) { setXlText(doc, cell, text); ctx.filled(what, text, a.role); done++; }
        else if (isBlankText(cur) && ctx.opts.highlight && styler) { cell.setAttribute('s', String(styler(+cell.getAttribute('s')))); ctx.hl(what, a.role); }
      };
      put(posCol, c.director_pos || '', a.role === 'executor' ? 'Должность (Сдал)' : 'Должность (Принял)');
      put(fioCol, c.director_fio ? (fioFull ? c.director_fio : shortFio(c.director_fio)) : '', a.role === 'executor' ? 'Подписант (Сдал)' : 'Подписант (Принял)');
    }
    return done;
  }
  async function fillXlsx(buf, parties, docInfo, opts) {
    const zip = await JSZip.loadAsync(buf), ctx = makeCtx(parties, docInfo, opts), sst = await readShared(zip);
    const sf = zip.file('xl/styles.xml'), stylesDoc = sf ? parseXml(await sf.async('string')) : null, styler = makeHighlightStyler(stylesDoc);
    const sheets = Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if (!sheets.length) throw new Error('В файле нет листов');
    for (const name of sheets) { const doc = parseXml(await zip.file(name).async('string')); ctx.role = null; processSheet(doc, sst, ctx, styler); if (ctx.opts.signers) processSigners(doc, sst, ctx, styler); zip.file(name, serialize(doc)); }
    if (stylesDoc) zip.file('xl/styles.xml', serialize(stylesDoc));
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), stats: ctx.stats, items: ctx.items };
  }

  /* =====================================================================
     КС-2 / КС-3 (Excel): отдельный, «структурный» заполнитель.
     Раньше реквизиты уходили в соседнюю узкую ячейку, в колонки кодов и в заголовки таблицы, а длинный текст
     обрезался соседними ячейками. Теперь:
       • строка стороны ищется по подписи «Заказчик / Подрядчик» и пишется в ту ячейку (или объединённую область),
         где в бланке стоит значение; формат «Заказчик - значение» в одной ячейке тоже понимается;
       • после записи ширина области считается заново, при необходимости включается перенос строк и
         увеличивается высота строки — ничего не обрезается;
       • рядом с «по ОКПО» старый код чужой организации убирается (в карточке ОКПО нет — ячейка подсвечивается);
       • подписанты: «Сдал / Принял» (КС-2) и «Заказчик / Подрядчик» над строкой подписи (КС-3).
     ===================================================================== */
  const KS_ROLE = w => /^(заказчик|генподрядчик)/.test(w) ? 'customer' : 'executor';
  const KS_LABEL_ONLY = /^(заказчик|генподрядчик|подрядчик|субподрядчик)(\s*\([^)]*\))?\s*[:.]?$/;
  const KS_COMBINED = /^((?:заказчик|генподрядчик|подрядчик|субподрядчик)[^–—:\-]{0,45}?)\s*[-–—:]\s+(\S[\s\S]*)$/i;
  const KS_CODE_LABEL = /окпо|окуд|окдп|окв|^форма|^код$|вид\s+(деятельности|операции)|^номер|^дата|^договор/;

  function ksOpen(doc, sst, stylesDoc) {
    const S = { doc, sst, cells: {}, rows: {}, merges: [], mAt: {}, cov: new Set(), colsDef: [], styles: stylesDoc };
    for (const row of byNS(doc, NS_S, 'row')) {
      const rn = +row.getAttribute('r'); S.rows[rn] = row; S.cells[rn] = {};
      for (const c of kids(row, 'c')) { const ref = splitRef(c.getAttribute('r') || ''); if (ref) S.cells[rn][ref.col] = c; }
    }
    for (const m of byNS(doc, NS_S, 'mergeCell')) {
      const [a, b] = (m.getAttribute('ref') || '').split(':'), A = splitRef(a), B = b && splitRef(b); if (!A || !B) continue;
      const mm = { r1: A.row, c1: A.col, r2: B.row, c2: B.col }; S.merges.push(mm); S.mAt[A.row + ':' + A.col] = mm;
      for (let r = mm.r1; r <= mm.r2; r++) for (let c = mm.c1; c <= mm.c2; c++) if (r !== mm.r1 || c !== mm.c1) S.cov.add(r + ':' + c);
    }
    const fmt = byNS(doc, NS_S, 'sheetFormatPr')[0];
    S.defW = fmt && fmt.getAttribute('defaultColWidth') ? parseFloat(fmt.getAttribute('defaultColWidth')) : (fmt && fmt.getAttribute('baseColWidth') ? parseFloat(fmt.getAttribute('baseColWidth')) + 0.71 : 8.43);
    S.defH = fmt && fmt.getAttribute('defaultRowHeight') ? parseFloat(fmt.getAttribute('defaultRowHeight')) : 15;
    S.colW = {}; for (const c of byNS(doc, NS_S, 'col')) { const w = c.getAttribute('hidden') === '1' ? 0 : parseFloat(c.getAttribute('width') || S.defW); for (let k = +c.getAttribute('min'); k <= Math.min(+c.getAttribute('max'), 200); k++) S.colW[k] = w; }
    // шрифты и форматы ячеек — чтобы оценить размер текста и включить перенос строк
    S.fonts = []; S.xfs = []; S.xfCache = {};
    if (stylesDoc) {
      const f = byNS(stylesDoc, NS_S, 'fonts')[0]; if (f) for (const x of kids(f, 'font')) { const sz = firstKid(x, 'sz'); S.fonts.push(sz ? parseFloat(sz.getAttribute('val')) : 11); }
      const cx = byNS(stylesDoc, NS_S, 'cellXfs')[0]; S.xfsEl = cx; if (cx) S.xfs = kids(cx, 'xf');
    }
    S.mdw = Math.max(5, Math.round((S.fonts[0] || 11) * 96 / 72 * 0.52));
    return S;
  }
  const ksText = (S, r, c) => { const el = S.cells[r] && S.cells[r][c]; return el ? xlCellText(el, S.sst) : ''; };
  const ksTextC = (S, r, c) => clean(ksText(S, r, c));
  function ksCell(S, r, c) {
    if (S.cells[r] && S.cells[r][c]) return S.cells[r][c];
    let row = S.rows[r];
    if (!row) {
      row = S.doc.createElementNS(NS_S, 'row'); row.setAttribute('r', String(r));
      const sd = byNS(S.doc, NS_S, 'sheetData')[0], after = Object.keys(S.rows).map(Number).filter(x => x > r).sort((a, b) => a - b)[0];
      if (after) sd.insertBefore(row, S.rows[after]); else sd.appendChild(row);
      S.rows[r] = row; S.cells[r] = {};
    }
    const c1 = S.doc.createElementNS(NS_S, 'c'); c1.setAttribute('r', numToCol(c) + r);
    const nxt = Object.keys(S.cells[r]).map(Number).filter(x => x > c).sort((a, b) => a - b)[0];
    if (nxt) row.insertBefore(c1, S.cells[r][nxt]); else row.appendChild(c1);
    S.cells[r][c] = c1; return c1;
  }
  const ksColPx = (S, c) => { const w = S.colW[c] != null ? S.colW[c] : S.defW; return w > 0 ? Math.round(w * S.mdw + 5) : 0; };
  const ksRowPt = (S, r) => { const row = S.rows[r]; if (row && row.getAttribute('hidden') === '1') return 0; return row && row.getAttribute('ht') ? parseFloat(row.getAttribute('ht')) : S.defH; };
  function ksXf(S, s) {
    s = s || 0; if (S.xfCache[s]) return S.xfCache[s];
    if (S.xfsEl && !S.xfs[s]) S.xfs = kids(S.xfsEl, 'xf');            // подсветка (жёлтая) тоже дописывает форматы — синхронизируем список
    const x = S.xfs[s] || S.xfs[0], al = x && firstKid(x, 'alignment');
    return (S.xfCache[s] = { sz: S.fonts[x ? (+x.getAttribute('fontId') || 0) : 0] || 10, wrap: !!al && al.getAttribute('wrapText') === '1', h: al ? al.getAttribute('horizontal') || '' : '', v: al ? al.getAttribute('vertical') || '' : '' });
  }
  /** вариант формата ячейки с переносом строк (и, если нужно, выравниванием) */
  function ksVariant(S, s, mods) {
    s = s || 0; if (!S.xfsEl || !S.xfs.length) return s;
    const key = s + '|' + JSON.stringify(mods); S.varCache = S.varCache || {}; if (S.varCache[key] != null) return S.varCache[key];
    S.xfs = kids(S.xfsEl, 'xf');
    const base = S.xfs[s] || S.xfs[0], xf = base.cloneNode(true); let al = firstKid(xf, 'alignment');
    if (!al) { al = S.styles.createElementNS(NS_S, 'alignment'); xf.appendChild(al); }
    if (mods.wrap) al.setAttribute('wrapText', '1');
    if (mods.v) al.setAttribute('vertical', mods.v);
    if (mods.h) al.setAttribute('horizontal', mods.h);
    xf.setAttribute('applyAlignment', '1'); S.xfsEl.appendChild(xf); S.xfs = kids(S.xfsEl, 'xf'); S.xfsEl.setAttribute('count', String(S.xfs.length));
    return (S.varCache[key] = S.xfs.length - 1);
  }
  /** сколько строк займёт текст в доступной ширине (жадный перенос по словам) */
  function ksLines(text, availPx, fontPt) {
    const cw = fontPt * 96 / 72 * 0.56; let lines = 0;
    for (const par of String(text).split(/\r?\n/)) {
      let cur = 0, n = 1;
      for (const w of par.split(/\s+/)) {
        const len = w.length * cw, sp = cur ? cw : 0;
        if (cur && cur + sp + len > availPx) { n++; cur = 0; }
        if (len > availPx) { n += Math.floor(len / availPx); cur = len % availPx; } else cur += (cur ? cw : 0) + len;
      }
      lines += n;
    }
    return Math.max(1, lines);
  }
  /** Запись текста в ячейку (merge-область) с проверкой: помещается ли текст; если нет — перенос строк и высота строки */
  function ksWrite(S, r, c, text, opts) {
    const cell = ksCell(S, r, c); setXlText(S.doc, cell, text);
    const m = S.mAt[r + ':' + c], c2 = m ? m.c2 : c, r2 = m ? m.r2 : r;
    const s0 = +cell.getAttribute('s') || 0, xf = ksXf(S, s0);
    let avail = 0; for (let k = c; k <= c2; k++) avail += ksColPx(S, k);
    if (!m && !xf.wrap) {                                   // текст без переноса вытекает в пустые соседние ячейки — как в Excel
      const free = (from, step) => { let w = 0; for (let k = from; k >= 1 && k <= 200; k += step) { if (S.cov.has(r + ':' + k) || S.mAt[r + ':' + k] || ksText(S, r, k).trim()) break; w += ksColPx(S, k); if (w > 4000) break; } return w; };
      if (xf.h === 'center') avail += 2 * Math.min(free(c - 1, -1), free(c + 1, 1)); else if (xf.h === 'right') avail += free(c - 1, -1); else avail += free(c + 1, 1);
    }
    avail = Math.max(20, avail - 6);
    const lines = ksLines(text, avail, xf.sz);
    if (lines <= 1 && !(opts && opts.forceWrap)) return;
    if (!xf.wrap || (opts && opts.top)) cell.setAttribute('s', String(ksVariant(S, s0, { wrap: true, v: xf.v === 'center' ? '' : (xf.v || 'top') })));
    let have = 0; for (let k = r; k <= r2; k++) have += ksRowPt(S, k);
    const need = lines * xf.sz * 1.27 + 3;
    if (need > have + 0.5) {
      const row = S.rows[r2] || S.rows[r]; if (row) { const ht = ksRowPt(S, +row.getAttribute('r')) + (need - have); row.setAttribute('ht', ht.toFixed(2)); row.setAttribute('customHeight', '1'); }
    }
  }
  const ksHl = (S, cell, styler) => { if (styler && cell) cell.setAttribute('s', String(styler(+cell.getAttribute('s') || 0))); };

  /** Где в строке стоит значение для подписи «Заказчик / Подрядчик»: первая непустая ячейка (или широкая объединённая область) правее подписи */
  function ksValueCell(S, r, labelCol) {
    const lm = S.mAt[r + ':' + labelCol], start = (lm ? lm.c2 : labelCol) + 1;
    let firstFilled = null, wide = null; const between = [];
    for (let c = start; c <= start + 40; c++) {
      if (S.cov.has(r + ':' + c)) continue;
      const m = S.mAt[r + ':' + c], t = ksTextC(S, r, c);
      if (t && KS_CODE_LABEL.test(lc(t)) && t.length < 40) break;           // дошли до «по ОКПО» и т. п.
      if (m && m.c2 - m.c1 >= 2) { wide = { r, c, m }; break; }              // широкая объединённая область — это и есть поле значения
      if (t && !/^\(.*\)$/.test(t)) { if (!firstFilled) firstFilled = { r, c, m }; between.push(c); }
      if (c > start + 14 && !t) break;
    }
    if (wide) { wide.junk = between; return wide; }   // непустые «одиночные» ячейки левее — остатки прежнего заполнения
    return firstFilled;
  }

  /** Код ОКПО в той же строке: ячейка справа от «по ОКПО» */
  function ksOkpoCell(S, r, fromCol) {
    for (let c = fromCol; c <= fromCol + 40; c++) {
      if (S.cov.has(r + ':' + c)) continue;
      if (/^по\s+окпо/.test(lc(ksTextC(S, r, c)))) { const m = S.mAt[r + ':' + c]; let k = (m ? m.c2 : c) + 1; while (S.cov.has(r + ':' + k)) k++; return { r, c: k }; }
    }
    return null;
  }

  function ksHeader(S, ctx, styler, done) {
    const ow = !!ctx.opts.overwrite;
    for (let r = 1; r <= 22; r++) {              // шапка акта / справки; ниже — таблица и подписи
      for (const c of Object.keys(S.cells[r] || {}).map(Number).sort((a, b) => a - b)) {
        if (S.cov.has(r + ':' + c)) continue;
        const raw = ksText(S, r, c), t = clean(raw), l = lc(t);
        if (!t || t.length > 900 || /^\(.*\)$/.test(t)) continue;
        let role = null, prefix = null, target = null;
        if (KS_LABEL_ONLY.test(l)) { role = KS_ROLE(l); target = ksValueCell(S, r, c); }
        else { const m = KS_COMBINED.exec(t); if (m && /^(заказчик|генподрядчик|подрядчик|субподрядчик)/i.test(m[1].trim())) { role = KS_ROLE(lc(m[1].trim())); prefix = m[1].replace(/\s+$/, ''); target = { r, c, m: S.mAt[r + ':' + c] }; } }
        if (!role) continue;
        const party = ctx.parties[role] || {}, key = role === 'customer' ? 'Заказчик (Генподрядчик)' : 'Подрядчик (Субподрядчик)';
        const line = ctx.render([ctx.opts.partyLine === false ? 'name_full' : 'party_line'], role).text;
        const lastCol = target ? (target.m ? target.m.c2 : target.c) : c;
        if (!target) continue;
        if (target.junk && target.junk.length && ow && line) for (const jc of target.junk) { const jel = S.cells[target.r] && S.cells[target.r][jc]; if (jel) { for (const ch of Array.from(jel.childNodes)) jel.removeChild(ch); jel.removeAttribute('t'); } }
        const cell = ksCell(S, target.r, target.c), cur = clean(xlCellText(cell, S.sst));
        if (!line) { if (ctx.opts.highlight && !cur) { ksHl(S, cell, styler); ctx.hl(key, role); } }
        else if (!cur || ow || prefix) {
          if (cur && !ow && prefix && !isBlankText(clean(raw).slice(prefix.length).replace(/^\s*[-–—:]\s*/, ''))) { /* в режиме «не заменять» оставляем как есть */ }
          else { ksWrite(S, target.r, target.c, prefix ? prefix + ' - ' + line : line, { top: true }); ctx.filled(key, line, role); done.n++; }
        }
        // код ОКПО прежней организации: убираем, чтобы не остался чужой код
        const ok = ksOkpoCell(S, r, lastCol + 1);
        if (ok) {
          const oc = ksCell(S, ok.r, ok.c), oldCode = clean(xlCellText(oc, S.sst));
          if (party.okpo) { if (!oldCode || ow) { setXlText(S.doc, oc, party.okpo); ctx.filled('ОКПО (' + key + ')', party.okpo, role); } }
          else if (oldCode && ow) { for (const ch of Array.from(oc.childNodes)) oc.removeChild(ch); oc.removeAttribute('t'); if (ctx.opts.highlight) { ksHl(S, oc, styler); ctx.hl('ОКПО (' + key + ') — в карточке нет, старый код убран', role); } }
          else if (!oldCode && ctx.opts.highlight) { ksHl(S, oc, styler); ctx.hl('ОКПО (' + key + ')', role); }
        }
      }
    }
  }

  const KS_POS_HINT = /должност/, KS_FIO_HINT = /расшифр|ф\.?\s?и\.?\s?о|фамил|инициал/;
  function ksSigners(S, ctx, styler, done) {
    const ow = !!ctx.opts.overwrite;
    const anchors = [];
    const rowsAll = Object.keys(S.cells).map(Number).sort((a, b) => a - b);
    for (const r of rowsAll) for (const c of Object.keys(S.cells[r]).map(Number).sort((a, b) => a - b)) {
      const t = lc(clean(ksText(S, r, c))).replace(/\s*\([^)]*\)/g, '').trim();
      if (!t || t.length > 20) continue;
      let m = /^(сдал|принял)[а-я]*\s*[:.]?$/.exec(t);
      if (m) anchors.push({ r, c, role: m[1] === 'сдал' ? 'executor' : 'customer', kind: 'A' });
      else if (r > 15 && (m = /^(заказчик|подрядчик)\s*[:.]?$/.exec(t))) anchors.push({ r, c, role: m[1] === 'заказчик' ? 'customer' : 'executor', kind: 'B' });
    }
    const used = new Set();
    for (const a of anchors) {
      const p = ctx.parties[a.role] || {}, who = a.role === 'executor' ? '(Сдал / Подрядчик)' : '(Принял / Заказчик)';
      const posT = p.director_pos || '', fioShort = p.director_fio ? shortFio(p.director_fio) : '';
      const nextSame = anchors.filter(x => x.r === a.r && x.c > a.c).map(x => x.c), cEnd = nextSame.length ? Math.min.apply(null, nextSame) - 1 : a.c + 40;
      const write = (r, c, text, what) => {
        const cell = ksCell(S, r, c), cur = clean(xlCellText(cell, S.sst));
        if (cur && !ow) return;
        if (text) { ksWrite(S, r, c, text, {}); ctx.filled(what + ' ' + who, text, a.role); done.n++; }
        else if (!cur && ctx.opts.highlight) { ksHl(S, cell, styler); ctx.hl(what + ' ' + who, a.role); }
      };
      if (a.kind === 'B') {
        // КС-3: под подписью «Заказчик / Подрядчик» одна ячейка «должность и фамилия с инициалами»
        let hit = null;
        for (let r = a.r; r <= a.r + 3 && !hit; r++) for (const c of Object.keys(S.cells[r] || {}).map(Number).sort((x, y) => x - y)) {
          if (c <= a.c || S.cov.has(r + ':' + c) || used.has(r + ':' + c)) continue;
          const t = clean(ksText(S, r, c)); if (!t || /^\(.*\)$/.test(t) || /^м\.?\s*п\.?$/i.test(t)) continue;
          if (/[А-ЯЁ]\.\s?[А-ЯЁ]\./.test(t) || /^(директор|руководитель|генеральный|председатель|индивидуальный|ип\s|представитель|управляющий)/i.test(t)) { hit = { r, c }; break; }
        }
        if (hit) { used.add(hit.r + ':' + hit.c); write(hit.r, hit.c, posT || fioShort ? [posT, fioShort].filter(Boolean).join(' ') : '', 'Подписант'); }
        continue;
      }
      // КС-2: строка «Сдал / Принял»
      const regs = [];
      for (let c = a.c + 1; c <= cEnd; c++) { if (S.cov.has(a.r + ':' + c)) continue; const m = S.mAt[a.r + ':' + c]; regs.push({ c1: c, c2: m ? m.c2 : c, m }); if (m) c = m.c2; }
      const cand = regs.filter(g => ksTextC(S, a.r, g.c1) || (g.m && g.c2 > g.c1)).filter(g => !/^м\.?\s*п\.?$/i.test(ksTextC(S, a.r, g.c1)));
      let posR = null, fioR = null, fioFull = false, sepHints = false;
      const hints = [];
      for (let r = a.r + 1; r <= a.r + 3; r++) for (const c of Object.keys(S.cells[r] || {}).map(Number)) {
        if (c <= a.c || c > cEnd) continue; const h = lc(ksTextC(S, r, c)); if (!h) continue;
        hints.push({ c, pos: KS_POS_HINT.test(h), fio: KS_FIO_HINT.test(h), t: h });
      }
      const both = hints.some(h => h.pos && h.fio);
      const hp = hints.find(h => h.pos && !h.fio), hf = hints.find(h => h.fio && !h.pos);
      const regAt = col => regs.find(g => g.c1 <= col && col <= g.c2);
      if (hp && hf && !both) { posR = regAt(hp.c); fioR = regAt(hf.c); fioFull = !/расшифр|инициал/.test(hf.t); sepHints = !!(posR && fioR && posR !== fioR); }
      if (!sepHints) {
        if (cand.length >= 2) { posR = cand[0]; fioR = cand[cand.length - 1]; }
        else if (cand.length === 1) { posR = null; fioR = cand[0]; }
        else { const free = regs.filter(g => !ksTextC(S, a.r, g.c1)); if (free.length >= 2) { posR = free[0]; fioR = free[free.length - 1]; } else if (free.length === 1) fioR = free[0]; }
        if (cand.length === 1 && !posR) { write(a.r, fioR.c1, posT || fioShort ? [posT, fioShort].filter(Boolean).join(' ') : '', 'Подписант'); continue; }
      }
      if (posR) write(a.r, posR.c1, posT, 'Должность');
      if (fioR) write(a.r, fioR.c1, p.director_fio ? (fioFull ? p.director_fio : fioShort) : '', 'Подписант');
    }
  }

  async function fillKs(buf, parties, docInfo, opts) {
    const zip = await JSZip.loadAsync(buf), ctx = makeCtx(parties, docInfo, Object.assign({ overwrite: true, partyLine: true, signers: true }, opts)), sst = await readShared(zip);
    const sf = zip.file('xl/styles.xml'), stylesDoc = sf ? parseXml(await sf.async('string')) : null, styler = makeHighlightStyler(stylesDoc);
    const sheets = Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if (!sheets.length) throw new Error('В файле нет листов');
    const done = { n: 0 };
    for (const name of sheets) {
      const doc = parseXml(await zip.file(name).async('string')), S = ksOpen(doc, sst, stylesDoc);
      ksHeader(S, ctx, ctx.opts.highlight ? styler : null, done);
      if (ctx.opts.signers !== false) ksSigners(S, ctx, ctx.opts.highlight ? styler : null, done);
      zip.file(name, serialize(doc));
    }
    if (stylesDoc) zip.file('xl/styles.xml', serialize(stylesDoc));
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), stats: ctx.stats, items: ctx.items };
  }

  /* ---------- чтение карты: DOCX / XLSX → текст ---------- */
  async function docxToText(buf) {
    const zip = await JSZip.loadAsync(buf), f = zip.file('word/document.xml'); if (!f) throw new Error('Не найден word/document.xml');
    const doc = parseXml(await f.async('string')), body = byNS(doc, NS_W, 'body')[0], out = [];
    const ptext = p => byNS(p, NS_W, 't').map(t => t.textContent).join('');
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue; const n = lname(c);
        if (n === 'p') { const t = ptext(c).trim(); if (t) out.push(t); }
        else if (n === 'tbl') for (const tr of kids(c, 'tr')) { const cells = kids(tr, 'tc').map(tc => kids(tc, 'p').map(ptext).join(' ').replace(/\s+/g, ' ').trim()); if (cells.some(Boolean)) out.push(cells.join(' | ')); }
        else if (n === 'sdt' || n === 'sdtContent') walk(c);
      }
    })(body);
    return out.join('\n');
  }
  async function xlsxToText(buf) {
    const zip = await JSZip.loadAsync(buf), sst = await readShared(zip), out = [];
    for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      const doc = parseXml(await zip.file(name).async('string'));
      for (const row of byNS(doc, NS_S, 'row')) { const cells = kids(row, 'c').map(c => xlCellText(c, sst).trim()).filter(Boolean); if (cells.length) out.push(cells.join(' | ')); }
    }
    return out.join('\n');
  }
  /** Старый формат .xls: переводим в .xlsx (модуль xls.js), дальше работаем как с обычным Excel */
  async function xlsToXlsx(buf) {
    const X2 = (typeof self !== 'undefined' && self.AssistantXls) || (typeof globalThis !== 'undefined' && globalThis.AssistantXls);
    if (!X2) throw new Error('Не загрузился модуль чтения .xls');
    return X2.toXlsx(buf);
  }
  async function fillFile(name, buf, parties, docInfo, opts) {
    const ext = (name.split('.').pop() || '').toLowerCase();
    if (ext === 'docx' || ext === 'dotx' || ext === 'docm') return fillDocx(buf, parties, docInfo, opts);
    if (ext === 'xls') {
      const cv = await xlsToXlsx(buf), r = await (opts && opts.ks ? fillKs : fillXlsx)(cv.data, parties, docInfo, opts);
      r.convertedFrom = 'xls'; r.convertMode = cv.mode; return r;
    }
    if (ext === 'xlsx' || ext === 'xlsm') return (opts && opts.ks) ? fillKs(buf, parties, docInfo, opts) : fillXlsx(buf, parties, docInfo, opts);
    const e = new Error('Формат .' + ext + ' не поддерживается. Сохраните бланк как .docx, .xlsx или .xls'); e.unsupported = true; throw e;
  }

  // внутренние функции — для модулей техпредложения и наложения печати (tech.js, overlay.js)
  const internals = {
    JSZip, X, NS_W, NS_S, NS_XML, parseXml, serialize, lname, kids, firstKid, byNS, lc, clean, parseNum, fmtMoney,
    looksFio, fioGen, posGen, shortFio, POS_GEN, mapParagraph, paraRuns, runText, setRunText, rangeRuns, splitAt, fillRange, highlightRange, addHighlight, cleanRPr, makeRun, mkW, lastTextRunRPr, headingRole, roleOfWord, firstKidRPr: r => firstKid(r, 'rPr'),
    readShared, xlCellText, setXlText, makeHighlightStyler, splitRef, numToCol, colToNum, cellText, NBSP6, isBlankText, writeCell, makeCtx, xlsToXlsx,
  };

  return { parseCardText, finishCard, applyIpRules, detectIp, kindOf, genderOf, fillDocx, fillXlsx, fillKs, fillFile, docxToText, xlsxToText, xlsToXlsx, classify, shortFio, fioGen, posGen, rubWords, fmtMoney, FIELD_TITLES, TOKEN_RE, TOKEN_ROLE, internals };
});
