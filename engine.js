/* =====================================================================
   Ассистент — движок заполнения форм (DOCX / XLSX)
   Работает целиком в браузере. Структуру, шрифты, таблицы не трогает:
   правится только текст в нужных местах, форматирование клонируется.
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('jszip'), require('@xmldom/xmldom'));
  } else {
    root.AssistantEngine = factory(root.JSZip, { DOMParser: root.DOMParser, XMLSerializer: root.XMLSerializer });
  }
})(typeof self !== 'undefined' ? self : this, function (JSZip, X) {
  'use strict';

  const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const NS_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_XML = 'http://www.w3.org/XML/1998/namespace';

  /* ---------------------------------------------------------------
     0. Общие утилиты XML
     --------------------------------------------------------------- */
  const parseXml = s => new X.DOMParser().parseFromString(s, 'application/xml');
  const serialize = d => {
    let s = new X.XMLSerializer().serializeToString(d);
    if (!/^<\?xml/.test(s)) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
    return s;
  };
  const lname = n => n.localName || String(n.nodeName).replace(/^.*:/, '');
  const kids = (n, name) => {
    const out = [];
    for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && lname(c) === name) out.push(c);
    return out;
  };
  const firstKid = (n, name) => kids(n, name)[0] || null;
  const byNS = (n, ns, name) => Array.from(n.getElementsByTagNameNS(ns, name));

  /* ---------------------------------------------------------------
     1. Справочник полей: как распознаём подпись поля в бланке
        Берётся совпадение, ближайшее к пустому месту; при равенстве — длиннее.
     --------------------------------------------------------------- */
  const NL = '(?<![а-яёa-z0-9])';   // слева не буква
  const NR = '(?![а-яёa-z])';        // справа не буква
  const rx = (s) => new RegExp(s, 'gi');

  const FIELDS = [
    ['inn_kpp', rx(NL + 'инн\\s*[/,\\\\и]\\s*кпп' + NR)],
    ['inn', rx(NL + 'инн' + NR)],
    ['kpp', rx(NL + 'кпп' + NR)],
    ['ogrn', rx(NL + 'огрн(?:ип)?' + NR + '|основной\\s+государственный\\s+регистрационный\\s+номер')],
    ['okpo', rx(NL + 'окпо' + NR)],
    ['oktmo', rx(NL + 'октмо' + NR)],
    ['okved', rx(NL + 'оквэд' + NR)],
    ['name_short', rx('сокращ[её]нн[а-яёa-z0-9]*\\s+(?:фирменн[а-яёa-z0-9]*\\s+)?(?:наименовани[а-яёa-z0-9]*|название)')],
    ['name_full', rx(
      '(?:полн[а-яёa-z0-9]*\\s+)?(?:фирменн[а-яёa-z0-9]*\\s+)?наименовани[еяю]\\s+(?:организации|юридического\\s+лица|участника(?:\\s+закупки)?|контрагента|претендента|заявителя|поставщика|подрядчика|исполнителя|компании|предприятия|партн[её]ра|субъекта)' +
      '|название\\s+(?:организации|компании)|полное\\s+(?:фирменное\\s+)?наименование|фирменное\\s+наименование|участник\\s+закупки|участник\\s+конкурса|наименование\\s+участника')],
    ['addr_post', rx('почтов[а-яёa-z0-9]*\\s+адрес|адрес\\s+для\\s+(?:корреспонденции|направления|почтовой)')],
    ['addr_fact', rx('фактическ[а-яёa-z0-9]*\\s+(?:адрес|местонахождение)|адрес\\s+фактическ[а-яёa-z0-9]*|место\\s+осуществления\\s+деятельности')],
    ['addr_legal', rx('юридическ[а-яёa-z0-9]*\\s+адрес|адрес\\s*(?:\\(\\s*)?(?:места\\s+нахождения|местонахождения|регистрации|юридическ[а-яёa-z0-9]*)|место\\s*нахождени[ея]|местонахождени[ея]')],
    ['addr_any', rx(NL + 'адрес' + NR + '(?!\\s+(?:электронн|эл\\.|e-?mail|сайта))')],
    ['email', rx('e-?mail|' + NL + 'эл[а-яёa-z0-9]*\\.?\\s*почт[а-яёa-z0-9]*|электронн[а-яёa-z0-9]+\\s+(?:почт[а-яёa-z0-9]*|адрес)|адрес\\s+электронной\\s+почты')],
    ['phone', rx('телефон\\s*[/,]\\s*факс|телефон[а-яёa-z0-9]*|' + NL + 'тел\\.?' + NR + '|контактный\\s+номер')],
    ['fax', rx('факс')],
    ['site', rx('сайт|web|веб-?сайт')],
    ['bank_name', rx('наименование\\s+банка|' + NL + 'банк' + NR + '|обслуживающ[а-яёa-z0-9]+\\s+банк')],
    ['bik', rx(NL + 'бик' + NR)],
    ['ks', rx(NL + 'к\\s*[/\\\\]\\s*с' + NR + '|корр[а-яёa-z0-9]*\\.?\\s*сч[её]т[а-яёa-z0-9]*|корреспондентск[а-яёa-z0-9]+')],
    ['rs', rx(NL + 'р\\s*[/\\\\]\\s*с' + NR + '|р\\.\\s*с\\.|расч[её]тн[а-яёa-z0-9]*\\s+сч[её]т(?:\\s+в\\s+банке)?|рас\\.?\\s*сч[её]т|номер\\s+сч[её]та')],
    ['signer_pos_fio', rx('должность\\s*[,/и]\\s*ф\\.?\\s*и\\.?\\s*о\\.?|ф\\.?\\s*и\\.?\\s*о\\.?\\s*[,/и]\\s*должность')],
    ['signer_fio', rx('ф\\.\\s*и\\.\\s*о\\.?|' + NL + 'фио' + NR + '|фамилия[, ]+имя[, ]+отчество')],
    ['signer_pos', rx(NL + 'должность' + NR)],
    ['inperson', rx('в\\s+лице')],
    ['basis', rx('действующ[а-яёa-z0-9]+\\s+на\\s+основании|на\\s+основании\\s+(?:устава|доверенности)')],
    ['tax', rx('систем[а-яёa-z0-9]+\\s+налогообложени[а-яёa-z0-9]+|налогообложени[а-яёa-z0-9]+|плательщик\\s+ндс|' + NL + 'ндс' + NR)],
    ['date', rx(NL + 'дата' + NR)],
  ];

  // «Голые» подписи — когда вся ячейка/строка состоит только из слова-подписи
  const BARE = [
    ['name_full', /^(?:наименование|организация|участник|исполнитель|подрядчик|поставщик|претендент|заявитель|контрагент|компания|предприятие|юридическое\s+лицо|фирма)[\s:,.()*№\-–—]*$/i],
    ['sig_title', /^(?:руководитель(?:\s+организации)?|подписант|уполномоченное\s+лицо|представитель|лицо,?\s+уполномоченное[^:]*|генеральный\s+директор|директор|президент|управляющий)[\s:,.()*\-–—]*$/i],
  ];

  function matchKeyEx(text, opt) {
    opt = opt || {};
    const t = String(text || '').replace(/\s+/g, ' ').trim();
    if (!t) return null;
    let best = null;
    for (const [key, re] of FIELDS) {
      for (const m of t.matchAll(re)) {
        const end = m.index + m[0].length, len = m[0].length;
        if (!best || end > best.end || (end === best.end && len > best.len)) best = { key, end, len };
      }
    }
    if (best) return { key: best.key, bare: false };
    if (!opt.noBare) for (const [key, re] of BARE) if (re.test(t)) return { key, bare: true };
    return null;
  }

  /* ---------------------------------------------------------------
     2. Работа с ФИО, должностями, датой
     --------------------------------------------------------------- */
  function shortFio(fio) {
    fio = (fio || '').trim();
    if (!fio) return '';
    if (/\./.test(fio)) return fio;
    const p = fio.split(/\s+/);
    if (p.length < 2) return fio;
    return p[0] + ' ' + p.slice(1, 3).map(x => x[0].toUpperCase() + '.').join('');
  }

  function genitivePosition(pos) {
    if (!pos) return '';
    return pos.trim().split(/\s+/).map(w => {
      const l = w.toLowerCase();
      if (/(ый|ой)$/.test(l)) return l.replace(/(ый|ой)$/, 'ого');
      if (/ий$/.test(l)) return l.replace(/ий$/, 'его');
      if (/ь$/.test(l)) return l.replace(/ь$/, 'я');
      if (/[бвгджзклмнпрстфхцчшщ]$/.test(l)) return l + 'а';
      return l;
    }).join(' ').replace(/^./, c => c.toUpperCase()).replace(/^(.)/, c => c.toLowerCase());
  }

  const MONTHS = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  function parseIso(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? { y: +m[1], mo: +m[2], d: +m[3] } : null;
  }
  const dateLong = iso => { const d = parseIso(iso); return d ? `«${String(d.d).padStart(2, '0')}» ${MONTHS[d.mo - 1]} ${d.y} г.` : ''; };
  const dateShort = iso => { const d = parseIso(iso); return d ? `${String(d.d).padStart(2, '0')}.${String(d.mo).padStart(2, '0')}.${d.y}` : ''; };

  const POS_RE = /^(генеральный\s+директор|исполнительный\s+директор|коммерческий\s+директор|директор|президент|руководитель|председатель|управляющий|индивидуальный\s+предприниматель|и\.\s*о\.\s*директора|главный\s+врач|ректор|начальник)\s+(.+)$/i;
  function splitPosFio(v) {
    const m = POS_RE.exec((v || '').trim());
    if (!m) return null;
    return { pos: m[1].replace(/^./, c => c.toUpperCase()), fio: m[2].trim() };
  }

  /* ---------------------------------------------------------------
     3. Разбор карты партнёра (текст → поля)
     --------------------------------------------------------------- */
  const OPF = [
    [/Общество с ограниченной ответственностью/gi, 'ООО'],
    [/Акционерное общество/gi, 'АО'],
    [/Публичное акционерное общество/gi, 'ПАО'],
    [/Индивидуальный предприниматель/gi, 'ИП'],
  ];

  function parseCardText(text) {
    const card = {};
    const set = (k, v) => { v = String(v || '').replace(/\s+/g, ' ').trim(); if (v && !card[k]) card[k] = v; };
    const lines = String(text || '').split(/\r?\n/).map(s => s.replace(/\u00A0/g, ' ').trim()).filter(Boolean);

    // пары «подпись → значение»
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
      if (ci > 0 && ci <= 70) {
        const lab = line.slice(0, ci), val = line.slice(ci + 1).trim();
        if (val) { pairs.push([lab, val]); continue; }
        if (lines[i + 1] && !matchKeyEx(lines[i + 1], { noBare: true })) pairs.push([lab, lines[i + 1]]);
        continue;
      }
      if (line.length <= 60 && matchKeyEx(line) && lines[i + 1] && !lines[i + 1].includes(' | ') && !matchKeyEx(lines[i + 1])) {
        pairs.push([line, lines[i + 1]]);
      }
    }

    const digits = (v, re) => { const m = re.exec(String(v).replace(/[\s-]/g, '')); return m ? m[0] : ''; };
    for (const [label, value] of pairs) {
      if (!value) continue;
      const r = matchKeyEx(label);
      if (!r) continue;
      switch (r.key) {
        case 'inn_kpp': { const m = String(value).match(/\d{9,12}/g) || []; set('inn', m.find(x => x.length >= 10)); set('kpp', m.find(x => x.length === 9)); break; }
        case 'inn': set('inn', digits(value, /\d{10,12}/)); break;
        case 'kpp': set('kpp', digits(value, /\d{9}/)); break;
        case 'ogrn': set('ogrn', digits(value, /\d{13,15}/)); break;
        case 'bik': set('bik', digits(value, /\d{9}/)); break;
        case 'rs': set('rs', digits(value, /\d{20}/)); break;
        case 'ks': set('ks', digits(value, /\d{20}/)); break;
        case 'signer_fio': set('director_fio', value); break;
        case 'signer_pos': set('director_pos', value); break;
        case 'signer_pos_fio':
        case 'sig_title': {
          const sp = splitPosFio(value);
          if (sp) { set('director_pos', sp.pos); set('director_fio', sp.fio); }
          else if (/^[А-ЯЁ][а-яё-]+\s+[А-ЯЁ]/.test(value)) {
            if (r.bare && !/^(руководитель|подписант|уполномоч|представитель)/i.test(label)) set('director_pos', label.replace(/[:\s]+$/, '').replace(/^./, c => c.toUpperCase()));
            set('director_fio', value);
          }
          break;
        }
        case 'addr_any': set('addr_legal', value); break;
        case 'inperson': case 'date': break;
        default: set(r.key, value);
      }
    }

    const all = lines.join('\n');
    // запасные способы по формату чисел
    if (!card.inn) { const m = all.match(/ИНН[^\d]{0,15}(\d{12}|\d{10})(?!\d)/i); if (m) set('inn', m[1]); }
    if (!card.kpp) { const m = all.match(/КПП[^\d]{0,15}(\d{9})(?!\d)/i); if (m) set('kpp', m[1]); }
    if (!card.ogrn) { const m = all.match(/ОГРН(?:ИП)?[^\d]{0,15}(\d{15}|\d{13})(?!\d)/i); if (m) set('ogrn', m[1]); }
    if (!card.bik) { const m = all.match(/БИК[^\d]{0,15}(\d{9})(?!\d)/i); if (m) set('bik', m[1]); }
    if (!card.ks || !card.rs) {
      for (const m of all.matchAll(/(?<!\d)(\d{20})(?!\d)/g)) {
        if (m[1].startsWith('301')) set('ks', m[1]); else set('rs', m[1]);
      }
    }
    if (!card.email) { const m = all.match(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/); if (m) set('email', m[0]); }
    if (!card.phone) { const m = all.match(/(?:\+7|8)[\s(-]*\d{3}[\s)-]*\d{3}[\s-]*\d{2}[\s-]*\d{2}/); if (m) set('phone', m[0]); }
    if (!card.name_full) {
      const m = all.match(/(?:Общество с ограниченной ответственностью|Акционерное общество|Публичное акционерное общество)\s*[«"“][^»"”\n]+[»"”]|(?:ООО|АО|ПАО|ЗАО|ОАО|НАО|АНО|МУП|ГУП|ФГУП)\s*[«"“][^»"”\n]+[»"”]|Индивидуальный предприниматель\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё-]+){1,2}|ИП\s+[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2}/);
      if (m) set('name_full', m[0]);
    }
    if (!card.director_fio) {
      const m = all.match(/(Генеральный директор|Исполнительный директор|Директор|Руководитель|Президент)[:\s–—-]+([А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ][а-яё.-]+){1,2})/);
      if (m) { set('director_pos', m[1]); set('director_fio', m[2]); }
    }

    // производные значения
    if (!card.name_short && card.name_full) {
      let s = card.name_full;
      for (const [re, to] of OPF) s = s.replace(re, to);
      set('name_short', s.replace(/\s+/g, ' ').replace(/[“”"]/g, m => m)); 
    }
    if (!card.name_full && card.name_short) set('name_full', card.name_short);
    const isIp = /^(ип|индивидуальный предприниматель)\s/i.test(card.name_full || card.name_short || '');
    if (isIp && !card.director_fio) {
      const m = (card.name_full || card.name_short).match(/^(?:ИП|Индивидуальный предприниматель)\s+(.+)$/i);
      if (m) { set('director_fio', m[1]); set('director_pos', 'Индивидуальный предприниматель'); }
    }
    if (card.director_pos && !card.director_pos_gen) set('director_pos_gen', genitivePosition(card.director_pos));
    if (!card.basis && card.director_fio && !isIp) set('basis', 'Устава');
    return card;
  }

  /* ---------------------------------------------------------------
     4. Контекст заполнения: значения, режимы, журнал
     --------------------------------------------------------------- */
  const ALLOWED_MIN = new Set(['name_full', 'name_short', 'signer_fio', 'signer_fio_short', 'signer_pos', 'signer_pos_fio', 'sig_title', 'inperson']);

  function makeCtx(card, opts) {
    opts = Object.assign({ mode: 'full', highlight: true, docDate: '', fallbackAddr: true }, opts || {});
    const ctx = {
      card, opts, handled: new Set(), items: [], stats: { filled: 0, highlighted: 0 },
      allowed(key) { return opts.mode !== 'min' || ALLOWED_MIN.has(key); },
      value(key) {
        const c = card;
        switch (key) {
          case 'name_full': return c.name_full || c.name_short || '';
          case 'name_short': return c.name_short || c.name_full || '';
          case 'inn_kpp': return c.inn ? (c.kpp ? c.inn + '/' + c.kpp : c.inn) : '';
          case 'addr_any': case 'addr_legal': return c.addr_legal || '';
          case 'addr_post': return c.addr_post || (opts.fallbackAddr ? c.addr_legal : '') || '';
          case 'addr_fact': return c.addr_fact || (opts.fallbackAddr ? c.addr_legal : '') || '';
          case 'signer_fio': return c.director_fio || '';
          case 'signer_fio_short': return shortFio(c.director_fio);
          case 'signer_pos': return c.director_pos || '';
          case 'signer_pos_fio': case 'sig_title': return c.director_fio ? [c.director_pos, c.director_fio].filter(Boolean).join(' ') : '';
          case 'inperson': return (c.director_pos_gen && c.director_fio_gen) ? c.director_pos_gen + ' ' + c.director_fio_gen : '';
          case 'date': return dateShort(opts.docDate);
          default: return c[key] || '';
        }
      },
      dateLong() { return dateLong(opts.docDate); },
      filled(key, val) { ctx.stats.filled++; ctx.items.push({ kind: 'fill', key, val }); },
      hl(what) { ctx.stats.highlighted++; ctx.items.push({ kind: 'hl', key: what }); },
    };
    return ctx;
  }

  const KEY_TITLES = {
    name_full: 'Наименование организации', name_short: 'Сокращённое наименование', inn: 'ИНН', kpp: 'КПП', inn_kpp: 'ИНН/КПП', ogrn: 'ОГРН', okpo: 'ОКПО',
    oktmo: 'ОКТМО', okved: 'ОКВЭД', addr_legal: 'Юридический адрес', addr_any: 'Адрес', addr_post: 'Почтовый адрес', addr_fact: 'Фактический адрес',
    email: 'E-mail', phone: 'Телефон', fax: 'Факс', site: 'Сайт', bank_name: 'Банк', bik: 'БИК', rs: 'Расчётный счёт', ks: 'Корр. счёт',
    signer_fio: 'ФИО подписанта', signer_fio_short: 'ФИО (сокр.)', signer_pos: 'Должность подписанта', signer_pos_fio: 'Должность и ФИО',
    sig_title: 'Руководитель', inperson: 'В лице (род. падеж)', basis: 'Основание', tax: 'Налогообложение', date: 'Дата', unknown: 'Неизвестное поле',
  };
  const title = k => KEY_TITLES[k] || k;

  /* ---------------------------------------------------------------
     5. DOCX: низкоуровневые операции с прогонами (w:r)
     --------------------------------------------------------------- */
  const mkW = (doc, name) => doc.createElementNS(NS_W, 'w:' + name);

  function runText(r) {
    let s = '';
    for (let c = r.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      const n = lname(c);
      if (n === 't') s += c.textContent; else if (n === 'tab') s += '\t';
    }
    return s;
  }
  function isSimpleRun(r) {
    for (let c = r.firstChild; c; c = c.nextSibling) {
      if (c.nodeType !== 1) continue;
      if (!['rPr', 't', 'tab'].includes(lname(c))) return false;
    }
    return true;
  }
  function setRunText(r, text) {
    const doc = r.ownerDocument;
    for (const c of Array.from(r.childNodes)) if (c.nodeType === 1 && ['t', 'tab'].includes(lname(c))) r.removeChild(c);
    text.split('\t').forEach((part, i) => {
      if (i > 0) r.appendChild(mkW(doc, 'tab'));
      if (part !== '') {
        const t = mkW(doc, 't');
        t.setAttributeNS(NS_XML, 'xml:space', 'preserve');
        t.appendChild(doc.createTextNode(part));
        r.appendChild(t);
      }
    });
  }
  const RECURSE = new Set(['hyperlink', 'ins', 'smartTag', 'fldSimple', 'sdt', 'sdtContent', 'customXml']);
  function paraRuns(p) {
    const out = [];
    (function walk(n) {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const nm = lname(c);
        if (nm === 'r') out.push(c); else if (RECURSE.has(nm)) walk(c);
      }
    })(p);
    return out;
  }
  function mapParagraph(p) {
    let pos = 0;
    const items = paraRuns(p).map(r => { const t = runText(r); const o = { r, t, s: pos, e: pos + t.length }; pos += t.length; return o; });
    return { items, text: items.map(i => i.t).join('') };
  }
  function splitAt(p, k) {
    for (const it of mapParagraph(p).items) {
      if (it.s < k && k < it.e) {
        if (!isSimpleRun(it.r)) return;
        const cl = it.r.cloneNode(true);
        setRunText(it.r, it.t.slice(0, k - it.s));
        setRunText(cl, it.t.slice(k - it.s));
        it.r.parentNode.insertBefore(cl, it.r.nextSibling);
        return;
      }
    }
  }
  function rangeRuns(p, s, e) {
    splitAt(p, s); splitAt(p, e);
    return mapParagraph(p).items.filter(it => it.t.length > 0 && it.s < e && it.e > s).map(it => it.r);
  }
  const HL_BEFORE = ['u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  function addHighlight(run) {
    const doc = run.ownerDocument;
    let rPr = firstKid(run, 'rPr');
    if (!rPr) { rPr = mkW(doc, 'rPr'); run.insertBefore(rPr, run.firstChild); }
    for (const h of kids(rPr, 'highlight')) rPr.removeChild(h);
    const hl = mkW(doc, 'highlight');
    hl.setAttributeNS(NS_W, 'w:val', 'yellow');
    const before = Array.from(rPr.childNodes).find(c => c.nodeType === 1 && HL_BEFORE.includes(lname(c)));
    if (before) rPr.insertBefore(hl, before); else rPr.appendChild(hl);
  }
  function stripFmt(rPr, names) {
    for (const c of Array.from(rPr.childNodes)) if (c.nodeType === 1 && names.includes(lname(c))) rPr.removeChild(c);
  }
  function cleanRPr(rPr, extra) {
    if (!rPr) return null;
    const c = rPr.cloneNode(true);
    stripFmt(c, ['ins', 'del', 'moveFrom', 'moveTo', 'rPrChange', 'highlight'].concat(extra || []));
    return c;
  }
  function makeRun(doc, text, rPr) {
    const r = mkW(doc, 'r');
    if (rPr) r.appendChild(rPr);
    setRunText(r, text);
    return r;
  }
  function fillRange(p, s, e, val) {
    const runs = rangeRuns(p, s, e);
    if (!runs.length) return false;
    setRunText(runs[0], val);
    const rp = firstKid(runs[0], 'rPr');
    if (rp) stripFmt(rp, ['highlight']);
    for (let i = 1; i < runs.length; i++) runs[i].parentNode.removeChild(runs[i]);
    return true;
  }
  function highlightRange(p, s, e) {
    const runs = rangeRuns(p, s, e);
    runs.forEach(addHighlight);
    return runs.length > 0;
  }
  const NBSP6 = '\u00A0\u00A0\u00A0\u00A0\u00A0\u00A0';
  function lastTextRunRPr(p) {
    const runs = paraRuns(p).filter(r => runText(r).trim());
    const r = runs[runs.length - 1];
    return r ? firstKid(r, 'rPr') : null;
  }
  function appendToParagraph(doc, p, text, highlight) {
    const rPr = cleanRPr(lastTextRunRPr(p), ['b', 'bCs']);
    const run = makeRun(doc, text, rPr);
    if (highlight) addHighlight(run);
    p.appendChild(run);
  }

  /* ---------------------------------------------------------------
     6. DOCX: пустые места в абзаце
     --------------------------------------------------------------- */
  const BLANK_RE = /_{3,}|\.{5,}|…{2,}/g;
  const PH_RE = /\[([^\]\n]{2,60})\]|\{\{([^}\n]{2,60})\}\}|(?<![\w/])<([^<>\n]{2,60})>/g;
  const DATE_RE = /[«"“]\s*_{1,}\s*[»"”]\s*_{3,}\s*20_{1,}\s*(?:г\.?|года)?/;

  function findBlanks(map, p) {
    const ranges = [];
    for (const m of map.text.matchAll(BLANK_RE)) ranges.push({ s: m.index, e: m.index + m[0].length });
    // подчёркнутые пробелы (линия для вписывания)
    for (const it of map.items) {
      if (it.t.length >= 3 && !it.t.trim()) {
        const rPr = firstKid(it.r, 'rPr');
        const u = rPr && firstKid(rPr, 'u');
        if (u && (u.getAttributeNS(NS_W, 'val') || u.getAttribute('w:val')) !== 'none') ranges.push({ s: it.s, e: it.e });
      }
    }
    for (const m of map.text.matchAll(PH_RE)) {
      const inner = m[1] || m[2] || m[3];
      ranges.push({ s: m.index, e: m.index + m[0].length, ph: true, inner });
    }
    ranges.sort((a, b) => a.s - b.s || b.e - a.e);
    const out = [];
    for (const r of ranges) {
      const last = out[out.length - 1];
      if (last && r.s < last.e) continue;
      out.push(r);
    }
    return out;
  }

  const trimEnd = s => s.replace(/[\s\u00A0]+$/, '');
  const trimStart = s => s.replace(/^[\s\u00A0]+/, '');

  function resolveContext(before, after, prevText, nextText) {
    let b = trimEnd(before), a = trimStart(after);
    const lineIsBlankOnly = !b && !a;       // абзац состоит только из линии
    if (!b && prevText && prevText.length < 160) b = trimEnd(prevText);
    if (lineIsBlankOnly && nextText) a = trimStart(nextText);   // подсказка «(подпись)» на следующей строке
    // подпись / печать — не заполняем и не подсвечиваем
    if (/^[\s)\]/(]*(?:подпис|м\.\s*п)/i.test(a)) return { act: 'skip' };
    if (/(?:м\.\s*п\.?|подпись\s*[:\-–]?)\s*$/i.test(b.slice(-30))) return { act: 'skip' };
    // подсказка в скобках под линией
    const hm = a.match(/^[)\]/]*\s*\(\s*([^)]{0,70})\)/);
    if (hm) {
      const h = hm[1].toLowerCase();
      if (/ф\.?\s*и\.?\s*о|фамилия|расшифровк|инициал/.test(h)) return { act: 'key', key: /расшифровк|инициал/.test(h) ? 'signer_fio_short' : 'signer_fio' };
      if (/должност/.test(h)) return { act: 'key', key: 'signer_pos' };
      if (/наименовани|организаци/.test(h)) return { act: 'key', key: 'name_full' };
      if (/подпис/.test(h)) return { act: 'skip' };
    }
    // «/ ______ /» — расшифровка подписи
    if (/\/\s*$/.test(b)) return { act: 'key', key: 'signer_fio_short' };
    if (/^\//.test(a)) return { act: 'skip' };      // линия подписи перед «/ФИО/»
    const r = matchKeyEx(b.slice(-120));
    if (!r) return { act: 'unknown', label: b.slice(-40) };
    if (r.key === 'sig_title') return { act: 'skip' };   // «Директор ________» в блоке подписей
    return { act: 'key', key: r.key };
  }

  function siblingParaText(p, dir) {
    let n = dir < 0 ? p.previousSibling : p.nextSibling;
    while (n && n.nodeType !== 1) n = dir < 0 ? n.previousSibling : n.nextSibling;
    return n && lname(n) === 'p' ? mapParagraph(n).text : '';
  }

  function applyResolution(p, b, r, ctx) {
    if (r.act === 'skip') return;
    if (r.act === 'unknown') {
      if (ctx.opts.mode !== 'min' && ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl(r.label || 'unknown'); }
      return;
    }
    if (!ctx.allowed(r.key)) return;
    const val = ctx.value(r.key);
    if (val) { fillRange(p, b.s, b.e, val); ctx.filled(r.key, val); }
    else if (ctx.opts.highlight) { highlightRange(p, b.s, b.e); ctx.hl(r.key); }
  }

  function processParagraph(doc, p, ctx) {
    let map = mapParagraph(p);
    if (!map.text.trim()) return;

    // дата вида «__» ______ 20__ г.
    let dateRange = null;
    const dm = map.text.match(DATE_RE);
    if (dm) {
      dateRange = { s: dm.index, e: dm.index + dm[0].length };
      const dv = ctx.dateLong();
      if (dv) { fillRange(p, dateRange.s, dateRange.e, dv); ctx.filled('date', dv); dateRange = null; }
      else if (ctx.opts.mode !== 'min' && ctx.opts.highlight) { highlightRange(p, dateRange.s, dateRange.e); ctx.hl('date'); }
      map = mapParagraph(p);
    }

    let blanks = findBlanks(map, p);
    if (dateRange) blanks = blanks.filter(b => !(b.s >= dateRange.s && b.e <= dateRange.e));
    const prevText = siblingParaText(p, -1), nextText = siblingParaText(p, +1);

    if (blanks.length) {
      for (let i = blanks.length - 1; i >= 0; i--) {
        const b = blanks[i];
        const cur = mapParagraph(p).text;
        let r;
        if (b.ph) { const k = matchKeyEx(b.inner, { noBare: true }); r = k ? { act: 'key', key: k.key } : { act: 'skip' }; }
        else r = resolveContext(cur.slice(0, b.s), cur.slice(b.e), prevText, nextText);
        applyResolution(p, b, r, ctx);
      }
      return;
    }

    // «Подпись:» без линии — дописываем значение после двоеточия
    const t = trimEnd(map.text);
    if (t.length <= 140 && /[:：]$/.test(t)) {
      const r = matchKeyEx(t.slice(0, -1));
      if (!r || r.key === 'sig_title' || !ctx.allowed(r.key)) return;
      const val = ctx.value(r.key);
      if (val) { appendToParagraph(doc, p, ' ' + val, false); ctx.filled(r.key, val); }
      else if (ctx.opts.highlight) { appendToParagraph(doc, p, ' ' + NBSP6, true); ctx.hl(r.key); }
    }
  }

  /* ---------------------------------------------------------------
     7. DOCX: таблицы
     --------------------------------------------------------------- */
  const cellText = tc => byNS(tc, NS_W, 't').map(t => t.textContent).join('');
  const isBlankText = s => /^[\s\u00A0_.…]*$/.test(s);
  const hasObjects = tc => byNS(tc, NS_W, 'drawing').length > 0 || byNS(tc, NS_W, 'tbl').length > 0;

  function writeCell(doc, tc, val, o) {
    o = o || {};
    const p = kids(tc, 'p')[0];
    if (!p) return;
    const runs = paraRuns(p);
    let rPr = null;
    if (runs.length) rPr = cleanRPr(firstKid(runs[0], 'rPr'));
    if (!rPr) { const pPr = firstKid(p, 'pPr'); const mark = pPr && firstKid(pPr, 'rPr'); if (mark) rPr = cleanRPr(mark); }
    if (!rPr && o.labelTc) {
      const lp = kids(o.labelTc, 'p')[0];
      const lr = lp && paraRuns(lp).find(r => runText(r).trim());
      const rp = lr && firstKid(lr, 'rPr');
      if (rp) rPr = cleanRPr(rp, ['b', 'bCs', 'u']);
    }
    for (const r of runs) r.parentNode.removeChild(r);
    const run = makeRun(doc, o.highlight ? NBSP6 : val, rPr);
    if (o.highlight) addHighlight(run);
    p.appendChild(run);
  }

  const looksLikeField = l => /[:：]\s*$/.test(l) || (l.split(/\s+/).length >= 2 && !/\d{3,}/.test(l));

  function handleTarget(doc, ctx, tcLabel, tcTarget, label, noBare) {
    const r = matchKeyEx(label, { noBare });
    if (r) {
      if (!ctx.allowed(r.key)) return;
      const val = ctx.value(r.key);
      if (val) { writeCell(doc, tcTarget, val, { labelTc: tcLabel }); ctx.filled(r.key, val); }
      else if (ctx.opts.highlight) { writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true }); ctx.hl(r.key); }
    } else if (ctx.opts.mode !== 'min' && ctx.opts.highlight && looksLikeField(label)) {
      writeCell(doc, tcTarget, '', { labelTc: tcLabel, highlight: true });
      ctx.hl(label.slice(0, 40));
    }
  }

  function processTables(doc, ctx) {
    const rows = byNS(doc, NS_W, 'tr');
    for (let ri = 0; ri < rows.length; ri++) {
      const tr = rows[ri];
      const cells = kids(tr, 'tc');
      // 1. «подпись | пустая ячейка»
      for (let i = 0; i < cells.length - 1; i++) {
        const label = cellText(cells[i]).trim();
        if (!label || isBlankText(label) || label.length > 200 || BLANK_RE.test(label)) { BLANK_RE.lastIndex = 0; continue; }
        BLANK_RE.lastIndex = 0;
        const next = cells[i + 1];
        const nextTxt = cellText(next);
        if (isBlankText(nextTxt) && !hasObjects(next)) {
          handleTarget(doc, ctx, cells[i], next, label, false);
          ctx.handled.add(cells[i]); ctx.handled.add(next);
          i++;
        } else if (matchKeyEx(label) || /[:：]\s*$/.test(label)) {
          ctx.handled.add(cells[i]);   // уже заполнено — не дописываем
        }
      }
      // 2. шапка в этой строке, пустая строка под ней
      const nextRow = rows[ri + 1];
      if (nextRow && nextRow.parentNode === tr.parentNode) {
        const below = kids(nextRow, 'tc');
        const allHeaders = cells.length >= 1 && cells.every(c => !isBlankText(cellText(c)));
        const belowEmpty = below.length === cells.length && below.every(c => { const t = cellText(c).trim(); return isBlankText(t) || /^\d+[.)]?$/.test(t); }) && below.some(c => isBlankText(cellText(c)));
        if (allHeaders && belowEmpty && !cells.some(c => ctx.handled.has(c))) {
          cells.forEach((hc, j) => {
            const tcb = below[j];
            if (!isBlankText(cellText(tcb)) || ctx.handled.has(tcb)) return;
            handleTarget(doc, ctx, hc, tcb, cellText(hc).trim(), true);
            ctx.handled.add(tcb);
          });
          cells.forEach(c => ctx.handled.add(c));
        }
      }
    }
  }

  function insideHandled(p, ctx) {
    for (let n = p.parentNode; n; n = n.parentNode) if (ctx.handled.has(n)) return true;
    return false;
  }

  /* ---------------------------------------------------------------
     8. DOCX: главная функция
     --------------------------------------------------------------- */
  async function fillDocx(buf, card, opts) {
    const zip = await JSZip.loadAsync(buf);
    if (!zip.file('word/document.xml')) throw new Error('Не найден word/document.xml — это не DOCX');
    const ctx = makeCtx(card, opts);
    const parts = ['word/document.xml'].concat(Object.keys(zip.files).filter(n => /^word\/(header|footer)\d*\.xml$/.test(n)));
    for (const part of parts) {
      const doc = parseXml(await zip.file(part).async('string'));
      processTables(doc, ctx);
      for (const p of byNS(doc, NS_W, 'p')) {
        if (insideHandled(p, ctx)) continue;
        try { processParagraph(doc, p, ctx); } catch (e) { ctx.items.push({ kind: 'err', key: String(e && e.message || e) }); }
      }
      zip.file(part, serialize(doc));
    }
    const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { data, stats: ctx.stats, items: ctx.items.map(i => ({ ...i, key: i.kind === 'hl' && !KEY_TITLES[i.key] ? i.key : title(i.key) })) };
  }

  /* ---------------------------------------------------------------
     9. XLSX
     --------------------------------------------------------------- */
  const colToNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const numToCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const splitRef = ref => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return m ? { col: colToNum(m[1]), row: +m[2] } : null; };

  async function readShared(zip) {
    const f = zip.file('xl/sharedStrings.xml');
    if (!f) return [];
    const doc = parseXml(await f.async('string'));
    return byNS(doc, NS_S, 'si').map(si => byNS(si, NS_S, 't').map(t => t.textContent).join(''));
  }
  function xlCellText(c, sst) {
    const t = c.getAttribute('t');
    const v = firstKid(c, 'v');
    if (t === 's') return v ? (sst[+v.textContent] || '') : '';
    if (t === 'inlineStr') return byNS(c, NS_S, 't').map(x => x.textContent).join('');
    if (t === 'str') return v ? v.textContent : '';
    return v ? v.textContent : '';
  }
  function setXlText(doc, c, text) {
    for (const ch of Array.from(c.childNodes)) c.removeChild(ch);
    c.setAttribute('t', 'inlineStr');
    const is = doc.createElementNS(NS_S, 'is');
    const t = doc.createElementNS(NS_S, 't');
    t.setAttributeNS(NS_XML, 'xml:space', 'preserve');
    t.appendChild(doc.createTextNode(text));
    is.appendChild(t); c.appendChild(is);
  }

  function makeHighlightStyler(stylesDoc) {
    if (!stylesDoc) return null;
    const fills = byNS(stylesDoc, NS_S, 'fills')[0];
    const xfsEl = byNS(stylesDoc, NS_S, 'cellXfs')[0];
    if (!fills || !xfsEl) return null;
    let fillId = null;
    const cache = {};
    return function (oldS) {
      oldS = oldS || 0;
      if (cache[oldS] != null) return cache[oldS];
      if (fillId == null) {
        const fill = stylesDoc.createElementNS(NS_S, 'fill');
        const pf = stylesDoc.createElementNS(NS_S, 'patternFill'); pf.setAttribute('patternType', 'solid');
        const fg = stylesDoc.createElementNS(NS_S, 'fgColor'); fg.setAttribute('rgb', 'FFFFFF00');
        const bg = stylesDoc.createElementNS(NS_S, 'bgColor'); bg.setAttribute('indexed', '64');
        pf.appendChild(fg); pf.appendChild(bg); fill.appendChild(pf); fills.appendChild(fill);
        fillId = kids(fills, 'fill').length - 1;
        fills.setAttribute('count', String(fillId + 1));
      }
      const xfs = kids(xfsEl, 'xf');
      const base = xfs[oldS] || xfs[0];
      const xf = base ? base.cloneNode(true) : stylesDoc.createElementNS(NS_S, 'xf');
      xf.setAttribute('fillId', String(fillId)); xf.setAttribute('applyFill', '1');
      xfsEl.appendChild(xf);
      const idx = kids(xfsEl, 'xf').length - 1;
      xfsEl.setAttribute('count', String(idx + 1));
      return (cache[oldS] = idx);
    };
  }

  function processSheet(doc, sst, ctx, styler) {
    const merges = {};
    for (const m of byNS(doc, NS_S, 'mergeCell')) {
      const [a, b] = (m.getAttribute('ref') || '').split(':');
      const A = splitRef(a), B = b && splitRef(b);
      if (A && B) merges[a] = B.col;
    }
    const rows = byNS(doc, NS_S, 'row');
    for (const row of rows) {
      const rn = +row.getAttribute('r');
      const cellEls = kids(row, 'c');
      const byCol = {};
      cellEls.forEach(c => { const ref = splitRef(c.getAttribute('r') || ''); if (ref) byCol[ref.col] = c; });
      for (const c of cellEls) {
        const ref = splitRef(c.getAttribute('r') || '');
        if (!ref) continue;
        const label = xlCellText(c, sst).trim();
        if (!label || !/[A-Za-zА-Яа-яЁё]/.test(label) || label.length > 200) continue;

        // «ИНН: ______» внутри одной ячейки
        const bm = /_{3,}\s*$/.exec(label);
        if (bm) {
          const r = matchKeyEx(label.slice(0, bm.index));
          if (r && r.key !== 'sig_title' && ctx.allowed(r.key)) {
            const val = ctx.value(r.key);
            if (val) { setXlText(doc, c, label.slice(0, bm.index).replace(/\s*$/, ' ') + val); ctx.filled(r.key, val); }
            else if (ctx.opts.highlight && styler) { c.setAttribute('s', String(styler(+c.getAttribute('s')))); ctx.hl(r.key); }
          }
          continue;
        }

        const targetCol = (merges[c.getAttribute('r')] || ref.col) + 1;
        let tgt = byCol[targetCol];
        const tgtBlank = !tgt || isBlankText(xlCellText(tgt, sst));
        if (!tgtBlank) continue;
        const r = matchKeyEx(label);
        if (r) {
          if (!ctx.allowed(r.key)) continue;
          const val = ctx.value(r.key);
          if (!tgt) {
            tgt = doc.createElementNS(NS_S, 'c');
            tgt.setAttribute('r', numToCol(targetCol) + rn);
            const after = cellEls.find(x => { const q = splitRef(x.getAttribute('r') || ''); return q && q.col > targetCol; });
            if (after) row.insertBefore(tgt, after); else row.appendChild(tgt);
            byCol[targetCol] = tgt;
          }
          if (val) { setXlText(doc, tgt, val); ctx.filled(r.key, val); }
          else if (ctx.opts.highlight && styler) { tgt.setAttribute('s', String(styler(+tgt.getAttribute('s')))); ctx.hl(r.key); }
        } else if (ctx.opts.mode !== 'min' && ctx.opts.highlight && tgt && styler && looksLikeField(label)) {
          tgt.setAttribute('s', String(styler(+tgt.getAttribute('s'))));
          ctx.hl(label.slice(0, 40));
        }
      }
    }
  }

  async function fillXlsx(buf, card, opts) {
    const zip = await JSZip.loadAsync(buf);
    const ctx = makeCtx(card, opts);
    const sst = await readShared(zip);
    const stylesFile = zip.file('xl/styles.xml');
    const stylesDoc = stylesFile ? parseXml(await stylesFile.async('string')) : null;
    const styler = makeHighlightStyler(stylesDoc);
    const sheets = Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    if (!sheets.length) throw new Error('В файле нет листов');
    for (const name of sheets) {
      const doc = parseXml(await zip.file(name).async('string'));
      processSheet(doc, sst, ctx, styler);
      zip.file(name, serialize(doc));
    }
    if (stylesDoc) zip.file('xl/styles.xml', serialize(stylesDoc));
    const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    return { data, stats: ctx.stats, items: ctx.items.map(i => ({ ...i, key: i.kind === 'hl' && !KEY_TITLES[i.key] ? i.key : title(i.key) })) };
  }

  /* ---------------------------------------------------------------
     10. Чтение карты партнёра из DOCX / XLSX → текст
     --------------------------------------------------------------- */
  async function docxToText(buf) {
    const zip = await JSZip.loadAsync(buf);
    const f = zip.file('word/document.xml');
    if (!f) throw new Error('Не найден word/document.xml');
    const doc = parseXml(await f.async('string'));
    const body = byNS(doc, NS_W, 'body')[0];
    const out = [];
    const ptext = p => byNS(p, NS_W, 't').map(t => t.textContent).join('');
    (function walk(node) {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        const n = lname(c);
        if (n === 'p') { const t = ptext(c).trim(); if (t) out.push(t); }
        else if (n === 'tbl') {
          for (const tr of kids(c, 'tr')) {
            const cells = kids(tr, 'tc').map(tc => kids(tc, 'p').map(ptext).join(' ').replace(/\s+/g, ' ').trim());
            if (cells.some(Boolean)) out.push(cells.join(' | '));
          }
        } else if (n === 'sdt' || n === 'sdtContent') walk(c);
      }
    })(body);
    return out.join('\n');
  }
  async function xlsxToText(buf) {
    const zip = await JSZip.loadAsync(buf);
    const sst = await readShared(zip);
    const out = [];
    for (const name of Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n))) {
      const doc = parseXml(await zip.file(name).async('string'));
      for (const row of byNS(doc, NS_S, 'row')) {
        const cells = kids(row, 'c').map(c => xlCellText(c, sst).trim()).filter(Boolean);
        if (cells.length) out.push(cells.join(' | '));
      }
    }
    return out.join('\n');
  }

  async function fillFile(name, buf, card, opts) {
    const ext = (name.split('.').pop() || '').toLowerCase();
    if (ext === 'docx' || ext === 'dotx' || ext === 'docm') return fillDocx(buf, card, opts);
    if (ext === 'xlsx' || ext === 'xlsm') return fillXlsx(buf, card, opts);
    const e = new Error('Формат .' + ext + ' не поддерживается. Сохраните бланк как .docx или .xlsx');
    e.unsupported = true;
    throw e;
  }

  return { parseCardText, fillDocx, fillXlsx, fillFile, docxToText, xlsxToText, matchKeyEx, shortFio, genitivePosition };
});
