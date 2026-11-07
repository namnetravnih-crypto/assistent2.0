/* =====================================================================
   Ассистент — «Техническое предложение»
   Алгоритмы: из ТЗ берутся показатели («не менее», «не более», «не шире», «от … до» …),
   в техническом предложении они заменяются на конкретные значения.
   Структура, шрифты, таблицы и стиль документа не меняются — правится только текст.
   ===================================================================== */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./engine.js'));
  else root.AssistantTech = factory(root.AssistantEngine);
})(typeof self !== 'undefined' ? self : this, function (Engine) {
  'use strict';
  const I = Engine.internals, { NS_W, NS_S, parseXml, serialize, kids, firstKid, byNS, lc, clean } = I;

  /* =====================================================================
     1. ПРАВИЛА ЗАМЕНЫ ПОКАЗАТЕЛЕЙ
     =====================================================================
     Правило закупок: участник указывает КОНКРЕТНОЕ значение, а не «не менее» / «не более».
       «не менее / не меньше / не ниже / не уже / не тоньше / не короче / не легче / минимум / ≥ / от N»
            → значение N (граница входит в допустимое — подходит само число);
       «не более / не больше / не выше / не шире / не толще / не длиннее / не тяжелее / максимум / ≤ / до N»
            → значение N (граница входит);
       «более / больше / выше / шире / толще / длиннее / тяжелее / свыше / > N»
            → значение СТРОГО больше N: N + шаг (для целых +1, для 2,5 → 2,6);
       «менее / меньше / ниже / уже / тоньше / короче / легче / < N»
            → значение СТРОГО меньше N: N − шаг;
       «от N до M», «N–M», «не менее N и не более M»
            → одно значение внутри диапазона (по настройке: нижняя / верхняя граница / середина),
              а для показателей-диапазонов («диапазон измерений», «диапазон регулирования») диапазон остаётся;
       «N ± d», «около N», «приблизительно N», «≈ N» → N.
     ===================================================================== */
  const NUM = '[+\\-−–]?\\d+(?:(?:[\\u00a0\\u202f ])\\d{3}(?!\\d))*(?:[.,]\\d+)?';
  const B = '(?<![а-яёa-z0-9])';
  const Q = {
    GE: 'не\\s+менее(?:\\s+чем)?|не\\s+меньше(?:\\s+чем)?|не\\s+ниже|не\\s+уже|не\\s+тоньше|не\\s+короче|не\\s+легче|не\\s+слабее|не\\s+мельче|не\\s+хуже|как\\s+минимум|по\\s+меньшей\\s+мере|минимум|мин\\.|≥|>=|=>',
    LE: 'не\\s+более(?:\\s+чем)?|не\\s+больше(?:\\s+чем)?|не\\s+выше|не\\s+шире|не\\s+толще|не\\s+длиннее|не\\s+тяжелее|не\\s+глубже|не\\s+крупнее|не\\s+превыша[а-яё]*|не\\s+превосход[а-яё]*|максимум|макс\\.|≤|<=|=<',
    GT: 'более(?:\\s+чем)?|больше(?:\\s+чем)?|выше|шире|толще|длиннее|тяжелее|глубже|крупнее|свыше|превыша[а-яё]*|>',
    LT: 'менее(?:\\s+чем)?|меньше(?:\\s+чем)?|ниже|уже|тоньше|короче|легче|мельче|<',
  };
  const MONTHS_RE = '(?:январ|феврал|март|апрел|ма[яй]|июн|июл|август|сентябр|октябр|ноябр|декабр)[а-яё]*';
  const UNIT_TOK = '[%°a-zа-яё/²³µмк.]{1,12}';
  // один общий разбор: порядок альтернатив = приоритет
  const MASTER = new RegExp([
    // диапазон «в диапазоне от N до M»
    `(?<range>${B}(?:(?:в\\s+)?(?:диапазоне|пределах|интервале)\\s+)?от\\s+(?<r1>${NUM})\\s*(?:${UNIT_TOK}\\s*)?(?:до|по|-|–|—)\\s*(?<r2>${NUM}))`,
    // «N ± d»
    `(?<tol>(?<t1>${NUM})\\s*(?:±|\\+\\s*/\\s*[-−–]|\\+[-−–])\\s*${NUM})`,
    // «N–M» через тире / многоточие
    `(?<dash>(?<![\\d.,/-])(?<d1>\\d+(?:[.,]\\d+)?)\\s*(?:–|—|…|\\.\\.\\.)\\s*(?<d2>\\d+(?:[.,]\\d+)?)(?![\\d.,]))`,
    // «N-M» через дефис — только если дальше идёт единица измерения
    `(?<hyph>(?<![\\d.,/-])(?<h1>\\d+(?:[.,]\\d+)?)-(?<h2>\\d+(?:[.,]\\d+)?)(?=\\s*(?:мм|см|м|км|кг|г|т|л|мл|вт|квт|в|а|ма|гц|кгц|мгц|°c|°с|%|шт|мин|час|ч|сек|с|мкм|бар|мпа|кпа|па|об|дб|лм|лк)(?![а-яёa-z])))`,
    // приблизительно
    `(?<approx>${B}(?:около|примерно|приблизительно|порядка|≈|~)\\s*(?<a1>${NUM}))`,
    // не менее / не более / более / менее
    `(?<ge>${B}(?:${Q.GE})[\\s:]*(?<g1>${NUM}))`,
    `(?<le>${B}(?:${Q.LE})[\\s:]*(?<l1>${NUM}))`,
    `(?<gt>${B}(?:${Q.GT})[\\s:]*(?<q1>${NUM}))`,
    `(?<lt>${B}(?:${Q.LT})[\\s:]*(?<w1>${NUM}))`,
    // «N и более», «N и менее»
    `(?<nge>(?<![\\d.,/-])(?<n1>${NUM})(?=\\s*(?:${UNIT_TOK}\\s*)?(?:и\\s+(?:более|больше|выше)))\\s*(?:${UNIT_TOK}\\s*)?и\\s+(?:более|больше|выше))`,
    `(?<nle>(?<![\\d.,/-])(?<m1>${NUM})(?=\\s*(?:${UNIT_TOK}\\s*)?(?:и\\s+(?:менее|меньше|ниже)))\\s*(?:${UNIT_TOK}\\s*)?и\\s+(?:менее|меньше|ниже))`,
    // «от N» / «до N» (без пары)
    `(?<from>${B}от\\s+(?<f1>${NUM})(?![.,]?\\d))`,
    `(?<upto>${B}до\\s+(?<u1>${NUM})(?![.,]?\\d)(?!\\s*(?:г\\b|года|${MONTHS_RE})))`,
  ].join('|'), 'giu');

  // отдельно: слова-«сироты» без числа, которые нужно выделить жёлтым (не удалось подставить число)
  const ORPHAN = new RegExp(`${B}(?:${Q.GE}|${Q.LE})(?![а-яёa-z])(?![\\s:]*[+\\-−–]?\\d)`, 'giu');

  const CHAR_RANGE = /диапазон|интервал|предел[ыа]?\s+(?:измерен|регулир|настрой|рабоч)|регулир|настройк[аи]|рабоч[а-яё]+\s+(?:температур|давлени|напряжени|частот)|угол\s+(?:поворот|наклон|обзор)/i;

  const decOf = s => { const m = /[.,](\d+)$/.exec(s); return m ? m[1].length : 0; };
  const toNum = s => parseFloat(String(s).replace(/[  \s]/g, '').replace(',', '.').replace(/[−–]/g, '-'));
  function fmtLike(n, like, dec) {
    const comma = /,/.test(like) || !/\./.test(like), grouped = /\d[   ]\d{3}/.test(like);
    let s = n.toFixed(dec);
    let [a, b] = s.split('.');
    if (grouped) a = a.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return b ? a + (comma ? ',' : '.') + b : a;
  }
  function bump(numTxt, dir) {
    const dec = decOf(numTxt), step = Math.pow(10, -dec), v = toNum(numTxt);
    return fmtLike(Math.round((v + dir * step) * Math.pow(10, dec)) / Math.pow(10, dec), numTxt, dec);
  }
  function midOf(a, b) {
    const da = decOf(a), db = decOf(b), v = (toNum(a) + toNum(b)) / 2;
    let dec = Math.max(da, db); if (!Number.isInteger(v * Math.pow(10, dec))) dec += 1;
    return fmtLike(v, a, dec);
  }
  function pickRange(lo, hi, o) {
    if (o.range === 'max') return hi;
    if (o.range === 'mid') return midOf(lo, hi);
    return lo;
  }

  /* findReplacements(text, opts) → [{ s, e, from, to, rule, keep }]
     opts: { range:'min'|'mid'|'max'|'keep', name:'название показателя' (для распознавания «диапазон …») } */
  function findReplacements(text, opts) {
    opts = Object.assign({ range: 'min', name: '' }, opts || {});
    const src = String(text).replace(/ё/g, 'е').replace(/Ё/g, 'Е');          // длина строки не меняется
    const keepRange = opts.range === 'keep' || CHAR_RANGE.test(opts.name || '') || CHAR_RANGE.test(src.slice(0, 40));
    const toks = [];
    for (const m of src.matchAll(MASTER)) {
      const g = m.groups || {}, s = m.index, e = m.index + m[0].length, from = text.slice(s, e);
      let t = null;
      if (g.range) t = { type: 'range', lo: g.r1, hi: g.r2 };
      else if (g.tol) t = { type: 'tol', v: g.t1 };
      else if (g.dash) t = { type: 'range', lo: g.d1, hi: g.d2 };
      else if (g.hyph) t = { type: 'range', lo: g.h1, hi: g.h2 };
      else if (g.approx) t = { type: 'eq', v: g.a1 };
      else if (g.ge) t = { type: 'ge', v: g.g1 };
      else if (g.le) t = { type: 'le', v: g.l1 };
      else if (g.gt) t = { type: 'gt', v: g.q1 };
      else if (g.lt) t = { type: 'lt', v: g.w1 };
      else if (g.nge) t = { type: 'ge', v: g.n1 };
      else if (g.nle) t = { type: 'le', v: g.m1 };
      else if (g.from) t = { type: 'ge', v: g.f1 };
      else if (g.upto) t = { type: 'le', v: g.u1 };
      if (t) toks.push(Object.assign(t, { s, e, from }));
    }
    // «не менее N и не более M» / «от N, но не более M» → диапазон
    for (let i = 0; i < toks.length - 1; i++) {
      const a = toks[i], b = toks[i + 1];
      const gap = src.slice(a.e, b.s);
      if (/^\s*(?:[%°a-zа-я/²³µ.]{0,12})?\s*[,;]?\s*(?:и|но|а|или)?\s*$/i.test(gap)) {
        if ((a.type === 'ge' && b.type === 'le') || (a.type === 'le' && b.type === 'ge')) {
          const lo = a.type === 'ge' ? a.v : b.v, hi = a.type === 'le' ? a.v : b.v;
          toks.splice(i, 2, { type: 'range', lo, hi, s: a.s, e: b.e, from: text.slice(a.s, b.e), combo: true });
        }
      }
    }
    const out = [];
    for (const t of toks) {
      let to = null, rule = '';
      switch (t.type) {
        case 'ge': to = t.v; rule = 'не менее → граница входит'; break;
        case 'le': to = t.v; rule = 'не более → граница входит'; break;
        case 'gt': to = bump(t.v, +1); rule = 'строго больше → ' + t.v + ' + шаг'; break;
        case 'lt': to = bump(t.v, -1); rule = 'строго меньше → ' + t.v + ' − шаг'; break;
        case 'eq': to = t.v; rule = 'приблизительно → точное значение'; break;
        case 'tol': to = t.v; rule = '± допуск → номинал'; break;
        case 'range':
          if (keepRange) { out.push({ s: t.s, e: t.e, from: t.from, to: t.from, rule: 'диапазон сохранён', keep: true }); continue; }
          to = pickRange(t.lo, t.hi, opts); rule = 'диапазон → ' + (opts.range === 'max' ? 'верхняя граница' : opts.range === 'mid' ? 'середина' : 'нижняя граница'); break;
      }
      if (to != null) {
        const rep = { s: t.s, e: t.e, from: t.from, to, rule }; out.push(rep);
        if (!rep.keep) { const fx = agreeNoun(src, rep); if (fx) (Array.isArray(fx) ? fx : [fx]).forEach(f => out.push({ s: f.s, e: f.e, from: text.slice(f.s, f.e), to: f.to, rule: 'согласование с числом' })); }
      }
    }
    return out;
  }
  /* ---------- согласование существительного с новым числом: «не менее 24 месяцев» → «24 месяца» ---------- */
  const NOUN_SETS = [['месяц', 'месяца', 'месяцев'], ['год', 'года', 'лет'], ['день', 'дня', 'дней'], ['час', 'часа', 'часов'], ['минута', 'минуты', 'минут'], ['секунда', 'секунды', 'секунд'],
    ['неделя', 'недели', 'недель'], ['штука', 'штуки', 'штук'], ['единица', 'единицы', 'единиц'], ['комплект', 'комплекта', 'комплектов'], ['упаковка', 'упаковки', 'упаковок'],
    ['лист', 'листа', 'листов'], ['рулон', 'рулона', 'рулонов'], ['пара', 'пары', 'пар'], ['место', 'места', 'мест'], ['позиция', 'позиции', 'позиций'], ['экземпляр', 'экземпляра', 'экземпляров'],
    ['сотрудник', 'сотрудника', 'сотрудников'], ['специалист', 'специалиста', 'специалистов'], ['раз', 'раза', 'раз'], ['этап', 'этапа', 'этапов'], ['ступень', 'ступени', 'ступеней']];
  const NOUN_MAP = {}; NOUN_SETS.forEach(set => set.forEach(f => { NOUN_MAP[f] = set; }));
  const nounForm = (set, n) => { const m100 = n % 100, m10 = n % 10; return (m100 > 10 && m100 < 20) ? set[2] : m10 === 1 ? set[0] : (m10 >= 2 && m10 <= 4) ? set[1] : set[2]; };
  const ADJ_RE = /^(рабоч|календарн|банковск|полн)[а-яё]*$/;
  function agreeNoun(src, rep) {
    if (!/^[+\-−]?\d+$/.test(String(rep.to).replace(/[\u00a0\u202f\s]/g, ''))) return null;
    const n = Math.abs(parseInt(String(rep.to).replace(/[\u00a0\u202f\s−]/g, ''), 10)); if (!isFinite(n)) return null;
    const m = /^(\s*(?:\([^)]{1,40}\)\s*)?)(?:([а-яё]+)\s+)?([а-яё]+)/i.exec(src.slice(rep.e)); if (!m) return null;
    let adj = m[2], noun = m[3];
    if (adj && !ADJ_RE.test(adj.toLowerCase())) { noun = adj; adj = null; if (!NOUN_MAP[noun.toLowerCase()]) return null; const off = rep.e + m[1].length; return fixNoun(src, off, off + noun.length, noun, n); }
    const set = NOUN_MAP[noun.toLowerCase()]; if (!set) return null;
    const out = [];
    const nounStart = rep.e + m[0].length - noun.length;
    if (adj && set[0] === 'день') {                            // «1 рабочий день» / «5 рабочих дней»
      const adjStart = rep.e + m[1].length, single = (n % 10 === 1 && n % 100 !== 11);
      const stemAdj = adj.toLowerCase().match(/^(рабоч|календарн|банковск|полн)/)[1], fixed = stemAdj === 'рабоч' ? (single ? 'рабочий' : 'рабочих') : stemAdj + (single ? 'ый' : 'ых');
      if (fixed !== adj.toLowerCase()) out.push({ s: adjStart, e: adjStart + adj.length, to: matchCase(adj, fixed) });
    }
    const f = fixNoun(src, nounStart, nounStart + noun.length, noun, n); if (f) out.push(f);
    return out.length ? out : null;
  }
  const matchCase = (orig, s) => /^[А-ЯЁ]/.test(orig) ? s.charAt(0).toUpperCase() + s.slice(1) : s;
  function fixNoun(src, s, e, noun, n) {
    const set = NOUN_MAP[noun.toLowerCase()]; if (!set) return null;
    const want = matchCase(noun, nounForm(set, n)); return want === noun ? null : { s, e, to: want };
  }
  function applyReplacements(text, reps) {
    let r = text; for (const x of reps.slice().sort((a, b) => b.s - a.s)) r = r.slice(0, x.s) + x.to + r.slice(x.e); return r;
  }
  function concretize(text, opts) { const reps = findReplacements(text, opts); return { text: applyReplacements(String(text), reps), reps, changed: reps.some(r => !r.keep && r.from !== r.to) }; }
  const hasQualifier = text => findReplacements(text, { range: 'min' }).some(r => !r.keep);
  const orphans = text => { const o = []; const src = String(text).replace(/ё/g, 'е'); for (const m of src.matchAll(ORPHAN)) o.push({ s: m.index, e: m.index + m[0].length }); return o; };

  /* =====================================================================
     2. ТАБЛИЦЫ: единый вид для DOCX и XLSX
     ===================================================================== */
  const STOP = new Set(['и', 'в', 'на', 'по', 'с', 'со', 'к', 'для', 'от', 'до', 'при', 'не', 'или', 'а', 'о', 'об', 'из', 'за', 'без', 'мм', 'кг', 'шт']);
  const stems = s => lc(s).replace(/\([^)]*\)/g, ' ').replace(/[^a-zа-я0-9\s]/g, ' ').split(/\s+/).filter(w => w.length > 2 && !STOP.has(w)).map(w => w.slice(0, 6));
  function similar(a, b) {
    const A = stems(a), B = stems(b); if (!A.length || !B.length) return 0;
    const sa = new Set(A), sb = new Set(B); let k = 0; sa.forEach(x => { if (sb.has(x)) k++; });
    return 2 * k / (sa.size + sb.size);
  }
  const isNumeric = s => /^[\s№#.\d,)(-]*$/.test(s);
  const blankText = s => !String(s || '').replace(/[\s_.… -]/g, '');

  const HDR = {
    name: /показател|характеристик|параметр|свойств|наименование\s+(?:показател|характеристик|параметр)/i,
    nameAlt: /наименование|название|описание\s+(?:товара|работ)/i,
    item: /наименование\s+(?:товар|продукц|работ|услуг|материал|оборудован)|товар|позиция|вид\s+(?:работ|товара)/i,
    prop: /предлага|предложен|участник|фактическ|конкретн|подтвержд/i,
    req: /требован|установлен|заказчик|норматив|значени|величин|описание/i,
    unit: /ед\.?\s*изм|единиц/i,
    num: /^\s*(?:№|n|п\/п|номер)/i,
  };

  /** Определение колонок по заголовку: { nameCol, itemCol, reqCol, propCol, unitCol, firstData } */
  function detectColumns(rows) {
    const out = { nameCol: -1, itemCol: -1, reqCol: -1, propCol: -1, unitCol: -1, firstData: 0, header: false };
    for (let r = 0; r < Math.min(rows.length, 5); r++) {
      const cells = rows[r].cells; if (cells.filter(c => !blankText(c.text)).length < 2) continue;
      let hits = 0; const o = { nameCol: -1, itemCol: -1, reqCol: -1, propCol: -1, unitCol: -1 };
      cells.forEach((c, j) => {
        const t = c.text; if (blankText(t) || t.length > 140) return;
        if (HDR.prop.test(t) && !(/установлен|требуем/i.test(t) && !/предлага|участник/i.test(t))) { if (o.propCol < 0) { o.propCol = j; hits++; } }
        else if (HDR.unit.test(t)) { if (o.unitCol < 0) o.unitCol = j; }
        else if (HDR.item.test(t) && !HDR.name.test(t)) { if (o.itemCol < 0) { o.itemCol = j; hits++; } }
        else if (HDR.name.test(t)) { if (o.nameCol < 0) { o.nameCol = j; hits++; } }
        else if (HDR.req.test(t)) { if (o.reqCol < 0) { o.reqCol = j; hits++; } }
        else if (HDR.nameAlt.test(t)) { if (o.nameCol < 0 && o.itemCol < 0) { o.nameCol = j; hits++; } }
      });
      if (hits >= 2) { Object.assign(out, o, { firstData: r + 1, header: true }); break; }
    }
    // таблица «Наименование | Значение» без слова «показатель»: наименование = характеристика
    if (out.header && out.nameCol < 0 && out.itemCol >= 0 && out.reqCol >= 0) { out.nameCol = out.itemCol; out.itemCol = -1; }
    if (out.header && out.reqCol < 0) {              // колонка требований — первая справа от названия, не равная колонке предложения/единиц
      for (let j = out.nameCol + 1; j < (rows[0] ? rows[0].cells.length : 0); j++) if (j !== out.propCol && j !== out.unitCol && j !== out.itemCol) { out.reqCol = j; break; }
    }
    return out;
  }
  /** Без заголовка: название — первая «текстовая» ячейка, значение — последняя непустая справа от неё */
  function guessColumns(rows) {
    const out = { nameCol: -1, itemCol: -1, reqCol: -1, propCol: -1, unitCol: -1, firstData: 0, header: false };
    const votes = {};
    for (const row of rows) {
      const cs = row.cells.map((c, j) => ({ j, t: c.text })).filter(x => !blankText(x.t) && !isNumeric(x.t));
      if (cs.length >= 2) { votes[cs[0].j] = (votes[cs[0].j] || 0) + 1; }
    }
    const best = Object.keys(votes).sort((a, b) => votes[b] - votes[a])[0];
    if (best != null) out.nameCol = +best;
    return out;
  }

  /* =====================================================================
     3. ЧТЕНИЕ ТЗ → список показателей
     ===================================================================== */
  /** table = { rows:[{cells:[{text,...}], section?:text}] } */
  function entriesFromTable(table) {
    const rows = table.rows; if (!rows.length) return [];
    let cols = detectColumns(rows); if (!cols.header) cols = guessColumns(rows);
    if (cols.nameCol < 0) return [];
    const out = []; let item = '';
    for (let r = cols.firstData; r < rows.length; r++) {
      const row = rows[r], cs = row.cells, nonEmpty = cs.filter(c => !blankText(c.text));
      if (!nonEmpty.length) continue;
      if (row.section) { item = clean(nonEmpty[0].text); continue; }
      const name = clean((cs[cols.nameCol] || {}).text || '');
      if (cols.itemCol >= 0 && !blankText((cs[cols.itemCol] || {}).text)) item = clean(cs[cols.itemCol].text);
      let reqCol = cols.reqCol;
      if (reqCol < 0) { for (let j = cs.length - 1; j > cols.nameCol; j--) if (!blankText(cs[j].text) && j !== cols.unitCol) { reqCol = j; break; } }
      if (blankText(name) || reqCol < 0) continue;
      const req = clean((cs[reqCol] || {}).text || ''); if (blankText(req)) continue;
      const unit = cols.unitCol >= 0 ? clean((cs[cols.unitCol] || {}).text || '') : '';
      out.push({ item, name, req, unit });
    }
    return out;
  }
  /** абзацы вида «Ширина: не менее 500 мм» или «Масса не более 3 кг» */
  function entriesFromParagraphs(lines) {
    const out = [];
    for (const raw of lines) {
      const t = clean(raw); if (t.length < 6 || t.length > 400) continue;
      const ci = t.indexOf(':');
      if (ci > 2 && ci < 90 && !blankText(t.slice(ci + 1))) { out.push({ item: '', name: t.slice(0, ci), req: t.slice(ci + 1).trim(), unit: '' }); continue; }
      const reps = findReplacements(t, { range: 'min' }).filter(r => !r.keep);
      if (reps.length) {
        const cut = t.slice(0, reps[0].s).replace(/\s*(?:должн[а-яё]*\s+(?:быть|составлять)|составля[а-яё]*|быть|не)\s*$/i, '').trim();
        if (cut.length > 2) out.push({ item: '', name: cut.replace(/[,:;–—-]+$/, ''), req: t.slice(reps[0].s), unit: '' });
      }
    }
    return out;
  }

  /* ---------- DOCX → таблицы и абзацы ---------- */
  const paraText = p => I.mapParagraph(p).text;
  const cellPlain = tc => kids(tc, 'p').map(paraText).join(' ').replace(/\s+/g, ' ').trim() + kids(tc, 'tbl').map(t => ' ' + byNS(t, NS_W, 't').map(x => x.textContent).join('')).join('');
  function gridInfo(tr) {
    let c = 0; const out = [];
    for (const tc of kids(tr, 'tc')) {
      const pr = firstKid(tc, 'tcPr'), gs = pr && firstKid(pr, 'gridSpan'), vm = pr && firstKid(pr, 'vMerge');
      const span = gs ? (parseInt(gs.getAttributeNS(NS_W, 'val') || gs.getAttribute('w:val'), 10) || 1) : 1;
      out.push({ tc, start: c, span, vcont: !!vm && (vm.getAttributeNS(NS_W, 'val') || vm.getAttribute('w:val') || 'continue') === 'continue' });
      c += span;
    }
    return { cells: out, width: c };
  }
  function docxTables(doc) {
    const tables = [];
    for (const tbl of byNS(doc, NS_W, 'tbl')) {
      if (tbl.parentNode && tbl.parentNode.nodeName.replace(/^.*:/, '') === 'tc') continue;      // вложенные таблицы пропускаем
      const trs = kids(tbl, 'tr'), infos = trs.map(gridInfo), width = Math.max.apply(null, infos.map(i => i.width).concat([1]));
      const rows = trs.map((tr, ri) => {
        const info = infos[ri], cells = [];
        for (let j = 0; j < width; j++) {
          const g = info.cells.find(x => j >= x.start && j < x.start + x.span);
          if (!g) { cells.push({ text: '', tc: null }); continue; }
          cells.push({ text: g.vcont ? '' : cellPlain(g.tc), tc: g.tc, first: g.start === j, span: g.span, vcont: g.vcont });
        }
        const uniq = info.cells.filter(g => !blankText(cellPlain(g.tc)));
        return { cells, tr, section: info.cells.length === 1 && width > 2 && !isNumeric(cellPlain(info.cells[0].tc)) ? true : false };
      });
      tables.push({ el: tbl, rows, kind: 'docx' });
    }
    return tables;
  }
  function docxBodyParagraphs(doc) {
    const body = byNS(doc, NS_W, 'body')[0], out = [];
    (function walk(n, inTbl) {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue; const nm = I.lname(c);
        if (nm === 'p') out.push({ p: c, inTbl });
        else if (nm === 'tbl') walk(c, true);
        else if (nm === 'tr' || nm === 'tc' || nm === 'sdt' || nm === 'sdtContent') walk(c, inTbl);
      }
    })(body, false);
    return out;
  }

  /* ---------- XLSX → таблицы ---------- */
  async function xlsxSheets(zip) {
    const sst = await I.readShared(zip), names = Object.keys(zip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort((a, b) => parseInt(a.replace(/\D/g, ''), 10) - parseInt(b.replace(/\D/g, ''), 10));
    const out = [];
    for (const name of names) {
      const doc = parseXml(await zip.file(name).async('string')); const rowsEls = byNS(doc, NS_S, 'row');
      let maxCol = 1; const rowsRaw = rowsEls.map(row => {
        const m = {}; for (const c of kids(row, 'c')) { const ref = I.splitRef(c.getAttribute('r') || ''); if (ref) { m[ref.col] = c; maxCol = Math.max(maxCol, ref.col); } }
        return { row, m, rn: +row.getAttribute('r') };
      });
      const wide = new Set();                       // строки, где ячейка объединена почти на всю ширину — это заголовок раздела
      for (const m of byNS(doc, NS_S, 'mergeCell')) { const [a, b] = (m.getAttribute('ref') || '').split(':'); const A = I.splitRef(a), B2 = b && I.splitRef(b); if (A && B2 && A.row === B2.row && B2.col - A.col + 1 >= Math.max(3, maxCol - 1)) wide.add(A.row); }
      const rows = rowsRaw.map(rr => {
        const cells = []; for (let j = 1; j <= maxCol; j++) { const el = rr.m[j]; cells.push({ text: el ? I.xlCellText(el, sst).trim() : '', el, col: j, rn: rr.rn, rowEl: rr.row }); }
        return { cells, rowEl: rr.row, rn: rr.rn, section: wide.has(rr.rn) };
      });
      out.push({ name, doc, rows, kind: 'xlsx', sst, maxCol });
    }
    return out;
  }

  async function readTz(name, buf) {
    const ext = (name.split('.').pop() || '').toLowerCase(), entries = [];
    const zip = await I.JSZip.loadAsync(buf);
    if (ext === 'docx' || ext === 'dotx') {
      const doc = parseXml(await zip.file('word/document.xml').async('string'));
      for (const t of docxTables(doc)) entries.push(...entriesFromTable(t));
      entries.push(...entriesFromParagraphs(docxBodyParagraphs(doc).filter(x => !x.inTbl).map(x => paraText(x.p))));
    } else if (ext === 'xlsx' || ext === 'xlsm') {
      for (const sh of await xlsxSheets(zip)) entries.push(...entriesFromTable(sh));
    } else throw new Error('ТЗ: формат .' + ext + ' не поддерживается. Сохраните как .docx или .xlsx');
    return entries;
  }

  /* =====================================================================
     4. ТЕХНИЧЕСКОЕ ПРЕДЛОЖЕНИЕ: подстановка значений
     ===================================================================== */
  function findEntry(entries, used, name, item) {
    let best = null, bestScore = 0;
    for (const e of entries) {
      let sc = similar(name, e.name); if (sc < 0.6) continue;
      if (item && e.item) sc += 0.15 * similar(item, e.item);
      if (used.has(e)) sc -= 0.2;
      if (sc > bestScore) { best = e; bestScore = sc; }
    }
    return best;
  }

  /** Обработка одной таблицы ТП. cellApi: { write(cell,text), replace(cell,fn), mark(cell) } */
  function processTable(table, entries, used, opts, api, report) {
    const rows = table.rows; if (!rows.length) return;
    let cols = detectColumns(rows); if (!cols.header) cols = guessColumns(rows);
    let item = '';
    for (let r = cols.firstData; r < rows.length; r++) {
      const row = rows[r], cs = row.cells, nonEmpty = cs.filter(c => !blankText(c.text));
      if (!nonEmpty.length) continue;
      if (row.section && nonEmpty.length === 1) { item = clean(nonEmpty[0].text); continue; }
      const name = cols.nameCol >= 0 ? clean((cs[cols.nameCol] || {}).text || '') : '';
      if (cols.itemCol >= 0 && !blankText((cs[cols.itemCol] || {}).text)) item = clean(cs[cols.itemCol].text);
      const entry = name ? findEntry(entries, used, name, item) : null;
      const rangeOpts = { range: opts.range, name };
      if (cols.propCol >= 0) {
        const cell = cs[cols.propCol]; if (!cell || cell.vcont) continue;
        if (blankText(cell.text)) {
          if (entry) {
            used.add(entry);
            const res = concretize(entry.req, { range: opts.range, name: entry.name });
            let val = res.text;
            if (entry.unit && cols.unitCol < 0 && /^[+\-−]?[\d\s.,]+$/.test(val)) val += ' ' + entry.unit;
            api.write(cell, val, cs[cols.reqCol >= 0 ? cols.reqCol : cols.nameCol]);
            report.push({ kind: 'fill', name: name || entry.name, from: entry.req, to: val, rule: res.reps.map(x => x.rule).join('; ') || 'значение из ТЗ без изменений' });
          } else if (opts.highlight && name) { api.mark(cell); report.push({ kind: 'hl', name: name || '(строка без названия)', why: 'в ТЗ не найден такой показатель' }); }
        } else if (hasQualifier(cell.text)) {
          const reps = api.replace(cell, t => findReplacements(t, rangeOpts));
          reps.filter(x => !x.keep).forEach(x => report.push({ kind: 'rep', name: name || cell.text.slice(0, 40), from: x.from, to: x.to, rule: x.rule }));
        }
        continue;
      }
      // колонки «предложение участника» нет — заменяем значения на месте (там, где они записаны)
      for (let j = 0; j < cs.length; j++) {
        const cell = cs[j]; if (!cell || j === cols.nameCol || j === cols.itemCol || cell.vcont || blankText(cell.text)) continue;
        if (cell.first === false) continue;
        const reps = api.replace(cell, t => findReplacements(t, rangeOpts));
        reps.filter(x => !x.keep).forEach(x => report.push({ kind: 'rep', name: name || clean(cell.text).slice(0, 40), from: x.from, to: x.to, rule: x.rule }));
        if (reps.length && entry) used.add(entry);
      }
    }
  }

  /* ---------- DOCX ---------- */
  function docxApi(doc, opts) {
    return {
      write(cell, text, likeCell) { I.writeCell(doc, cell.tc, text, { labelTc: likeCell && likeCell.tc }); },
      mark(cell) { I.writeCell(doc, cell.tc, '', { highlight: true }); },
      replace(cell, fn) {
        const all = [];
        for (const p of kids(cell.tc, 'p')) {
          const map = I.mapParagraph(p), reps = fn(map.text); all.push(...reps);
          for (const x of reps.filter(y => !y.keep && y.from !== y.to).sort((a, b) => b.s - a.s)) I.fillRange(p, x.s, x.e, x.to);
          if (opts.highlight) markOrphans(p);
        }
        return all;
      },
    };
  }
  function markOrphans(p) {
    const map = I.mapParagraph(p), os = orphans(map.text);
    for (const o of os.sort((a, b) => b.s - a.s)) I.highlightRange(p, o.s, o.e);
    return os.length;
  }
  async function fillDocx(buf, entries, opts) {
    const zip = await I.JSZip.loadAsync(buf), report = [], used = new Set();
    const f = zip.file('word/document.xml'); if (!f) throw new Error('Не найден word/document.xml — это не DOCX');
    const doc = parseXml(await f.async('string')), api = docxApi(doc, opts);
    for (const t of docxTables(doc)) processTable(t, entries, used, opts, api, report);
    if (opts.replaceInText) {
      for (const { p, inTbl } of docxBodyParagraphs(doc)) {
        if (inTbl) continue;
        const map = I.mapParagraph(p); if (!map.text.trim()) continue;
        const reps = findReplacements(map.text, { range: opts.range, name: map.text.slice(0, 60) }).filter(x => !x.keep && x.from !== x.to);
        for (const x of reps.sort((a, b) => b.s - a.s)) { I.fillRange(p, x.s, x.e, x.to); report.push({ kind: 'rep', name: 'текст документа', from: x.from, to: x.to, rule: x.rule }); }
      }
    }
    if (opts.highlight) {                          // остались «не менее / не более» без числа — подсвечиваем
      let left = 0; for (const { p } of docxBodyParagraphs(doc)) left += markOrphans(p);
      if (left) report.push({ kind: 'hl', name: 'формулировки «не менее / не более» без числа', why: 'число не найдено рядом — впишите значение вручную (' + left + ')' });
    }
    zip.file('word/document.xml', serialize(doc));
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), report };
  }

  /* ---------- XLSX ---------- */
  async function fillXlsx(buf, entries, opts) {
    const zip = await I.JSZip.loadAsync(buf), report = [], used = new Set();
    const sf = zip.file('xl/styles.xml'), stylesDoc = sf ? parseXml(await sf.async('string')) : null, styler = I.makeHighlightStyler(stylesDoc);
    for (const sh of await xlsxSheets(zip)) {
      const doc = sh.doc;
      const ensure = cell => {
        if (cell.el) return cell.el;
        const row = cell.rowEl, c = doc.createElementNS(NS_S, 'c'); c.setAttribute('r', I.numToCol(cell.col) + cell.rn);
        const after = kids(row, 'c').find(x => { const q = I.splitRef(x.getAttribute('r') || ''); return q && q.col > cell.col; });
        if (after) row.insertBefore(c, after); else row.appendChild(c); cell.el = c; return c;
      };
      const api = {
        write(cell, text) { I.setXlText(doc, ensure(cell), text); cell.text = text; },
        mark(cell) { if (styler) { const el = ensure(cell); el.setAttribute('s', String(styler(+el.getAttribute('s')))); } },
        replace(cell, fn) {
          const reps = fn(cell.text), real = reps.filter(x => !x.keep && x.from !== x.to);
          if (real.length) { const nt = applyReplacements(cell.text, real); I.setXlText(doc, cell.el, nt); cell.text = nt; }
          if (opts.highlight && orphans(cell.text).length && styler) api.mark(cell);
          return reps;
        },
      };
      processTable(sh, entries, used, opts, api, report);
      zip.file(sh.name, serialize(doc));
    }
    if (stylesDoc) zip.file('xl/styles.xml', serialize(stylesDoc));
    return { data: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), report };
  }

  async function fillProposal(name, buf, entries, opts) {
    opts = Object.assign({ range: 'min', highlight: true, replaceInText: true }, opts || {});
    const ext = (name.split('.').pop() || '').toLowerCase();
    let res;
    if (ext === 'docx' || ext === 'dotx') res = await fillDocx(buf, entries, opts);
    else if (ext === 'xlsx' || ext === 'xlsm') res = await fillXlsx(buf, entries, opts);
    else throw new Error('Техническое предложение: формат .' + ext + ' не поддерживается. Сохраните как .docx или .xlsx');
    const stats = { replaced: res.report.filter(r => r.kind === 'rep' || r.kind === 'fill').length, highlighted: res.report.filter(r => r.kind === 'hl').length };
    return { data: res.data, report: res.report, stats };
  }

  return { findReplacements, concretize, applyReplacements, hasQualifier, readTz, fillProposal, entriesFromTable, entriesFromParagraphs, similar };
});
