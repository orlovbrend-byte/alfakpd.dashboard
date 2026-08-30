/* ============================================================
   Альфа-КПД · интерфейс дэшборда
   ============================================================ */
(function () {
  'use strict';
  const C = window.CORE;
  const $ = (id) => document.getElementById(id);
  const els = (s, r) => [...(r || document).querySelectorAll(s)];
  const dash = '—';

  let D = null, BM = null;

  /* Код Apps Script отдаём как есть, одной строкой-константой: он
     должен копироваться в таблицу без правок. */
  const GS_CODE = [
    'function doGet(e) {',
    '  // Лист, который читаем. Название должно совпадать ТОЧНО.',
    "  var SHEET_NAME = 'Недельный отчёт по трафику';",
    '',
    '  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);',
    '  if (!sh) {',
    "    return out({ error: 'Лист не найден: ' + SHEET_NAME });",
    '  }',
    '  var v = sh.getDataRange().getValues();',
    '  var rows = [];',
    '',
    '  // Первая строка — заголовки, данные идут с второй.',
    '  for (var i = 1; i < v.length; i++) {',
    '    var r = v[i];',
    '    if (!r[0] || !r[2]) continue;           // нет даты или направления',
    '    rows.push({',
    '      period:        iso(r[0]),             // A · начало периода',
    '      end:           iso(r[1]),             // B · конец периода',
    '      dir:           str(r[2]),             // C · направление',
    '      src:           str(r[3]),             // D · источник',
    '      page:          str(r[4]),             // E · посадочная страница',
    '      spend:         num(r[5]),             // F · расход',
    '      leads:         num(r[6]),             // G · лиды',
    '      kl:            num(r[8]),             // I · квал-лиды',
    '      inwork:        num(r[9]),             // J · в работе',
    '      badlead:       num(r[10]),            // K · не кач. лид',
    '      spam:          num(r[11]),            // L · спам',
    '      existing:      num(r[13]),            // N · существующие',
    '      deals_created: num(r[17]),            // R · создано сделок',
    '      pipeline:      num(r[18]),            // S · сумма созданных',
    '      deals_lost:    num(r[19]),            // T · проиграно сделок',
    '      revenue_lost:  num(r[20]),            // U · сумма проигранных',
    '      deals_won:     num(r[21]),            // V · успешных сделок',
    '      revenue_won:   num(r[22])             // W · сумма успешных',
    '    });',
    '  }',
    '  return out({ weekly: rows, count: rows.length, sheet: SHEET_NAME });',
    '}',
    '',
    'function out(obj) {',
    '  return ContentService.createTextOutput(JSON.stringify(obj))',
    '    .setMimeType(ContentService.MimeType.JSON);',
    '}',
    '',
    '// Дата → YYYY-MM-DD. Дэшборд сортирует периоды как строки,',
    '// поэтому формат обязателен именно такой.',
    'function iso(d) {',
    '  if (d instanceof Date) {',
    "    return Utilities.formatDate(d, 'Asia/Vladivostok', 'yyyy-MM-dd');",
    '  }',
    "  return d ? String(d).slice(0, 10) : null;",
    '}',
    '',
    '// Числа приходят и текстом с пробелами-разделителями («11 328 914»),',
    '// поэтому чистим пробелы и неразрывные пробелы, иначе получим 0.',
    'function num(x) {',
    "  if (x === '' || x === null || x === undefined) return 0;",
    "  if (typeof x === 'number') return x;",
    "  var s = String(x).replace(/\\u00A0/g, '').replace(/\\s/g, '').replace(',', '.');",
    '  var n = parseFloat(s);',
    '  return isNaN(n) ? 0 : n;',
    '}',
    '',
    'function str(x) {',
    "  if (x === null || x === undefined) return null;",
    "  var s = String(x).replace(/\\u00A0/g, ' ').trim();",
    "  return (s === '' || s === '-' || s === '—') ? null : s;",
    '}'
  ].join('\n');
  const CH = {};
  const S = { view: 'over', level: 'monthly', from: null, to: null, dirs: [], srcs: [], pages: [],
              heatDim: 'dir', heatMetric: 'cpql', fcMethod: 'last3', pfPeriod: null,
              shDim: 'src', shMetric: 'kl', gsUrl: '' };
  /* Планы, введённые руками, живут в localStorage и накладываются
     поверх выгрузки. Так правка не теряется при перезагрузке и не
     требует записи в саму таблицу. */
  const PLANS_KEY = 'alfakpd.plans.v1';
  let PLANS = {};
  const CAL = { open: false, year: null, pick: null };

  /* ---------- ФОРМАТ ---------- */
  const nf = new Intl.NumberFormat('ru-RU');
  const fmtN = (v) => v == null || !isFinite(v) ? dash : nf.format(Math.round(v));
  const fmtM = (v) => v == null || !isFinite(v) ? dash : nf.format(Math.round(v)) + ' ₽';
  const fmtP = (v, d) => v == null || !isFinite(v) ? dash : (v * 100).toFixed(d == null ? 1 : d) + '%';
  const fmtX = (v) => v == null || !isFinite(v) ? dash : v.toFixed(1) + '×';
  function fmtBig(v) {
    if (v == null || !isFinite(v)) return dash;
    const a = Math.abs(v);
    if (a >= 1e9) return (v / 1e9).toFixed(2) + ' млрд ₽';
    if (a >= 1e6) return (v / 1e6).toFixed(1) + ' млн ₽';
    if (a >= 1e3) return nf.format(Math.round(v / 1e3)) + ' тыс ₽';
    return fmtM(v);
  }
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g,
    (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /* ---------- ГРАФИКИ ----------
     Ключевое: Chart.js должен получить размеры контейнера. Если карточка
     скрыта (display:none) в момент создания, полотно получает 0×0 и
     остаётся пустым навсегда. Поэтому рисуем ТОЛЬКО активную вкладку,
     а после показа делаем resize. */
  const COL = { acc: '#DFFF4B', mid: '#A8A8B3', dim: '#6E6E7A', red: '#F4534F',
                amber: '#FBBF24', blue: '#5EC8F2', line: '#26262B',
                accSoft: 'rgba(223,255,75,.14)' };

  function axis(extra) {
    return Object.assign({
      grid: { color: COL.line, drawBorder: false },
      border: { display: false },
      ticks: { color: COL.dim, font: { family: "'JetBrains Mono',monospace", size: 10 }, maxRotation: 0, autoSkipPadding: 14 }
    }, extra || {});
  }
  function opts(extra) {
    return Object.assign({
      maintainAspectRatio: false, responsive: true,
      animation: { duration: 620, easing: 'easeOutQuart' },
      /* Наведение в любой точке подсвечивает ВСЕ ряды этого периода —
         иначе на линии приходится попадать курсором в точку. */
      interaction: { mode: 'index', intersect: false },
      hover: { mode: 'index', intersect: false },
      plugins: {
        legend: { labels: { color: COL.mid, boxWidth: 10, boxHeight: 10, usePointStyle: true,
                            padding: 14, font: { family: "'Manrope',sans-serif", size: 12 } } },
        tooltip: {
          enabled: true, backgroundColor: '#08080A', borderColor: COL.line, borderWidth: 1,
          titleColor: COL.acc, bodyColor: COL.fg || '#F2F2F4', padding: 11, cornerRadius: 9,
          displayColors: true, boxPadding: 4, caretSize: 6,
          titleFont: { family: "'JetBrains Mono',monospace", size: 11.5 },
          bodyFont: { family: "'JetBrains Mono',monospace", size: 12 }
        }
      },
      scales: { x: axis(), y: axis({ beginAtZero: true }) }
    }, extra || {});
  }
  const pctAxis = () => axis({ beginAtZero: true, ticks: { color: COL.dim,
    font: { family: "'JetBrains Mono',monospace", size: 10 }, callback: (v) => v + '%' } });
  const moneyTip = (it) => ' ' + it.dataset.label + ': ' + fmtM(it.raw);

  const line = (label, data, color, fill) => ({
    label, data, borderColor: color, backgroundColor: fill || 'transparent', fill: !!fill,
    tension: .3, borderWidth: 2.2, pointRadius: data.length > 24 ? 0 : 3, pointHoverRadius: 6,
    pointBackgroundColor: color, pointBorderColor: '#0E0E0F', pointBorderWidth: 1.5, spanGaps: true
  });
  const bars = (label, data, color) => ({
    label, data, backgroundColor: color, borderRadius: 5, borderSkipped: false,
    maxBarThickness: 36, hoverBackgroundColor: COL.acc
  });

  function draw(id, cfg) {
    const cv = $(id);
    if (!cv) return;
    if (CH[id]) { CH[id].destroy(); delete CH[id]; }
    CH[id] = new Chart(cv, cfg);
  }
  /** После показа вкладки полотна уже имеют размеры — досчитываем. */
  function resizeAll() {
    Object.values(CH).forEach(c => { try { c.resize(); } catch (e) {} });
  }

  /* ---------- KPI ---------- */
  function kpi(o) {
    let d = '<span class="flat">' + dash + '</span>';
    if (o.value != null && o.prev != null && o.prev !== 0 && isFinite(o.value) && isFinite(o.prev)) {
      const rel = (o.value - o.prev) / Math.abs(o.prev);
      const good = o.lower ? rel < 0 : rel > 0;
      const cls = Math.abs(rel) < .005 ? 'flat' : good ? 'up' : 'dn';
      d = `<span class="${cls}">${rel > 0 ? '▲ +' : '▼ '}${Math.abs(rel * 100).toFixed(1)}%</span>`;
    } else if (o.prev == null && o.value != null) {
      d = '<span class="flat">нет прошлого периода</span>';
    }
    return `<div class="kpi${o.acc ? ' acc' : ''}">
      <div class="kpi-l">${esc(o.label)}${o.hint ? `<i class="hint" title="${esc(o.hint)}">?</i>` : ''}</div>
      <div class="kpi-v">${o.display}</div>
      <div class="kpi-d">${d}</div>
      ${o.sub ? `<div class="kpi-s">${o.sub}</div>` : ''}</div>`;
  }

  /* ---------- ТАБЛИЦА с сортировкой ---------- */
  const SORT = {};
  function table(host, cols, rows, o) {
    o = o || {};
    const node = $(host);
    if (!node) return;
    const key = o.key || host;
    const st = SORT[key] || (SORT[key] = o.sort || { col: 0, dir: 1 });
    const col = cols[st.col];
    if (col && col.sortVal) {
      rows = rows.slice().sort((a, b) => {
        const x = col.sortVal(a), y = col.sortVal(b);
        if (x == null) return 1;
        if (y == null) return -1;
        return typeof x === 'string' ? st.dir * x.localeCompare(y, 'ru') : st.dir * (x - y);
      });
    }
    const head = cols.map((c, i) =>
      `<th data-si="${i}"${c.num ? ' style="text-align:right"' : ''}>${esc(c.t)}${st.col === i ? (st.dir > 0 ? ' ↑' : ' ↓') : ''}</th>`).join('');
    const body = rows.map(r => '<tr>' + cols.map(c =>
      `<td${c.num ? ' class="num"' : ''}>${c.v(r)}</td>`).join('') + '</tr>').join('');
    const foot = o.total ? '<tfoot><tr>' + cols.map(c =>
      `<td${c.num ? ' class="num"' : ''}>${c.f ? c.f(o.total) : ''}</td>`).join('') + '</tr></tfoot>' : '';
    node.innerHTML = `<div class="tw"><table><thead><tr>${head}</tr></thead><tbody>${body}</tbody>${foot}</table></div>`;
    els('th[data-si]', node).forEach(th => th.onclick = () => {
      const i = +th.dataset.si;
      SORT[key] = { col: i, dir: st.col === i ? -st.dir : (cols[i].num ? -1 : 1) };
      table(host, cols, rows, o);
    });
  }

  /* ---------- ПУСТОЙ СРЕЗ ----------
     Ноль и «нет данных» — разные вещи. Если срез пуст, показываем это
     словами: нули читались бы как факт «расход нулевой». */
  function emptyHtml(hint) {
    return `<div class="empty"><b>В этом срезе нет данных</b>
      ${esc(hint || 'Комбинация фильтров не даёт ни одной строки.')}
      <div style="margin-top:14px"><button class="btn" data-act="reset">Сбросить фильтры</button></div></div>`;
  }
  function showEmpty(hosts, hint) {
    Object.keys(CH).forEach(id => { CH[id].destroy(); delete CH[id]; });
    hosts.forEach((h, i) => { const n = $(h); if (n) n.innerHTML = i === 0 ? emptyHtml(hint) : ''; });
    els('[data-act="reset"]').forEach(b => b.onclick = () => {
      S.dirs = []; S.srcs = []; S.pages = []; syncFilters(); render();
    });
  }

  /* ============================================================
     ОБЗОР
     ============================================================ */
  function renderOver() {
    const rows = C.filterRows(D, S);
    if (!rows.length) return showEmpty(['ovKpi', 'ovFlags', 'ovCharts', 'ovTable']);
    $('ovCharts').innerHTML = `
      <div class="card"><h3>Лиды и квал-лиды по периодам</h3>
        <p class="sub">Наведите курсор — покажет точные значения за период. Заливка — квал-лиды.</p>
        <div class="ch tall"><canvas id="c1"></canvas></div></div>
      <div class="card"><h3>CPL и CPQL</h3>
        <p class="sub">Стоимость лида и квал-лида. Рост — плохо: платим больше за то же.</p>
        <div class="ch tall"><canvas id="c2"></canvas></div></div>`;
    const a = C.agg(rows), p = C.previousWindow(D, S), s = C.series(D, S);
    const L = s.map(x => x.label);

    $('ovKpi').innerHTML = [
      kpi({ label: 'Расход', display: fmtBig(a.spend), value: a.spend, prev: p && p.spend, lower: true }),
      kpi({ label: 'Лиды', display: fmtN(a.leads), value: a.leads, prev: p && p.leads }),
      kpi({ label: 'Квал-лиды', display: fmtN(a.kl), value: a.kl, prev: p && p.kl, acc: true }),
      kpi({ label: 'CPL', display: fmtM(a.cpl), value: a.cpl, prev: p && p.cpl, lower: true,
            sub: a.cplClean != null ? 'без спама ' + fmtM(a.cplClean) : '' }),
      kpi({ label: 'CPQL', display: fmtM(a.cpql), value: a.cpql, prev: p && p.cpql, lower: true, acc: true,
            hint: 'Расход ÷ квал-лиды' }),
      kpi({ label: 'CR лид → квал', display: fmtP(a.cr), value: a.cr, prev: p && p.cr }),
      kpi({ label: 'Доля спама', display: fmtP(a.spamShare, 0), value: a.spamShare, prev: p && p.spamShare, lower: true }),
      kpi({ label: 'Выручка (успешные)', display: fmtBig(a.revenue_won), value: a.revenue_won, prev: p && p.revenue_won,
            sub: 'только закрытые сделки' })
    ].join('');

    draw('c1', { type: 'line', data: { labels: L, datasets: [
      line('Лиды', s.map(x => x.leads), COL.mid),
      line('Квал-лиды', s.map(x => x.kl), COL.acc, COL.accSoft)] },
      options: opts({ plugins: { tooltip: { callbacks: {
        title: (i) => s[i[0].dataIndex].labelFull,
        label: (it) => ' ' + it.dataset.label + ': ' + fmtN(it.raw),
        afterBody: (i) => { const x = s[i[0].dataIndex];
          return ['CR: ' + fmtP(x.cr) + '   CPQL: ' + fmtM(x.cpql)]; } } } } }) });

    draw('c2', { type: 'line', data: { labels: L, datasets: [
      line('CPQL', s.map(x => x.cpql == null ? null : Math.round(x.cpql)), COL.acc, COL.accSoft),
      line('CPL', s.map(x => x.cpl == null ? null : Math.round(x.cpl)), COL.mid)] },
      options: opts({ plugins: { tooltip: { callbacks: {
        title: (i) => s[i[0].dataIndex].labelFull, label: moneyTip } } } }) });

    $('ovFlags').innerHTML = flagsHtml(buildFlags(a, p, s));

    const g = C.groupBy(D, S, 'src').filter(r => r.spend > 0 || r.leads > 0);
    table('ovTable', [
      { t: 'Источник', v: r => esc(r.key), sortVal: r => r.key },
      { t: 'Расход', num: 1, v: r => fmtM(r.spend), sortVal: r => r.spend, f: t => fmtM(t.spend) },
      { t: 'Лиды', num: 1, v: r => fmtN(r.leads), sortVal: r => r.leads, f: t => fmtN(t.leads) },
      { t: 'Квал', num: 1, v: r => fmtN(r.kl), sortVal: r => r.kl, f: t => fmtN(t.kl) },
      { t: 'CPL', num: 1, v: r => fmtM(r.cpl), sortVal: r => r.cpl, f: t => fmtM(t.cpl) },
      { t: 'CPQL', num: 1, v: r => fmtM(r.cpql), sortVal: r => r.cpql, f: t => fmtM(t.cpql) },
      { t: 'CR', num: 1, v: r => fmtP(r.cr), sortVal: r => r.cr, f: t => fmtP(t.cr) },
      { t: 'Спам', num: 1, v: r => fmtP(r.spamShare, 0), sortVal: r => r.spamShare, f: t => fmtP(t.spamShare, 0) }
    ], g, { key: 'ov', sort: { col: 1, dir: -1 }, total: a });
  }

  /** Базовые выводы: только то, что видно в цифрах, без домыслов. */
  function buildFlags(a, p, s) {
    const out = [];
    if (a.spamShare != null && a.spamShare > .3)
      out.push({ t: 'bad', h: 'Спам ' + fmtP(a.spamShare, 0) + ' от всех лидов',
        s: 'CPL и конверсия по этому срезу занижены: часть «лидов» — мусор. Смотрите CPL без спама: ' + fmtM(a.cplClean) + '.' });
    if (p && p.cpql && a.cpql && (a.cpql - p.cpql) / p.cpql > .25)
      out.push({ t: 'warn', h: 'CPQL вырос на ' + (((a.cpql / p.cpql) - 1) * 100).toFixed(0) + '%',
        s: fmtM(p.cpql) + ' → ' + fmtM(a.cpql) + '. Проверьте, что именно подорожало — расход вырос или квал-лидов стало меньше.' });
    if (p && p.cpql && a.cpql && (a.cpql - p.cpql) / p.cpql < -.15)
      out.push({ t: 'ok', h: 'CPQL снизился на ' + ((1 - a.cpql / p.cpql) * 100).toFixed(0) + '%',
        s: fmtM(p.cpql) + ' → ' + fmtM(a.cpql) + '. Если объём квал-лидов не упал — это чистое улучшение.' });
    if (a.cr != null && a.cr < .12 && a.leads > 60)
      out.push({ t: 'warn', h: 'Конверсия в квал ' + fmtP(a.cr),
        s: 'На ' + fmtN(a.leads) + ' лидов всего ' + fmtN(a.kl) + ' квал. Либо трафик нецелевой, либо лиды не доходят до обработки.' });
    if (a.revenue_won > 0 && a.pipeline > 0 && a.pipeline / a.revenue_won > 20)
      out.push({ t: 'warn', h: 'Создано сделок на ' + fmtBig(a.pipeline) + ', закрыто на ' + fmtBig(a.revenue_won),
        s: 'Разрыв ' + (a.pipeline / a.revenue_won).toFixed(0) + '×. Либо сделки не закрываются в CRM, либо цикл сделки длиннее периода — ROMI по этому окну считать нельзя.' });
    if (a.revenue_won === 0 && a.spend > 0)
      out.push({ t: 'warn', h: 'Выручки в срезе нет, расход ' + fmtBig(a.spend),
        s: 'ROMI и ДРР не считаются. Успешные сделки в таблице не заполнены за этот период.' });
    const dirs = C.groupBy(D, S, 'dir').filter(r => r.kl >= 5 && r.cpql != null).sort((x, y) => y.cpql - x.cpql);
    if (dirs.length >= 2) {
      const w = dirs[0], b = dirs[dirs.length - 1];
      if (w.cpql / b.cpql > 2)
        out.push({ t: 'warn', h: 'Разброс CPQL между направлениями ' + (w.cpql / b.cpql).toFixed(1) + '×',
          s: esc(b.key) + ' — ' + fmtM(b.cpql) + ', ' + esc(w.key) + ' — ' + fmtM(w.cpql) + '. Бюджет стоит смотреть по направлениям, а не в среднем.' });
    }
    if (!out.length) out.push({ t: 'ok', h: 'Аномалий в срезе не видно', s: 'Спам, конверсия и CPQL в пределах обычного.' });
    return out;
  }
  const flagsHtml = (fl) => fl.map(f =>
    `<div class="flag ${f.t}"><i></i><div><b>${f.h}</b><span>${f.s}</span></div></div>`).join('');

  /* ============================================================
     ДИНАМИКА
     ============================================================ */
  function renderDyn() {
    const rows = C.filterRows(D, S);
    if (!rows.length) return showEmpty(['dynKpi', 'dynCharts', 'dynTable']);
    $('dynCharts').innerHTML = `
      <div class="card"><h3>Лиды и квал-лиды</h3>
        <p class="sub">Серая линия — все лиды, лаймовая с заливкой — квал. Расхождение линий и есть качество трафика.</p>
        <div class="ch"><canvas id="d0"></canvas></div></div>
      <div class="card"><h3>CPL и CPQL</h3>
        <p class="sub">Стоимость лида и квал-лида в одном масштабе. Рост — платим больше за то же.</p>
        <div class="ch"><canvas id="d5"></canvas></div></div>
      <div class="card"><h3>Расход по периодам</h3>
        <p class="sub">Столбцы — расход. Наведите — точная сумма и что за неё получили.</p>
        <div class="ch"><canvas id="d1"></canvas></div></div>
      <div class="card"><h3>Конверсия лид → квал</h3>
        <p class="sub">Две линии: от всей базы и от базы без спама. Вторая — настоящее качество обработки.</p>
        <div class="ch"><canvas id="d2"></canvas></div></div>
      <div class="card"><h3>Сделки: создано и закрыто</h3>
        <p class="sub">Столбцы — созданные (это потенциал, не деньги), линия — успешно закрытые.</p>
        <div class="ch"><canvas id="d3"></canvas></div></div>
      <div class="card"><h3>Доля спама в лидах</h3>
        <p class="sub">Там, где столбец высокий, CPL и конверсия за период недостоверны.</p>
        <div class="ch"><canvas id="d4"></canvas></div></div>
      <div class="card"><h3>Качество базы: из чего состоят лиды</h3>
        <p class="sub">Один столбец — 100% лидов периода: квал, некачественные, спам и прочие. Видно, куда уходит поток.</p>
        <div class="ch"><canvas id="d6"></canvas></div></div>
      <div class="card"><h3>Стоимость успешной сделки</h3>
        <p class="sub">Расход ÷ успешные сделки. Пропуски — периоды, где закрытых сделок в таблице нет.</p>
        <div class="ch"><canvas id="d7"></canvas></div></div>`;
    const a = C.agg(rows), p = C.previousWindow(D, S), s = C.series(D, S);
    const L = s.map(x => x.label);

    $('dynKpi').innerHTML = [
      kpi({ label: 'Периодов в срезе', display: fmtN(s.length), value: null, prev: null,
            sub: C.periodLabel(S.from, S.level) + ' → ' + C.periodLabel(S.to, S.level) }),
      kpi({ label: 'Расход', display: fmtBig(a.spend), value: a.spend, prev: p && p.spend, lower: true }),
      kpi({ label: 'Квал-лиды', display: fmtN(a.kl), value: a.kl, prev: p && p.kl, acc: true }),
      kpi({ label: 'CPQL', display: fmtM(a.cpql), value: a.cpql, prev: p && p.cpql, lower: true, acc: true }),
      kpi({ label: 'Прошлое окно', display: p ? C.periodShort(p.from, S.level) + ' → ' + C.periodShort(p.to, S.level) : dash,
            value: null, prev: null, sub: p ? 'та же длина — дельты сравнимы' : 'в данных нет окна той же длины' })
    ].join('');

    const tip = (fmt, extra) => ({ callbacks: {
      title: (i) => s[i[0].dataIndex].labelFull,
      label: (it) => ' ' + it.dataset.label + ': ' + fmt(it.raw),
      afterBody: extra ? (i) => extra(s[i[0].dataIndex]) : undefined } });

    draw('d0', { type: 'line', data: { labels: L, datasets: [
      line('Лиды', s.map(x => x.leads), COL.mid),
      line('Квал-лиды', s.map(x => x.kl), COL.acc, COL.accSoft)] },
      options: opts({ plugins: { tooltip: tip(fmtN, (x) => ['CR: ' + fmtP(x.cr) + '   CPQL: ' + fmtM(x.cpql)]) } }) });

    draw('d5', { type: 'line', data: { labels: L, datasets: [
      line('CPQL', s.map(x => x.cpql == null ? null : Math.round(x.cpql)), COL.acc, COL.accSoft),
      line('CPL', s.map(x => x.cpl == null ? null : Math.round(x.cpl)), COL.mid)] },
      options: opts({ plugins: { tooltip: tip(fmtM) } }) });

    draw('d6', { type: 'bar', data: { labels: L, datasets: [
      bars('Квал-лиды', s.map(x => x.kl), COL.acc),
      bars('Некачественные', s.map(x => x.badlead), 'rgba(251,191,36,.75)'),
      bars('Спам', s.map(x => x.spam), 'rgba(244,83,79,.75)'),
      bars('Прочие', s.map(x => Math.max(0, x.leads - x.kl - x.badlead - x.spam)), 'rgba(168,168,179,.35)')] },
      options: opts({ scales: { x: axis({ stacked: true }), y: axis({ stacked: true, beginAtZero: true }) },
        plugins: { tooltip: tip(fmtN, (x) => ['всего лидов: ' + fmtN(x.leads)]) } }) });

    draw('d7', { type: 'line', data: { labels: L, datasets: [
      line('Стоимость сделки', s.map(x => x.cpd == null ? null : Math.round(x.cpd)), COL.blue, 'rgba(94,200,242,.12)')] },
      options: opts({ plugins: { legend: { display: false },
        tooltip: tip(fmtM, (x) => ['успешных сделок: ' + fmtN(x.deals_won)]) } }) });

    draw('d1', { type: 'bar', data: { labels: L, datasets: [bars('Расход', s.map(x => x.spend), COL.acc)] },
      options: opts({ plugins: { legend: { display: false },
        tooltip: tip(fmtM, (x) => ['лидов: ' + fmtN(x.leads) + '   квал: ' + fmtN(x.kl), 'CPQL: ' + fmtM(x.cpql)]) } }) });

    draw('d2', { type: 'line', data: { labels: L, datasets: [
      line('CR лид → квал', s.map(x => x.cr == null ? null : +(x.cr * 100).toFixed(1)), COL.mid),
      line('CR без спама', s.map(x => x.crClean == null ? null : +(x.crClean * 100).toFixed(1)), COL.acc, COL.accSoft)] },
      options: opts({ scales: { x: axis(), y: pctAxis() },
        plugins: { tooltip: tip((v) => v + '%') } }) });

    draw('d3', { data: { labels: L, datasets: [
      Object.assign(bars('Создано сделок', s.map(x => x.deals_created), 'rgba(168,168,179,.35)'), { type: 'bar', order: 2 }),
      Object.assign(line('Успешно закрыто', s.map(x => x.deals_won), COL.acc), { type: 'line', order: 1 })] },
      options: opts({ plugins: { tooltip: tip(fmtN) } }) });

    draw('d4', { type: 'bar', data: { labels: L, datasets: [
      bars('Доля спама', s.map(x => x.spamShare == null ? null : +(x.spamShare * 100).toFixed(1)), 'rgba(244,83,79,.7)')] },
      options: opts({ plugins: { legend: { display: false }, tooltip: tip((v) => v + '%') },
        scales: { x: axis(), y: pctAxis() } }) });

    const M = [
      ['Расход', 'spend', fmtM, 1], ['Лиды', 'leads', fmtN, 0], ['Спам', 'spam', fmtN, 1],
      ['Квал-лиды', 'kl', fmtN, 0], ['CPL', 'cpl', fmtM, 1], ['CPL без спама', 'cplClean', fmtM, 1],
      ['CPQL', 'cpql', fmtM, 1], ['CR в квал', 'cr', (v) => fmtP(v), 0],
      ['Создано сделок', 'deals_created', fmtN, 0], ['Успешных сделок', 'deals_won', fmtN, 0],
      ['Выручка (успешные)', 'revenue_won', fmtBig, 0]
    ];
    table('dynTable', [
      { t: 'Метрика', v: r => esc(r.n) },
      { t: 'Текущее окно', num: 1, v: r => r.fmt(r.cur) },
      { t: 'Прошлое окно', num: 1, v: r => p ? r.fmt(r.prev) : dash },
      { t: 'Δ', num: 1, v: r => {
          if (!p || r.prev == null || r.cur == null || !r.prev) return `<span class="flat">${dash}</span>`;
          const d = (r.cur - r.prev) / Math.abs(r.prev);
          const good = r.lower ? d < 0 : d > 0;
          return `<span class="${Math.abs(d) < .005 ? 'flat' : good ? 'up' : 'dn'}">${d > 0 ? '▲ +' : '▼ '}${Math.abs(d * 100).toFixed(1)}%</span>`;
        } }
    ], M.map(([n, k, fmt, lower]) => ({ n, cur: a[k], prev: p ? p[k] : null, fmt, lower })), { key: 'dyn' });
  }

  /* ============================================================
     ДОЛЯ ПО ИСТОЧНИКАМ ЛИДОГЕНЕРАЦИИ
     ============================================================ */
  const SH_DIMN = { src: 'источникам', dir: 'направлениям', page: 'страницам' };
  const SH_DIMT = { src: 'Источник', dir: 'Направление', page: 'Страница' };
  function renderShare() {
    els('#shDim .chip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.d === S.shDim)));
    els('#shMetric .chip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.m === S.shMetric)));
    const rows0 = C.filterRows(D, S);
    if (!rows0.length) return showEmpty(['shKpi', 'shCharts', 'shTable', 'shFlags']);
    const sh = C.shares(D, S, S.shDim);
    const rows = sh.rows.filter(r => r.leads > 0 || r.kl > 0);
    if (!rows.length) return showEmpty(['shKpi', 'shCharts', 'shTable', 'shFlags']);
    const metricName = S.shMetric === 'kl' ? 'квал-лидов' : 'лидов';
    $('shSub').textContent = 'Кто и сколько приносит ' + metricName + ' в текущем срезе. Проценты считаются от итога СРЕЗА: '
      + 'выберите направление — и доли пересчитаются внутри него. Сортировка по клику на заголовок.';

    const top = rows.slice().sort((a, b) => (b[S.shMetric] || 0) - (a[S.shMetric] || 0));
    const first = top[0];
    const need80 = rows.filter((r, i) => r.cumKl <= .8 || i === 0).length;
    $('shKpi').innerHTML = [
      kpi({ label: SH_DIMT[S.shDim] + 'ов в срезе', display: fmtN(rows.length), value: null, prev: null,
            sub: 'всего лидов ' + fmtN(sh.raw.leads) + ' · квал ' + fmtN(sh.raw.kl) }),
      kpi({ label: 'Лидер по ' + metricName, display: esc(first.key), value: null, prev: null, acc: true,
            sub: fmtN(first[S.shMetric]) + ' — ' + fmtP(S.shMetric === 'kl' ? first.shareKl : first.shareLeads, 1) + ' среза' }),
      kpi({ label: 'Дают 80% квала', display: fmtN(need80) + ' из ' + fmtN(rows.length), value: null, prev: null,
            sub: 'на них и приходится основная работа' }),
      kpi({ label: 'CPQL лидера', display: fmtM(first.cpql), value: null, prev: null,
            sub: 'по срезу ' + fmtM(sh.total.cpql) })
    ].join('');

    $('shCharts').innerHTML = `
      <div class="card"><h3>Доля в лидах и квал-лидах</h3>
        <p class="sub">Две полосы рядом: слева доля в лидах, справа — в квал-лидах. Если вторая заметно выше первой, ${SH_DIMT[S.shDim].toLowerCase()} даёт качество, а не объём.</p>
        <div class="ch tall"><canvas id="s1"></canvas></div></div>
      <div class="card"><h3>Структура квал-лидов</h3>
        <p class="sub">Кольцо — вклад каждого в квал-лиды среза. Наведите на сектор: покажет штуки, процент и CPQL.</p>
        <div class="ch tall"><canvas id="s2"></canvas></div></div>
      <div class="card"><h3>Доля квала против доли расхода</h3>
        <p class="sub">Зелёное вправо — приносит квала больше, чем съедает бюджета. Красное влево — наоборот. Это подсказка, куда переложить деньги.</p>
        <div class="ch tall"><canvas id="s3"></canvas></div></div>
      <div class="card"><h3>Как менялась доля по периодам</h3>
        <p class="sub">Топ-6 по квал-лидам. Видно, кто набирает вес, а кто теряет.</p>
        <div class="ch tall"><canvas id="s4"></canvas></div></div>`;

    const L = top.map(r => r.key);
    draw('s1', { type: 'bar', data: { labels: L, datasets: [
      bars('Доля в лидах, %', top.map(r => r.shareLeads == null ? null : +(r.shareLeads * 100).toFixed(1)), 'rgba(168,168,179,.45)'),
      bars('Доля в квал-лидах, %', top.map(r => r.shareKl == null ? null : +(r.shareKl * 100).toFixed(1)), COL.acc)] },
      options: opts({ indexAxis: L.length > 8 ? 'y' : 'x',
        scales: L.length > 8 ? { x: pctAxis(), y: axis() } : { x: axis(), y: pctAxis() },
        plugins: { tooltip: { callbacks: {
          label: (it) => ' ' + it.dataset.label + ': ' + it.raw + '%',
          afterBody: (i) => { const r = top[i[0].dataIndex];
            return ['лидов: ' + fmtN(r.leads) + '   квал: ' + fmtN(r.kl), 'CPQL: ' + fmtM(r.cpql) + '   CR: ' + fmtP(r.cr)]; } } } } }) });

    const PAL = ['#DFFF4B', '#5EC8F2', '#FBBF24', '#F4534F', '#A78BFA', '#34D399', '#F472B6', '#A8A8B3'];
    const donut = top.slice(0, 8);
    const other = top.slice(8).reduce((a, r) => a + (r.kl || 0), 0);
    const dl = donut.map(r => r.key).concat(other > 0 ? ['прочие'] : []);
    const dv = donut.map(r => r.kl).concat(other > 0 ? [other] : []);
    draw('s2', { type: 'doughnut', data: { labels: dl, datasets: [{ data: dv,
      backgroundColor: dl.map((_, i) => PAL[i % PAL.length]), borderColor: '#141416', borderWidth: 2, hoverOffset: 10 }] },
      options: opts({ cutout: '58%', scales: {},
        plugins: { legend: { position: 'right', labels: { color: COL.mid, boxWidth: 10, usePointStyle: true, padding: 10,
          font: { family: "'Manrope',sans-serif", size: 11.5 } } },
          tooltip: { callbacks: {
            label: (it) => { const tot = dv.reduce((a, b) => a + b, 0);
              return ' ' + it.label + ': ' + fmtN(it.raw) + ' квал (' + (tot ? (it.raw / tot * 100).toFixed(1) : 0) + '%)'; },
            afterBody: (i) => { const r = donut[i[0].dataIndex];
              return r ? ['CPQL: ' + fmtM(r.cpql) + '   расход: ' + fmtM(r.spend)] : []; } } } } }) });

    const lift = rows.filter(r => r.lift != null).sort((a, b) => b.lift - a.lift);
    draw('s3', { type: 'bar', data: { labels: lift.map(r => r.key), datasets: [{
      label: 'Доля квала − доля расхода, п.п.', data: lift.map(r => +(r.lift * 100).toFixed(1)),
      backgroundColor: lift.map(r => r.lift >= 0 ? 'rgba(223,255,75,.8)' : 'rgba(244,83,79,.8)'),
      borderRadius: 5, borderSkipped: false, maxBarThickness: 26 }] },
      options: opts({ indexAxis: 'y', scales: { x: axis(), y: axis() },
        plugins: { legend: { display: false }, tooltip: { callbacks: {
          label: (it) => ' перекос: ' + (it.raw > 0 ? '+' : '') + it.raw + ' п.п.',
          afterBody: (i) => { const r = lift[i[0].dataIndex];
            return ['доля квала: ' + fmtP(r.shareKl) + '   доля расхода: ' + fmtP(r.shareSpend),
                    'квал ' + fmtN(r.kl) + ' шт · расход ' + fmtM(r.spend)]; } } } } }) });

    const ss = C.shareSeries(D, S, S.shDim, S.shMetric, 6);
    draw('s4', { type: 'line', data: { labels: ss.labels,
      datasets: ss.rows.map((r, i) => line(r.key, r.data, PAL[i % PAL.length])) },
      options: opts({ scales: { x: axis(), y: pctAxis() },
        plugins: { tooltip: { callbacks: {
          title: (i) => C.periodLabel(ss.periods[i[0].dataIndex], S.level),
          label: (it) => ' ' + it.dataset.label + ': ' + (it.raw == null ? dash : it.raw + '%') } } } }) });

    table('shTable', [
      { t: SH_DIMT[S.shDim], v: r => esc(r.key), sortVal: r => r.key },
      { t: 'Лиды', num: 1, v: r => fmtN(r.leads), sortVal: r => r.leads, f: t => fmtN(t.leads) },
      { t: 'Доля лидов', num: 1, v: r => fmtP(r.shareLeads), sortVal: r => r.shareLeads, f: () => '100%' },
      { t: 'Квал-лиды', num: 1, v: r => fmtN(r.kl), sortVal: r => r.kl, f: t => fmtN(t.kl) },
      { t: 'Доля квала', num: 1, v: r => `<b style="color:var(--acc)">${fmtP(r.shareKl)}</b>`, sortVal: r => r.shareKl, f: () => '100%' },
      { t: 'Накопл. квал', num: 1, v: r => fmtP(r.cumKl, 0), sortVal: r => r.cumKl },
      { t: 'Расход', num: 1, v: r => fmtM(r.spend), sortVal: r => r.spend, f: t => fmtM(t.spend) },
      { t: 'Доля расхода', num: 1, v: r => fmtP(r.shareSpend), sortVal: r => r.shareSpend },
      { t: 'Перекос', num: 1, v: r => r.lift == null ? dash :
          `<span class="${r.lift >= 0 ? 'up' : 'dn'}">${r.lift > 0 ? '+' : ''}${(r.lift * 100).toFixed(1)} п.п.</span>`,
        sortVal: r => r.lift },
      { t: 'CPQL', num: 1, v: r => fmtM(r.cpql), sortVal: r => r.cpql, f: t => fmtM(t.cpql) },
      { t: 'CR', num: 1, v: r => fmtP(r.cr), sortVal: r => r.cr, f: t => fmtP(t.cr) }
    ], rows, { key: 'sh', sort: { col: 3, dir: -1 }, total: sh.total });

    const fl = [];
    if (need80 <= 3 && rows.length > 4)
      fl.push({ t: 'warn', h: '80% квал-лидов дают всего ' + need80 + ' ' + (S.shDim === 'src' ? 'источника' : 'направления'),
        s: 'Концентрация высокая: сбой в любом из них сразу срежет поток. Стоит понимать, есть ли замена.' });
    const best = lift.filter(r => r.kl >= 5)[0];
    if (best && best.lift > .05)
      fl.push({ t: 'ok', h: esc(best.key) + ' даёт квала больше, чем стоит',
        s: 'доля квала ' + fmtP(best.shareKl) + ' против доли расхода ' + fmtP(best.shareSpend) +
           ' (перекос +' + (best.lift * 100).toFixed(1) + ' п.п.), CPQL ' + fmtM(best.cpql) + '. Кандидат на увеличение бюджета.' });
    const worst = lift.filter(r => r.kl >= 5 || r.spend > 0).slice(-1)[0];
    if (worst && worst.lift < -.05)
      fl.push({ t: 'bad', h: esc(worst.key) + ' съедает бюджета больше, чем приносит квала',
        s: 'доля расхода ' + fmtP(worst.shareSpend) + ' против доли квала ' + fmtP(worst.shareKl) +
           ' (перекос ' + (worst.lift * 100).toFixed(1) + ' п.п.), CPQL ' + fmtM(worst.cpql) + '. Здесь и надо разбираться.' });
    const noKl = rows.filter(r => r.spend > 20000 && r.kl === 0);
    if (noKl.length)
      fl.push({ t: 'bad', h: 'Есть расход без единого квал-лида: ' + noKl.map(r => esc(r.key)).join(', '),
        s: 'суммарно ' + fmtM(noKl.reduce((a, r) => a + r.spend, 0)) + ' в этом срезе.' });
    $('shFlags').innerHTML = fl.length ? flagsHtml(fl) : '';
  }

  /* ============================================================
     ТЕПЛОВАЯ КАРТА
     ============================================================ */
  function heatColor(pos) {
    const st = [[223, 255, 75], [251, 191, 36], [244, 83, 79]];
    const t = Math.max(0, Math.min(1, pos)) * 2;
    const i = Math.min(1, Math.floor(t)), f = t - i;
    const c = st[i].map((v, j) => Math.round(v + (st[i + 1][j] - v) * f));
    return `rgba(${c[0]},${c[1]},${c[2]},${(0.2 + 0.62 * Math.max(0, Math.min(1, pos))).toFixed(3)})`;
  }
  function renderHeat() {
    const M = { cpql: 'CPQL', cpl: 'CPL', cplClean: 'CPL без спама', cpd: 'Стоимость успешной сделки' };
    const DIMN = { dir: 'направлениям', src: 'источникам', page: 'страницам' };
    els('#heatMetric .chip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.m === S.heatMetric)));
    els('#heatDim .chip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.d === S.heatDim)));
    const h = C.heatmap(D, S, S.heatDim, S.heatMetric);
    $('heatSub').textContent = M[S.heatMetric] + ' по ' + DIMN[S.heatDim] +
      ' и периодам. Зелёное — дешевле, красное — дороже. Наведите на ячейку: покажет точное значение, расход и число лидов.';
    if (!h.rows.length || h.lo == null) {
      $('heat').innerHTML = emptyHtml('Для этой метрики в срезе нет значений: нужен ненулевой расход и хотя бы один квал-лид.');
      $('heatLegend').innerHTML = ''; $('heatTop').innerHTML = '';
      els('[data-act="reset"]').forEach(b => b.onclick = () => { S.dirs = []; S.srcs = []; S.pages = []; syncFilters(); render(); });
      return;
    }
    /* Логарифмическая шкала: CPQL различается в сотни раз, на линейной
       шкале всё кроме максимума слилось бы в один цвет. */
    const lo = Math.max(1, h.lo), hi = Math.max(lo * 1.0001, h.hi);
    const pos = (v) => (Math.log(Math.max(1, v)) - Math.log(lo)) / (Math.log(hi) - Math.log(lo));
    const head = '<tr><th class="rowh">' + esc({ dir: 'Направление', src: 'Источник', page: 'Страница' }[S.heatDim]) +
      '</th>' + h.labels.map(l => `<th>${esc(l)}</th>`).join('') + '<th>Итого</th></tr>';
    const body = h.rows.map(r => '<tr><th class="rowh" title="' + esc(r.key) + '">' + esc(r.key) + '</th>' +
      r.cells.map(c => {
        if (!c) return '<td><div class="hc nil">—</div></td>';
        const t = `${esc(r.key)} · ${C.periodLabel(c.period, S.level)}\n${M[S.heatMetric]}: ${fmtM(c.value)}\nрасход ${fmtM(c.spend)} · лидов ${fmtN(c.leads)} · квал ${fmtN(c.kl)}`;
        return `<td><div class="hc" style="background:${heatColor(pos(c.value))}" title="${esc(t)}">${Math.round(c.value / (c.value >= 10000 ? 1000 : 1))}${c.value >= 10000 ? 'k' : ''}</div></td>`;
      }).join('') +
      `<td><div class="hc" style="background:${r.total != null ? heatColor(pos(r.total)) : 'transparent'};font-weight:700" title="${esc('Итого за окно: ' + fmtM(r.total))}">${r.total == null ? '—' : Math.round(r.total / (r.total >= 10000 ? 1000 : 1)) + (r.total >= 10000 ? 'k' : '')}</div></td></tr>`).join('');
    $('heat').innerHTML = `<div class="heat-wrap"><table class="heat"><thead>${head}</thead><tbody>${body}</tbody></table></div>`;
    $('heatLegend').innerHTML = `<span>${fmtM(h.lo)}</span><div class="heat-bar"></div><span>${fmtM(h.hi)}</span>
      <span style="margin-left:8px">значения в ячейках: k = тысяч ₽ · шкала логарифмическая</span>`;

    const best = h.rows.filter(r => r.total != null && r.kl >= 3).sort((x, y) => x.total - y.total);
    if (best.length >= 2) {
      const b = best[0], w = best[best.length - 1];
      $('heatTop').innerHTML = flagsHtml([
        { t: 'ok', h: 'Дешевле всех: ' + esc(b.key) + ' — ' + fmtM(b.total),
          s: 'расход ' + fmtM(b.spend) + ', квал-лидов ' + fmtN(b.kl) + '. Считается только там, где квал-лидов ≥ 3.' },
        { t: 'bad', h: 'Дороже всех: ' + esc(w.key) + ' — ' + fmtM(w.total),
          s: 'расход ' + fmtM(w.spend) + ', квал-лидов ' + fmtN(w.kl) + '. Разница с лучшим — ' + (w.total / b.total).toFixed(1) + '×.' }
      ]);
    } else $('heatTop').innerHTML = '';
  }

  /* ============================================================
     ПЛАН-ФАКТ
     ============================================================ */
  /* ---------- хранение планов ---------- */
  function plansLoad() {
    try { PLANS = JSON.parse(localStorage.getItem(PLANS_KEY) || '{}') || {}; }
    catch (e) { PLANS = {}; }
  }
  function plansSave() {
    try { localStorage.setItem(PLANS_KEY, JSON.stringify(PLANS)); return true; }
    catch (e) { return false; }
  }
  /** План периода = план из выгрузки, поверх которого ручные правки. */
  function planFor(per, dir) {
    const base = (D.plan_fact[per] || []).find(r => r.dir === dir) || {};
    const man = (PLANS[per] || {})[dir] || {};
    const pick = (k) => man[k] != null ? man[k] : base[k];
    return { kl_plan: pick('kl_plan'), leads_plan: pick('leads_plan'),
             spend_plan: pick('spend_plan'), revenue_plan: pick('revenue_plan'),
             manual: Object.keys(man).length > 0 };
  }
  function planSet(per, dir, key, val) {
    PLANS[per] = PLANS[per] || {};
    PLANS[per][dir] = PLANS[per][dir] || {};
    if (val === '' || val == null || isNaN(+val)) delete PLANS[per][dir][key];
    else PLANS[per][dir][key] = +val;
    if (!Object.keys(PLANS[per][dir]).length) delete PLANS[per][dir];
    if (!Object.keys(PLANS[per] || {}).length) delete PLANS[per];
    plansSave();
  }

  /** Факт по направлению за период — из тех же строк, что весь дэшборд. */
  function factFor(per, dir) {
    return C.agg(C.filterRows(D, { level: 'monthly', from: per, to: per,
      dirs: [dir], srcs: S.srcs, pages: S.pages }));
  }

  function renderPF() {
    const pfPers = Object.keys(D.plan_fact || {});
    const planPers = Object.keys(PLANS);
    const pers = [...new Set(pfPers.concat(planPers))].sort();
    if (!pers.length) { $('pfBody').innerHTML = emptyHtml('Плановых периодов в данных нет.'); return; }
    if (!S.pfPeriod || !pers.includes(S.pfPeriod)) S.pfPeriod = pers[pers.length - 1];
    const per = S.pfPeriod;
    $('pfPeriods').innerHTML = pers.map(p =>
      `<button class="chip" data-p="${p}" aria-pressed="${p === per}">${esc(C.periodLabel(p, 'monthly'))}${(PLANS[p] ? ' •' : '')}</button>`).join('');
    els('#pfPeriods .chip').forEach(b => b.onclick = () => { S.pfPeriod = b.dataset.p; renderPF(); });
    $('pfCopy').innerHTML = '<option value="">скопировать план из…</option>' +
      pers.filter(p => p !== per).map(p => `<option value="${p}">${esc(C.periodLabel(p, 'monthly'))}</option>`).join('');
    $('pfClear').classList.toggle('hide', !PLANS[per]);

    /* Направления: все, у кого есть план ИЛИ факт в этом месяце. */
    const monthRows = C.filterRows(D, { level: 'monthly', from: per, to: per, dirs: [], srcs: S.srcs, pages: S.pages });
    const dirsFact = [...new Set(monthRows.map(r => r.dir).filter(Boolean))];
    const dirsPlan = (D.plan_fact[per] || []).map(r => r.dir).concat(Object.keys(PLANS[per] || {}));
    const dirs = [...new Set(dirsFact.concat(dirsPlan))].sort((a, b) => a.localeCompare(b, 'ru'));

    const rows = dirs.map(d => {
      const p = planFor(per, d), f = factFor(per, d);
      const r = { dir: d, manual: p.manual,
        kl_plan: p.kl_plan, kl_fact: f.kl,
        leads_plan: p.leads_plan, leads_fact: f.leads,
        spend_plan: p.spend_plan, spend_fact: f.spend,
        revenue_plan: p.revenue_plan, revenue_fact: f.revenue_won,
        deals_fact: f.deals_won, cpql_fact: f.cpql };
      r.kl_pct = r.kl_plan ? r.kl_fact / r.kl_plan : null;
      r.leads_pct = r.leads_plan ? r.leads_fact / r.leads_plan : null;
      r.spend_pct = r.spend_plan ? r.spend_fact / r.spend_plan : null;
      r.cpql_plan = (r.kl_plan && r.spend_plan) ? r.spend_plan / r.kl_plan : null;
      return r;
    });

    /* Итог считается только по направлениям, где план ЕСТЬ: иначе
       процент выполнения делился бы на неполный план и врал. */
    const withPlan = rows.filter(r => r.kl_plan != null || r.leads_plan != null || r.spend_plan != null);
    const T = withPlan.reduce((a, r) => {
      a.kl_plan += r.kl_plan || 0; a.kl_fact += r.kl_plan != null ? r.kl_fact : 0;
      a.leads_plan += r.leads_plan || 0; a.leads_fact += r.leads_plan != null ? r.leads_fact : 0;
      a.spend_plan += r.spend_plan || 0; a.spend_fact += r.spend_plan != null ? r.spend_fact : 0;
      return a; }, { kl_plan: 0, kl_fact: 0, leads_plan: 0, leads_fact: 0, spend_plan: 0, spend_fact: 0 });
    const allFact = rows.reduce((a, r) => { a.deals += r.deals_fact; a.rev += r.revenue_fact; return a; }, { deals: 0, rev: 0 });

    /* Предупреждения ровно про то, что искажает проценты. */
    const noPlan = rows.filter(r => r.kl_plan == null && r.leads_plan == null && r.spend_plan == null && (r.kl_fact || r.leads_fact || r.spend_fact));
    const noFact = rows.filter(r => (r.kl_plan != null || r.spend_plan != null) && !r.kl_fact && !r.leads_fact && !r.spend_fact);
    let notes = '';
    if (noPlan.length) notes += `<div class="note warn"><b>Направления без плана:</b> ${noPlan.map(r => esc(r.dir)).join(', ')}.
      Факт по ним есть, план не задан — в проценты выполнения они не попадают. Задайте план в таблице ниже, и они включатся в итог.</div>`;
    if (noFact.length) notes += `<div class="note"><b>План задан, факта нет:</b> ${noFact.map(r => esc(r.dir)).join(', ')}.</div>`;
    if (PLANS[per]) notes += `<div class="note"><b>В этом месяце есть ручные правки плана</b> (отмечены точкой). Они хранятся в этом браузере и накладываются поверх выгрузки — сама таблица не меняется.</div>`;
    if (!withPlan.length) notes = `<div class="note warn"><b>На этот месяц плана нет.</b> Показан только факт. Введите план в таблице ниже — проценты и полосы появятся сразу.</div>` + notes;
    $('pfNote').innerHTML = notes;

    /* ---------- инфографика выполнения ---------- */
    const gauge = (label, fact, plan, fmt, lowerBetter) => {
      const p = plan ? fact / plan : null;
      const cls = p == null ? 'b-nil' : lowerBetter
        ? (p <= 1.05 ? 'b-ok' : p <= 1.2 ? 'b-warn' : 'b-bad')
        : (p >= .95 ? 'b-ok' : p >= .7 ? 'b-warn' : 'b-bad');
      const col = cls === 'b-ok' ? COL.acc : cls === 'b-warn' ? COL.amber : cls === 'b-bad' ? COL.red : COL.dim;
      const w = p == null ? 0 : Math.max(1.5, Math.min(100, p * 100));
      return `<div class="kpi"><div class="kpi-l">${esc(label)}</div>
        <div class="kpi-v" style="color:${col}">${p == null ? dash : (p * 100).toFixed(0) + '%'}</div>
        <div class="kpi-s">${fmt(fact)} из ${plan ? fmt(plan) : dash}</div>
        <div class="pf-bar" style="height:8px;margin-top:9px"><i style="width:${w}%;background:${col}"></i></div></div>`;
    };
    $('pfKpi').innerHTML =
      gauge('Квал-лиды', T.kl_fact, T.kl_plan, fmtN) +
      gauge('Лиды', T.leads_fact, T.leads_plan, fmtN) +
      gauge('Расход (освоение)', T.spend_fact, T.spend_plan, fmtBig, true) +
      `<div class="kpi acc"><div class="kpi-l">Успешные сделки</div>
        <div class="kpi-v">${fmtN(allFact.deals)}</div>
        <div class="kpi-s">выручка ${fmtBig(allFact.rev)}</div></div>`;

    /* ---------- таблица с редактируемым планом ---------- */
    const num = (v) => v == null ? '' : (Math.round(v * 100) / 100);
    const inp = (r, k, step) => `<input class="pin mono" type="number" min="0" step="${step}"
      data-dir="${esc(r.dir)}" data-k="${k}" value="${num(r[k])}" placeholder="—"
      aria-label="${esc(r.dir)} · план"/>`;
    const bar = (v, lowerBetter) => {
      if (v == null) return '<span class="badge b-nil">нет плана</span>';
      const cls = lowerBetter ? (v <= 1.05 ? 'ok' : v <= 1.2 ? 'warn' : 'bad')
                              : (v >= .95 ? 'ok' : v >= .7 ? 'warn' : 'bad');
      const col = cls === 'ok' ? COL.acc : cls === 'warn' ? COL.amber : COL.red;
      return `<div style="display:flex;align-items:center;gap:8px;justify-content:flex-end">
        <div class="pf-bar" style="width:62px"><i style="width:${Math.min(100, v * 100).toFixed(0)}%;background:${col}"></i></div>
        <span class="mono" style="color:${col};font-size:12px">${(v * 100).toFixed(0)}%</span></div>`;
    };
    const cols = [
      { t: 'Направление', v: r => esc(r.dir) + (r.manual ? ' <span style="color:var(--acc)">•</span>' : ''), sortVal: r => r.dir },
      { t: 'Квал план', num: 1, v: r => inp(r, 'kl_plan', 1), sortVal: r => r.kl_plan },
      { t: 'Квал факт', num: 1, v: r => fmtN(r.kl_fact), sortVal: r => r.kl_fact },
      { t: 'Выполнение квал', num: 1, v: r => bar(r.kl_pct), sortVal: r => r.kl_pct },
      { t: 'Лиды план', num: 1, v: r => inp(r, 'leads_plan', 1), sortVal: r => r.leads_plan },
      { t: 'Лиды факт', num: 1, v: r => fmtN(r.leads_fact), sortVal: r => r.leads_fact },
      { t: 'Выполнение лидов', num: 1, v: r => bar(r.leads_pct), sortVal: r => r.leads_pct },
      { t: 'Расход план', num: 1, v: r => inp(r, 'spend_plan', 1000), sortVal: r => r.spend_plan },
      { t: 'Расход факт', num: 1, v: r => fmtM(r.spend_fact), sortVal: r => r.spend_fact },
      { t: 'Освоение', num: 1, v: r => bar(r.spend_pct, true), sortVal: r => r.spend_pct },
      { t: 'CPQL план', num: 1, v: r => fmtM(r.cpql_plan), sortVal: r => r.cpql_plan },
      { t: 'CPQL факт', num: 1, v: r => fmtM(r.cpql_fact), sortVal: r => r.cpql_fact },
      { t: 'Выручка', num: 1, v: r => fmtBig(r.revenue_fact), sortVal: r => r.revenue_fact }
    ];
    $('pfBody').innerHTML = '<div id="pfT"></div>';
    table('pfT', cols, rows, { key: 'pf', sort: { col: 2, dir: -1 } });
    /* Правка плана сохраняется по уходу с поля, а не по каждой цифре:
       иначе таблица перерисовывалась бы на каждое нажатие клавиши. */
    els('#pfT .pin').forEach(i => {
      i.onkeydown = (e) => { if (e.key === 'Enter') i.blur(); };
      i.onchange = () => { planSet(per, i.dataset.dir, i.dataset.k, i.value); renderPF(); };
    });

    /* ---------- графики ---------- */
    $('pfCharts').innerHTML = `
      <div class="card"><h3>Квал-лиды: план и факт</h3>
        <p class="sub">Серое — план, цвет — факт: зелёный при выполнении, красный при провале.</p>
        <div class="ch"><canvas id="p1"></canvas></div></div>
      <div class="card"><h3>Расход: план и факт</h3>
        <p class="sub">Перерасход при недоборе квал-лидов — худший из сценариев.</p>
        <div class="ch"><canvas id="p2"></canvas></div></div>
      <div class="card"><h3>Выполнение плана по квал-лидам, %</h3>
        <p class="sub">Линия 100% — план. Всё, что ниже, — недобор.</p>
        <div class="ch"><canvas id="p3"></canvas></div></div>
      <div class="card"><h3>CPQL: план и факт</h3>
        <p class="sub">План CPQL = план расхода ÷ план квал-лидов. Факт выше плана — квал дороже ожидаемого.</p>
        <div class="ch"><canvas id="p4"></canvas></div></div>`;
    const dl = rows.map(r => r.dir);
    const colFor = (v, lower) => v == null ? 'rgba(168,168,179,.5)'
      : (lower ? (v <= 1.05 ? COL.acc : v <= 1.2 ? COL.amber : COL.red)
               : (v >= .95 ? COL.acc : v >= .7 ? COL.amber : COL.red));
    const ptip = (fmt) => ({ callbacks: {
      label: (it) => ' ' + it.dataset.label + ': ' + fmt(it.raw),
      afterBody: (i) => { const r = rows[i[0].dataIndex];
        return ['выполнение квал: ' + (r.kl_pct == null ? 'нет плана' : (r.kl_pct * 100).toFixed(0) + '%'),
                'освоение бюджета: ' + (r.spend_pct == null ? 'нет плана' : (r.spend_pct * 100).toFixed(0) + '%')]; } } });
    draw('p1', { type: 'bar', data: { labels: dl, datasets: [
      bars('План', rows.map(r => r.kl_plan), 'rgba(168,168,179,.3)'),
      { label: 'Факт', data: rows.map(r => r.kl_fact), backgroundColor: rows.map(r => colFor(r.kl_pct)),
        borderRadius: 5, borderSkipped: false, maxBarThickness: 36 }] },
      options: opts({ plugins: { tooltip: ptip(fmtN) } }) });
    draw('p2', { type: 'bar', data: { labels: dl, datasets: [
      bars('План', rows.map(r => r.spend_plan), 'rgba(168,168,179,.3)'),
      { label: 'Факт', data: rows.map(r => r.spend_fact), backgroundColor: rows.map(r => colFor(r.spend_pct, true)),
        borderRadius: 5, borderSkipped: false, maxBarThickness: 36 }] },
      options: opts({ plugins: { tooltip: ptip(fmtM) } }) });
    draw('p3', { type: 'bar', data: { labels: dl, datasets: [{
      label: 'Выполнение, %', data: rows.map(r => r.kl_pct == null ? null : +(r.kl_pct * 100).toFixed(0)),
      backgroundColor: rows.map(r => colFor(r.kl_pct)), borderRadius: 5, borderSkipped: false, maxBarThickness: 36 }] },
      options: opts({ scales: { x: axis(), y: pctAxis() }, plugins: { legend: { display: false },
        tooltip: { callbacks: { label: (it) => it.raw == null ? ' нет плана' : ' выполнение: ' + it.raw + '%' } } } }) });
    draw('p4', { data: { labels: dl, datasets: [
      Object.assign(bars('CPQL план', rows.map(r => r.cpql_plan), 'rgba(168,168,179,.3)'), { type: 'bar' }),
      Object.assign(bars('CPQL факт', rows.map(r => r.cpql_fact), COL.blue), { type: 'bar' })] },
      options: opts({ plugins: { tooltip: { callbacks: { label: (it) => ' ' + it.dataset.label + ': ' + fmtM(it.raw) } } } }) });

    /* ---------- выводы ---------- */
    const fl = [];
    rows.forEach(r => {
      if (r.kl_pct != null && r.spend_pct != null && r.kl_pct < .7 && r.spend_pct > 1)
        fl.push({ t: 'bad', h: esc(r.dir) + ': бюджет перерасходован при недоборе квал-лидов',
          s: 'квал ' + (r.kl_pct * 100).toFixed(0) + '% плана (' + fmtN(r.kl_fact) + ' из ' + fmtN(r.kl_plan) +
             '), расход ' + (r.spend_pct * 100).toFixed(0) + '% (' + fmtM(r.spend_fact) + ' из ' + fmtM(r.spend_plan) + ').' });
      else if (r.kl_pct != null && r.kl_pct >= 1 && r.spend_pct != null && r.spend_pct <= 1)
        fl.push({ t: 'ok', h: esc(r.dir) + ': план выполнен в рамках бюджета',
          s: 'квал ' + fmtN(r.kl_fact) + ' из ' + fmtN(r.kl_plan) + ' при ' + (r.spend_pct * 100).toFixed(0) + '% бюджета.' });
      if (r.cpql_plan && r.cpql_fact && r.cpql_fact / r.cpql_plan > 1.5)
        fl.push({ t: 'warn', h: esc(r.dir) + ': квал-лид дороже плана в ' + (r.cpql_fact / r.cpql_plan).toFixed(1) + '×',
          s: 'план ' + fmtM(r.cpql_plan) + ', факт ' + fmtM(r.cpql_fact) + '.' });
    });
    $('pfFlags').innerHTML = fl.length ? flagsHtml(fl.slice(0, 8)) : '';
  }

  /** Копирование плана из другого месяца — чтобы не вбивать заново. */
  function pfCopyFrom(srcPer) {
    const per = S.pfPeriod;
    if (!srcPer || srcPer === per) return;
    const dirs = new Set((D.plan_fact[srcPer] || []).map(r => r.dir).concat(Object.keys(PLANS[srcPer] || {})));
    dirs.forEach(d => {
      const p = planFor(srcPer, d);
      ['kl_plan', 'leads_plan', 'spend_plan', 'revenue_plan'].forEach(k => {
        if (p[k] != null) planSet(per, d, k, p[k]);
      });
    });
    renderPF();
  }
  function pfAddDir(name) {
    const d = String(name || '').trim();
    if (!d) return;
    PLANS[S.pfPeriod] = PLANS[S.pfPeriod] || {};
    PLANS[S.pfPeriod][d] = PLANS[S.pfPeriod][d] || { kl_plan: 0 };
    plansSave(); renderPF();
  }
  function pfClearManual() {
    delete PLANS[S.pfPeriod]; plansSave(); renderPF();
  }
  function pfExport() {
    const per = S.pfPeriod;
    const rows = els('#pfT tbody tr').length ? null : null;
    const dirs = [...new Set((D.plan_fact[per] || []).map(r => r.dir)
      .concat(Object.keys(PLANS[per] || {}))
      .concat(C.filterRows(D, { level: 'monthly', from: per, to: per, dirs: [], srcs: [], pages: [] }).map(r => r.dir)))].filter(Boolean).sort();
    const head = ['Месяц','Направление','Квал план','Квал факт','% квал','Лиды план','Лиды факт','% лиды',
                  'Расход план','Расход факт','% расход','CPQL план','CPQL факт','Успешных сделок','Выручка'];
    const q = (v) => v == null ? '' : /[";\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
    const pc = (f, p) => p ? (f / p * 100).toFixed(0) + '%' : '';
    const body = dirs.map(d => {
      const p = planFor(per, d), f = factFor(per, d);
      const cpqlP = (p.kl_plan && p.spend_plan) ? p.spend_plan / p.kl_plan : null;
      return [C.periodLabel(per, 'monthly'), d, p.kl_plan, f.kl, pc(f.kl, p.kl_plan),
        p.leads_plan, f.leads, pc(f.leads, p.leads_plan),
        p.spend_plan, Math.round(f.spend), pc(f.spend, p.spend_plan),
        cpqlP == null ? '' : Math.round(cpqlP), f.cpql == null ? '' : Math.round(f.cpql),
        f.deals_won, Math.round(f.revenue_won)].map(q).join(';');
    });
    const csv = '\uFEFF' + [head.join(';')].concat(body).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = 'alfakpd-plan-fakt-' + per + '.csv';
    a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  /* ============================================================
     ПРОГНОЗ
     ============================================================ */
  function renderFC() {
    const s = C.series(D, S);
    if (s.length < 3) return showEmpty(['fcKpi', 'fcCharts', 'fcNote'],
      'Для прогноза нужно минимум 3 периода в срезе — расширьте окно.');
    els('#fcMethod .chip').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.f === S.fcMethod)));
    const MN = { last3: 'среднее за 3 последних периода', trend: 'линейный тренд по всему окну', mom: 'средний темп прироста' };
    $('fcNote').innerHTML = `<div class="note">Метод: <b>${MN[S.fcMethod]}</b>. Прогноз — арифметика по истории среза, а не модель: он не знает про сезон, акции и изменения бюджета. Пунктир на графике — прогнозная точка.</div>`;
    const F = (k) => C.forecast(s.map(x => x[k]), S.fcMethod);
    const f = { spend: F('spend'), leads: F('leads'), kl: F('kl') };
    f.cpql = f.kl > 0 ? f.spend / f.kl : null;
    f.cpl = f.leads > 0 ? f.spend / f.leads : null;
    const last = s[s.length - 1];
    $('fcKpi').innerHTML = [
      kpi({ label: 'Расход · прогноз', display: fmtBig(f.spend), value: f.spend, prev: last.spend, lower: true,
            sub: 'последний факт ' + fmtBig(last.spend) }),
      kpi({ label: 'Квал-лиды · прогноз', display: fmtN(f.kl), value: f.kl, prev: last.kl, acc: true,
            sub: 'последний факт ' + fmtN(last.kl) }),
      kpi({ label: 'CPQL · прогноз', display: fmtM(f.cpql), value: f.cpql, prev: last.cpql, lower: true, acc: true,
            sub: 'последний факт ' + fmtM(last.cpql) }),
      kpi({ label: 'CPL · прогноз', display: fmtM(f.cpl), value: f.cpl, prev: last.cpl, lower: true,
            sub: 'последний факт ' + fmtM(last.cpl) })
    ].join('');
    $('fcCharts').innerHTML = `
      <div class="card"><h3>Квал-лиды: факт и прогноз</h3>
        <p class="sub">Сплошная — факт, пунктир — прогнозная точка.</p>
        <div class="ch tall"><canvas id="f1"></canvas></div></div>
      <div class="card"><h3>CPQL: факт и прогноз</h3>
        <p class="sub">Если прогноз CPQL выше факта — при том же бюджете квал-лидов будет меньше.</p>
        <div class="ch tall"><canvas id="f2"></canvas></div></div>`;
    const L = s.map(x => x.label).concat(['прогноз']);
    const withF = (vals, fv) => vals.concat([fv == null ? null : Math.round(fv)]);
    const mk = (label, vals, fv, col) => ({ labels: L, datasets: [
      Object.assign(line(label + ' · факт', vals.concat([null]), col), {}),
      { label: label + ' · прогноз', data: vals.map(() => null).slice(0, -1)
          .concat([vals[vals.length - 1], fv == null ? null : Math.round(fv)]),
        borderColor: col, borderDash: [6, 4], borderWidth: 2, pointRadius: 4,
        pointBackgroundColor: '#0E0E0F', pointBorderColor: col, pointBorderWidth: 2, tension: 0, fill: false } ] });
    draw('f1', { type: 'line', data: mk('Квал-лиды', s.map(x => x.kl), f.kl, COL.acc),
      options: opts({ plugins: { tooltip: { callbacks: { label: (it) => ' ' + it.dataset.label + ': ' + fmtN(it.raw) } } } }) });
    draw('f2', { type: 'line', data: mk('CPQL', s.map(x => x.cpql == null ? null : Math.round(x.cpql)), f.cpql, COL.amber),
      options: opts({ plugins: { tooltip: { callbacks: { label: (it) => ' ' + it.dataset.label + ': ' + fmtM(it.raw) } } } }) });
  }

  /* ============================================================
     ДАННЫЕ
     ============================================================ */
  function renderTable() {
    const rows = C.filterRows(D, S);
    if (!rows.length) return showEmpty(['tblBody']);
    $('tblCount').textContent = fmtN(rows.length) + ' строк в срезе';
    const shown = rows.slice(0, 3000);
    $('tblBody').innerHTML = '<div id="tblT"></div>';
    table('tblT', [
      { t: 'Период', v: r => esc(C.periodLabel(r.period, S.level)), sortVal: r => r.period },
      { t: 'Направление', v: r => esc(r.dir || dash), sortVal: r => r.dir || '' },
      { t: 'Источник', v: r => esc(r.src || dash), sortVal: r => r.src || '' },
      { t: 'Страница', v: r => esc(r.page || dash), sortVal: r => r.page || '' },
      { t: 'Расход', num: 1, v: r => fmtM(r.spend), sortVal: r => r.spend },
      { t: 'Лиды', num: 1, v: r => fmtN(r.leads), sortVal: r => r.leads },
      { t: 'Спам', num: 1, v: r => fmtN(r.spam), sortVal: r => r.spam },
      { t: 'Квал', num: 1, v: r => fmtN(r.kl), sortVal: r => r.kl },
      { t: 'CPQL', num: 1, v: r => { const v = r.kl ? r.spend / r.kl : null; return fmtM(v); }, sortVal: r => r.kl ? r.spend / r.kl : null },
      { t: 'Сделок', num: 1, v: r => fmtN(r.deals_won), sortVal: r => r.deals_won },
      { t: 'Выручка', num: 1, v: r => fmtBig(r.revenue_won), sortVal: r => r.revenue_won }
    ], shown, { key: 'tbl', sort: { col: 4, dir: -1 }, total: C.agg(rows) });
    if (rows.length > 3000)
      $('tblBody').insertAdjacentHTML('afterbegin',
        '<div class="note warn">Показаны первые 3 000 строк из ' + fmtN(rows.length) + ' — сузьте срез фильтрами.</div>');
  }
  function exportCsv() {
    const rows = C.filterRows(D, S);
    const cols = ['period','dir','src','page','spend','leads','spam','kl','deals_created','deals_won','revenue_won'];
    const head = ['Период','Направление','Источник','Страница','Расход','Лиды','Спам','Квал','Создано сделок','Успешных сделок','Выручка'];
    const q = (v) => v == null ? '' : /[";\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
    const csv = '\uFEFF' + [head.join(';')].concat(rows.map(r => cols.map(c => q(r[c])).join(';'))).join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
    a.download = 'alfakpd-' + S.level + '-' + (S.from || '') + '_' + (S.to || '') + '.csv';
    a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  /* ============================================================
     ГУГЛ-ТАБЛИЦА: подключение и проверка
     ============================================================ */
  const GS_KEY = 'alfakpd.gsurl.v1';
  const GS = { rows: null, log: [] };
  const SHEET_ID = (D && D.meta && D.meta.source && D.meta.source.sheetId) || '';

  function gsLog(ok, msg, detail) {
    GS.log.push({ ok, msg, detail });
    $('gsLog').innerHTML = GS.log.map(l =>
      `<div class="flag ${l.ok === true ? 'ok' : l.ok === false ? 'bad' : 'warn'}"><i></i><div>
        <b>${esc(l.msg)}</b>${l.detail ? `<span>${esc(l.detail)}</span>` : ''}</div></div>`).join('');
  }

  /** Проверка связи: тянем данные и сверяем с тем, что уже в дэшборде.
      Проверяем не «ответил ли сервер», а сходятся ли ЦИФРЫ — иначе
      «связь есть» ничего не значит. */
  async function gsCheck() {
    const url = ($('gsUrl').value || '').trim();
    GS.log = []; GS.rows = null;
    if (!url) { gsLog(false, 'Ссылка не указана', 'Вставьте адрес веб-приложения из инструкции.'); return; }
    try { localStorage.setItem(GS_KEY, url); } catch (e) {}
    S.gsUrl = url;
    gsLog(null, 'Шаг 1 · Отправляю запрос', url.slice(0, 90) + (url.length > 90 ? '…' : ''));
    let res, txt;
    try {
      res = await fetch(url, { method: 'GET', redirect: 'follow' });
    } catch (e) {
      gsLog(false, 'Шаг 1 · Соединение не установлено', e.message +
        ' — почти всегда это значит, что в развёртывании скрипта не выбрано «Доступ: у кого есть ссылка», либо ссылка не та.');
      return;
    }
    gsLog(res.ok, 'Шаг 2 · Ответ получен: HTTP ' + res.status, res.ok ? '' : 'Скрипт вернул ошибку. Откройте ссылку в браузере — увидите текст ошибки.');
    if (!res.ok) return;
    try { txt = await res.text(); } catch (e) { gsLog(false, 'Шаг 2 · Тело ответа не прочитано', e.message); return; }
    if (/^\s*</.test(txt)) {
      gsLog(false, 'Шаг 3 · Пришла HTML-страница вместо данных',
        'Это страница входа Google. В развёртывании нужно «Выполнять от имени: я» и «Доступ: у кого есть ссылка».');
      return;
    }
    let json;
    try { json = JSON.parse(txt); } catch (e) { gsLog(false, 'Шаг 3 · Ответ не разобран как JSON', txt.slice(0, 160)); return; }
    const arr = Array.isArray(json) ? json : (json.weekly || json.rows || json.data);
    if (!Array.isArray(arr)) { gsLog(false, 'Шаг 3 · В ответе нет массива строк', 'Ожидались поля weekly / rows / data.'); return; }
    gsLog(true, 'Шаг 3 · Данные разобраны: строк ' + fmtN(arr.length));

    const need = ['period', 'dir', 'src', 'spend', 'leads', 'kl'];
    const miss = need.filter(k => !(k in (arr[0] || {})));
    gsLog(!miss.length, 'Шаг 4 · Обязательные поля', miss.length ? 'не хватает: ' + miss.join(', ') : 'все на месте: ' + need.join(', '));
    if (miss.length) return;

    const sum = (rs, k) => rs.reduce((a, r) => a + (+r[k] || 0), 0);
    const gsSpend = sum(arr, 'spend'), gsLeads = sum(arr, 'leads'), gsKl = sum(arr, 'kl');
    const mine = C.agg(D.weekly);
    const near = (a, b, tol) => b === 0 ? a === 0 : Math.abs(a - b) / Math.abs(b) <= tol;
    const okS = near(gsSpend, mine.spend, .01), okL = near(gsLeads, mine.leads, .01), okK = near(gsKl, mine.kl, .01);
    gsLog(okS, 'Шаг 5 · Расход сходится', 'из таблицы ' + fmtM(gsSpend) + ' · в дэшборде ' + fmtM(mine.spend));
    gsLog(okL, 'Шаг 5 · Лиды сходятся', 'из таблицы ' + fmtN(gsLeads) + ' · в дэшборде ' + fmtN(mine.leads));
    gsLog(okK, 'Шаг 5 · Квал-лиды сходятся', 'из таблицы ' + fmtN(gsKl) + ' · в дэшборде ' + fmtN(mine.kl));

    const per = [...new Set(arr.map(r => String(r.period)))].sort();
    gsLog(true, 'Шаг 6 · Покрытие периодов', per.length + ' периодов, с ' + per[0] + ' по ' + per[per.length - 1] +
      ' · в дэшборде последняя неделя ' + D.meta.coverage.weekTo);
    /* Свежее = в таблице появились недели ПОЗЖЕ последней в дэшборде.
       Это не ошибка, а повод нажать «Обновить данные». */
    const newer = per[per.length - 1] > D.meta.coverage.weekTo;
    gsLog(true, newer ? 'Шаг 6 · В таблице есть более свежие недели' : 'Шаг 6 · Данные в дэшборде не отстают',
      newer ? 'Последняя в таблице ' + per[per.length - 1] + ', в дэшборде ' + D.meta.coverage.weekTo +
              ' — нажмите «Обновить данные».' : 'Обновление не требуется.');
    GS.rows = arr;
    const allOk = okS && okL && okK && !miss.length;
    gsLog(allOk, allOk ? 'ИТОГ · Интеграция работает, цифры совпадают' : 'ИТОГ · Связь есть, но цифры расходятся',
      allOk ? 'Можно обновлять данные этой кнопкой.' : 'Сверьте, что скрипт читает лист «Недельный отчёт по трафику» целиком, а не фильтрованный вид.');
    $('gsApply').disabled = !GS.rows;
  }

  /** Замена недельных данных на пришедшие из таблицы. Месячные не
      трогаем: скрипт из инструкции отдаёт недельный лист. */
  function gsApply() {
    if (!GS.rows) return;
    const norm = (r) => ({ period: String(r.period).slice(0, 10), end: r.end || null,
      dir: r.dir || null, src: r.src || null, page: r.page || null,
      spend: +r.spend || 0, leads: +r.leads || 0, kl: +r.kl || 0,
      inwork: +r.inwork || 0, badlead: +r.badlead || 0, spam: +r.spam || 0, existing: +r.existing || 0,
      deals_created: +r.deals_created || 0, pipeline: +r.pipeline || 0,
      deals_lost: +r.deals_lost || 0, revenue_lost: +r.revenue_lost || 0,
      deals_won: +r.deals_won || 0, revenue_won: +r.revenue_won || 0 });
    D.weekly = GS.rows.map(norm);
    const ps = [...new Set(D.weekly.map(r => r.period))].sort();
    D.meta.coverage.weeks = ps.length;
    D.meta.coverage.weekFrom = ps[0]; D.meta.coverage.weekTo = ps[ps.length - 1];
    D.meta.builtAt = new Date().toISOString();
    window.__D = D;   // наружу отдаём тот же объект, что использует модуль
    gsLog(true, 'Данные обновлены', 'недель ' + ps.length + ', последняя ' + ps[ps.length - 1] +
      '. Обновление живёт до перезагрузки страницы: сам файл не перезаписывается.');
    buildPeriods(false); syncFilters(); render();
  }

  function renderGS() {
    const url = SHEET_ID ? 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit' : '';
    $('gsBody').innerHTML = `
      <div class="note"><b>Зачем нужен скрипт.</b> Браузер не может читать вашу таблицу напрямую:
        она закрыта для посторонних, а межсайтовые запросы к Google запрещены политикой безопасности.
        Поэтому данные отдаёт маленький скрипт, который живёт внутри самой таблицы —
        доступ остаётся у вас, наружу уходит только то, что вы разрешили.</div>

      <div class="card" style="margin-bottom:14px">
        <h3>Шаг 1 · Создать скрипт в таблице</h3>
        <p class="sub">В таблице${url ? ' (<a href="' + url + '" target="_blank" rel="noopener">открыть</a>)' : ''}:
          меню <b>Расширения → Apps Script</b>. Удалите содержимое файла и вставьте это:</p>
        <div class="tw" style="max-height:290px"><pre class="code" id="gsCode">${esc(GS_CODE)}</pre></div>
        <div style="margin-top:10px"><button class="btn" id="gsCopy">Скопировать код</button></div>
      </div>

      <div class="card" style="margin-bottom:14px">
        <h3>Шаг 2 · Опубликовать</h3>
        <p class="sub">Кнопка <b>Начать развёртывание → Новое развёртывание</b>.
          Тип — <b>Веб-приложение</b>. Дальше два параметра, и оба важны:</p>
        <ul class="ul">
          <li><b>Выполнять от имени: я</b> — иначе скрипт не увидит закрытую таблицу.</li>
          <li><b>Доступ: у кого есть ссылка</b> — иначе дэшборд получит страницу входа Google вместо данных.</li>
        </ul>
        <p class="sub">Google попросит разрешение — это нормально, скрипт читает только эту таблицу.
          После развёртывания скопируйте <b>URL веб-приложения</b> (заканчивается на <span class="mono">/exec</span>).</p>
      </div>

      <div class="card" style="margin-bottom:14px">
        <h3>Шаг 3 · Подключить и проверить</h3>
        <p class="sub">Вставьте ссылку и нажмите «Проверить связь». Проверка сверяет не факт ответа,
          а совпадение сумм с тем, что уже в дэшборде — иначе «связь есть» ничего не значит.</p>
        <div style="display:flex;gap:9px;flex-wrap:wrap;align-items:center">
          <input type="text" id="gsUrl" placeholder="https://script.google.com/macros/s/.../exec"
            style="flex:1;min-width:280px" value="${esc(S.gsUrl || '')}"/>
          <button class="btn solid" id="gsRun">Проверить связь</button>
          <button class="btn" id="gsApply" disabled>Обновить данные</button>
        </div>
        <div class="flags" id="gsLog" style="margin-top:14px"></div>
      </div>

      <div class="card">
        <h3>Чек-лист: что проверить, если не работает</h3>
        <table class="chk"><tbody>
          ${[['Ссылка заканчивается на /exec', 'Адрес, заканчивающийся на /dev, работает только для вас и только в открытой сессии.'],
             ['В развёртывании «Доступ: у кого есть ссылка»', 'Самая частая причина. Признак — в проверке приходит HTML вместо данных.'],
             ['В развёртывании «Выполнять от имени: я»', 'Иначе скрипт запускается от имени читателя и не видит закрытую таблицу.'],
             ['Имя листа в скрипте совпадает с таблицей', 'В коде это SHEET_NAME. Лишний пробел в названии листа — и придёт пустой массив.'],
             ['После правки кода сделано НОВОЕ развёртывание', 'Правка кода сама по себе не меняет то, что отдаёт старая ссылка. Это ловит почти всех.'],
             ['Открыли ссылку в браузере и увидели JSON', 'Если там текст ошибки — он же и есть ответ, что починить.'],
             ['Суммы в проверке совпали', 'Если связь есть, а цифры разошлись — скрипт читает не тот лист или фильтрованный вид.']]
            .map(([a, b]) => `<tr><td class="chk-c">☐</td><td><b>${esc(a)}</b><span>${esc(b)}</span></td></tr>`).join('')}
        </tbody></table>
        <p class="sub" style="margin-top:12px">Обновление через эту вкладку действует до перезагрузки страницы:
          сам файл дэшборда не перезаписывается. Чтобы зафиксировать данные в файле, пересоберите его из выгрузки.</p>
      </div>`;
    $('gsCopy').onclick = () => {
      const t = GS_CODE;
      if (navigator.clipboard) navigator.clipboard.writeText(t).then(
        () => { $('gsCopy').textContent = 'Скопировано'; setTimeout(() => $('gsCopy').textContent = 'Скопировать код', 1600); },
        () => {});
    };
    $('gsRun').onclick = () => { gsCheck(); };
    $('gsApply').onclick = () => gsApply();
    /* Лог прошлой проверки не теряем при возврате на вкладку. */
    if (GS.log.length) { const l = GS.log; GS.log = []; l.forEach(x => gsLog(x.ok, x.msg, x.detail)); }
    $('gsApply').disabled = !GS.rows;
  }

  /* ============================================================
     МЕТОДИКА
     ============================================================ */
  function renderMethod() {
    const cov = D.meta.coverage, src = D.meta.source;
    const f = [
      ['CPL', 'Расход ÷ Лиды', 'Сколько платим за обращение. Если спама много, CPL занижен — смотрите CPL без спама.'],
      ['CPL без спама', 'Расход ÷ (Лиды − Спам)', 'Честная стоимость живого обращения.'],
      ['CPQL', 'Расход ÷ Квал-лиды', 'Главная метрика: сколько стоит лид, который реально может купить.'],
      ['CR лид → квал', 'Квал-лиды ÷ Лиды', 'Качество трафика и обработки вместе.'],
      ['CR без спама', 'Квал-лиды ÷ (Лиды − Спам)', 'Качество обработки отдельно от качества трафика.'],
      ['Доля спама', 'Спам ÷ Лиды', 'Выше 30% — CPL и конверсия по этому срезу недостоверны.'],
      ['Выручка', 'Сумма УСПЕШНЫХ сделок', 'Только закрытые деньги. Сумма созданных сделок — это pipeline, потенциал, и в выручку не идёт.'],
      ['Стоимость сделки', 'Расход ÷ Успешные сделки', 'Считается только там, где успешные сделки заполнены.'],
      ['ROMI', '(Выручка − Расход) ÷ Расход', 'Считается от УСПЕШНЫХ сделок. Если их нет — метрика не показывается, а не рисуется нулём.'],
      ['ДРР', 'Расход ÷ Выручка', 'Доля рекламных расходов в выручке.'],
      ['Отдача pipeline', 'Сумма созданных сделок ÷ Расход', 'Потенциал на рубль. Не деньги — только вероятность.'],
      ['Доля в квал-лидах', 'Квал источника ÷ Квал всего среза', 'Знаменатель — итог СРЕЗА. Выбрали направление — доли пересчитались внутри него.'],
      ['Накопленная доля', 'Сумма долей сверху вниз', 'Показывает, сколько источников дают 80% результата. Чем их меньше, тем выше риск.'],
      ['Перекос', 'Доля в квал-лидах − Доля в расходе', 'Плюс = источник приносит квала больше, чем съедает бюджета. Подсказка, куда переложить деньги.'],
      ['Выполнение плана', 'Факт ÷ План', 'Считается только там, где план задан. Направления без плана в общий процент не попадают — иначе он делился бы на неполный план.'],
      ['CPQL план', 'План расхода ÷ План квал-лидов', 'Во сколько планировали уложиться. Сравнение с фактом показывает, где просчитались.']
    ];
    $('mBody').innerHTML = `
      <div class="note"><b>Источник данных.</b> Google-таблица, листы «${esc(src.tabs[0])}» и «${esc(src.tabs[1])}».
        Покрытие: ${cov.weeks} недель (${esc(C.periodLabel(cov.weekFrom, 'weekly'))} … ${esc(C.periodLabel(cov.weekTo, 'weekly'))}),
        ${cov.months} месяцев (${esc(C.periodLabel(cov.monthFrom, 'monthly'))} … ${esc(C.periodLabel(cov.monthTo, 'monthly'))}).
        Собрано: ${new Date(D.meta.builtAt).toLocaleString('ru-RU')}.</div>
      <div class="note warn"><b>Лист «янв-июнь 2026» сознательно не используется.</b>
        В нём первый квартал посчитан дважды (формула складывает «1 кв 2026» и «Месячный 2026», где Q1 уже есть),
        а колонки «Расход», «Лиды всего» и «CRQL» содержат одну и ту же формулу. Полугодие считается здесь
        из недельного и месячного листов.</div>
      <div class="grid g2">${f.map(([n, formula, why]) => `
        <div class="card"><h3>${esc(n)}</h3>
          <p class="sub mono" style="color:var(--acc);margin-bottom:8px">${esc(formula)}</p>
          <p class="sub" style="margin:0">${esc(why)}</p></div>`).join('')}</div>
      <div class="card" style="margin-top:16px"><h3>Правила, по которым дэшборд отказывается считать</h3>
        <p class="sub">Пустой срез показывает «данных нет», а не нули: ноль читается как факт «расход нулевой».
          Метрики с делением на ноль показывают «—», а не 0. Пустой знаменатель никогда не заменяется единицей.
          Оценочная выручка не подставляется: если успешных сделок нет, ROMI и ДРР просто не выводятся.</p></div>
      ${(D.meta.issuesTotal || 0) ? `<div class="note bad" style="margin-top:14px"><b>Замечаний к данным в таблице: ${fmtN(D.meta.issuesTotal)}</b>
        — строки, где квал-лидов больше, чем лидов, или в числовой колонке лежит текст. Это расхождения в самой
        таблице, дэшборд их не правит, а показывает.</div>` : ''}`;
  }

  /* ============================================================
     КАЛЕНДАРЬ НЕДЕЛЬ
     ============================================================ */
  function calWeeks(year) {
    return C.periodList(D, 'weekly').filter(p => +p.slice(0, 4) === year);
  }
  function calYears() {
    return [...new Set(C.periodList(D, 'weekly').map(p => +p.slice(0, 4)))].sort();
  }
  function calOpen(on) {
    CAL.open = !!on;
    $('cal').classList.toggle('hide', !on);
    $('calBack').classList.toggle('hide', !on);
    $('calBtn').setAttribute('aria-expanded', String(!!on));
    if (on) {
      CAL.pick = null;
      CAL.year = +(S.to || S.from || C.periodList(D, 'weekly').slice(-1)[0]).slice(0, 4);
      calRender();
      $('calClose').focus();
    }
  }
  function calRender() {
    const ys = calYears();
    if (CAL.year == null) CAL.year = ys[ys.length - 1];
    $('calYear').textContent = CAL.year;
    $('calPrev').disabled = CAL.year <= ys[0];
    $('calNext').disabled = CAL.year >= ys[ys.length - 1];
    $('calHint').textContent = CAL.pick
      ? 'Начало: ' + C.periodLabel(CAL.pick, 'weekly') + '. Теперь выберите неделю конца.'
      : 'Кликните неделю начала, затем неделю конца. Или возьмите готовый интервал внизу.';
    const ws = calWeeks(CAL.year);
    const byMon = new Map();
    ws.forEach(p => {
      const m = +p.slice(5, 7);
      if (!byMon.has(m)) byMon.set(m, []);
      byMon.get(m).push(p);
    });
    $('calBody').innerHTML = [...byMon.entries()].map(([m, list]) => `
      <div class="cal-mon"><h4>${esc(C.MONRU[m - 1])}</h4><div class="cal-weeks">${list.map(p => {
        const inR = S.from && S.to && p >= S.from && p <= S.to;
        const edge = p === S.from || p === S.to || p === CAL.pick;
        const d = new Date(p + 'T00:00:00Z'), e = new Date(d.getTime() + 6 * 864e5);
        const f = (x) => String(x.getUTCDate()).padStart(2, '0') + '.' + String(x.getUTCMonth() + 1).padStart(2, '0');
        return `<button class="cw${edge ? ' edge' : inR ? ' in' : ''}" data-w="${p}">${f(d)}<small>по ${f(e)}</small></button>`;
      }).join('')}</div></div>`).join('') || '<div class="empty">В этом году недель нет</div>';
    els('#calBody .cw').forEach(b => b.onclick = () => calPick(b.dataset.w));
    $('calSel').textContent = S.from && S.to
      ? C.periodLabel(S.from, 'weekly') + ' → ' + C.periodLabel(S.to, 'weekly')
      : 'интервал не выбран';
  }
  function calPick(p) {
    if (!CAL.pick) { CAL.pick = p; calRender(); return; }
    let a = CAL.pick, b = p;
    if (a > b) { const t = a; a = b; b = t; }
    S.from = a; S.to = b; CAL.pick = null;
    calRender(); calSyncBtn(); render();
  }
  function calQuick(n) {
    const ps = C.periodList(D, 'weekly');
    if (n === 'all') { S.from = ps[0]; S.to = ps[ps.length - 1]; }
    else { S.to = ps[ps.length - 1]; S.from = ps[Math.max(0, ps.length - +n)]; }
    CAL.pick = null; calRender(); calSyncBtn(); render();
  }
  function calSyncBtn() {
    $('calBtn').textContent = S.from && S.to
      ? C.periodLabel(S.from, 'weekly') + '  →  ' + C.periodLabel(S.to, 'weekly')
      : 'выбрать недели';
  }

  /* ============================================================
     ФИЛЬТРЫ
     ============================================================ */
  function dims() {
    const rs = S.level === 'weekly' ? D.weekly : D.monthly;
    const d = new Set(), s = new Set(), p = new Set();
    rs.forEach(r => { if (r.dir) d.add(r.dir); if (r.src) s.add(r.src); if (r.page) p.add(r.page); });
    const srt = (x) => [...x].sort((a, b) => a.localeCompare(b, 'ru'));
    return { dir: srt(d), src: srt(s), page: srt(p) };
  }
  function chipRow(host, key, items) {
    $(host).innerHTML = items.map(v =>
      `<button class="chip sm" data-v="${esc(v)}" aria-pressed="${S[key].includes(v)}">${esc(v)}</button>`).join('');
    els('#' + host + ' .chip').forEach(b => b.onclick = () => {
      const v = b.dataset.v, i = S[key].indexOf(v);
      i >= 0 ? S[key].splice(i, 1) : S[key].push(v);
      b.setAttribute('aria-pressed', String(i < 0));
      syncCount(); render();
    });
  }
  function syncCount() {
    $('cntDir').textContent = S.dirs.length ? S.dirs.length + ' выбрано' : 'все';
    $('cntSrc').textContent = S.srcs.length ? S.srcs.length + ' выбрано' : 'все';
    $('cntPage').textContent = S.pages.length ? S.pages.length + ' выбрано' : 'все';
  }
  function syncFilters() {
    const dm = dims();
    S.dirs = S.dirs.filter(v => dm.dir.includes(v));
    S.srcs = S.srcs.filter(v => dm.src.includes(v));
    S.pages = S.pages.filter(v => dm.page.includes(v));
    chipRow('fDir', 'dirs', dm.dir);
    chipRow('fSrc', 'srcs', dm.src);
    chipRow('fPage', 'pages', dm.page);
    syncCount();
  }
  function buildPeriods(keep) {
    const ps = C.periodList(D, S.level);
    const o = (p) => `<option value="${p}">${esc(C.periodLabel(p, S.level))}</option>`;
    $('selFrom').innerHTML = ps.map(o).join('');
    $('selTo').innerHTML = ps.map(o).join('');
    if (!keep || !ps.includes(S.from) || !ps.includes(S.to)) {
      S.from = ps[Math.max(0, ps.length - (S.level === 'weekly' ? 13 : 12))];
      S.to = ps[ps.length - 1];
    }
    if (S.from > S.to) S.from = S.to;
    $('selFrom').value = S.from; $('selTo').value = S.to;
    calSyncBtn();
  }
  function syncLevelUI() {
    const w = S.level === 'weekly';
    $('fldCal').classList.toggle('hide', !w);
    $('fldFrom').classList.toggle('hide', w);
    $('fldTo').classList.toggle('hide', w);
    if (!w && CAL.open) calOpen(false);
    els('#segLevel button').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.l === S.level)));
  }

  /* ============================================================
     РЕНДЕР
     ============================================================ */
  const VIEWS = { over: renderOver, dyn: renderDyn, share: renderShare, heat: renderHeat,
                  pf: renderPF, fc: renderFC, table: renderTable, gs: renderGS, method: renderMethod };
  function render() {
    els('.view').forEach(v => v.classList.toggle('on', v.id === 'v-' + S.view));
    els('.tab').forEach(t => t.setAttribute('aria-selected', String(t.dataset.v === S.view)));
    $('filters').classList.toggle('hide', S.view === 'method' || S.view === 'gs');
    /* Уничтожаем графики НЕактивных вкладок: их полотна ушли в
       display:none, живые экземпляры там только держат память и ловят
       события resize. Держим ровно графики текущего экрана. */
    const live = $('v-' + S.view);
    Object.keys(CH).forEach(id => {
      const cv = document.getElementById(id);
      if (!cv || !live || !live.contains(cv)) { CH[id].destroy(); delete CH[id]; }
    });
    const fn = VIEWS[S.view];
    if (fn) fn();
    /* Полотна уже в живой раскладке — только теперь у них есть размеры.
       Без этого график, созданный в скрытой карточке, остаётся пустым. */
    requestAnimationFrame(resizeAll);
    const cov = D.meta.coverage;
    $('topMeta').innerHTML = `<b>${cov.weeks}</b> недель · <b>${cov.months}</b> месяцев · последняя неделя <b>${esc(C.periodLabel(cov.weekTo, 'weekly'))}</b>`;
  }

  function bind() {
    els('.tab').forEach(t => t.onclick = () => { S.view = t.dataset.v; render(); window.scrollTo({ top: 0, behavior: 'smooth' }); });
    els('#segLevel button').forEach(b => b.onclick = () => {
      S.level = b.dataset.l; syncLevelUI(); buildPeriods(false); syncFilters(); render();
    });
    $('selFrom').onchange = () => { S.from = $('selFrom').value; if (S.from > S.to) { S.to = S.from; $('selTo').value = S.to; } render(); };
    $('selTo').onchange = () => { S.to = $('selTo').value; if (S.to < S.from) { S.from = S.to; $('selFrom').value = S.from; } render(); };
    $('btnReset').onclick = () => { S.dirs = []; S.srcs = []; S.pages = []; syncFilters(); buildPeriods(false); render(); };
    $('btnCsv').onclick = exportCsv;
    /* Любая смена состояния идёт через render(): он сам решает, какой
       экран активен. Прямой вызов renderHeat() с другой вкладки рисовал
       бы в скрытый DOM — источник «то работает, то нет». */
    els('#heatMetric .chip').forEach(b => b.onclick = () => { S.heatMetric = b.dataset.m; render(); });
    els('#heatDim .chip').forEach(b => b.onclick = () => { S.heatDim = b.dataset.d; render(); });
    els('#fcMethod .chip').forEach(b => b.onclick = () => { S.fcMethod = b.dataset.f; render(); });
    els('#shDim .chip').forEach(b => b.onclick = () => { S.shDim = b.dataset.d; render(); });
    els('#shMetric .chip').forEach(b => b.onclick = () => { S.shMetric = b.dataset.m; render(); });
    $('pfSave').onclick = () => {
      const ok = plansSave();
      $('pfSave').textContent = ok ? 'Планы сохранены' : 'Не удалось сохранить';
      setTimeout(() => $('pfSave').textContent = 'Сохранить планы', 1800);
    };
    $('pfCsv').onclick = () => pfExport();
    $('pfAdd').onclick = () => {
      const v = ($('pfNewDir').value || '').trim();
      if (!v) { $('pfNewDir').focus(); return; }
      pfAddDir(v); $('pfNewDir').value = '';
    };
    $('pfNewDir').onkeydown = (e) => { if (e.key === 'Enter') $('pfAdd').click(); };
    $('pfCopy').onchange = () => { const v = $('pfCopy').value; $('pfCopy').value = ''; pfCopyFrom(v); };
    $('pfClear').onclick = () => { pfClearManual(); };
    /* Календарь: три независимых способа закрыть — кнопка, фон, Esc. */
    $('calBtn').onclick = () => calOpen(!CAL.open);
    $('calClose').onclick = () => calOpen(false);
    $('calBack').onclick = () => calOpen(false);
    $('calPrev').onclick = () => { CAL.year--; calRender(); };
    $('calNext').onclick = () => { CAL.year++; calRender(); };
    els('#calQuick .chip').forEach(b => b.onclick = () => calQuick(b.dataset.q));
    const onEsc = (e) => { if (e.key === 'Escape' && CAL.open) calOpen(false); };
    addEventListener('keydown', onEsc);
    document.addEventListener('keydown', onEsc);
    let rt; addEventListener('resize', () => { clearTimeout(rt); rt = setTimeout(resizeAll, 200); });
  }

  function boot() {
    const el = document.getElementById('data');
    try { D = JSON.parse(el.textContent); }
    catch (e) { document.body.innerHTML = '<div class="empty"><b>Данные повреждены</b>' + esc(e.message) + '</div>'; return; }
    D.plan_fact = D.plan_fact || {};
    plansLoad();
    try { S.gsUrl = localStorage.getItem(GS_KEY) || ''; } catch (e) {}
    Chart.defaults.font.family = "'Manrope',sans-serif";
    Chart.defaults.color = COL.mid;
    bind();
    syncLevelUI(); buildPeriods(false); syncFilters();
    render();
    window.__D = D; window.__S = S;
    window.__API = { render, C, calOpen, calQuick, VIEWS, CH, PLANS: () => PLANS,
                     planFor, planSet, factFor, pfAddDir, pfCopyFrom, pfClearManual,
                     gsCheck, gsApply, GS, GS_CODE };
  }
  document.readyState === 'loading' ? addEventListener('DOMContentLoaded', boot) : boot();
})();
