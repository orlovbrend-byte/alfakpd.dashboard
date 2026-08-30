/* ============================================================
   Альфа-КПД · расчётное ядро
   Чистые функции: на вход данные + срез, на выход числа.
   Никакого DOM — чтобы это можно было проверить тестами.
   ============================================================ */
(function (root) {
  'use strict';

  const MONRU = ['Январь','Февраль','Март','Апрель','Май','Июнь',
                 'Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
  const MONSHORT = ['янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек'];

  const rowsOf = (D, level) => (level === 'weekly' ? D.weekly : D.monthly) || [];

  /** Все периоды уровня, отсортированные. Период — ISO-строка,
      поэтому лексикографическая сортировка = хронологическая. */
  function periodList(D, level) {
    const s = new Set();
    rowsOf(D, level).forEach(r => s.add(r.period));
    return [...s].sort();
  }

  function periodLabel(p, level) {
    if (!p) return '—';
    if (level === 'weekly') {
      const d = new Date(p + 'T00:00:00Z');
      const e = new Date(d.getTime() + 6 * 864e5);
      const f = (x) => String(x.getUTCDate()).padStart(2, '0') + '.' + String(x.getUTCMonth() + 1).padStart(2, '0');
      return f(d) + '–' + f(e);
    }
    const [y, m] = p.split('-');
    return MONRU[+m - 1] + ' ' + y;
  }
  function periodShort(p, level) {
    if (level === 'weekly') {
      const d = new Date(p + 'T00:00:00Z');
      return String(d.getUTCDate()).padStart(2, '0') + '.' + String(d.getUTCMonth() + 1).padStart(2, '0');
    }
    const [y, m] = p.split('-');
    return MONSHORT[+m - 1] + ' ' + y.slice(2);
  }

  /** Фильтр по срезу. Пустой список измерения = «все», а не «ничего». */
  function filterRows(D, S) {
    const rs = rowsOf(D, S.level);
    const dirs = S.dirs && S.dirs.length ? new Set(S.dirs) : null;
    const srcs = S.srcs && S.srcs.length ? new Set(S.srcs) : null;
    const pages = S.pages && S.pages.length ? new Set(S.pages) : null;
    const out = [];
    for (const r of rs) {
      if (S.from && r.period < S.from) continue;
      if (S.to && r.period > S.to) continue;
      if (dirs && !dirs.has(r.dir)) continue;
      if (srcs && !srcs.has(r.src)) continue;
      if (pages && !pages.has(r.page)) continue;
      out.push(r);
    }
    return out;
  }

  const F = ['spend','leads','kl','spam','badlead','existing',
             'deals_created','pipeline','deals_lost','revenue_lost','deals_won','revenue_won'];

  /** Сумма + производные. Производные считаются ОДИН раз здесь,
      чтобы одна метрика не считалась двумя способами на двух экранах. */
  function agg(rows) {
    const a = {};
    for (const k of F) a[k] = 0;
    a.rows = rows.length;
    for (const r of rows) for (const k of F) a[k] += r[k] || 0;
    return derive(a);
  }

  function derive(a) {
    const leadsClean = a.leads - a.spam;
    a.leadsClean = leadsClean;
    a.cpl = a.leads > 0 ? a.spend / a.leads : null;
    a.cplClean = leadsClean > 0 ? a.spend / leadsClean : null;
    a.cpql = a.kl > 0 ? a.spend / a.kl : null;
    a.cr = a.leads > 0 ? a.kl / a.leads : null;
    a.crClean = leadsClean > 0 ? a.kl / leadsClean : null;
    a.spamShare = a.leads > 0 ? a.spam / a.leads : null;
    a.crWon = a.kl > 0 ? a.deals_won / a.kl : null;
    a.cpd = a.deals_won > 0 ? a.spend / a.deals_won : null;
    a.romi = a.spend > 0 ? (a.revenue_won - a.spend) / a.spend : null;
    a.drr = a.revenue_won > 0 ? a.spend / a.revenue_won : null;
    a.pipelineReturn = a.spend > 0 ? a.pipeline / a.spend : null;
    return a;
  }

  /** Ряд по периодам внутри среза. Ось — сам период (ISO),
      подпись отдельно: хронология не зависит от подписи. */
  function series(D, S) {
    const by = new Map();
    for (const r of filterRows(D, S)) {
      let b = by.get(r.period);
      if (!b) { b = {}; for (const k of F) b[k] = 0; by.set(r.period, b); }
      for (const k of F) b[k] += r[k] || 0;
    }
    return [...by.keys()].sort().map(p => {
      const o = derive(Object.assign({ period: p }, by.get(p)));
      o.label = periodShort(p, S.level);
      o.labelFull = periodLabel(p, S.level);
      return o;
    });
  }

  /** Предыдущее окно той же длины — для корректной дельты. */
  function previousWindow(D, S) {
    const ps = periodList(D, S.level);
    const a = ps.indexOf(S.from), b = ps.indexOf(S.to);
    if (a < 0 || b < 0) return null;
    const len = b - a + 1;
    if (a - len < 0) return null;
    const from = ps[a - len], to = ps[a - 1];
    const w = agg(filterRows(D, Object.assign({}, S, { from, to })));
    w.from = from; w.to = to;
    return w;
  }

  /** Группировка по измерению. */
  function groupBy(D, S, dim) {
    const by = new Map();
    for (const r of filterRows(D, S)) {
      const k = r[dim] || '—';
      if (!by.has(k)) by.set(k, []);
      by.get(k).push(r);
    }
    return [...by.entries()].map(([k, rs]) => Object.assign({ key: k }, agg(rs)));
  }

  /** Матрица тепловой карты: измерение × период по одной метрике.
      Шкала считается по фактическому разбросу, иначе всё одного цвета. */
  function heatmap(D, S, dim, metric) {
    const ps = [];
    const seen = new Set();
    const by = new Map();
    for (const r of filterRows(D, S)) {
      if (!seen.has(r.period)) { seen.add(r.period); ps.push(r.period); }
      const k = r[dim] || '—';
      if (!by.has(k)) by.set(k, new Map());
      const m = by.get(k);
      let c = m.get(r.period);
      if (!c) { c = {}; for (const f of F) c[f] = 0; m.set(r.period, c); }
      for (const f of F) c[f] += r[f] || 0;
    }
    ps.sort();
    const rows = [];
    let lo = Infinity, hi = -Infinity;
    for (const [key, m] of by) {
      const cells = ps.map(p => {
        const c = m.get(p);
        if (!c) return null;
        const v = derive(Object.assign({}, c))[metric];
        if (v == null || !isFinite(v)) return null;
        if (v < lo) lo = v;
        if (v > hi) hi = v;
        return { period: p, value: v, spend: c.spend, leads: c.leads, kl: c.kl };
      });
      const tot = agg([...m.values()].map(c => Object.assign({}, c)));
      rows.push({ key, cells, total: tot[metric], spend: tot.spend, kl: tot.kl });
    }
    // Тяжёлые каналы сверху: смотреть на дорогой канал с 1 лидом смысла нет.
    rows.sort((x, y) => y.spend - x.spend);
    return { periods: ps, rows, lo: isFinite(lo) ? lo : null, hi: isFinite(hi) ? hi : null,
             labels: ps.map(p => periodShort(p, S.level)) };
  }

  /** Доля измерения в срезе: по лидам и по квал-лидам одновременно.
      Знаменатель — сумма ПО СРЕЗУ, поэтому при смене фильтра проценты
      пересчитываются от нового целого, а не от всей базы. */
  function shares(D, S, dim) {
    const g = groupBy(D, S, dim);
    const tot = g.reduce((a, r) => {
      a.leads += r.leads; a.kl += r.kl; a.spend += r.spend; a.deals_won += r.deals_won;
      return a; }, { leads: 0, kl: 0, spend: 0, deals_won: 0 });
    const rows = g.map(r => Object.assign({}, r, {
      shareLeads: tot.leads > 0 ? r.leads / tot.leads : null,
      shareKl: tot.kl > 0 ? r.kl / tot.kl : null,
      shareSpend: tot.spend > 0 ? r.spend / tot.spend : null,
      /* Перекос: доля квала минус доля расхода. Плюс = источник даёт
         квала больше, чем съедает бюджета. Это и есть ответ на вопрос
         «куда переложить деньги». */
      lift: (tot.kl > 0 && tot.spend > 0) ? (r.kl / tot.kl) - (r.spend / tot.spend) : null
    }));
    rows.sort((a, b) => b.kl - a.kl || b.leads - a.leads);
    // Накопленная доля квала — для правила Парето.
    let acc = 0;
    rows.forEach(r => { acc += r.shareKl || 0; r.cumKl = acc; });
    return { rows, total: derive(Object.assign({}, tot)), raw: tot };
  }

  /** Прогноз на следующий период. Три честных метода, без магии. */
  function forecast(vals, method) {
    const v = vals.filter(x => x != null && isFinite(x));
    if (v.length < 2) return null;
    if (method === 'last3') {
      const t = v.slice(-3);
      return t.reduce((a, b) => a + b, 0) / t.length;
    }
    if (method === 'trend') {
      const n = v.length;
      let sx = 0, sy = 0, sxy = 0, sxx = 0;
      for (let i = 0; i < n; i++) { sx += i; sy += v[i]; sxy += i * v[i]; sxx += i * i; }
      const d = n * sxx - sx * sx;
      if (!d) return v[n - 1];
      const b = (n * sxy - sx * sy) / d, a = (sy - b * sx) / n;
      return a + b * n;
    }
    // momentum: средний темп прироста
    let g = 0, c = 0;
    for (let i = 1; i < v.length; i++) if (v[i - 1]) { g += (v[i] - v[i - 1]) / Math.abs(v[i - 1]); c++; }
    return c ? v[v.length - 1] * (1 + g / c) : v[v.length - 1];
  }

  /** Ряд доли одного измерения по периодам — для графика «как менялась
      доля источника». Возвращает по одному ряду на значение измерения. */
  function shareSeries(D, S, dim, metric, topN) {
    const per = new Map();          // period -> key -> value
    const totByPer = new Map();     // period -> total
    for (const r of filterRows(D, S)) {
      const k = r[dim] || '—';
      if (!per.has(r.period)) per.set(r.period, new Map());
      const m = per.get(r.period);
      m.set(k, (m.get(k) || 0) + (r[metric] || 0));
      totByPer.set(r.period, (totByPer.get(r.period) || 0) + (r[metric] || 0));
    }
    const ps = [...per.keys()].sort();
    const sh = shares(D, S, dim).rows.slice(0, topN || 6).map(r => r.key);
    return {
      periods: ps, labels: ps.map(p => periodShort(p, S.level)), keys: sh,
      rows: sh.map(k => ({ key: k, data: ps.map(p => {
        const t = totByPer.get(p) || 0;
        const v = (per.get(p) || new Map()).get(k) || 0;
        return t > 0 ? +(v / t * 100).toFixed(1) : null;
      }) }))
    };
  }

  root.CORE = { MONRU, periodList, periodLabel, periodShort, filterRows,
                agg, derive, series, previousWindow, groupBy, heatmap,
                shares, shareSeries, forecast, FIELDS: F };
})(typeof window !== 'undefined' ? window : globalThis);
