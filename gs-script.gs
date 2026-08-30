function doGet(e) {
  // Лист, который читаем. Название должно совпадать ТОЧНО.
  var SHEET_NAME = 'Недельный отчёт по трафику';

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) {
    return out({ error: 'Лист не найден: ' + SHEET_NAME });
  }
  var v = sh.getDataRange().getValues();
  var rows = [];

  // Первая строка — заголовки, данные идут с второй.
  for (var i = 1; i < v.length; i++) {
    var r = v[i];
    if (!r[0] || !r[2]) continue;           // нет даты или направления
    rows.push({
      period:        iso(r[0]),             // A · начало периода
      end:           iso(r[1]),             // B · конец периода
      dir:           str(r[2]),             // C · направление
      src:           str(r[3]),             // D · источник
      page:          str(r[4]),             // E · посадочная страница
      spend:         num(r[5]),             // F · расход
      leads:         num(r[6]),             // G · лиды
      kl:            num(r[8]),             // I · квал-лиды
      inwork:        num(r[9]),             // J · в работе
      badlead:       num(r[10]),            // K · не кач. лид
      spam:          num(r[11]),            // L · спам
      existing:      num(r[13]),            // N · существующие
      deals_created: num(r[17]),            // R · создано сделок
      pipeline:      num(r[18]),            // S · сумма созданных
      deals_lost:    num(r[19]),            // T · проиграно сделок
      revenue_lost:  num(r[20]),            // U · сумма проигранных
      deals_won:     num(r[21]),            // V · успешных сделок
      revenue_won:   num(r[22])             // W · сумма успешных
    });
  }
  return out({ weekly: rows, count: rows.length, sheet: SHEET_NAME });
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Дата → YYYY-MM-DD. Дэшборд сортирует периоды как строки,
// поэтому формат обязателен именно такой.
function iso(d) {
  if (d instanceof Date) {
    return Utilities.formatDate(d, 'Asia/Vladivostok', 'yyyy-MM-dd');
  }
  return d ? String(d).slice(0, 10) : null;
}

// Числа приходят и текстом с пробелами-разделителями («11 328 914»),
// поэтому чистим пробелы и неразрывные пробелы, иначе получим 0.
function num(x) {
  if (x === '' || x === null || x === undefined) return 0;
  if (typeof x === 'number') return x;
  var s = String(x).replace(/\u00A0/g, '').replace(/\s/g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function str(x) {
  if (x === null || x === undefined) return null;
  var s = String(x).replace(/\u00A0/g, ' ').trim();
  return (s === '' || s === '-' || s === '—') ? null : s;
}
