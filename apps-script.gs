// Серверная часть сайта. Вставляется в Google-таблицу: Расширения → Apps Script (подробно — README.md).
// Читать данные может любой, кто знает адрес; записывать — только с паролем из свойства скрипта PASSWORD.

// раздел сайта → id листа (число после gid= в адресе таблицы). По id, чтобы переименование листов ничего не ломало.
var SHEETS = { games: 0, preorders: 489614522, players: 515997837 };
var WRITABLE = ['preorders', 'players'];  // эти листы сайт сохраняет целиком; на Лист1 меняет только ячейки «Партий»
var GAME_NAME = 'Название', PLAYS_COL = 'Партий';  // столбцы Лист1
var OLD_PLAYS_SHEET = 297173692;  // бывший Лист2 — нужен только для migratePlays
var MAX_ROWS = 5000;

function doGet() {
  try {
    var ss = SpreadsheetApp.getActive(), data = {};
    Object.keys(SHEETS).forEach(function (key) { data[key] = readSheet(sheetById(ss, SHEETS[key])); });
    // сайт ждёт партии отдельным списком, как раньше на Лист2
    data.plays = data.games.map(function (g) {
      var p = {};
      p[GAME_NAME] = g[GAME_NAME];
      p[PLAYS_COL] = g[PLAYS_COL] || '0';
      return p;
    });
    return json(data);
  } catch (err) {
    return json({ error: String(err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents);
    var expected = PropertiesService.getScriptProperties().getProperty('PASSWORD');
    if (!expected || req.password !== expected) return json({ error: 'password' });
    if (req.op === 'addPlay') {
      if (typeof req.name !== 'string' || [1, -1].indexOf(req.delta) < 0) return json({ error: 'неверный запрос' });
      lock.waitLock(20000);
      return json({ ok: true, value: addPlay(sheetById(SpreadsheetApp.getActive(), SHEETS.games), req.name, req.delta) });
    }
    if (WRITABLE.indexOf(req.key) < 0 || !validRows(req.rows)) return json({ error: 'неверный запрос' });
    lock.waitLock(20000);  // два сохранения одновременно не перемешаются
    writeSheet(sheetById(SpreadsheetApp.getActive(), SHEETS[req.key]), req.rows);
    return json({ ok: true });
  } catch (err) {
    return json({ error: String(err.message || err) });
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function sheetById(ss, id) {
  var sheet = ss.getSheets().filter(function (s) { return s.getSheetId() === id; })[0];
  if (!sheet) throw new Error('В таблице нет листа с id ' + id);
  return sheet;
}

// Лист → массив строк { "Заголовок столбца": "значение как в ячейке" }
function readSheet(sheet) {
  var values = sheet.getDataRange().getDisplayValues(), header = values[0];
  return values.slice(1).map(function (row) {
    var obj = {};
    header.forEach(function (h, i) { if (h) obj[h] = row[i]; });
    return obj;
  }).filter(function (obj) {
    return Object.keys(obj).some(function (k) { return String(obj[k]).trim(); });
  });
}

// Заменяет все строки листа под заголовком. Заголовок, оформление и столбцы правее заголовков не трогает.
function writeSheet(sheet, rows) {
  var header = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getDisplayValues()[0];
  var width = header.length;
  while (width && !header[width - 1]) width--;
  if (!width) throw new Error('На листе «' + sheet.getName() + '» нет строки заголовков');
  header = header.slice(0, width);

  var old = sheet.getLastRow() - 1;
  if (old > 0) {
    var range = sheet.getRange(2, 1, old, width);
    var hasFormulas = range.getFormulas().some(function (r) { return r.some(Boolean); });
    if (hasFormulas) throw new Error('На листе «' + sheet.getName() + '» есть формулы — сайт не будет его перезаписывать');
    range.clearContent();
  }
  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, width).setValues(rows.map(function (r) {
    return header.map(function (h) {
      var v = r[h] === undefined ? '' : r[h];
      return typeof v === 'string' && v.charAt(0) === '=' ? "'" + v : v;  // текст «=...» не должен стать формулой
    });
  }));
}

// Лист1: все значения + номера столбцов «Название» и «Партий» (−1, если столбца нет)
function gameColumns(sheet) {
  var values = sheet.getDataRange().getDisplayValues();
  return { values: values, name: values[0].indexOf(GAME_NAME), plays: values[0].indexOf(PLAYS_COL) };
}

// +1 / −1 партия: меняется одна ячейка, остальной Лист1 (ручные правки, ссылки, описания) не трогается.
// Прибавляет сам скрипт, поэтому партии, записанные с двух устройств одновременно, не теряются.
function addPlay(sheet, name, delta) {
  var c = gameColumns(sheet), key = name.trim().toLowerCase();
  if (c.name < 0 || c.plays < 0) throw new Error('На листе «' + sheet.getName() + '» нет столбца «' + (c.name < 0 ? GAME_NAME : PLAYS_COL) + '»');
  for (var r = 1; r < c.values.length; r++) {
    if (c.values[r][c.name].trim().toLowerCase() !== key) continue;
    var cell = sheet.getRange(r + 1, c.plays + 1);
    if (cell.getFormula()) throw new Error('В ячейке партий у «' + name.trim() + '» формула — сайт её не перезаписывает');
    var value = Math.max(0, (Number(String(c.values[r][c.plays]).replace(',', '.')) || 0) + delta);
    cell.setValue(value);
    return value;
  }
  throw new Error('Игры «' + name.trim() + '» нет на листе «' + sheet.getName() + '»');
}

function validRows(rows) {
  return Array.isArray(rows) && rows.length <= MAX_ROWS && rows.every(function (r) {
    return r && typeof r === 'object' && !Array.isArray(r) && Object.keys(r).every(function (k) {
      var v = r[k];
      return typeof v === 'number' || typeof v === 'boolean' || (typeof v === 'string' && v.length <= 1000);
    });
  });
}

// Обнулить статистику: в редакторе Apps Script выберите resetStats в списке функций вверху и нажмите «Выполнить».
// «Партий» на Лист1 и «Партий»/«Побед» на Лист5 станут 0; остальные столбцы и имена останутся. Откат — «Файл → История версий».
function resetStats() {
  var ss = SpreadsheetApp.getActive(), games = sheetById(ss, SHEETS.games), c = gameColumns(games);
  if (c.plays >= 0 && c.values.length > 1) {
    var range = games.getRange(2, c.plays + 1, c.values.length - 1, 1);
    if (range.getFormulas().some(function (r) { return r[0]; })) throw new Error('В столбце «' + PLAYS_COL + '» есть формулы — обнулять не буду');
    range.setValues(c.values.slice(1).map(function (r) { return [c.name >= 0 && r[c.name].trim() ? 0 : r[c.plays]]; }));
  }
  var players = sheetById(ss, SHEETS.players), rows = readSheet(players);
  rows.forEach(function (r) { ['Партий', 'Побед'].forEach(function (col) { if (col in r) r[col] = 0; }); });
  writeSheet(players, rows);
}

// Один раз, ДО удаления Лист2: переносит партии с Лист2 в столбец «Партий» на Лист1 (создаёт его в конце, если нет).
// Заполняет только пустые ячейки, так что повторный запуск ничего не испортит. Итог — в «Журнале выполнения».
function migratePlays() {
  var ss = SpreadsheetApp.getActive(), games = sheetById(ss, SHEETS.games), counts = {};
  readSheet(sheetById(ss, OLD_PLAYS_SHEET)).forEach(function (r) { counts[String(r[GAME_NAME]).trim().toLowerCase()] = r[PLAYS_COL]; });
  var c = gameColumns(games);
  if (c.name < 0) throw new Error('На Лист1 нет столбца «' + GAME_NAME + '»');
  if (c.plays < 0) {
    c.plays = games.getLastColumn();
    games.getRange(1, c.plays + 1).setValue(PLAYS_COL);
  } else if (c.values.length > 1 && games.getRange(2, c.plays + 1, c.values.length - 1, 1).getFormulas().some(function (r) { return r[0]; })) {
    throw new Error('В столбце «' + PLAYS_COL + '» есть формулы — переносить не буду');
  }
  var moved = {};
  var column = c.values.slice(1).map(function (r) {
    var key = r[c.name].trim().toLowerCase(), cur = r[c.plays] === undefined ? '' : r[c.plays];
    if (!key || String(cur).trim()) return [cur];
    moved[key] = true;
    return [Number(counts[key]) || 0];
  });
  if (column.length) games.getRange(2, c.plays + 1, column.length, 1).setValues(column);
  var lost = Object.keys(counts).filter(function (k) { return !moved[k] && Number(counts[k]) > 0; });
  Logger.log(lost.length
    ? 'Не перенесены (такой игры нет на Лист1 или ячейка уже заполнена): ' + lost.join(', ')
    : 'Готово: партии перенесены на Лист1. Лист2 можно удалять.');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
