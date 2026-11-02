// Серверная часть сайта. Вставляется в Google-таблицу: Расширения → Apps Script (подробно — README.md).
// Читать данные может любой, кто знает адрес; записывать — только с паролем из свойства скрипта PASSWORD.

// раздел сайта → id листа (число после gid= в адресе таблицы). По id, чтобы переименование листов ничего не ломало.
var SHEETS = { games: 0, preorders: 489614522, players: 515997837 };
var WRITABLE = ['preorders', 'players'];  // эти листы сайт сохраняет целиком
// Журнал и оценки — отдельные листы; скрипт создаёт их сам при первой записи. Не переименовывайте их.
var LOG_SHEET = 'Журнал', LOG_COLS = ['ID', 'Дата', 'Игра', 'Игроки', 'Победители', 'Заметка'];
var RATINGS_SHEET = 'Оценки', RATING_COLS = ['Игра', 'Игрок', 'Оценка'];
// Столбцы-счётчики: скрипт пересчитывает их из журнала, руками их менять бесполезно
var GAME_NAME = 'Название', GAME_PLAYS = 'Партий';               // Лист1
var PLAYER_NAME = 'Игрок', PLAYER_PLAYS = 'Партий', PLAYER_WINS = 'Побед';  // Лист5
var MAX_ROWS = 5000;

function doGet() {
  try {
    var ss = SpreadsheetApp.getActive(), data = {};
    Object.keys(SHEETS).forEach(function (key) { data[key] = readSheet(sheetById(ss, SHEETS[key])); });
    data.log = readNamed(ss, LOG_SHEET);
    data.ratings = readNamed(ss, RATINGS_SHEET);
    return json(data);
  } catch (err) {
    return json({ error: String(err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActive();
    var expected = PropertiesService.getScriptProperties().getProperty('PASSWORD');
    if (!expected || req.password !== expected) return json({ error: 'password' });
    var bad = json({ error: 'неверный запрос' });
    if (req.op === 'log') {
      if (!text(req.game, 200) || !text(req.date, 20) || !names(req.players) || !req.players.length || !names(req.winners) || !(req.note === undefined || text(req.note, 500, true))) return bad;
      lock.waitLock(20000);
      var id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      getOrCreate(ss, LOG_SHEET, LOG_COLS).appendRow([id, req.date, req.game, req.players.join(', '), req.winners.join(', '), req.note || ''].map(safe));
      syncCounters(ss);
      return json({ ok: true, id: id });
    }
    if (req.op === 'unlog') {
      if (!text(req.id, 40)) return bad;
      lock.waitLock(20000);
      var log = ss.getSheetByName(LOG_SHEET), row = log && findRow(log, ['ID'], [req.id]);
      if (!row) return json({ error: 'Такой записи в журнале нет — обновите страницу' });
      log.deleteRow(row);
      syncCounters(ss);
      return json({ ok: true });
    }
    if (req.op === 'rate') {
      if (!text(req.game, 200) || !text(req.player, 100) || [0, 1, 2, 3, 4, 5].indexOf(req.value) < 0) return bad;
      lock.waitLock(20000);
      var sheet = getOrCreate(ss, RATINGS_SHEET, RATING_COLS), r = findRow(sheet, ['Игра', 'Игрок'], [req.game, req.player]);
      if (r && req.value) sheet.getRange(r, RATING_COLS.indexOf('Оценка') + 1).setValue(req.value);
      else if (r) sheet.deleteRow(r);
      else if (req.value) sheet.appendRow([safe(req.game), safe(req.player), req.value]);
      return json({ ok: true });
    }
    if (WRITABLE.indexOf(req.key) < 0 || !validRows(req.rows)) return bad;
    lock.waitLock(20000);  // два сохранения одновременно не перемешаются
    writeSheet(sheetById(ss, SHEETS[req.key]), req.rows);
    if (req.key === 'players') syncCounters(ss);  // в присланной таблице могли быть устаревшие счётчики
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

function readNamed(ss, name) {
  var sheet = ss.getSheetByName(name);
  return sheet ? readSheet(sheet) : [];
}

function getOrCreate(ss, name, cols) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (!sheet.getLastRow()) sheet.appendRow(cols);
  return sheet;
}

// Лист → массив строк { "Заголовок столбца": "значение как в ячейке" }
function readSheet(sheet) {
  if (!sheet.getLastRow()) return [];
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
    return header.map(function (h) { return safe(r[h] === undefined ? '' : r[h]); });
  }));
}

// Номер строки листа (с 1), где в столбцах cols стоят значения vals (без учёта регистра и пробелов по краям), или 0
function findRow(sheet, cols, vals) {
  if (sheet.getLastRow() < 2) return 0;
  var values = sheet.getDataRange().getDisplayValues(), idx = cols.map(function (c) { return values[0].indexOf(c); });
  if (idx.indexOf(-1) >= 0) throw new Error('На листе «' + sheet.getName() + '» нет столбца «' + cols[idx.indexOf(-1)] + '»');
  for (var r = 1; r < values.length; r++) {
    if (idx.every(function (i, k) { return norm(values[r][i]) === norm(vals[k]); })) return r + 1;
  }
  return 0;
}

// Пересчитать счётчики из журнала: «Партий» на Лист1, «Партий»/«Побед» на Лист5. Новые игроки из журнала дописываются на Лист5.
function syncCounters(ss) {
  var games = {}, plays = {}, wins = {}, seen = {};
  readNamed(ss, LOG_SHEET).forEach(function (e) {
    inc(games, e['Игра']);
    splitNames(e['Игроки']).forEach(function (n) { inc(plays, n); seen[norm(n)] = n; });
    splitNames(e['Победители']).forEach(function (n) { inc(wins, n); seen[norm(n)] = n; });
  });
  setCounts(sheetById(ss, SHEETS.games), GAME_NAME, [[GAME_PLAYS, games]]);
  var players = sheetById(ss, SHEETS.players);
  var known = {};
  readSheet(players).forEach(function (p) { known[norm(p[PLAYER_NAME])] = true; });
  var header = players.getRange(1, 1, 1, Math.max(1, players.getLastColumn())).getDisplayValues()[0];
  Object.keys(seen).forEach(function (k) {
    if (!known[k]) players.appendRow(header.map(function (h) { return h === PLAYER_NAME ? safe(seen[k]) : ''; }));
  });
  setCounts(players, PLAYER_NAME, [[PLAYER_PLAYS, plays], [PLAYER_WINS, wins]]);
}

// В столбцах specs ([заголовок, {имя: число}]) у строк с именем ставит число (0, если имени нет в словаре). Меняются только эти столбцы.
function setCounts(sheet, nameCol, specs) {
  if (sheet.getLastRow() < 2) return;
  var values = sheet.getDataRange().getDisplayValues(), nameIdx = values[0].indexOf(nameCol);
  if (nameIdx < 0) return;
  specs.forEach(function (spec) {
    var col = values[0].indexOf(spec[0]);
    if (col < 0) return;  // столбца нет — считать некуда, сайт всё равно посчитает сам
    var range = sheet.getRange(2, col + 1, values.length - 1, 1);
    if (range.getFormulas().some(function (r) { return r[0]; })) throw new Error('В столбце «' + spec[0] + '» на листе «' + sheet.getName() + '» есть формулы — счётчики не обновлены');
    range.setValues(values.slice(1).map(function (r) {
      var key = norm(r[nameIdx]);
      return [key ? spec[1][key] || 0 : r[col]];
    }));
  });
}

// Обнулить статистику = очистить журнал: в редакторе выберите resetStats и нажмите «Выполнить». Оценки остаются.
// Откат — «Файл → История версий».
function resetStats() {
  var ss = SpreadsheetApp.getActive(), log = ss.getSheetByName(LOG_SHEET);
  if (log && log.getLastRow() > 1) log.deleteRows(2, log.getLastRow() - 1);
  syncCounters(ss);
}

function norm(s) { return String(s === undefined || s === null ? '' : s).trim().toLowerCase(); }
function inc(map, name) { var k = norm(name); if (k) map[k] = (map[k] || 0) + 1; }
function splitNames(s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function safe(v) { return typeof v === 'string' && v.charAt(0) === '=' ? "'" + v : v; }  // текст «=...» не должен стать формулой
function text(v, max, allowEmpty) { return typeof v === 'string' && v.length <= max && (allowEmpty || v.trim() !== ''); }
function names(list) {
  return Array.isArray(list) && list.length <= 30 && list.every(function (n) { return text(n, 100) && n.indexOf(',') < 0; });
}

function validRows(rows) {
  return Array.isArray(rows) && rows.length <= MAX_ROWS && rows.every(function (r) {
    return r && typeof r === 'object' && !Array.isArray(r) && Object.keys(r).every(function (k) {
      var v = r[k];
      return typeof v === 'number' || typeof v === 'boolean' || (typeof v === 'string' && v.length <= 1000);
    });
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
