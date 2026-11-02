// Серверная часть сайта. Вставляется в Google-таблицу: Расширения → Apps Script (подробно — README.md).
// Читать данные может любой, кто знает адрес; записывать — только с паролем из свойства скрипта PASSWORD.

// раздел сайта → id листа (число после gid= в адресе таблицы). По id, чтобы переименование листов ничего не ломало.
var SHEETS = { games: 0, plays: 297173692, preorders: 489614522, players: 515997837 };
var WRITABLE = ['plays', 'preorders', 'players'];  // список игр (Лист1) сайт только читает
var MAX_ROWS = 5000;

function doGet() {
  try {
    var ss = SpreadsheetApp.getActive(), data = {};
    Object.keys(SHEETS).forEach(function (key) { data[key] = readSheet(sheetById(ss, SHEETS[key])); });
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

function validRows(rows) {
  return Array.isArray(rows) && rows.length <= MAX_ROWS && rows.every(function (r) {
    return r && typeof r === 'object' && !Array.isArray(r) && Object.keys(r).every(function (k) {
      var v = r[k];
      return typeof v === 'number' || typeof v === 'boolean' || (typeof v === 'string' && v.length <= 1000);
    });
  });
}

// Обнулить статистику: в редакторе Apps Script выберите resetStats в списке функций вверху и нажмите «Выполнить».
// «Партий» на Лист2 и «Партий»/«Побед» на Лист5 станут 0; названия игр и имена игроков останутся. Откат — «Файл → История версий».
function resetStats() {
  var ss = SpreadsheetApp.getActive();
  [[SHEETS.plays, ['Партий']], [SHEETS.players, ['Партий', 'Побед']]].forEach(function (job) {
    var sheet = sheetById(ss, job[0]), rows = readSheet(sheet);
    rows.forEach(function (r) { job[1].forEach(function (col) { if (col in r) r[col] = 0; }); });
    writeSheet(sheet, rows);
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
