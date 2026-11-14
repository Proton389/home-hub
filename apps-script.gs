// Серверная часть сайта. Вставляется в Google-таблицу: Расширения → Apps Script (подробно — README.md).
// Читать данные может любой, кто знает адрес; записывать — только с паролем из свойства скрипта PASSWORD.

// раздел сайта → id листа (число после gid= в адресе таблицы). По id, чтобы переименование листов ничего не ломало.
var SHEETS = { games: 0, preorders: 489614522, players: 515997837 };
var WRITABLE = ['preorders', 'players'];  // эти листы сайт сохраняет целиком
// Журнал и оценки — отдельные листы; скрипт создаёт их сам при первой записи. Не переименовывайте их.
// «Очки» — «Мэри: 120, Инна: 95», только у тех, кому их вписали (в играх на победу/поражение столбец пустой)
var LOG_SHEET = 'Журнал', LOG_SCORES = 'Очки', LOG_COLS = ['ID', 'Дата', 'Игра', 'Игроки', 'Победители', 'Заметка', LOG_SCORES];
var RATINGS_SHEET = 'Оценки', RATING_COLS = ['Игра', 'Игрок', 'Оценка'];
// Столбцы-счётчики: скрипт пересчитывает их из журнала, руками их менять бесполезно
var GAME_NAME = 'Название', GAME_PLAYS = 'Партий';               // Лист1
var PLAYER_NAME = 'Игрок', PLAYER_PLAYS = 'Партий', PLAYER_WINS = 'Побед';  // Лист5
var MAX_ROWS = 5000;

function doGet() {
  try {
    var ss = SpreadsheetApp.getActive(), data = {};
    Object.keys(SHEETS).forEach(function (key) { data[key] = readSheet(sheetById(ss, SHEETS[key])); });
    var games = sheetById(ss, SHEETS.games);
    data.gameColumns = games.getRange(1, 1, 1, Math.max(1, games.getLastColumn())).getDisplayValues()[0].filter(Boolean);  // поля формы «Новая игра»
    data.log = readNamed(ss, LOG_SHEET);
    data.ratings = readNamed(ss, RATINGS_SHEET);
    return json(data);
  } catch (err) {
    return json({ error: String(err.message || err) });
  }
}

function doPost(e) {
  if (e.parameter && e.parameter.tg !== undefined) return telegramWebhook(e);  // сообщения Telegram-бота
  var lock = LockService.getScriptLock();
  try {
    var req = JSON.parse(e.postData.contents), ss = SpreadsheetApp.getActive();
    var expected = PropertiesService.getScriptProperties().getProperty('PASSWORD');
    if (!expected || req.password !== expected) return json({ error: 'password' });
    var bad = json({ error: 'неверный запрос' });
    if (req.op === 'check') return json({ ok: true });  // сайт проверяет пароль до того, как открыть форму
    if (req.op === 'saveGame') {
      var f = req.fields;
      if (!f || typeof f !== 'object' || Array.isArray(f) || !text(f[GAME_NAME], 200) || !(req.original == null || text(req.original, 200)) ||
          !Object.keys(f).every(function (k) { return typeof f[k] === 'string' && f[k].length <= 5000; })) return bad;
      lock.waitLock(20000);
      saveGame(ss, req.original, f);
      return json({ ok: true });
    }
    if (req.op === 'renamePlayer') {
      if (!text(req.from, 100) || !names([req.to]) || !req.to.trim()) return bad;
      lock.waitLock(20000);
      renamePlayer(ss, req.from.trim(), req.to.trim());
      return json({ ok: true });
    }
    if (req.op === 'deleteGame') {
      if (!text(req.name, 200)) return bad;
      lock.waitLock(20000);
      deleteGame(ss, req.name);
      return json({ ok: true });
    }
    if (req.op === 'log') {
      if (!validPlay(req)) return bad;
      lock.waitLock(20000);
      var id = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      var sh = getOrCreate(ss, LOG_SHEET, LOG_COLS), row = playFields(req);
      row.ID = id;
      sh.appendRow(sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0].map(function (h) { return h in row ? safe(row[h]) : ''; }));
      syncCounters(ss);
      return json({ ok: true, id: id, scores: true });  // scores — сайт так видит, что скрипт новый и очки записаны
    }
    if (req.op === 'editLog') {  // исправить запись журнала на месте (ID и порядок записей сохраняются)
      if (!text(req.id, 40) || !validPlay(req)) return bad;
      lock.waitLock(20000);
      var lg = getOrCreate(ss, LOG_SHEET, LOG_COLS), at = findRow(lg, ['ID'], [req.id]);
      if (!at) return json({ error: 'Такой записи в журнале нет — обновите страницу' });
      var hdr = lg.getRange(1, 1, 1, lg.getLastColumn()).getDisplayValues()[0], vals = playFields(req);
      Object.keys(vals).forEach(function (h) { var c = hdr.indexOf(h); if (c >= 0) lg.getRange(at, c + 1).setValue(safe(vals[h])); });
      syncCounters(ss);
      return json({ ok: true, scores: true });
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
      if (req.value && !playedGame(ss, req.game, req.player)) return json({ error: 'Оценку ставит только тот, кто играл в эту игру' });
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
  var hdr = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getDisplayValues()[0];
  var missing = cols.filter(function (c) { return hdr.indexOf(c) < 0; });  // например, «Очки» у журнала, созданного раньше
  if (missing.length) sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length).setValues([missing]);
  return sheet;
}

// Запись журнала из запроса сайта: столбец → значение
function playFields(r) {
  var s = r.scores || {};
  return { 'Дата': r.date, 'Игра': r.game, 'Игроки': r.players.join(', '), 'Победители': r.winners.join(', '), 'Заметка': r.note || '',
    'Очки': r.players.filter(function (n) { return Object.prototype.hasOwnProperty.call(s, n); }).map(function (n) { return n + ': ' + s[n]; }).join(', ') };
}
// «Мэри: 120» → ['Мэри', ': 120']; без очков — [текст, '']
function scorePart(x) {
  var m = String(x).match(/^(.*):\s*(-?\d+(?:\.\d+)?)$/);
  return m ? [m[1], ': ' + m[2]] : [x, ''];
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

// Добавить игру (original пустой) или изменить игру original. Пишутся только изменённые ячейки; «Партий» считает журнал.
// Переименование заодно меняет название игры в журнале и оценках, чтобы статистика не разделилась на две игры.
function saveGame(ss, original, fields) {
  var sheet = sheetById(ss, SHEETS.games), values = sheet.getDataRange().getDisplayValues(), header = values[0];
  var nameIdx = header.indexOf(GAME_NAME), name = fields[GAME_NAME].trim(), row = 0;
  if (nameIdx < 0) throw new Error('На листе «' + sheet.getName() + '» нет столбца «' + GAME_NAME + '»');
  fields[GAME_NAME] = name;
  for (var r = 1; r < values.length; r++) {
    var k = norm(values[r][nameIdx]);
    if (original && !row && k === norm(original)) row = r + 1;
    else if (k && k === norm(name)) throw new Error('Игра «' + name + '» уже есть в коллекции');
  }
  if (original && !row) throw new Error('Игры «' + original + '» нет в таблице — обновите страницу');
  var editable = function (h) { return h && h !== GAME_PLAYS && h in fields; };
  if (!row) {
    sheet.appendRow(header.map(function (h) { return editable(h) ? safe(fields[h]) : ''; }));
  } else {
    var changed = [];
    header.forEach(function (h, i) { if (editable(h) && values[row - 1][i] !== fields[h]) changed.push(i); });
    var cells = changed.map(function (i) { return sheet.getRange(row, i + 1); });
    if (cells.some(function (c) { return c.getFormula(); })) throw new Error('В изменяемых ячейках есть формулы — сайт их не перезаписывает');
    cells.forEach(function (c, n) { c.setValue(safe(fields[header[changed[n]]])); });
    if (norm(original) !== norm(name)) {
      [[LOG_SHEET, 'Игра'], [RATINGS_SHEET, 'Игра']].forEach(function (job) {
        var s = ss.getSheetByName(job[0]);
        if (!s || s.getLastRow() < 2) return;
        var v = s.getDataRange().getDisplayValues(), col = v[0].indexOf(job[1]);
        if (col < 0 || !v.slice(1).some(function (x) { return norm(x[col]) === norm(original); })) return;
        s.getRange(2, col + 1, v.length - 1, 1).setValues(v.slice(1).map(function (x) { return [norm(x[col]) === norm(original) ? safe(name) : x[col]]; }));
      });
    }
  }
  syncCounters(ss);  // новой игре — 0 партий, переименованной — её партии из журнала
}

// Переименовать игрока на Лист5, в журнале и оценках. Если новое имя уже занято — игроки объединяются
// (если оба оценили одну игру, остаётся оценка того, в кого объединяют).
function renamePlayer(ss, from, to) {
  var same = norm(from) === norm(to), players = sheetById(ss, SHEETS.players);
  if (players.getLastRow() > 1) {
    var v = players.getDataRange().getDisplayValues(), col = v[0].indexOf(PLAYER_NAME), fromRow = 0, merge = false;
    if (col < 0) throw new Error('На листе игроков нет столбца «' + PLAYER_NAME + '»');
    for (var r = 1; r < v.length; r++) {
      if (!fromRow && norm(v[r][col]) === norm(from)) fromRow = r + 1;
      else if (!same && norm(v[r][col]) === norm(to)) merge = true;
    }
    if (fromRow && merge) players.deleteRow(fromRow);
    else if (fromRow) players.getRange(fromRow, col + 1).setValue(safe(to));
  }
  var log = ss.getSheetByName(LOG_SHEET);
  if (log && log.getLastRow() > 1) {
    var lv = log.getDataRange().getDisplayValues();
    ['Игроки', 'Победители', LOG_SCORES].forEach(function (h) {
      var c = lv[0].indexOf(h);
      if (c < 0) return;
      var out = lv.slice(1).map(function (row) {
        var seen = {};
        return [safe(splitNames(row[c]).map(function (x) {
          var p = h === LOG_SCORES ? scorePart(x) : [x, ''];
          return [norm(p[0]) === norm(from) ? to : p[0], p[1]];
        }).filter(function (p) { var k = norm(p[0]); return seen[k] ? false : (seen[k] = true); })
          .map(function (p) { return p[0] + p[1]; }).join(', '))];
      });
      if (out.some(function (x, i) { return x[0] !== lv[i + 1][c]; })) log.getRange(2, c + 1, out.length, 1).setValues(out);
    });
  }
  var ratings = ss.getSheetByName(RATINGS_SHEET);
  if (ratings && ratings.getLastRow() > 1) {
    var rv = ratings.getDataRange().getDisplayValues(), gi = rv[0].indexOf('Игра'), pi = rv[0].indexOf('Игрок'), has = {};
    rv.slice(1).forEach(function (row) { if (norm(row[pi]) === norm(to)) has[norm(row[gi])] = true; });
    for (var k = rv.length - 1; k >= 1; k--) {  // снизу вверх: удаление строки не сдвигает ещё не просмотренные
      if (norm(rv[k][pi]) !== norm(from)) continue;
      if (!same && has[norm(rv[k][gi])]) ratings.deleteRow(k + 1);
      else ratings.getRange(k + 1, pi + 1).setValue(safe(to));
    }
  }
  syncCounters(ss);
}

// Удалить игру из коллекции. Записи журнала остаются (это история), оценки игры удаляются.
function deleteGame(ss, name) {
  var sheet = sheetById(ss, SHEETS.games), row = findRow(sheet, [GAME_NAME], [name]);
  if (!row) throw new Error('Игры «' + name.trim() + '» нет в таблице — обновите страницу');
  sheet.deleteRow(row);
  var ratings = ss.getSheetByName(RATINGS_SHEET);
  if (ratings && ratings.getLastRow() > 1) {
    var v = ratings.getDataRange().getDisplayValues(), gi = v[0].indexOf('Игра');
    for (var r = v.length - 1; r >= 1; r--) if (norm(v[r][gi]) === norm(name)) ratings.deleteRow(r + 1);
  }
  syncCounters(ss);
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

// Резервная копия всей таблицы в папку на Google Диске; хранятся последние BACKUP_KEEP копий, старые уходят в корзину.
// Один раз выберите setupMonthlyBackup и нажмите «Выполнить» — дальше копия делается сама 1-го числа каждого месяца.
var BACKUP_FOLDER = 'Домашний хаб — резервные копии', BACKUP_KEEP = 12;
function backupMonthly() {
  var ss = SpreadsheetApp.getActive(), found = DriveApp.getFoldersByName(BACKUP_FOLDER);
  var folder = found.hasNext() ? found.next() : DriveApp.createFolder(BACKUP_FOLDER);
  var stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  DriveApp.getFileById(ss.getId()).makeCopy(ss.getName() + ' — копия ' + stamp, folder);
  var files = [], it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });
}
function setupMonthlyBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'backupMonthly') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupMonthly').timeBased().onMonthDay(1).atHour(4).create();
  backupMonthly();  // первая копия сразу — заодно Google один раз спросит разрешение на Диск
}

/* ---------- Telegram-бот ---------- */
// Настройка — README.md, раздел «Telegram-бот». Свойства скрипта: TG_TOKEN — ключ бота от @BotFather,
// WEBAPP_URL — адрес веб-приложения (/exec), TG_ADMINS — Telegram ID тех, кто может выбирать чат для уведомлений (через запятую).
// TG_SECRET и TG_CHAT скрипт заполняет сам.
var SITE_URL = 'https://proton389.github.io/home-hub/';
var TYPE_LABEL = { Coop: 'Кооператив', Noncoop: 'Соревновательная', Traitor: 'С предателем' };
var TG_ALIAS = { 'помощь': 'help', 'случайная': 'random', 'игра': 'game', 'стат': 'stats', 'статистика': 'stats', 'предзаказы': 'preorders', 'уведомления': 'notify' };
var MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];

function prop(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function telegramWebhook(e) {
  var ok = HtmlService.createHtmlOutput('ok');  // HtmlService отвечает сразу, без переадресации, на которую Telegram жалуется
  if (!prop('TG_TOKEN') || !prop('TG_SECRET') || e.parameter.tg !== prop('TG_SECRET')) return ok;  // не от Telegram — молча игнорируем
  var u = JSON.parse(e.postData.contents), cache = CacheService.getScriptCache(), key = 'tg' + u.update_id;
  // Google отвечает Telegram переадресацией, и тот может прислать то же сообщение ещё раз — обрабатываем один раз
  if (cache.get(key)) return ok;
  cache.put(key, '1', 21600);
  try { handleTelegram(u); } catch (err) { console.error(err); }
  return ok;
}

function handleTelegram(u) {
  if (u.callback_query) {  // кнопка «🎲 Ещё»
    var q = u.callback_query, part = String(q.data || '').split('|');
    tg('answerCallbackQuery', { callback_query_id: q.id });
    if (part[0] === 'rnd' && q.message) tgReply(q.message.chat.id, randomGame(Number(part[1]) || 0, Number(part[2]) || 0));
    return;
  }
  var m = u.message;
  if (!m || !m.text || m.text.charAt(0) !== '/') return;
  var words = m.text.trim().split(/\s+/), cmd = words[0].slice(1).split('@')[0].toLowerCase(), arg = words.slice(1).join(' ');
  cmd = TG_ALIAS[cmd] || cmd;
  var r = cmd === 'start' || cmd === 'help' ? tgHelp(m)
    : cmd === 'random' ? randomGame(parseInt(words[1], 10) || 0, parseInt(words[2], 10) || 0)
    : cmd === 'game' ? (arg ? findGame(arg) : { text: 'Напишите часть названия: /game Ханаби' })
    : cmd === 'stats' ? statsText()
    : cmd === 'preorders' ? preordersText()
    : cmd === 'notify' ? notifyHere(m)
    : null;
  if (r) tgReply(m.chat.id, r);
}

function tg(method, payload) {
  var res = UrlFetchApp.fetch('https://api.telegram.org/bot' + prop('TG_TOKEN') + '/' + method,
    { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true });
  return JSON.parse(res.getContentText());
}
function tgReply(chat, r) {
  var body = { chat_id: chat, text: r.text, parse_mode: 'HTML', disable_web_page_preview: true };
  if (r.buttons) body.reply_markup = { inline_keyboard: r.buttons };
  return tg('sendMessage', body);
}
function h(s) { return String(s === undefined || s === null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }

// Сегодня по часовому поясу таблицы (у скрипта он может быть другим)
function today() {
  var t = Utilities.formatDate(new Date(), SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd').split('-');
  return new Date(Number(t[0]), Number(t[1]) - 1, Number(t[2]));
}
function parseDay(s) {  // «01.12.2026», «01/12/2026» или «2026-12-01»
  var t = String(s || '').trim(), m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/), d;
  if (m) d = [m[3], m[2], m[1]]; else { m = t.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/); if (m) d = [m[1], m[2], m[3]]; }
  return d ? new Date(Number(d[2]), Number(d[1]) - 1, Number(d[0])) : null;
}
function fmtDay(d) { return ('0' + d.getDate()).slice(-2) + '.' + ('0' + (d.getMonth() + 1)).slice(-2) + '.' + d.getFullYear(); }
function daysUntil(d) { return Math.round((d - today()) / 864e5); }
function plural(n, one, few, many) { var a = Math.abs(n) % 100, b = a % 10; return a > 10 && a < 20 ? many : b === 1 ? one : b >= 2 && b <= 4 ? few : many; }

// Всё, что нужно боту, — так же, как считает сайт: партии и победы из журнала, оценки только от тех, кто играл
function hubData() {
  var ss = SpreadsheetApp.getActive(), log = readNamed(ss, LOG_SHEET), plays = {}, played = {}, players = {}, rating = {};
  var player = function (n) { return players[norm(n)] = players[norm(n)] || { name: n, plays: 0, wins: 0 }; };
  log.forEach(function (e) {
    var k = norm(e['Игра']), d = parseDay(e['Дата']), g = plays[k] = plays[k] || { name: String(e['Игра']).trim(), n: 0, last: null };
    g.n++;
    if (d && (!g.last || d > g.last)) g.last = d;
    played[k] = played[k] || {};
    splitNames(e['Игроки']).forEach(function (n) { player(n).plays++; played[k][norm(n)] = true; });
    splitNames(e['Победители']).forEach(function (n) { player(n).wins++; });
  });
  readNamed(ss, RATINGS_SHEET).forEach(function (r) {
    var k = norm(r['Игра']), v = Number(r['Оценка']) || 0;
    if (!v || !(played[k] || {})[norm(r['Игрок'])]) return;
    var x = rating[k] = rating[k] || { sum: 0, n: 0 };
    x.sum += v; x.n++;
  });
  return { games: readSheet(sheetById(ss, SHEETS.games)).filter(function (g) { return String(g[GAME_NAME]).trim(); }),
    preorders: readSheet(sheetById(ss, SHEETS.preorders)), log: log, plays: plays, players: players, rating: rating };
}

function rulesLinks(text) {  // как на сайте: по ссылке на строку, подпись перед ссылкой необязательна
  var out = [], re = /https?:\/\/[^\s,;]+/gi, m, from = 0, t = String(text || '');
  while ((m = re.exec(t))) {
    var start = Math.max(from, t.lastIndexOf('\n', m.index) + 1);
    out.push({ label: t.slice(start, m.index).replace(/^[\s,;:|—–-]+|[\s,;:|—–-]+$/g, ''), url: m[0] });
    from = re.lastIndex;
  }
  return out.map(function (r, i) { return { url: r.url, label: r.label || (out.length > 1 ? 'Буклет ' + (i + 1) : 'Правила') }; });
}

function gameCard(g, d) {
  var k = norm(g[GAME_NAME]), min = String(g['Мин игроков'] || ''), max = String(g['Макс игроков'] || ''), facts = [];
  if (min || max) facts.push('👥 ' + (!max || min === max ? min : min + '–' + max));
  if (String(g['Время'] || '').trim()) facts.push('⏱ ' + g['Время'] + ' мин');
  var types = Object.keys(TYPE_LABEL).filter(function (t) { return Number(g[t]) === 1; }).map(function (t) { return TYPE_LABEL[t]; });
  if (types.length) facts.push(types.join(', '));
  var lines = ['<b>' + h(g[GAME_NAME]) + '</b>'];
  if (facts.length) lines.push(h(facts.join(' · ')));
  if (String(g['Где искать'] || '').trim()) lines.push('📍 ' + h(g['Где искать']));
  var r = d.rating[k], p = d.plays[k];
  lines.push((r ? '★ ' + (r.sum / r.n).toFixed(1) + ' (' + r.n + ' ' + plural(r.n, 'оценка', 'оценки', 'оценок') + ') · ' : '') +
    'партий: ' + (p ? p.n : 0) + (p && p.last ? ', последний раз ' + fmtDay(p.last) : ''));
  var rules = rulesLinks(g['Правила']);
  if (rules.length) lines.push('📖 ' + rules.map(function (x) { return '<a href="' + h(x.url) + '">' + h(x.label) + '</a>'; }).join(' · '));
  return lines.join('\n');
}

function randomGame(n, t) {
  var d = hubData(), pool = d.games.filter(function (g) {
    var min = Number(g['Мин игроков']) || 0, max = Number(g['Макс игроков']) || 0, time = Number(g['Время']) || 0;
    return (!n || (min <= n && n <= max)) && (!t || time <= t);
  });
  if (!pool.length) return { text: 'Под такие условия игр в коллекции нет.' };
  var g = pool[Math.floor(Math.random() * pool.length)];
  var cond = (n ? n + ' ' + plural(n, 'игрок', 'игрока', 'игроков') : '') + (n && t ? ', ' : '') + (t ? 'до ' + t + ' мин' : '');
  return { text: '🎲 ' + (cond ? 'Для «' + cond + '» выпала' : 'Выпала') + ' — из ' + pool.length + ' подходящих:\n\n' + gameCard(g, d),
    buttons: [[{ text: '🎲 Ещё', callback_data: 'rnd|' + n + '|' + t }]] };
}

function findGame(q) {
  var d = hubData(), key = norm(q);
  var hits = d.games.filter(function (g) { return norm(g[GAME_NAME]).indexOf(key) >= 0 || norm(g['Где искать']).indexOf(key) >= 0; });
  var exact = hits.filter(function (g) { return norm(g[GAME_NAME]) === key; });
  if (exact.length || hits.length === 1) return { text: gameCard(exact[0] || hits[0], d) };
  if (!hits.length) return { text: 'Не нашёл «' + h(q) + '». Попробуйте часть названия.' };
  return { text: 'Нашлось ' + hits.length + ':\n' + hits.slice(0, 15).map(function (g) { return '• ' + h(g[GAME_NAME]); }).join('\n') +
    (hits.length > 15 ? '\n…' : '') + '\n\nУточните название: /game …' };
}

function statsText() {
  var d = hubData();
  if (!d.log.length) return { text: 'В журнале пока нет партий.' };
  var ps = Object.keys(d.players).map(function (k) { return d.players[k]; }).filter(function (p) { return p.plays; })
    .sort(function (a, b) { return b.wins - a.wins || b.plays - a.plays; });
  var top = Object.keys(d.plays).map(function (k) { return d.plays[k]; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 5);
  return { text: '📊 <b>Сыграно партий: ' + d.log.length + '</b>\n\n<b>Игроки</b> (победы / партии)\n' +
    ps.map(function (p) { return h(p.name) + ' — ' + p.wins + ' / ' + p.plays + ' (' + Math.round(p.wins / p.plays * 100) + '%)'; }).join('\n') +
    '\n\n<b>Чаще всего играли</b>\n' + top.map(function (g) { return h(g.name) + ' — ' + g.n; }).join('\n') + '\n\n' + SITE_URL };
}

function preordersText() {
  var items = hubData().preorders.filter(function (p) { return String(p['Название'] || '').trim(); })
    .map(function (p) { return { name: p['Название'], d: parseDay(p['ДД/ММ/ГГГГ']) }; })
    .sort(function (a, b) { return (a.d || 9e15) - (b.d || 9e15); });
  if (!items.length) return { text: 'Предзаказов нет.' };
  return { text: '📦 <b>Предзаказы</b>\n' + items.map(function (x) {
    if (!x.d) return '• ' + h(x.name) + ' — дата не указана';
    var n = daysUntil(x.d);
    return '• ' + h(x.name) + ' — ' + fmtDay(x.d) + (n < 0 ? ' (срок прошёл)' : n === 0 ? ' (сегодня)' : ' (через ' + n + ' ' + plural(n, 'день', 'дня', 'дней') + ')');
  }).join('\n') };
}

function isAdmin(from) { return !!from && String(prop('TG_ADMINS') || '').split(/[\s,]+/).indexOf(String(from.id)) >= 0; }
function notifyHere(m) {
  if (!isAdmin(m.from)) return { text: 'Выбрать чат для уведомлений может только администратор бота. Ваш Telegram ID: <code>' + (m.from ? m.from.id : '?') + '</code> — его можно добавить в свойство скрипта TG_ADMINS.' };
  PropertiesService.getScriptProperties().setProperty('TG_CHAT', String(m.chat.id));
  return { text: '🔔 Уведомления будут приходить сюда: о предзаказах — накануне и в день прихода, итоги месяца — 1-го числа.' };
}
function tgHelp(m) {
  return { text: '<b>Домашний хаб</b> — коллекция настольных игр\n\n' +
    '/random 4 60 — случайная игра (на 4 игроков, до 60 минут; числа необязательны)\n' +
    '/game Ханаби — карточка игры: где лежит, правила, оценка, партии\n' +
    '/stats — статистика игроков и самые частые игры\n' +
    '/preorders — предзаказы и когда придут\n' +
    '/notify — присылать уведомления в этот чат (для администратора)\n\n' +
    'Сайт: ' + SITE_URL + '\nВаш Telegram ID: <code>' + (m.from ? m.from.id : '?') + '</code>' };
}

// Каждый день (расписание ставит setupTelegram): напоминания о предзаказах, 1-го числа — итоги прошлого месяца
function tgDaily() {
  var chat = prop('TG_CHAT');
  if (!chat || !prop('TG_TOKEN')) return;
  var d = hubData(), now = today(), msgs = [];
  d.preorders.forEach(function (p) {
    var x = parseDay(p['ДД/ММ/ГГГГ']), n = x ? daysUntil(x) : null;
    if (n === 0) msgs.push('📦 Сегодня приходит <b>' + h(p['Название']) + '</b>');
    if (n === 1) msgs.push('📦 Завтра приходит <b>' + h(p['Название']) + '</b>');
  });
  if (now.getDate() === 1) {
    var from = new Date(now.getFullYear(), now.getMonth() - 1, 1), month = d.log.filter(function (e) { var x = parseDay(e['Дата']); return x && x >= from && x < now; });
    if (month.length) {
      var games = {}, wins = {};
      month.forEach(function (e) {
        games[e['Игра']] = (games[e['Игра']] || 0) + 1;
        splitNames(e['Победители']).forEach(function (n) { wins[n] = (wins[n] || 0) + 1; });
      });
      var best = function (o) { return Object.keys(o).sort(function (a, b) { return o[b] - o[a]; })[0]; }, g = best(games), w = best(wins);
      msgs.push('🗓 <b>Итоги ' + MONTHS_GEN[from.getMonth()] + '</b>: ' +
        month.length + ' ' + plural(month.length, 'партия', 'партии', 'партий') + '. Чаще всего — ' + h(g) + ' (' + games[g] + ')' +
        (w ? '. Больше всех побед — ' + h(w) + ' (' + wins[w] + ')' : '') + '.');
    }
  }
  if (msgs.length) tgReply(chat, { text: msgs.join('\n\n') });
}

// Один раз после того, как добавили TG_TOKEN: подключает бота к этому веб-приложению, задаёт меню команд и ежедневное расписание
function setupTelegram() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('TG_TOKEN')) throw new Error('Сначала добавьте свойство скрипта TG_TOKEN — ключ бота от @BotFather');
  var secret = props.getProperty('TG_SECRET') || Utilities.getUuid().replace(/-/g, '');
  props.setProperty('TG_SECRET', secret);
  var base = props.getProperty('WEBAPP_URL') || ScriptApp.getService().getUrl();
  if (!/\/exec$/.test(base || '')) throw new Error('Не знаю адрес веб-приложения (получил «' + base + '»). Добавьте свойство скрипта WEBAPP_URL — адрес, который заканчивается на /exec');
  var r = tg('setWebhook', { url: base + '?tg=' + secret, allowed_updates: ['message', 'callback_query'], drop_pending_updates: true });
  if (!r.ok) throw new Error('Telegram ответил ошибкой: ' + r.description + ' — проверьте TG_TOKEN');
  tg('setMyCommands', { commands: [
    { command: 'random', description: 'Случайная игра: /random 4 60' }, { command: 'game', description: 'Карточка игры: /game Ханаби' },
    { command: 'stats', description: 'Статистика игроков' }, { command: 'preorders', description: 'Предзаказы' },
    { command: 'notify', description: 'Уведомления в этот чат' }, { command: 'help', description: 'Что умеет бот' }] });
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tgDaily') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tgDaily').timeBased().everyDays(1).atHour(10).create();
  Logger.log('Готово: бот подключён к ' + base + '. Напишите ему /start.');
}

function norm(s) { return String(s === undefined || s === null ? '' : s).trim().toLowerCase(); }
function inc(map, name) { var k = norm(name); if (k) map[k] = (map[k] || 0) + 1; }
function splitNames(s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean); }
function safe(v) { return typeof v === 'string' && v.charAt(0) === '=' ? "'" + v : v; }  // текст «=...» не должен стать формулой
function text(v, max, allowEmpty) { return typeof v === 'string' && v.length <= max && (allowEmpty || v.trim() !== ''); }
function validPlay(r) {
  return text(r.game, 200) && text(r.date, 20) && names(r.players) && r.players.length > 0 && names(r.winners) && (r.note === undefined || text(r.note, 500, true)) &&
    (r.scores === undefined || (!!r.scores && typeof r.scores === 'object' && !Array.isArray(r.scores) && Object.keys(r.scores).every(function (k) {
      return r.players.indexOf(k) >= 0 && typeof r.scores[k] === 'number' && Math.abs(r.scores[k]) < 1e9;  // очки — только у игравших и только числа
    })));
}
// Есть ли игрок среди игравших в эту игру по журналу (оценивать могут только они)
function playedGame(ss, game, player) {
  return readNamed(ss, LOG_SHEET).some(function (e) {
    return norm(e['Игра']) === norm(game) && splitNames(e['Игроки']).some(function (n) { return norm(n) === norm(player); });
  });
}
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
