// node test_apps_script.mjs — проверка apps-script.gs на поддельной таблице (без Google).
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

class FakeSheet {
  constructor(id, name, rows = []) { Object.assign(this, { id, name, v: rows.map(r => [...r]), f: rows.map(r => r.map(() => '')) }); }
  getSheetId() { return this.id; }
  getName() { return this.name; }
  getLastRow() { let n = this.v.length; while (n && !this.v[n - 1].some(x => String(x ?? '').trim())) n--; return n; }
  getLastColumn() { return Math.max(0, ...this.v.map(r => r.length)); }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  appendRow(row) { this.v.splice(this.getLastRow(), 0, [...row]); this.f.splice(this.getLastRow() - 1, 0, row.map(() => '')); }
  deleteRow(r) { this.deleteRows(r, 1); }
  deleteRows(r, n) { this.v.splice(r - 1, n); this.f.splice(r - 1, n); }
  getRange(row, col, nr = 1, nc = 1) {
    const grid = src => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => src[row - 1 + i]?.[col - 1 + j] ?? ''));
    const put = (i, j, x) => { (this.v[row - 1 + i] ??= [])[col - 1 + j] = x; };
    return {
      getDisplayValues: () => grid(this.v).map(r => r.map(String)),
      getFormulas: () => grid(this.f),
      getFormula: () => grid(this.f)[0][0],
      clearContent: () => grid(this.v).forEach((r, i) => r.forEach((_, j) => { if (this.v[row - 1 + i]) put(i, j, ''); })),
      setValue: x => put(0, 0, x),
      setValues: vals => {
        assert.equal(vals.length, nr); vals.forEach(r => assert.equal(r.length, nc));
        vals.forEach((r, i) => r.forEach((x, j) => put(i, j, x)));
      },
    };
  }
}

function setup({ password = 'secret', props: extra = {} } = {}) {
  const sheets = [
    new FakeSheet(0, 'Лист1', [['Название', 'Coop', 'Партий', 'Где искать'], ['Ханаби', '1', '0', 'Шкаф F'], ['Skull', '0', '0', ''], ['', '', '', '']]),
    new FakeSheet(489614522, 'Лист3', [['Название', 'ДД/ММ/ГГГГ', 'Заметка', ''], ['TES', '01.08.2027', 'коробка', 'без заголовка']]),
    new FakeSheet(515997837, 'Лист5', [['Игрок', 'Партий', 'Побед'], ['Оверлорд', '0', '0'], ['Азраэль ', '0', '0']]),
  ];
  const drive = { folders: [], copies: [], trashed: [], triggers: [], today: null, tg: [], cache: {}, logs: [] };
  const props = { PASSWORD: password, ...extra };
  const ctx = {
    DriveApp: {
      getFoldersByName: n => { const f = drive.folders.filter(x => x.name === n); return { hasNext: () => f.length > 0, next: () => f.shift() }; },
      createFolder: n => { const f = { name: n, files: [], getFiles() { const a = [...this.files]; return { hasNext: () => a.length > 0, next: () => a.shift() }; } }; drive.folders.push(f); return f; },
      getFileById: () => ({ makeCopy: (name, folder) => {
        const c = { name, n: drive.copies.length, getDateCreated() { return new Date(2026, 0, 1 + this.n); }, setTrashed() { drive.trashed.push(this.n); } };
        drive.copies.push(c); folder.files.push(c); return c;
      } }),
    },
    ScriptApp: {
      getProjectTriggers: () => drive.triggers,
      deleteTrigger: tr => drive.triggers.splice(drive.triggers.indexOf(tr), 1),
      newTrigger: fn => { const tr = { fn, getHandlerFunction: () => fn }; const b = { timeBased: () => b, onMonthDay: d => { tr.day = d; return b; }, everyDays: n => { tr.every = n; return b; }, atHour: h => { tr.hour = h; return b; }, create: () => drive.triggers.push(tr) }; return b; },
      getService: () => ({ getUrl: () => drive.serviceUrl ?? 'https://script.google.com/macros/s/DEPLOY/exec' }),
    },
    Utilities: { formatDate: d => drive.today ?? d.toISOString().slice(0, 10), getUuid: () => 'aaaa-bbbb-cccc' },
    UrlFetchApp: { fetch: (url, o) => { drive.tg.push({ url, method: url.split('/').pop(), body: JSON.parse(o.payload) }); return { getContentText: () => '{"ok":true}' }; } },
    CacheService: { getScriptCache: () => ({ get: k => drive.cache[k] ?? null, put: (k, v) => { drive.cache[k] = v; } }) },
    HtmlService: { createHtmlOutput: s => ({ s }) },
    Logger: { log: m => drive.logs.push(m) },
    SpreadsheetApp: { getActive: () => ({
      getId: () => 'sheet-id', getName: () => 'Хаб', getSpreadsheetTimeZone: () => 'Europe/Moscow',
      getSheets: () => sheets,
      getSheetByName: name => sheets.find(s => s.name === name) ?? null,
      insertSheet: name => { const s = new FakeSheet(1000 + sheets.length, name); sheets.push(s); return s; },
    }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => { let held = false; return { waitLock: () => { held = true; }, hasLock: () => held, releaseLock: () => { held = false; } }; } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return this; }, s }) },
  };
  vm.runInNewContext(readFileSync(new URL('./apps-script.gs', import.meta.url), 'utf8'), ctx);
  const get = () => JSON.parse(ctx.doGet().s);
  const post = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ password: 'secret', ...body }) } }).s);
  const sheet = name => sheets.find(s => s.name === name);
  const col = (name, header) => { const s = sheet(name); const i = s.v[0].indexOf(header); return s.v.slice(1).map(r => String(r[i] ?? '')); };
  return { sheets, get, post, ctx, sheet, col, drive, props };
}
const play = (game, players, winners = [], extra = {}) => ({ op: 'log', date: '10.11.2026', game, players, winners, ...extra });

// чтение: журнала и оценок ещё нет — пустые списки
let t = setup();
let data = t.get();
assert.deepEqual(data.games[0], { Название: 'Ханаби', Coop: '1', Партий: '0', 'Где искать': 'Шкаф F' });
assert.deepEqual([data.log, data.ratings], [[], []]);
assert.deepEqual(data.preorders, [{ Название: 'TES', 'ДД/ММ/ГГГГ': '01.08.2027', Заметка: 'коробка' }]);

// пароль обязателен; мусор не принимается
assert.equal(t.post({ ...play('Ханаби', ['Оверлорд']), password: 'nope' }).error, 'password');
assert.equal(setup({ password: null }).post(play('Ханаби', ['Оверлорд'])).error, 'password');
for (const bad of [play('', ['Оверлорд']), play('Ханаби', []), play('Ханаби', ['Имя, с запятой']), play('Ханаби', 'Оверлорд'),
  play('Ханаби', ['Оверлорд'], ['x'], { note: 5 }), { op: 'unlog', id: 3 }, { op: 'rate', game: 'Ханаби', player: 'Оверлорд', value: 7 },
  { key: 'games', rows: [] }, { key: 'players', rows: 'x' }, { key: 'players', rows: [{ a: { nested: 1 } }] }]) {
  assert.ok(t.post(bad).error, JSON.stringify(bad));
}
assert.equal(t.sheet('Журнал'), undefined);  // отказы ничего не создали

// запись в журнал: лист создаётся сам, счётчики пересчитаны, новый игрок из журнала дописан на Лист5
const r1 = t.post(play('Ханаби', ['Оверлорд', 'Азраэль', 'Новенький'], ['Оверлорд', 'Азраэль'], { note: 'кооп, победили' }));
assert.ok(r1.ok && r1.id);
const r2 = t.post(play(' ханаби ', ['Оверлорд', 'Новенький'], ['Новенький']));
t.post(play('Skull', ['Оверлорд'], []));
data = t.get();
assert.equal(data.log.length, 3);
assert.deepEqual(data.log[0], { ID: r1.id, Дата: '10.11.2026', Игра: 'Ханаби', Игроки: 'Оверлорд, Азраэль, Новенький', Победители: 'Оверлорд, Азраэль', Заметка: 'кооп, победили' });
assert.deepEqual(t.col('Лист1', 'Партий'), ['2', '1', '']);  // пустая строка игр осталась пустой
assert.deepEqual(data.players, [
  { Игрок: 'Оверлорд', Партий: '3', Побед: '1' },
  { Игрок: 'Азраэль ', Партий: '1', Побед: '1' },
  { Игрок: 'Новенький', Партий: '2', Побед: '1' },
]);
assert.equal(t.col('Лист1', 'Где искать')[0], 'Шкаф F');  // другие столбцы Лист1 не тронуты

// удаление записи пересчитывает счётчики; несуществующая запись — понятная ошибка
assert.deepEqual(t.post({ op: 'unlog', id: r2.id }), { ok: true });
assert.deepEqual(t.col('Лист1', 'Партий'), ['1', '1', '']);
assert.deepEqual(t.get().players.map(p => [p['Игрок'], p['Партий'], p['Побед']]), [['Оверлорд', '2', '1'], ['Азраэль ', '1', '1'], ['Новенький', '1', '0']]);
assert.match(t.post({ op: 'unlog', id: r2.id }).error, /нет/);

// игроки: сохранение листа целиком возвращает правильные счётчики, даже если прислали устаревшие
const stale = t.get().players.map(p => ({ ...p, Партий: 99 }));
assert.deepEqual(t.post({ key: 'players', rows: stale }), { ok: true });
assert.equal(t.get().players[0]['Партий'], '2');

// оценки: создание, изменение (без учёта регистра), снятие оценкой 0
assert.deepEqual(t.post({ op: 'rate', game: 'Ханаби', player: 'Оверлорд', value: 4 }), { ok: true });
t.post({ op: 'rate', game: 'ханаби', player: 'оверлорд ', value: 5 });
assert.match(t.post({ op: 'rate', game: 'Skull', player: 'Азраэль', value: 2 }).error, /играл/);  // Азраэль в Skull не играл
t.post({ op: 'rate', game: 'Skull', player: 'Оверлорд', value: 2 });
assert.deepEqual(t.get().ratings, [{ Игра: 'Ханаби', Игрок: 'Оверлорд', Оценка: '5' }, { Игра: 'Skull', Игрок: 'Оверлорд', Оценка: '2' }]);
t.post({ op: 'rate', game: 'Skull', player: 'Оверлорд', value: 0 });  // снять оценку можно всегда
assert.equal(t.get().ratings.length, 1);

// правка записи журнала: поля меняются на месте, ID тот же, счётчики пересчитаны
assert.deepEqual(t.post({ ...play('Skull', ['Оверлорд', 'Новенький'], ['Новенький'], { note: 'исправлено', date: '11.11.2026' }), op: 'editLog', id: r1.id }), { ok: true });
data = t.get();
assert.deepEqual(data.log.find(e => e.ID === r1.id), { ID: r1.id, Дата: '11.11.2026', Игра: 'Skull', Игроки: 'Оверлорд, Новенький', Победители: 'Новенький', Заметка: 'исправлено' });
assert.equal(data.log.findIndex(e => e.ID === r1.id), 0);  // место в журнале не сменилось
assert.deepEqual(t.col('Лист1', 'Партий'), ['0', '2', '']);
assert.deepEqual(data.players.map(p => [p['Игрок'], p['Партий'], p['Побед']]), [['Оверлорд', '2', '0'], ['Азраэль ', '0', '0'], ['Новенький', '1', '1']]);
assert.match(t.post({ ...play('Skull', ['Оверлорд']), op: 'editLog', id: 'нет-такой' }).error, /нет/);
assert.ok(t.post({ ...play('', ['Оверлорд']), op: 'editLog', id: r1.id }).error);
assert.ok(t.post({ ...play('Skull', []), op: 'editLog', id: r1.id }).error);

// «=...» сохраняется как текст, а не формула
t.post(play('=HACK()', ['=X']));
assert.equal(t.get().log.at(-1)['Игра'], "'=HACK()");

// формулы в столбце-счётчике: запись в журнал остаётся, а ошибка объясняет, почему счётчики не обновились
t = setup();
t.sheet('Лист1').f[1][2] = '=1+1';
assert.match(t.post(play('Ханаби', ['Оверлорд'])).error, /формулы/);

// предзаказы как делает сайт: прочитал, поменял, сохранил — «Заметка» на месте, столбец без заголовка не тронут
t = setup();
const pre = t.get().preorders;
pre[0]['ДД/ММ/ГГГГ'] = '01.12.2026';
pre.push({ Название: 'Stonesaga', 'ДД/ММ/ГГГГ': '01.03.2027' });
assert.deepEqual(t.post({ key: 'preorders', rows: pre }), { ok: true });
assert.deepEqual(t.sheet('Лист3').v.slice(1), [['TES', '01.12.2026', 'коробка', 'без заголовка'], ['Stonesaga', '01.03.2027', '']]);

// обнуление = очистка журнала; оценки остаются
t = setup();
t.post(play('Ханаби', ['Оверлорд'], ['Оверлорд']));
t.post({ op: 'rate', game: 'Ханаби', player: 'Оверлорд', value: 3 });
t.ctx.resetStats();
data = t.get();
assert.deepEqual([data.log.length, data.ratings.length], [0, 1]);
assert.deepEqual(t.col('Лист1', 'Партий'), ['0', '0', '']);
assert.deepEqual(data.players.map(p => p['Побед']), ['0', '0']);

// игры: проверка пароля, добавление, правка, переименование вместе с журналом и оценками
t = setup();
assert.deepEqual(t.get().gameColumns, ['Название', 'Coop', 'Партий', 'Где искать']);
assert.deepEqual(t.post({ op: 'check' }), { ok: true });
assert.equal(t.post({ op: 'check', password: 'nope' }).error, 'password');
for (const bad of [{ fields: { Coop: '1' } }, { fields: { Название: '  ' } }, { fields: { Название: 'X', Coop: 1 } }, { fields: 'x' }, { fields: { Название: 'X' }, original: 5 }]) {
  assert.ok(t.post({ op: 'saveGame', ...bad }).error, JSON.stringify(bad));
}
assert.deepEqual(t.post({ op: 'saveGame', fields: { Название: ' Новая ', Coop: '1', 'Где искать': 'Шкаф A', Партий: '99', Лишний: 'x' } }), { ok: true });
assert.deepEqual(t.sheet('Лист1').v[3], ['Новая', '1', 0, 'Шкаф A']);  // «Партий» не взят из формы, неизвестный столбец не создан
assert.match(t.post({ op: 'saveGame', fields: { Название: 'ханаби' } }).error, /уже есть/);
assert.match(t.post({ op: 'saveGame', original: 'Нет такой', fields: { Название: 'Y' } }).error, /нет в таблице/);
assert.deepEqual(t.post({ op: 'saveGame', original: 'Skull', fields: { Название: 'Skull', 'Где искать': 'Полка 3' } }), { ok: true });
assert.deepEqual(t.sheet('Лист1').v[2], ['Skull', '0', 0, 'Полка 3']);
assert.match(t.post({ op: 'saveGame', original: 'Skull', fields: { Название: 'Новая' } }).error, /уже есть/);  // переименовать в существующую нельзя
t.post(play('Skull', ['Оверлорд'], ['Оверлорд']));
t.post({ op: 'rate', game: 'Skull', player: 'Оверлорд', value: 4 });
assert.deepEqual(t.post({ op: 'saveGame', original: 'skull', fields: { Название: 'Skull Deluxe', Coop: '0', 'Где искать': 'Полка 3' } }), { ok: true });
data = t.get();
assert.equal(data.games[1]['Название'], 'Skull Deluxe');
assert.equal(data.games[1]['Партий'], '1');  // партии переехали вместе с названием
assert.deepEqual([data.log[0]['Игра'], data.ratings[0]['Игра']], ['Skull Deluxe', 'Skull Deluxe']);
t.sheet('Лист1').f[1][3] = '=A1';  // формула в изменяемой ячейке — отказ, ничего не меняется
assert.match(t.post({ op: 'saveGame', original: 'Ханаби', fields: { Название: 'Ханаби', 'Где искать': 'другое' } }).error, /формулы/);
assert.equal(t.sheet('Лист1').v[1][3], 'Шкаф F');

// переименование игрока: Лист5, журнал (игроки и победители), оценки
t = setup();
t.post(play('Ханаби', ['Оверлорд', 'Азраэль'], ['Азраэль']));
t.post(play('Skull', ['Азраэль'], []));
t.post({ op: 'rate', game: 'Ханаби', player: 'Азраэль', value: 5 });
t.post({ op: 'rate', game: 'Skull', player: 'Азраэль', value: 2 });
t.post({ op: 'rate', game: 'Ханаби', player: 'Оверлорд', value: 3 });
for (const bad of [{ from: '', to: 'X' }, { from: 'Азраэль', to: '' }, { from: 'Азраэль', to: 'С, запятой' }]) assert.ok(t.post({ op: 'renamePlayer', ...bad }).error, JSON.stringify(bad));
assert.deepEqual(t.post({ op: 'renamePlayer', from: 'азраэль ', to: 'Азраэль Великий' }), { ok: true });
data = t.get();
assert.deepEqual(data.players.map(p => [p['Игрок'], p['Партий'], p['Побед']]), [['Оверлорд', '1', '0'], ['Азраэль Великий', '2', '1']]);
assert.deepEqual(data.log.map(e => [e['Игроки'], e['Победители']]), [['Оверлорд, Азраэль Великий', 'Азраэль Великий'], ['Азраэль Великий', '']]);
assert.deepEqual(data.ratings.map(r => r['Игрок']), ['Азраэль Великий', 'Азраэль Великий', 'Оверлорд']);
// объединение с существующим: строка игрока уходит, партии складываются, совпавшая оценка — у того, в кого объединяют
assert.deepEqual(t.post({ op: 'renamePlayer', from: 'Азраэль Великий', to: 'Оверлорд' }), { ok: true });
data = t.get();
assert.deepEqual(data.players.map(p => [p['Игрок'], p['Партий'], p['Побед']]), [['Оверлорд', '2', '1']]);
assert.deepEqual(data.log.map(e => e['Игроки']), ['Оверлорд', 'Оверлорд']);  // дубль в одной партии убран
assert.deepEqual(data.ratings.map(r => [r['Игра'], r['Игрок'], r['Оценка']]).sort(), [['Skull', 'Оверлорд', '2'], ['Ханаби', 'Оверлорд', '3']]);
// только регистр/пробелы — просто исправление имени, оценки не теряются
t.post({ op: 'renamePlayer', from: 'Оверлорд', to: 'оверлорд' });
assert.deepEqual([t.get().players[0]['Игрок'], t.get().ratings.length], ['оверлорд', 2]);

// удаление игры: строка уходит, оценки игры удаляются, журнал остаётся
assert.deepEqual(t.post({ op: 'deleteGame', name: 'skull' }), { ok: true });
data = t.get();
assert.deepEqual(data.games.map(g => g['Название']), ['Ханаби']);
assert.deepEqual(data.ratings.map(r => r['Игра']), ['Ханаби']);
assert.equal(data.log.length, 2);
assert.match(t.post({ op: 'deleteGame', name: 'skull' }).error, /нет в таблице/);

// резервные копии: папка одна, расписание одно (повторная настройка не плодит), хранятся 12 последних
t = setup();
t.ctx.setupMonthlyBackup();
t.ctx.setupMonthlyBackup();
assert.deepEqual(t.drive.triggers.map(x => [x.fn, x.day, x.hour]), [['backupMonthly', 1, 4]]);
assert.equal(t.drive.folders.length, 1);
for (let i = 0; i < 12; i++) t.ctx.backupMonthly();
assert.equal(t.drive.copies.length, 14);
assert.deepEqual([...new Set(t.drive.trashed)].sort((a, b) => a - b), [0, 1]);  // в корзину ушли две самые старые

// Telegram-бот: подключение
t = setup({ props: { TG_TOKEN: '123:ABC', TG_ADMINS: '111, 222' } });
const sh1 = t.sheet('Лист1');
sh1.v = [['Название', 'Мин игроков', 'Макс игроков', 'Время', 'Coop', 'Noncoop', 'Партий', 'Где искать', 'Правила'],
  ['Ханаби', '2', '5', '30', '1', '0', '0', 'Шкаф F', 'https://drive.google.com/a'],
  ['Skull', '3', '6', '20', '0', '1', '0', '', 'База https://x.ru/1\nДополнение https://x.ru/2'],
  ['Сумерки <империи>', '3', '6', '480', '0', '1', '0', '', '']];
sh1.f = sh1.v.map(r => r.map(() => ''));
t.post(play('Ханаби', ['Оверлорд', 'Азраэль'], ['Азраэль']));
t.post(play('Ханаби', ['Оверлорд'], [], { date: '12.11.2026' }));
t.post({ op: 'rate', game: 'Ханаби', player: 'Оверлорд', value: 4 });
t.post({ op: 'rate', game: 'Ханаби', player: 'Азраэль', value: 5 });
t.sheet('Оценки').appendRow(['Ханаби', 'Волчара', '1']);  // оценка того, кто не играл, — не считается
t.ctx.setupTelegram();
assert.equal(t.props.TG_SECRET, 'aaaabbbbcccc');
const hook = t.drive.tg.find(c => c.method === 'setWebhook');
assert.equal(hook.url, 'https://api.telegram.org/bot123:ABC/setWebhook');
assert.equal(hook.body.url, 'https://script.google.com/macros/s/DEPLOY/exec?tg=aaaabbbbcccc');
assert.ok(t.drive.tg.some(c => c.method === 'setMyCommands'));
assert.match(t.drive.logs[0], /Готово/);
t.ctx.setupTelegram();  // повторно — секрет тот же, расписание одно
assert.equal(t.props.TG_SECRET, 'aaaabbbbcccc');
assert.deepEqual(t.drive.triggers.map(x => [x.fn, x.every, x.hour]), [['tgDaily', 1, 10]]);
t.drive.serviceUrl = 'https://script.google.com/macros/s/SCRIPT/dev';  // тестовый адрес Telegram не подойдёт
assert.throws(() => t.ctx.setupTelegram(), /WEBAPP_URL/);
delete t.drive.serviceUrl;

// команды
let upd = 0;
const tgSay = (text, from = 111, chat = from) => {
  t.drive.tg = [];
  const r = t.ctx.doPost({ parameter: { tg: t.props.TG_SECRET }, postData: { contents: JSON.stringify({ update_id: ++upd, message: { text, from: { id: from }, chat: { id: chat } } }) } });
  assert.equal(r.s, 'ok');
  return t.drive.tg.filter(c => c.method === 'sendMessage').map(c => c.body);
};
let [msg] = tgSay('/start', 555);
assert.equal(msg.chat_id, 555);
assert.equal(msg.parse_mode, 'HTML');
assert.match(msg.text, /<code>555<\/code>/);
[msg] = tgSay('/game ханаби');
assert.match(msg.text, /<b>Ханаби<\/b>/);
assert.match(msg.text, /👥 2–5 · ⏱ 30 мин · Кооператив/);
assert.match(msg.text, /📍 Шкаф F/);
assert.match(msg.text, /★ 4\.5 \(2 оценки\) · партий: 2, последний раз 12\.11\.2026/);
[msg] = tgSay('/game@HubBot skull');
assert.match(msg.text, /<a href="https:\/\/x\.ru\/1">База<\/a> · <a href="https:\/\/x\.ru\/2">Дополнение<\/a>/);
assert.match(msg.text, /партий: 0$/m);
[msg] = tgSay('/игра сумерки');  // русская команда тоже работает; < > экранированы
assert.match(msg.text, /Сумерки &lt;империи&gt;/);
assert.match(tgSay('/game и')[0].text, /Нашлось 2:\n• Ханаби\n• Сумерки/);  // «и» — в Ханаби и в Сумерках
assert.match(tgSay('/game zzz')[0].text, /Не нашёл/);
assert.match(tgSay('/game')[0].text, /Напишите часть названия/);
for (let i = 0; i < 20; i++) {
  [msg] = tgSay('/random 2 60');
  assert.match(msg.text, /Ханаби/);  // на двоих до часа — только Ханаби
  assert.match(msg.text, /из 1 подходящих/);
}
assert.deepEqual(msg.reply_markup.inline_keyboard[0][0], { text: '🎲 Ещё', callback_data: 'rnd|2|60' });
assert.match(tgSay('/random 9')[0].text, /игр в коллекции нет/);
assert.match(tgSay('/random')[0].text, /из 3 подходящих/);
[msg] = tgSay('/stats');
assert.match(msg.text, /Сыграно партий: 2/);
assert.match(msg.text, /Азраэль — 1 \/ 1 \(100%\)\nОверлорд — 0 \/ 2 \(0%\)/);
assert.match(msg.text, /Ханаби — 2/);
t.drive.today = '2026-11-09';
[msg] = tgSay('/preorders');
assert.match(msg.text, /TES — 01\.08\.2027 \(через 265 дней\)/);
assert.deepEqual(tgSay('просто текст'), []);  // не команда — молчим
assert.deepEqual(tgSay('/unknown'), []);

// уведомления: выбрать чат может только администратор
assert.match(tgSay('/notify', 999, -100)[0].text, /только администратор/);
assert.equal(t.props.TG_CHAT, undefined);
assert.match(tgSay('/notify', 222, -100)[0].text, /будут приходить сюда/);
assert.equal(t.props.TG_CHAT, '-100');

// чужие запросы и повторы игнорируются; сайт работает как раньше
t.drive.tg = [];
for (const parameter of [{ tg: 'wrong' }, { tg: '' }]) {
  assert.equal(t.ctx.doPost({ parameter, postData: { contents: JSON.stringify({ update_id: 900, message: { text: '/start', from: { id: 1 }, chat: { id: 1 } } }) } }).s, 'ok');
}
assert.equal(t.drive.tg.length, 0);
const same = { parameter: { tg: t.props.TG_SECRET }, postData: { contents: JSON.stringify({ update_id: 901, message: { text: '/start', from: { id: 1 }, chat: { id: 1 } } }) } };
t.ctx.doPost(same); t.ctx.doPost(same);
assert.equal(t.drive.tg.filter(c => c.method === 'sendMessage').length, 1);
assert.equal(t.post({ op: 'check' }).ok, true);

// кнопка «🎲 Ещё»
t.drive.tg = [];
t.ctx.doPost({ parameter: { tg: t.props.TG_SECRET }, postData: { contents: JSON.stringify({ update_id: 902, callback_query: { id: 'cb1', data: 'rnd|2|60', message: { chat: { id: -100 } } } }) } });
assert.deepEqual(t.drive.tg.map(c => c.method), ['answerCallbackQuery', 'sendMessage']);
assert.match(t.drive.tg[1].body.text, /Ханаби/);

// ежедневная рассылка: предзаказ завтра и сегодня; 1-го числа — итоги прошлого месяца
const pre3 = t.get().preorders;
pre3.push({ Название: 'Stonesaga', 'ДД/ММ/ГГГГ': '01.12.2026' }, { Название: 'Сегодняшний', 'ДД/ММ/ГГГГ': '30.11.2026' });
t.post({ key: 'preorders', rows: pre3 });
t.drive.today = '2026-11-30'; t.drive.tg = [];
t.ctx.tgDaily();
assert.equal(t.drive.tg.length, 1);
assert.equal(t.drive.tg[0].body.chat_id, '-100');
assert.match(t.drive.tg[0].body.text, /Завтра приходит <b>Stonesaga<\/b>/);
assert.match(t.drive.tg[0].body.text, /Сегодня приходит <b>Сегодняшний<\/b>/);
t.drive.today = '2026-12-01'; t.drive.tg = [];
t.ctx.tgDaily();
assert.match(t.drive.tg[0].body.text, /Сегодня приходит <b>Stonesaga<\/b>/);
assert.match(t.drive.tg[0].body.text, /Итоги ноября<\/b>: 2 партии\. Чаще всего — Ханаби \(2\)\. Больше всех побед — Азраэль \(1\)\./);
t.drive.today = '2026-12-05'; t.drive.tg = [];
t.ctx.tgDaily();
assert.equal(t.drive.tg.length, 0);  // нечего сказать — молчим
delete t.props.TG_CHAT;
t.drive.today = '2026-11-30';
t.ctx.tgDaily();
assert.equal(t.drive.tg.length, 0);  // чат не выбран — не шлём

// без ключа бота webhook ничего не делает, а настройка объясняет, чего не хватает
t = setup();
assert.equal(t.ctx.doPost({ parameter: { tg: '' }, postData: { contents: '{"update_id":1}' } }).s, 'ok');
assert.equal(t.drive.tg.length, 0);
assert.throws(() => t.ctx.setupTelegram(), /TG_TOKEN/);

console.log('apps-script.gs: все проверки прошли');
