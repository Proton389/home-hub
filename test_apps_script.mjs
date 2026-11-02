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

function setup({ password = 'secret' } = {}) {
  const sheets = [
    new FakeSheet(0, 'Лист1', [['Название', 'Coop', 'Партий', 'Где искать'], ['Ханаби', '1', '0', 'Шкаф F'], ['Skull', '0', '0', ''], ['', '', '', '']]),
    new FakeSheet(489614522, 'Лист3', [['Название', 'ДД/ММ/ГГГГ', 'Заметка', ''], ['TES', '01.08.2027', 'коробка', 'без заголовка']]),
    new FakeSheet(515997837, 'Лист5', [['Игрок', 'Партий', 'Побед'], ['Оверлорд', '0', '0'], ['Азраэль ', '0', '0']]),
  ];
  const ctx = {
    SpreadsheetApp: { getActive: () => ({
      getSheets: () => sheets,
      getSheetByName: name => sheets.find(s => s.name === name) ?? null,
      insertSheet: name => { const s = new FakeSheet(1000 + sheets.length, name); sheets.push(s); return s; },
    }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => password }) },
    LockService: { getScriptLock: () => { let held = false; return { waitLock: () => { held = true; }, hasLock: () => held, releaseLock: () => { held = false; } }; } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return this; }, s }) },
  };
  vm.runInNewContext(readFileSync(new URL('./apps-script.gs', import.meta.url), 'utf8'), ctx);
  const get = () => JSON.parse(ctx.doGet().s);
  const post = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ password: 'secret', ...body }) } }).s);
  const sheet = name => sheets.find(s => s.name === name);
  const col = (name, header) => { const s = sheet(name); const i = s.v[0].indexOf(header); return s.v.slice(1).map(r => String(r[i] ?? '')); };
  return { sheets, get, post, ctx, sheet, col };
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
t.post({ op: 'rate', game: 'Skull', player: 'Азраэль', value: 2 });
assert.deepEqual(t.get().ratings, [{ Игра: 'Ханаби', Игрок: 'Оверлорд', Оценка: '5' }, { Игра: 'Skull', Игрок: 'Азраэль', Оценка: '2' }]);
t.post({ op: 'rate', game: 'Skull', player: 'Азраэль', value: 0 });
assert.equal(t.get().ratings.length, 1);

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

console.log('apps-script.gs: все проверки прошли');
