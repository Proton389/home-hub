// node test_apps_script.mjs — проверка apps-script.gs на поддельной таблице (без Google).
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

class FakeSheet {
  constructor(id, name, rows) { Object.assign(this, { id, name, v: rows.map(r => [...r]), f: rows.map(r => r.map(() => '')) }); }
  getSheetId() { return this.id; }
  getName() { return this.name; }
  getLastRow() { return this.v.length; }
  getLastColumn() { return Math.max(0, ...this.v.map(r => r.length)); }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
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

function setup({ password = 'secret', games } = {}) {
  const sheets = [
    new FakeSheet(0, 'Лист1', games ?? [['Название', 'Coop', 'Партий'], ['Ханаби', '1', '2'], ['Skull', '0', ''], ['', '', '']]),
    new FakeSheet(297173692, 'Лист2', [['Название', 'Партий'], ['Ханаби ', '3'], ['Skull', '0'], ['Удалённая', '2']]),
    new FakeSheet(489614522, 'Лист3', [['Название', 'ДД/ММ/ГГГГ', 'Заметка', ''], ['TES', '01.08.2027', 'коробка', 'без заголовка']]),
    new FakeSheet(515997837, 'Лист5', [['Игрок', 'Партий', 'Побед'], ['Оверлорд', '19', '19']]),
  ];
  const logs = [];
  const ctx = {
    SpreadsheetApp: { getActive: () => ({ getSheets: () => sheets }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => password }) },
    LockService: { getScriptLock: () => { let held = false; return { waitLock: () => { held = true; }, hasLock: () => held, releaseLock: () => { held = false; } }; } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return this; }, s }) },
    Logger: { log: s => logs.push(s) },
  };
  vm.runInNewContext(readFileSync(new URL('./apps-script.gs', import.meta.url), 'utf8'), ctx);
  const get = () => JSON.parse(ctx.doGet().s);
  const post = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).s);
  return { sheets, get, post, ctx, logs };
}
const play = (t, name, delta, password = 'secret') => t.post({ op: 'addPlay', name, delta, password });

// чтение: заголовки — ключи, пустые строки отброшены, партии собраны из Лист1 (Лист2 больше не читается)
let t = setup();
let data = t.get();
assert.deepEqual(data.games, [{ Название: 'Ханаби', Coop: '1', Партий: '2' }, { Название: 'Skull', Coop: '0', Партий: '' }]);
assert.deepEqual(data.plays, [{ Название: 'Ханаби', Партий: '2' }, { Название: 'Skull', Партий: '0' }]);
assert.deepEqual(data.preorders, [{ Название: 'TES', 'ДД/ММ/ГГГГ': '01.08.2027', Заметка: 'коробка' }]);
t.sheets.splice(1, 1);  // Лист2 удалён — сайт продолжает работать
assert.equal(t.get().plays.length, 2);

// без пароля / с чужим паролем / без PASSWORD в настройках — отказ, ничего не меняется
assert.equal(play(t, 'Ханаби', 1, 'nope').error, 'password');
assert.equal(play(setup({ password: null }), 'Ханаби', 1, '').error, 'password');
assert.equal(t.get().plays[0]['Партий'], '2');

// мусор и Лист1 целиком не принимаются
for (const bad of [{ key: 'games', rows: [] }, { key: 'plays', rows: [] }, { key: 'players', rows: [{ a: { nested: 1 } }] }, { key: 'players', rows: 'x' },
  { op: 'addPlay', name: 'Ханаби', delta: 5 }, { op: 'addPlay', name: 42, delta: 1 }]) {
  assert.ok(t.post({ ...bad, password: 'secret' }).error, JSON.stringify(bad));
}

// партия: меняется одна ячейка, имя ищется без учёта пробелов и регистра, ниже нуля не уходит
assert.deepEqual(play(t, 'Ханаби', 1), { ok: true, value: 3 });
assert.deepEqual(play(t, ' ханаби ', 1), { ok: true, value: 4 });
assert.deepEqual(play(t, 'Skull', 1), { ok: true, value: 1 });
assert.deepEqual(play(t, 'Skull', -1), { ok: true, value: 0 });
assert.deepEqual(play(t, 'Skull', -1), { ok: true, value: 0 });
assert.deepEqual(t.sheets[0].v.slice(1, 3), [['Ханаби', '1', 4], ['Skull', '0', 0]]);
assert.match(play(t, 'Нет такой', 1).error, /нет на листе/);
assert.match(play(setup({ games: [['Название', 'Coop'], ['Ханаби', '1']] }), 'Ханаби', 1).error, /нет столбца «Партий»/);
t.sheets[0].f[1][2] = '=1+1';
assert.match(play(t, 'Ханаби', 1).error, /формула/);

// игроки: «=...» сохраняется как текст; лист с формулами не перезаписывается
t = setup();
assert.deepEqual(t.post({ key: 'players', rows: [{ Игрок: '=HACK()', Партий: 1, Побед: 0 }], password: 'secret' }), { ok: true });
assert.equal(t.get().players[0]['Игрок'], "'=HACK()");
t.sheets[3].f[1][0] = '=A1';
assert.match(t.post({ key: 'players', rows: [], password: 'secret' }).error, /формулы/);

// предзаказы как делает сайт: прочитал, поменял, сохранил — «Заметка» на месте, столбец без заголовка не тронут
t = setup();
const pre = t.get().preorders;
pre[0]['ДД/ММ/ГГГГ'] = '01.12.2026';
pre.push({ Название: 'Stonesaga', 'ДД/ММ/ГГГГ': '01.03.2027' });
assert.deepEqual(t.post({ key: 'preorders', rows: pre, password: 'secret' }), { ok: true });
assert.deepEqual(t.sheets[2].v.slice(1), [['TES', '01.12.2026', 'коробка', 'без заголовка'], ['Stonesaga', '01.03.2027', '']]);

// обнуление: «Партий» на Лист1 и статистика игроков — 0, остальное не тронуто, пустые строки игроков убраны
t = setup();
t.sheets[3].v.splice(2, 0, ['', '', '']);
t.ctx.resetStats();
assert.deepEqual(t.sheets[0].v.slice(1), [['Ханаби', '1', 0], ['Skull', '0', 0], ['', '', '']]);
data = t.get();
assert.deepEqual(data.players, [{ Игрок: 'Оверлорд', Партий: '0', Побед: '0' }]);
assert.equal(data.preorders[0]['ДД/ММ/ГГГГ'], '01.08.2027');

// перенос с Лист2: столбец создаётся, партии переезжают по названию, ненайденные — в журнал; повтор ничего не меняет
t = setup({ games: [['Название', 'Coop'], ['Ханаби', '1'], ['Skull', '0'], ['Новая', '1']] });
t.ctx.migratePlays();
assert.deepEqual(t.sheets[0].v, [['Название', 'Coop', 'Партий'], ['Ханаби', '1', 3], ['Skull', '0', 0], ['Новая', '1', 0]]);
assert.match(t.logs[0], /Не перенесены.*удалённая/);
t.sheets[0].v[1][2] = 7;  // уже заполненное не меняется (записывается то же значение)
t.ctx.migratePlays();
assert.equal(String(t.sheets[0].v[1][2]), '7');
t.sheets[0].f[2][2] = '=1';  // формулы в столбце — отказ
assert.throws(() => t.ctx.migratePlays(), /формулы/);
t = setup();  // столбец «Партий» уже есть и заполнен у Ханаби — берётся только пустая ячейка Skull
t.ctx.migratePlays();
assert.deepEqual(t.sheets[0].v.slice(1, 3), [['Ханаби', '1', '2'], ['Skull', '0', 0]]);

console.log('apps-script.gs: все проверки прошли');
