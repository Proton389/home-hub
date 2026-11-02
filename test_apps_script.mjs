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
  getRange(row, col, nr, nc) {
    const grid = src => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => src[row - 1 + i]?.[col - 1 + j] ?? ''));
    return {
      getDisplayValues: () => grid(this.v).map(r => r.map(String)),
      getFormulas: () => grid(this.f),
      clearContent: () => grid(this.v).forEach((r, i) => r.forEach((_, j) => { if (this.v[row - 1 + i]) this.v[row - 1 + i][col - 1 + j] = ''; })),
      setValues: vals => {
        assert.equal(vals.length, nr); vals.forEach(r => assert.equal(r.length, nc));
        vals.forEach((r, i) => r.forEach((x, j) => { (this.v[row - 1 + i] ??= [])[col - 1 + j] = x; }));
      },
    };
  }
}

function setup(password = 'secret') {
  const sheets = [
    new FakeSheet(0, 'Лист1', [['Название', 'Coop'], ['Ханаби', '1'], ['', '']]),
    new FakeSheet(297173692, 'Лист2', [['Название', 'Партий'], ['Ханаби', '2'], ['Skull', '1'], ['Old', '5']]),
    new FakeSheet(489614522, 'Лист3', [['Название', 'ДД/ММ/ГГГГ', 'Заметка', ''], ['TES', '01.08.2027', 'коробка', 'без заголовка']]),
    new FakeSheet(515997837, 'Лист5', [['Игрок', 'Партий', 'Побед'], ['Оверлорд', '19', '19']]),
  ];
  const ctx = {
    SpreadsheetApp: { getActive: () => ({ getSheets: () => sheets }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => password }) },
    LockService: { getScriptLock: () => { let held = false; return { waitLock: () => { held = true; }, hasLock: () => held, releaseLock: () => { held = false; } }; } },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return this; }, s }) },
  };
  vm.runInNewContext(readFileSync(new URL('./apps-script.gs', import.meta.url), 'utf8'), ctx);
  const get = () => JSON.parse(ctx.doGet().s);
  const post = body => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).s);
  return { sheets, get, post };
}

// чтение: заголовки становятся ключами, пустые строки отброшены
let t = setup();
let data = t.get();
assert.deepEqual(data.games, [{ Название: 'Ханаби', Coop: '1' }]);
assert.deepEqual(data.plays.length, 3);
assert.deepEqual(data.preorders, [{ Название: 'TES', 'ДД/ММ/ГГГГ': '01.08.2027', Заметка: 'коробка' }]);

// запись без пароля / с чужим паролем / без PASSWORD в настройках — отказ
assert.equal(t.post({ key: 'plays', rows: [], password: 'nope' }).error, 'password');
assert.equal(setup(null).post({ key: 'plays', rows: [], password: '' }).error, 'password');
assert.equal(t.get().plays.length, 3);

// список игр и мусор не принимаются
assert.ok(t.post({ key: 'games', rows: [], password: 'secret' }).error);
assert.ok(t.post({ key: 'plays', rows: [{ a: { nested: 1 } }], password: 'secret' }).error);
assert.ok(t.post({ key: 'plays', rows: 'x', password: 'secret' }).error);

// запись: строки заменены, лишние старые очищены, «=...» сохранён как текст
assert.deepEqual(t.post({ key: 'plays', rows: [{ Название: 'Ханаби', Партий: 3 }, { Название: '=HACK()', Партий: 1 }], password: 'secret' }), { ok: true });
assert.deepEqual(t.get().plays, [{ Название: 'Ханаби', Партий: '3' }, { Название: "'=HACK()", Партий: '1' }]);

// как делает сайт: прочитал, поменял дату, сохранил — столбец «Заметка» сохранился, столбец без заголовка не тронут
const pre = t.get().preorders;
pre[0]['ДД/ММ/ГГГГ'] = '01.12.2026';
pre.push({ Название: 'Stonesaga', 'ДД/ММ/ГГГГ': '01.03.2027' });
assert.deepEqual(t.post({ key: 'preorders', rows: pre, password: 'secret' }), { ok: true });
assert.deepEqual(t.sheets[2].v.slice(1), [['TES', '01.12.2026', 'коробка', 'без заголовка'], ['Stonesaga', '01.03.2027', '']]);

// лист с формулами не перезаписывается
t = setup();
t.sheets[1].f[1][0] = '=Лист1!A2';
assert.match(t.post({ key: 'plays', rows: [], password: 'secret' }).error, /формулы/);
assert.equal(t.get().plays.length, 3);

console.log('apps-script.gs: все проверки прошли');
