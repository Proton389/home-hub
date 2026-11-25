// node test_sleeves.mjs — протекторы из index.html: разбор столбца «Карты» и какие игры запас упакует целиком.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');
const line = start => html.split('\n').find(l => l.startsWith(start));
const block = html.slice(html.indexOf('/* ---------- Протекторы'), html.indexOf('/* ---------- Полезное'));
const ctx = vm.createContext({ db: { games: [], sleeves: [] }, save: () => Promise.resolve(true), send: body => { ctx.sent = body(); return Promise.resolve(ctx.ok); } });
vm.runInContext([line('const COL ='), line('const SLV ='), line('const num ='), line('const norm ='), line('const byName ='), block].join('\n'), ctx);
const { sizeKey, cardNeeds, sleeveFits, sleeveGroups, addSleeves, packGame } = vm.runInContext('({ sizeKey, cardNeeds, sleeveFits, sleeveGroups, addSleeves, packGame })', ctx);
const plain = x => JSON.parse(JSON.stringify(x));  // массивы из vm — другого «мира», deepEqual без этого их не сравнит

// размер — без учёта ориентации, запятой/точки и вида «×»
assert.equal(sizeKey('63,5×88'), '63.5x88');
assert.equal(sizeKey('88x63.5'), '63.5x88');
assert.equal(sizeKey('кот'), '');
// одинаковые размеры в одной игре складываются, «≈» не мешает
assert.deepEqual(plain([...cardNeeds({ Карты: '≈ 35 шт. 57×89 · 11 шт. 70×120 · 5 шт. 89×57' })]), [['57x89', 40], ['70x120', 11]]);
assert.equal(cardNeeds({}).size, 0);

const game = (Название, Карты, Протекторы = '') => ({ Название, Карты, Протекторы });
ctx.db.games = [
  game('Хватит', '100 шт. 63,5×88'),
  game('Два размера', '50 шт. 63,5×88 · 10 шт. 70×120'),
  game('Много карт', '120 шт. 63,5×88'),
  game('Уже в протекторах', '10 шт. 63,5×88', '1'),
  game('Другой размер', '10 шт. 44×68'),
  game('Без карт', ''),
];
const group = name => sleeveGroups().find(x => x.name === name);
const fits = (name = 'Mayday') => { const g = group(name); const r = sleeveFits(g.stock); return plain([r.enough.map(x => x.Название), r.short.map(x => [x.g.Название, x.lack])]); };

addSleeves(100, '63,5', '88', 'Mayday');
addSleeves(5, '70', '120', ' mayday ');  // знакомое название — пишется как уже записано
assert.deepEqual(plain(ctx.db.sleeves.map(r => r.Название)), ['Mayday', 'Mayday']);
assert.deepEqual(fits(), [['Хватит'], [['Два размера', 5], ['Много карт', 20]]]);

addSleeves(10, '63.5', '88', 'MAYDAY');  // те же протекторы ещё раз — прибавились к строке, а не новой строкой
assert.equal(ctx.db.sleeves.length, 2);
assert.equal(ctx.db.sleeves[0].Количество, 110);
assert.deepEqual(fits(), [['Хватит'], [['Два размера', 5], ['Много карт', 10]]]);

// другое название — отдельная группа: размеры двух марок для одной игры не складываются
addSleeves(10, '70', '120', 'Hobby World');
assert.deepEqual(plain(sleeveGroups().map(g => g.name)), ['Hobby World', 'Mayday']);
assert.equal(sleeveFits(sleeveGroups()[0].stock).short.length, 0);

// упаковали игру: её карты уходят из запаса (в т. ч. из второй строки того же размера), игра — «в протекторах»
ctx.db.sleeves.push({ Название: 'Mayday', Размер: '88×63,5', Количество: '30' });  // вписано в таблицу руками, ориентация другая
assert.deepEqual(fits(), [['Хватит', 'Много карт'], [['Два размера', 5]]]);
const many = ctx.db.games.find(g => g.Название === 'Много карт'), counts = () => plain(ctx.db.sleeves.map(r => Number(r.Количество)));
ctx.ok = false;  // отметка игры не сохранилась — всё как было
assert.equal(await packGame(group('Mayday'), many), false);
assert.deepEqual(counts(), [110, 5, 10, 30]);
assert.equal(many.Протекторы, '');
ctx.ok = true;
assert.equal(await packGame(group('Mayday'), many), true);
assert.deepEqual(plain(ctx.sent), { op: 'saveGame', original: 'Много карт', fields: { Название: 'Много карт', Протекторы: '1' } });
assert.deepEqual(counts(), [0, 5, 10, 20]);
assert.equal(many.Протекторы, '1');
assert.deepEqual(fits(), [[], [['Два размера', 35], ['Хватит', 80]]]);

console.log('протекторы: все проверки прошли');
