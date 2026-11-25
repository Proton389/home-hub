// node test_sleeves.mjs — протекторы из index.html: разбор столбца «Карты» и какие игры запас упакует целиком.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

const html = readFileSync(new URL('./index.html', import.meta.url), 'utf8');
const line = start => html.split('\n').find(l => l.startsWith(start));
const block = html.slice(html.indexOf('/* ---------- Протекторы'), html.indexOf('/* ---------- Полезное'));
const ctx = vm.createContext({ db: { games: [], sleeves: [] }, save: () => Promise.resolve(true), send: body => { ctx.sent = body(); return Promise.resolve(ctx.ok); } });
vm.runInContext([line('const COL ='), line('const SLV ='), line('const num ='), line('const norm ='), line('const byName ='), block].join('\n'), ctx);
const { sizeKey, cardNeeds, sleeveFits, sleeveGroups, addSleeves, packGame, wholePlan, mixedPlan, mixedChoices, buyOptions } = vm.runInContext('({ sizeKey, cardNeeds, sleeveFits, sleeveGroups, addSleeves, packGame, wholePlan, mixedPlan, mixedChoices, buyOptions })', ctx);
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
assert.equal(await packGame(many, wholePlan(group('Mayday'), many)), false);
assert.deepEqual(counts(), [110, 5, 10, 30]);
assert.equal(many.Протекторы, '');
ctx.ok = true;
assert.equal(await packGame(many, wholePlan(group('Mayday'), many)), true);
assert.deepEqual(plain(ctx.sent), { op: 'saveGame', original: 'Много карт', fields: { Название: 'Много карт', Протекторы: '1' } });
assert.deepEqual(counts(), [0, 5, 10, 20]);
assert.equal(many.Протекторы, '1');
assert.deepEqual(fits(), [[], [['Два размера', 35], ['Хватит', 80]]]);

// «Смешанные»: разные размеры — разными названиями, но каждый размер целиком одним названием
ctx.db.sleeves = [
  { Название: 'A', Размер: '56×40', Количество: '100' },
  { Название: 'B', Размер: '34×100', Количество: '50' },
  { Название: 'B', Размер: '40×56', Количество: '10' },
  { Название: 'C', Размер: '100×34', Количество: '100' },
  { Название: 'D', Размер: '70×120', Количество: '5' },
  { Название: 'D', Размер: '80×80', Количество: '5' },
];
ctx.db.games = [
  game('Смесь', '60 шт. 40×56 · 40 шт. 34×100'),
  game('Одним названием', '5 шт. 40×56 · 5 шт. 34×100'),     // B хватает на всё, но можно и смешать — тоже здесь
  game('Только D', '1 шт. 70×120 · 1 шт. 80×80'),           // оба размера есть только у D — смешивать нечего
  game('Делить нельзя', '105 шт. 40×56 · 1 шт. 34×100'),     // 100 у A + 10 у B = 110, но один размер не делится
  game('Один размер', '10 шт. 40×56'),
  game('Смесь, уже упакована', '1 шт. 40×56 · 1 шт. 34×100', '1'),
];
const groups = sleeveGroups(), plans = ctx.db.games.map(g => mixedPlan(g, groups));
assert.deepEqual(plans.map(p => p && plain([...p].map(([k, gr]) => [k, gr.name]))), [[['40x56', 'A'], ['34x100', 'C']], [['40x56', 'A'], ['34x100', 'C']], null, null, null, null]);  // по умолчанию — где размера больше
// на выбор — только названия, закрывающие размер целиком; где больше — первым
assert.deepEqual(plain([...mixedChoices(ctx.db.games[1], groups)].map(([k, ok]) => [k, ok.map(gr => gr.name)])), [['40x56', ['A', 'B']], ['34x100', ['C', 'B']]]);
const mix = ctx.db.games[0];
assert.equal(await packGame(mix, plans[0]), true);
assert.deepEqual(counts(), [40, 50, 10, 60, 5, 5]);
assert.equal(mix.Протекторы, '1');
assert.equal(mixedPlan(mix, sleeveGroups()), null);  // упакованная — больше не предлагается

// «Сколько докупить»: по варианту на марку; запас других марок идёт в дело, размер игры не делится; «Неизвестные» не в счёт
ctx.db.sleeves = [
  { Название: 'P', Размер: '63,5×88', Количество: '100' },
  { Название: 'U', Размер: '63,5×88', Количество: '30' },
  { Название: 'M', Размер: '44×68', Количество: '10' },
  { Название: 'Неизвестные', Размер: '63,5×88', Количество: '1000' },
];
ctx.db.games = [
  game('Г1', '80 шт. 63,5×88'),
  game('Г2', '40 шт. 63,5×88 · 20 шт. 44×68'),
  game('Г3', '25 шт. 63,5×88'),
  game('Упакована', '500 шт. 63,5×88', '1'),
];
const buy = () => plain(buyOptions(sleeveGroups()).map(o => [o.gr.name, o.total, Object.fromEntries(o.plan)]));
// P: 25 карт Г3 — в 30 шт. U, Г1 и Г2 (120) — из своих 100 → докупить 20; 44×68 у M всего 10, Г2 нужно 20 → 20 P
// M: Г1 → 100 шт. P, Г3 → 30 шт. U, Г2 (40) некуда → 40 M; 44×68: свои 10 + докупить 10
// U: Г1 → P (остаётся 20), Г2 и Г3 (65) — свои 30 + 35; 44×68 — все 20
assert.deepEqual(buy(), [['P', 40, { '63.5x88': 20, '44x68': 20 }], ['M', 50, { '63.5x88': 40, '44x68': 10 }], ['U', 55, { '63.5x88': 35, '44x68': 20 }]]);
ctx.db.sleeves.push({ Название: 'P', Размер: '44×68', Количество: '20' });
assert.deepEqual(buy()[0], ['P', 20, { '63.5x88': 20 }]);

console.log('протекторы: все проверки прошли');
