# Дозаполняет пустые ячейки на Лист1 данными с Tesera и BoardGameGeek. Заполненные ячейки не трогает.
# Запуск: python fill_table.py [--dry] [столбцы…] — без столбцов заполняет всё, что умеет (список — FILL внизу).
#   --dry — только показать, что будет записано: в таблицу и в covers/ ничего не пишется.
# Ключ к игре — ссылка на Tesera в «Ссылка». Если ячейка пустая, игра ищется на Tesera по названию (только точное совпадение).
# Обложки сохраняются в covers/ (до 800 px) и covers/s/ (до 240 px) — после запуска их нужно закоммитить и выложить.
import getpass, html, json, os, re, sys, time, urllib.parse, urllib.request
from functools import cached_property

HERE = os.path.dirname(os.path.abspath(__file__))
API = re.search(r"const API = '([^']+)'", open(os.path.join(HERE, 'index.html'), encoding='utf-8').read()).group(1)
TESERA = 'https://api.tesera.ru/'


def fetch(url, body=None, raw=False):
    req = urllib.request.Request(url, body and json.dumps(body).encode('utf-8'),
                                 {'User-Agent': 'home-hub-fill/1.0', 'Content-Type': 'text/plain;charset=utf-8'})
    for attempt in range(3):  # повтор безопасен: запись тех же ячеек теми же значениями
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                data = r.read()
                return data if raw else json.loads(data.decode('utf-8'))
        except OSError:
            if attempt == 2: raise
            time.sleep(3)


def fmt(w):
    return f'{w:.1f}'.replace('.', ',')  # запятая: в русской таблице «2.9» может стать датой


def num(v):
    return str(v) if str(v or '0') != '0' else None  # Tesera пишет 0 там, где данных нет


def text_of(desc):  # HTML описания Tesera → обычный текст с абзацами
    t = re.sub(r'(?i)<br\s*/?>|</p>|</li>', '\n', desc or '')
    t = html.unescape(re.sub(r'<[^>]+>', '', t)).replace('\xa0', ' ')
    return re.sub(r'\n{3,}', '\n\n', '\n'.join(s.strip() for s in t.splitlines())).strip()[:5000] or None  # 5000 — предел скрипта таблицы


class Game:
    def __init__(self, row):
        self.row, self.name = row, row['Название'].strip()

    @cached_property
    def alias(self):
        m = re.search(r'tesera\.ru/game/([^/?#]+)', self.row.get('Ссылка', ''))
        if m: return m.group(1)
        if self.row.get('Ссылка', '').strip(): return None  # ссылка не на Tesera — не гадаем
        found = [x for x in fetch(TESERA + 'search/games?query=' + urllib.parse.quote(self.name))
                 if x.get('type') == 'Game' and self.name.lower() in (str(x.get('title', '')).strip().lower(), str(x.get('value', '')).strip().lower())]
        return found[0]['alias'] if len(found) == 1 else None

    @cached_property
    def tesera(self):
        return fetch(TESERA + 'games/' + urllib.parse.quote(self.alias)) if self.alias else {}

    @property
    def t(self):
        return self.tesera.get('game') or {}

    @cached_property
    def bgg(self):
        bid = self.t.get('bggId')
        return fetch(f'https://api.geekdo.com/api/dynamicinfo?objectid={bid}&objecttype=thing')['item'] if num(bid) else {}


def rules(g):  # PDF правил, выложенные на Tesera, — по строке «подпись — ссылка», как в README
    files = [f for f in g.tesera.get('files', []) if str(f.get('photoUrl', '')).lower().endswith('.pdf') and re.search('(?i)правил', f.get('title', ''))]
    # quote: пробелы и запятые в имени файла оборвали бы ссылку на сайте
    return '\n'.join(f"{f['title'].strip()} — {urllib.parse.quote(f['photoUrl'], safe=':/%')}" for f in files) or None


def cover(g, dry):
    if not g.t.get('photoUrl') or not re.fullmatch(r'[\w.-]+', g.alias): return None
    name = g.alias.lower() + '.jpg'
    if not dry and not os.path.exists(os.path.join(HERE, 'covers', name)):
        import io
        from PIL import Image  # только для обложек: pip install pillow
        img = Image.open(io.BytesIO(fetch(g.t['photoUrl'], raw=True))).convert('RGB')
        for folder, size in (('covers', 800), (os.path.join('covers', 's'), 240)):
            copy = img.copy()
            copy.thumbnail((size, size))
            copy.save(os.path.join(HERE, folder, name), quality=88)
    return 'covers/' + name


FILL = {  # столбец → откуда берётся значение (None — данных нет, ячейка остаётся пустой)
    'Ссылка': lambda g, dry: g.alias and 'https://tesera.ru/game/' + g.alias,
    'Мин игроков': lambda g, dry: num(g.t.get('playersMin')),
    'Макс игроков': lambda g, dry: num(g.t.get('playersMax')),
    'Время': lambda g, dry: num(g.t.get('playtimeMax')) or num(g.t.get('playtimeMin')),
    'Описание': lambda g, dry: text_of(g.t.get('description')) or g.t.get('descriptionShort') or None,
    'Правила': lambda g, dry: rules(g),
    'Обложка': cover,
    'Сложность': lambda g, dry: (lambda w: w and fmt(w))((g.bgg.get('polls') or {}).get('boardgameweight', {}).get('averageweight')),
}

if __name__ == '__main__':
    assert fmt(2.8978) == '2,9' and num('0') is None and num(4) == '4'
    assert text_of('<p>Раз&nbsp;два<br>три</p><p></p><p>четыре</p>') == 'Раз два\nтри\n\nчетыре'
    dry = '--dry' in sys.argv
    sheet = fetch(API)
    cols = [c for c in sys.argv[1:] if not c.startswith('--')] or [c for c in FILL if c in sheet['gameColumns']]
    if set(cols) - set(FILL): sys.exit(f'Не умею заполнять: {", ".join(set(cols) - set(FILL))}. Умею: {", ".join(FILL)}')
    todo = [(r, [c for c in cols if not str(r.get(c, '')).strip()]) for r in sheet['games'] if r.get('Название', '').strip()]
    todo = [(r, empty) for r, empty in todo if empty]
    print(f'Игр с пустыми ячейками ({", ".join(cols)}): {len(todo)}')
    pw = '' if dry or not todo else os.environ.get('HUB_PASSWORD') or getpass.getpass('Пароль сайта: ')
    if pw and fetch(API, {'op': 'check', 'password': pw}).get('ok') is not True:
        sys.exit('Пароль не подошёл')
    for row, empty in todo:
        g, found, missing = Game(row), {}, []
        for c in empty:
            try:
                v = FILL[c](g, dry)
            except Exception as e:
                v = None; print(f'  {g.name}: {c} — ошибка: {e}')
            if v: found[c] = v
            else: missing.append(c)
        if found and not dry:
            res = fetch(API, {'op': 'saveGame', 'password': pw, 'original': g.name, 'fields': {'Название': g.name, **found}})
            if res.get('ok') is not True: print(f'  {g.name}: не записано — {res}'); continue
        print(f'{g.name}: ' + '; '.join([f'{c} = {str(v)[:60]!r}' for c, v in found.items()] + ([f'не нашлось: {", ".join(missing)}'] if missing else [])))
        time.sleep(0.5)
