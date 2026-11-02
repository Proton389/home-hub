"""Домашний хаб на своём компьютере. Запуск: start.bat (или python server.py), затем открыть напечатанный адрес.

Это тот же сайт, что на GitHub Pages: сервер раздаёт файлы из этой папки (копии репозитория),
данные — из той же Google-таблицы. start.bat перед запуском подтягивает свежую версию с GitHub.
"""
import socket
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

PORT = 8080
SITE = Path(__file__).resolve().parent


class Handler(SimpleHTTPRequestHandler):
    extensions_map = {**SimpleHTTPRequestHandler.extensions_map,
                      '.webmanifest': 'application/manifest+json', '.js': 'text/javascript', '.mjs': 'text/javascript'}

    def send_head(self):
        # .git и прочие служебные файлы наружу не отдаём
        if any(part.startswith('.') for part in self.path.split('?')[0].split('/')):
            self.send_error(404)
            return None
        return super().send_head()

    def end_headers(self):
        self.send_header('Cache-Control', 'no-cache')  # чтобы после обновления телефон не показывал старую страницу
        super().end_headers()


def lan_ips():
    # все адреса компьютера: при включённом VPN «основной» адрес — это VPN, а не Wi-Fi
    ips = {a[4][0] for a in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET)}
    return sorted(ip for ip in ips if not ip.startswith(('127.', '169.254.'))) or ['IP-этого-компьютера']


if __name__ == '__main__':
    print(f'На этом компьютере:  http://localhost:{PORT}')
    print('С телефона в Wi-Fi — один из адресов (обычно адрес Wi-Fi-адаптера):')
    for ip in lan_ips():
        print(f'    http://{ip}:{PORT}')
    ThreadingHTTPServer(('0.0.0.0', PORT), partial(Handler, directory=str(SITE))).serve_forever()
