"""python -m unittest test_server — локальный сервер отдаёт сайт и не отдаёт служебные файлы."""
import threading
import unittest
import urllib.error
import urllib.request
from functools import partial
from http.server import ThreadingHTTPServer

import server


def get(url):
    try:
        with urllib.request.urlopen(url) as r:
            return r.status, r.headers.get('Content-Type', ''), r.read()
    except urllib.error.HTTPError as e:
        return e.code, '', b''


class ServerTest(unittest.TestCase):
    def setUp(self):
        self.httpd = ThreadingHTTPServer(('127.0.0.1', 0), partial(server.Handler, directory=str(server.SITE)))
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        self.url = f'http://127.0.0.1:{self.httpd.server_port}'

    def tearDown(self):
        self.httpd.shutdown()
        self.httpd.server_close()

    def test_serves_site(self):
        code, _, body = get(self.url + '/')
        self.assertEqual(code, 200)
        self.assertIn('Домашний хаб'.encode(), body)
        self.assertEqual(get(self.url + '/manifest.webmanifest')[1], 'application/manifest+json')
        self.assertEqual(get(self.url + '/sw.js')[0], 200)

    def test_hides_service_files(self):
        for path in ('/.git/config', '/.git/HEAD', '/a/.hidden', '/../NewrRep/.git/config'):
            self.assertEqual(get(self.url + path)[0], 404, path)


if __name__ == '__main__':
    unittest.main()
