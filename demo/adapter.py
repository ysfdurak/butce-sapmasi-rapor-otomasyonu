"""Tarayıcı demosu köprüsü.

Masaüstü sürümündeki app.Handler yönlendirmesini olduğu gibi kullanır; yalnız
HTTP soketi yerine tarayıcıdaki Web Worker'dan gelen istekleri işler.
Hesap kodu (model.py, butce.py, analiz.py ...) değiştirilmeden çalışır.
"""

import io, json, os, sys

sys.path.insert(0, "/app")
os.chdir("/app")

import bootstrap  # noqa: E402

bootstrap.dependencies()

import app as A  # noqa: E402


class SyncExecutor:
    """Pyodide iş parçacığı açamaz; işler istek içinde sırayla çalışır."""

    def submit(self, fn):
        fn()

    def shutdown(self, wait=True):
        pass


APP = A.Application(state_dir="/state")
APP.executor = SyncExecutor()
HOST = "demo.local"


class _Server:
    application = APP
    origin = "http://" + HOST

    def shutdown(self):
        pass


class _Request(A.Handler):
    def __init__(self, method, path, body):
        self.server = _Server()
        self.command = method
        self.path = path
        self.headers = {
            "Host": HOST,
            "Cookie": "butce_session=" + APP.token,
            "Content-Type": "application/json",
            "Content-Length": str(len(body)),
        }
        self.rfile = io.BytesIO(body)
        self.result = (500, "application/json", b'{"error":"Yanit yok"}', {})

    def send(self, code, body=b"", kind="application/json; charset=utf-8", headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False, default=A.serial, allow_nan=False).encode()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.result = (code, kind, bytes(body), dict(headers or {}))


def handle(method, path, body):
    request = _Request(method, path, bytes(body or b""))
    if method == "GET":
        request.do_GET()
    else:
        request.do_POST()
    return request.result
