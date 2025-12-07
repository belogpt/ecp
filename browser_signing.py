"""Вспомогательный сервер для подписания через браузер с плагином CryptoPro."""
from __future__ import annotations

import base64
import json
import logging
import os
import secrets
import socket
import threading
from dataclasses import dataclass
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Optional
from urllib.parse import parse_qs, urlparse

logger = logging.getLogger(__name__)


@dataclass
class BrowserSigningResult:
    signature: bytes
    message: str = ""


class BrowserSigningError(RuntimeError):
    pass


class _BrowserSigningHandler(BaseHTTPRequestHandler):
    server_version = "CryptoProBrowserSigner/1.0"
    session: "BrowserSigningSession"  # type: ignore[assignment]

    def log_message(self, fmt, *args):  # pragma: no cover - только для отладки
        logger.debug("BrowserSigningServer: " + fmt, *args)

    def _reject(self, status=HTTPStatus.FORBIDDEN, message: str = "Forbidden"):
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(message.encode("utf-8"))

    def _check_client(self) -> bool:
        host = self.client_address[0]
        if host not in {"127.0.0.1", "::1"}:
            self._reject(HTTPStatus.FORBIDDEN, "Localhost only")
            return False
        return True

    def do_GET(self):
        if not self._check_client():
            return

        parsed = urlparse(self.path)
        if parsed.path != "/":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        params = parse_qs(parsed.query)
        nonce = (params.get("nonce") or [None])[0]
        if nonce != self.session.nonce:
            self._reject(HTTPStatus.FORBIDDEN, "Invalid nonce")
            return

        body = self.session.render_page()
        data = body.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        if not self._check_client():
            return

        parsed = urlparse(self.path)
        if parsed.path != "/result":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            self._reject(HTTPStatus.BAD_REQUEST, "Invalid JSON")
            return

        if payload.get("nonce") != self.session.nonce:
            self._reject(HTTPStatus.FORBIDDEN, "Invalid nonce")
            return

        status = payload.get("status")
        if status != "ok":
            error_message = payload.get("error") or "Неизвестная ошибка браузерной подписи"
            self.session.set_error(error_message)
            self._respond_ok()
            return

        signature_b64 = payload.get("signature")
        if not signature_b64:
            self._reject(HTTPStatus.BAD_REQUEST, "No signature")
            return
        try:
            signature = base64.b64decode(signature_b64)
        except Exception:
            self._reject(HTTPStatus.BAD_REQUEST, "Bad signature encoding")
            return

        self.session.set_result(BrowserSigningResult(signature=signature, message="Плагин вернул подпись"))
        self._respond_ok()

    def _respond_ok(self):
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(b"{\"status\":\"accepted\"}")


class BrowserSigningSession:
    """Организует цикл ожидания подписи через браузерный плагин CryptoPro."""

    def __init__(self, pdf_path: str):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")
        self.pdf_path = pdf_path
        self._server: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._result: Optional[BrowserSigningResult] = None
        self._error: Optional[str] = None
        self._event = threading.Event()
        self.nonce = secrets.token_urlsafe(24)
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
        self._pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")
        self._port = None

    def start(self):
        if self._server:
            return
        self._port = self._find_free_port()
        handler = self._build_handler()
        self._server = ThreadingHTTPServer(("127.0.0.1", self._port), handler)
        self._server.session = self  # type: ignore[attr-defined]
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()
        logger.info("Браузерный сервер подписи запущен на 127.0.0.1:%s", self._port)

    def stop(self):
        if self._server:
            self._server.shutdown()
            self._server.server_close()
            self._server = None
        if self._thread:
            self._thread.join(timeout=2.0)
            self._thread = None

    def url(self) -> str:
        if self._port is None:
            raise BrowserSigningError("Сервер ещё не запущен")
        return f"http://127.0.0.1:{self._port}/?nonce={self.nonce}"

    def wait(self, timeout: float = 180.0) -> BrowserSigningResult:
        finished = self._event.wait(timeout=timeout)
        if not finished:
            raise BrowserSigningError("Не получили ответ из браузера за отведённое время")
        if self._error:
            raise BrowserSigningError(self._error)
        if not self._result:
            raise BrowserSigningError("Ответ из браузера пуст")
        return self._result

    def is_finished(self) -> bool:
        return self._event.is_set()

    def set_error(self, message: str):
        logger.error("Ошибка при подписи через браузер: %s", message)
        self._error = message
        self._event.set()

    def set_result(self, result: BrowserSigningResult):
        self._result = result
        self._event.set()

    def render_page(self) -> str:
        return f"""
<!doctype html>
<html lang=\"ru\">
<head>
  <meta charset=\"UTF-8\" />
  <title>Подпись PDF через CryptoPro</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 24px; }}
    .panel {{ max-width: 760px; margin: 0 auto; padding: 16px; border: 1px solid #d0d0d0; border-radius: 8px; }}
    .log {{ background: #f7f7f7; border: 1px solid #d0d0d0; padding: 12px; border-radius: 6px; height: 200px; overflow: auto; white-space: pre-wrap; }}
    button {{ padding: 10px 16px; font-size: 14px; }}
    .error {{ color: #a40000; }}
  </style>
</head>
<body>
  <div class=\"panel\">
    <h2>Подпись PDF через браузерный плагин CryptoPro</h2>
    <p>Файл: <b>{os.path.basename(self.pdf_path)}</b></p>
    <p>Плагин запросит выбор сертификата и ввод PIN/пароля при необходимости. После успешной подписи окно можно закрыть.</p>
    <button id=\"startBtn\">Выбрать сертификат и подписать</button>
    <div id=\"status\" class=\"log\"></div>
    <div id=\"error\" class=\"error\"></div>
  </div>
  <script>
    const nonce = {json.dumps(self.nonce)};
    const postUrl = '/result';
    const pdfBase64 = {json.dumps(self._pdf_b64)};

    function log(msg) {{
      const box = document.getElementById('status');
      box.textContent += msg + '\n';
      box.scrollTop = box.scrollHeight;
    }}
    function showError(msg) {{
      document.getElementById('error').textContent = msg;
      log('Ошибка: ' + msg);
    }}

    async function sendResult(payload) {{
      try {{
        await fetch(postUrl, {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify(payload),
        }});
        log('Результат отправлен приложению.');
      }} catch (e) {{
        showError('Не удалось отправить результат: ' + e);
      }}
    }}

    async function sign() {{
      document.getElementById('error').textContent = '';
      log('Проверяем наличие плагина CryptoPro...');
      if (!window.cadesplugin) {{
        showError('Плагин CryptoPro не найден в браузере.');
        await sendResult({{nonce, status: 'error', error: 'Плагин CryptoPro не найден'}});
        return;
      }}
      try {{
        await window.cadesplugin;
        const plugin = window.cadesplugin;
        log('Открываем хранилище сертификатов...');
        const store = await plugin.CreateObjectAsync('CAdESCOM.Store');
        await store.Open();
        const certs = await store.Certificates;
        const selected = await certs.Select();
        const count = await selected.Count;
        if (!count) {{
          showError('Выбор сертификата отменён.');
          await sendResult({{nonce, status: 'error', error: 'Выбор сертификата отменён'}});
          return;
        }}
        const cert = await selected.Item(1);
        const signer = await plugin.CreateObjectAsync('CAdESCOM.CPSigner');
        await signer.propset_Certificate(cert);
        log('Сертификат выбран, формируем подпись...');
        const sd = await plugin.CreateObjectAsync('CAdESCOM.CadesSignedData');
        await sd.propset_ContentEncoding(plugin.CADESCOM_BASE64_TO_BINARY);
        await sd.propset_Content(pdfBase64);
        const signature = await sd.SignCades(signer, plugin.CADESCOM_CADES_BES, true);
        log('Подпись сформирована, отправляем обратно в приложение...');
        await sendResult({{nonce, status: 'ok', signature}});
        log('Готово. Теперь можно вернуться в приложение.');
      }} catch (e) {{
        console.error(e);
        const msg = (e && e.message) ? e.message : String(e);
        showError(msg);
        await sendResult({{nonce, status: 'error', error: msg}});
      }}
    }}

    document.getElementById('startBtn').addEventListener('click', () => {{
      sign();
    }});
  </script>
</body>
</html>
"""

    @staticmethod
    def _find_free_port() -> int:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("127.0.0.1", 0))
            return s.getsockname()[1]

    def _build_handler(self):
        session = self

        class Handler(_BrowserSigningHandler):
            pass

        Handler.session = session  # type: ignore[assignment]
        return Handler

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb):
        self.stop()


__all__ = [
    "BrowserSigningSession",
    "BrowserSigningError",
    "BrowserSigningResult",
]
