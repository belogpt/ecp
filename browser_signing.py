"""Вспомогательный сервер для подписания через браузер с плагином CryptoPro."""
from __future__ import annotations

import base64
import json
import logging
import mimetypes
import os
import secrets
import socket
import threading
from dataclasses import dataclass
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Optional
from urllib.parse import parse_qs, unquote, urlparse

from paths import get_resource_path

PLUGIN_SCRIPT_SOURCES = [
    "https://www.cryptopro.ru/sites/default/files/products/cades/cadesplugin_api.js",
    "chrome-extension://iifchhfnnmpdbibifmljnfjhpififfog/nmcades_plugin_api.js",
    "chrome-extension://epiejncknlhcgcanmnmnjnmghjkpgkdd/nmcades_plugin_api.js",
]

logger = logging.getLogger(__name__)


class _SessionLogHandler(logging.Handler):
    def __init__(self, session: "BrowserSigningSession"):
        super().__init__()
        self.session = session
        self.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
        )

    def emit(self, record):  # pragma: no cover - используется только в рантайме
        try:
            message = self.format(record)
        except Exception:
            return
        self.session._append_log(message)


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
        logger.warning(
            "Отклонён запрос от %s: %s (%s)",
            self.client_address[0],
            message,
            status,
        )
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
        logger.debug("GET %s от %s", parsed.path, self.client_address[0])
        if parsed.path == "/logs":
            self._handle_logs(parsed)
            return
        if parsed.path == "/config":
            self._handle_config(parsed)
            return

        if parsed.path in {"/", "/index.html"}:
            if not self._validate_nonce(parsed):
                return
            self._serve_static_file("index.html")
            return

        if self._serve_static_file(parsed.path.lstrip("/")):
            return

        self.send_error(HTTPStatus.NOT_FOUND)

    def _handle_logs(self, parsed):
        params = parse_qs(parsed.query)
        nonce = (params.get("nonce") or [None])[0]
        if nonce != self.session.nonce:
            logger.warning(
                "Запрос логов с неверным nonce от %s: %s", self.client_address[0], nonce
            )
            self._reject(HTTPStatus.FORBIDDEN, "Invalid nonce")
            return

        if not self.session.log_to_page:
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        after_raw = (params.get("after") or ["0"])[0]
        try:
            after = int(after_raw)
        except ValueError:
            logger.debug("Некорректный параметр after в запросе логов: %s", after_raw)
            after = 0

        last_id, items = self.session.get_logs_since(after)
        payload = json.dumps({"last": last_id, "items": items}, ensure_ascii=False)
        data = payload.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _validate_nonce(self, parsed) -> bool:
        params = parse_qs(parsed.query)
        nonce = (params.get("nonce") or [None])[0]
        if nonce != self.session.nonce:
            logger.warning(
                "Запрос с неверным nonce от %s: %s", self.client_address[0], nonce
            )
            self._reject(HTTPStatus.FORBIDDEN, "Invalid nonce")
            return False
        return True

    def _serve_static_file(self, relative_path: str) -> bool:
        clean_path = os.path.normpath(unquote(relative_path)).lstrip(os.sep)
        if not clean_path:
            clean_path = "index.html"
        full_path = (self.session.static_root / clean_path).resolve()
        try:
            full_path.relative_to(self.session.static_root)
        except ValueError:
            self._reject(HTTPStatus.FORBIDDEN, "Path is outside static root")
            return True

        if not full_path.exists() or not full_path.is_file():
            return False

        content_type, _ = mimetypes.guess_type(full_path.name)
        content_type = content_type or "application/octet-stream"
        if content_type.startswith("text/") or content_type in {"application/javascript", "application/json"}:
            content_type = f"{content_type}; charset=utf-8"
        data = full_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)
        return True

    def _handle_config(self, parsed):
        if not self._validate_nonce(parsed):
            return

        initial_last_id, initial_logs = self.session.get_logs_since(0)
        payload = json.dumps(
            {
                "nonce": self.session.nonce,
                "pdfName": os.path.basename(self.session.pdf_path),
                "pdfBase64": self.session.pdf_b64,
                "logEnabled": self.session.log_to_page,
                "initialLogs": initial_logs,
                "lastLogId": initial_last_id,
                "pluginScriptSources": self.session.plugin_script_sources,
            },
            ensure_ascii=False,
        )
        data = payload.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):
        if not self._check_client():
            return

        parsed = urlparse(self.path)
        logger.debug("POST %s от %s", parsed.path, self.client_address[0])
        if parsed.path != "/result":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        try:
            payload = json.loads(raw.decode("utf-8"))
        except Exception:
            logger.exception("Не удалось разобрать JSON из браузера")
            self._reject(HTTPStatus.BAD_REQUEST, "Invalid JSON")
            return

        if payload.get("nonce") != self.session.nonce:
            logger.warning(
                "POST с неверным nonce от %s: %s", self.client_address[0], payload
            )
            self._reject(HTTPStatus.FORBIDDEN, "Invalid nonce")
            return

        status = payload.get("status")
        if status != "ok":
            error_message = payload.get("error") or "Неизвестная ошибка браузерной подписи"
            logger.error(
                "Браузер сообщил об ошибке: %s (payload: %s)",
                error_message,
                {k: v for k, v in payload.items() if k != "signature"},
            )
            self.session.set_error(error_message)
            self._respond_ok()
            return

        signature_b64 = payload.get("signature")
        if not signature_b64:
            logger.error("Ответ из браузера без подписи: %s", payload)
            self._reject(HTTPStatus.BAD_REQUEST, "No signature")
            return
        try:
            signature = base64.b64decode(signature_b64)
        except Exception:
            logger.exception("Не удалось декодировать подпись из base64")
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

    def __init__(self, pdf_path: str, log_to_page: bool = True):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF не найден: {pdf_path}")
        self.pdf_path = pdf_path
        self._server: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None
        self._result: Optional[BrowserSigningResult] = None
        self._error: Optional[str] = None
        self._event = threading.Event()
        self._log_lock = threading.Lock()
        self._log_entries: list[str] = []
        self._log_handler: Optional[_SessionLogHandler] = None
        self.log_to_page = log_to_page
        self.nonce = secrets.token_urlsafe(24)
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
        self._pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")
        self.static_root = Path(get_resource_path("web/signing")).resolve()
        if not self.static_root.exists():
            raise FileNotFoundError(
                f"Каталог статических файлов для браузерной подписи не найден: {self.static_root}"
            )
        self.plugin_script_sources = list(PLUGIN_SCRIPT_SOURCES)
        self._port = None

    def start(self):
        if self._server:
            return
        self._port = self._find_free_port()
        handler = self._build_handler()
        self._server = ThreadingHTTPServer(("127.0.0.1", self._port), handler)
        self._server.session = self  # type: ignore[attr-defined]
        if self.log_to_page and not self._log_handler:
            handler = _SessionLogHandler(self)
            handler.setLevel(logging.DEBUG)
            self._log_handler = handler
            logging.getLogger().addHandler(handler)
            # Явно фиксируем запуск в логе браузера, даже если уровень логгера повысился.
            self._append_log(
                f"Браузерный сервер подписи запущен на 127.0.0.1:{self._port}"
            )
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()
        logger.info("Браузерный сервер подписи запущен на 127.0.0.1:%s", self._port)

    def stop(self):
        if self._server:
            self._server.shutdown()
            self._server.server_close()
            logger.info("Браузерный сервер подписи остановлен")
            self._server = None
        if self._thread:
            self._thread.join(timeout=2.0)
            self._thread = None
        if self._log_handler:
            logging.getLogger().removeHandler(self._log_handler)
            self._log_handler = None

    def url(self) -> str:
        if self._port is None:
            raise BrowserSigningError("Сервер ещё не запущен")
        return f"http://127.0.0.1:{self._port}/?nonce={self.nonce}"

    @property
    def pdf_b64(self) -> str:
        return self._pdf_b64

    def wait(self, timeout: float = 180.0) -> BrowserSigningResult:
        finished = self._event.wait(timeout=timeout)
        if not finished:
            logger.error(
                "Ожидание ответа из браузера истекло через %s секунд", timeout
            )
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
        logger.info("Успешная подпись через браузер получена: %s", result.message)
        self._result = result
        self._event.set()

    def _append_log(self, message: str):
        with self._log_lock:
            self._log_entries.append(message)
            if len(self._log_entries) > 500:
                self._log_entries = self._log_entries[-500:]

    def get_logs_since(self, last_id: int) -> tuple[int, list[str]]:
        with self._log_lock:
            next_id = max(0, last_id)
            items = self._log_entries[next_id:]
            return len(self._log_entries), items

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
