import base64
import datetime
import logging
import os
import sys
import shutil
import subprocess
from dataclasses import dataclass
from typing import List, Optional

try:
    import win32com.client  # type: ignore
    import pywintypes  # type: ignore
except Exception:  # pragma: no cover
    win32com = None
    pywintypes = None

logger = logging.getLogger(__name__)

# -------------------------------
# CAPICOM / CAdESCOM constants
# -------------------------------

CAPICOM_LOCAL_MACHINE_STORE = 1
CAPICOM_CURRENT_USER_STORE = 2
CAPICOM_MY_STORE = "My"

CAPICOM_STORE_OPEN_READ_ONLY = 0
CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = 2

CAPICOM_CERTIFICATE_FIND_SHA1_HASH = 0

CADESCOM_CADES_BES = 1

CADESCOM_ENCODE_BASE64 = 0
CADESCOM_ENCODE_BINARY = 1

# For binary content (PDF):
CADESCOM_BASE64_TO_BINARY = 1


class SignerCadescomError(RuntimeError):
    """Исключение с человекопонятным текстом для UI."""


@dataclass
class CertificateSummary:
    subject: str
    issuer: str
    not_before: datetime.datetime
    not_after: datetime.datetime
    thumbprint: str
    has_private_key: bool
    is_valid: bool

    @property
    def common_name(self) -> str:
        parts = [p.strip() for p in self.subject.split(",")]
        for part in parts:
            if part.upper().startswith("CN="):
                return part.split("=", 1)[1].strip()
        return self.subject


# -------------------------------
# Internal helpers
# -------------------------------

def _ensure_com_available():
    if sys.platform != "win32":
        raise SignerCadescomError("Подпись через CAdESCOM доступна только в Windows.")
    if win32com is None:
        raise SignerCadescomError(
            "Не установлен модуль pywin32. Установите pywin32 и убедитесь в доступности CAdESCOM."
        )


def _dispatch(prog_id: str):
    """Создаёт COM-объект по ProgID.

    Предпочитаем dynamic.Dispatch, чтобы не зависеть от gen_py/makepy-кэша,
    который может возвращать неправильные интерфейсы.
    """
    try:
        from win32com.client import dynamic  # type: ignore
        return dynamic.Dispatch(prog_id)
    except Exception:
        return win32com.client.Dispatch(prog_id)


def _open_store(location: int, open_mode: int = CAPICOM_STORE_OPEN_READ_ONLY):
    """Открывает хранилище сертификатов Windows с безопасным fallback."""
    last_exc: Optional[Exception] = None

    for pid in ("CAdESCOM.Store", "CAdESCOM.Store.1", "CAPICOM.Store", "CAPICOM.Store.1"):
        try:
            store = _dispatch(pid)
            store.Open(location, CAPICOM_MY_STORE, open_mode)
            return store
        except AttributeError as exc:
            last_exc = exc
            continue
        except Exception as exc:  # pragma: no cover
            last_exc = exc
            continue

    raise SignerCadescomError(
        "Не удалось открыть хранилище сертификатов через CAdESCOM/CAPICOM."
    ) from last_exc


def _safe_is_valid(cert) -> bool:
    """Мягкая проверка валидности ТОЛЬКО по датам.

    Важно: НЕ вызываем cert.IsValid(), чтобы не провоцировать ERROR_MORE_DATA.
    """
    try:
        now = datetime.datetime.now(datetime.timezone.utc)

        not_before = cert.ValidFromDate
        not_after = cert.ValidToDate

        nb = not_before if getattr(not_before, "tzinfo", None) else not_before.replace(
            tzinfo=datetime.timezone.utc
        )
        na = not_after if getattr(not_after, "tzinfo", None) else not_after.replace(
            tzinfo=datetime.timezone.utc
        )

        return nb <= now <= na
    except Exception:
        return True


def _collect_store(location: int) -> List[CertificateSummary]:
    certificates: List[CertificateSummary] = []
    store = _open_store(location, CAPICOM_STORE_OPEN_READ_ONLY)

    try:
        for cert in list(store.Certificates):
            try:
                thumbprint = str(cert.Thumbprint).replace(" ", "")
                has_private_key = bool(getattr(cert, "HasPrivateKey", False))
                is_valid = _safe_is_valid(cert)

                certificates.append(
                    CertificateSummary(
                        subject=str(cert.SubjectName),
                        issuer=str(cert.IssuerName),
                        not_before=cert.ValidFromDate,
                        not_after=cert.ValidToDate,
                        thumbprint=thumbprint,
                        has_private_key=has_private_key,
                        is_valid=is_valid,
                    )
                )
            except Exception:  # pragma: no cover
                logger.exception("Ошибка при разборе сертификата из хранилища")
                continue
    finally:
        try:
            store.Close()
        except Exception:
            pass

    return certificates


# -------------------------------
# Public API
# -------------------------------

def list_certificates() -> List[CertificateSummary]:
    """Список сертификатов из CurrentUser\My и LocalMachine\My."""
    _ensure_com_available()

    certificates: List[CertificateSummary] = []
    for location in (CAPICOM_CURRENT_USER_STORE, CAPICOM_LOCAL_MACHINE_STORE):
        try:
            certificates.extend(_collect_store(location))
        except SignerCadescomError:
            continue

    certificates.sort(
        key=lambda c: (
            not c.has_private_key,
            not c.is_valid,
            c.not_after,
        )
    )
    return certificates


def _find_certificate(store, thumbprint: str):
    try:
        found = store.Certificates.Find(CAPICOM_CERTIFICATE_FIND_SHA1_HASH, thumbprint)
        if getattr(found, "Count", 0) > 0:
            return found.Item(1)
    except Exception:
        return None
    return None


def _select_certificate(store, thumbprint: Optional[str]):
    if thumbprint:
        cert = _find_certificate(store, thumbprint)
        if not cert:
            raise SignerCadescomError("Сертификат с указанным отпечатком не найден")
        return cert

    summaries = list_certificates()
    for s in summaries:
        if s.has_private_key and s.is_valid:
            cert = _find_certificate(store, s.thumbprint)
            if cert:
                return cert

    if summaries:
        cert = _find_certificate(store, summaries[0].thumbprint)
        if cert:
            return cert

    raise SignerCadescomError("Не найден подходящий сертификат в хранилище Windows")


def _encode_signature(signature, encoding: str) -> bytes:
    if encoding == "base64":
        return signature.encode("utf-8") if isinstance(signature, str) else bytes(signature)

    # der
    if isinstance(signature, str):
        try:
            return base64.b64decode(signature)
        except Exception:
            return signature.encode("utf-8")
    return bytes(signature)


def _com_error_to_message(exc) -> str:
    hres = getattr(exc, "hresult", None)
    exinfo = getattr(exc, "excepinfo", None)

    source = None
    desc = None
    scode = None

    if exinfo and isinstance(exinfo, (list, tuple)):
        # excepinfo: (wCode, source, description, helpfile, helpid, scode)
        if len(exinfo) > 1:
            source = exinfo[1]
        if len(exinfo) > 2:
            desc = exinfo[2]
        if len(exinfo) > 5:
            scode = exinfo[5]

    parts = []
    if hres is not None:
        parts.append(f"hresult={hex(hres & 0xFFFFFFFF)}")
    if scode is not None:
        parts.append(f"scode={hex(scode & 0xFFFFFFFF)}")
    if source:
        parts.append(str(source))
    if desc:
        parts.append(str(desc))

    blob = " | ".join(parts) if parts else str(exc)
    lower = blob.lower()

    if "nte_bad_keyset" in lower or "0x80090016" in lower:
        return f"Носитель ключа не найден. Подключите флешку/токен. ({blob})"
    if "0x8009000d" in lower:
        return f"Нет доступа к закрытому ключу. Проверьте ключ и PIN. ({blob})"
    if "scard" in lower:
        return f"Ошибка токена/смарт-карты. ({blob})"
    if "license" in lower or "лиценз" in lower:
        return f"Проблема лицензии CryptoPro CSP. ({blob})"

    return f"Ошибка подписи через CAdESCOM: {blob}"


# -------------------------------
# cryptcp fallback
# -------------------------------

def _find_cryptcp() -> Optional[str]:
    path = shutil.which("cryptcp")
    if path:
        return path

    candidates = [
        r"C:\Program Files\Crypto Pro\CSP\cryptcp.exe",
        r"C:\Program Files (x86)\Crypto Pro\CSP\cryptcp.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def _sign_file_cryptcp(
    input_path: str,
    output_path: Optional[str],
    thumbprint: Optional[str],
    detached: bool,
    encoding: str,
) -> str:
    cryptcp = _find_cryptcp()
    if not cryptcp:
        raise SignerCadescomError("cryptcp не найден в системе для резервного подписания.")

    if not detached:
        raise SignerCadescomError(
            "Резервное подписывание через cryptcp сейчас поддерживает только отсоединённую подпись."
        )

    # Попытка 1: если есть thumbprint и задан output_path — попробуем явный вывод.
    if thumbprint and output_path:
        cmd = [cryptcp, "-sign", "-detached", "-thumbprint", thumbprint]
        if encoding == "der":
            cmd.append("-der")
        cmd += [input_path, output_path]

        logger.info("Пробуем cryptcp fallback (thumbprint+out): %s", " ".join(cmd))
        p = subprocess.run(cmd, capture_output=True, text=True)

        if p.returncode == 0 and os.path.exists(output_path):
            return output_path

    # Попытка 2: режим с автосозданием <file>.sgn рядом c файлом.
    # По описаниям использования cryptcp для отсоединённой подписи.
    cmd = [cryptcp, "-sign", "-detached", "-q"]
    if encoding == "der":
        cmd.append("-der")
    if thumbprint:
        # попробуем передать thumbprint и тут
        cmd += ["-thumbprint", thumbprint]
    cmd += [input_path]

    logger.info("Пробуем cryptcp fallback (auto-out): %s", " ".join(cmd))
    p = subprocess.run(cmd, capture_output=True, text=True)

    expected = input_path + ".sgn"
    if p.returncode == 0 and os.path.exists(expected):
        if output_path and os.path.abspath(output_path) != os.path.abspath(expected):
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
                os.replace(expected, output_path)
                return output_path
            except Exception:
                return expected
        return expected

    stderr = (p.stderr or "").strip()
    stdout = (p.stdout or "").strip()
    tail = (stderr or stdout)[:500]

    raise SignerCadescomError(
        "Не удалось создать подпись через cryptcp."
        + (f" Сообщение cryptcp: {tail}" if tail else "")
    )


# -------------------------------
# Signing
# -------------------------------

def sign_file(
    input_path: str,
    output_path: Optional[str] = None,
    thumbprint: Optional[str] = None,
    detached: bool = True,
    encoding: str = "base64",
) -> str:
    """Подписывает файл через COM-интерфейс CAdESCOM.

    В случае COM-ошибок пытается использовать cryptcp как резервный путь.
    """
    _ensure_com_available()

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Файл для подписи не найден: {input_path}")

    if encoding not in ("base64", "der"):
        raise ValueError("encoding должен быть 'base64' или 'der'")

    store = _open_store(CAPICOM_CURRENT_USER_STORE, CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED)

    try:
        certificate = _select_certificate(store, thumbprint)

        if not getattr(certificate, "HasPrivateKey", False):
            raise SignerCadescomError("Сертификат без доступа к закрытому ключу")

        signer = _dispatch("CAdESCOM.CPSigner")
        signer.Certificate = certificate

        signed_data = _dispatch("CAdESCOM.CadesSignedData")

        with open(input_path, "rb") as f:
            content_bytes = f.read()

        signed_data.ContentEncoding = CADESCOM_BASE64_TO_BINARY
        signed_data.Content = base64.b64encode(content_bytes).decode("ascii")

        encoding_type = (
            CADESCOM_ENCODE_BASE64 if encoding == "base64" else CADESCOM_ENCODE_BINARY
        )

        # Двухшаговый вызов SignCades:
        # 1) без 4-го параметра (опциональный)
        # 2) с encoding_type
        try:
            raw_signature = signed_data.SignCades(
                signer, CADESCOM_CADES_BES, detached
            )
        except Exception:
            raw_signature = signed_data.SignCades(
                signer, CADESCOM_CADES_BES, detached, encoding_type
            )

    except Exception as exc:
        # COM -> cryptcp fallback
        if pywintypes and isinstance(exc, pywintypes.com_error):
            msg = _com_error_to_message(exc)
            logger.exception("COM ошибка подписи: %s", msg)

            try:
                return _sign_file_cryptcp(
                    input_path, output_path, thumbprint, detached, encoding
                )
            except Exception:
                raise SignerCadescomError(msg) from exc

        message = str(exc)
        if "Class not registered" in message:
            raise SignerCadescomError(
                "CADESCOM не зарегистрирован/несовпадение разрядности"
            ) from exc

        raise

    finally:
        try:
            store.Close()
        except Exception:
            pass

    signature_bytes = _encode_signature(raw_signature, encoding)

    if not output_path:
        output_path = f"{input_path}.p7s"

    with open(output_path, "wb") as f:
        f.write(signature_bytes)

    logger.info("Подпись создана через CAdESCOM: %s", output_path)
    return output_path
