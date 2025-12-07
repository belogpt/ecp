import base64
import datetime
import logging
import os
import sys
from dataclasses import dataclass
from typing import List, Optional

try:
    import win32com.client  # type: ignore
    import pywintypes  # type: ignore
except Exception:  # pragma: no cover - win32com может отсутствовать в окружении CI
    win32com = None
    pywintypes = None

if win32com is not None:  # pragma: no cover - зависит от среды Windows
    try:
        ensure_dispatch = win32com.client.gencache.EnsureDispatch
        com_constants = win32com.client.constants
    except Exception:
        ensure_dispatch = win32com.client.Dispatch
        com_constants = None
else:
    ensure_dispatch = None
    com_constants = None

logger = logging.getLogger(__name__)


def _const(name: str, fallback):  # pragma: no cover - простая функция
    if com_constants is None:
        return fallback
    try:
        return getattr(com_constants, name)
    except Exception:
        return fallback


CAPICOM_LOCAL_MACHINE_STORE = _const("CAPICOM_LOCAL_MACHINE_STORE", 1)
CAPICOM_CURRENT_USER_STORE = _const("CAPICOM_CURRENT_USER_STORE", 2)
CAPICOM_MY_STORE = _const("CAPICOM_MY_STORE", "My")
CAPICOM_STORE_OPEN_READ_ONLY = _const("CAPICOM_STORE_OPEN_READ_ONLY", 0)
CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED = _const("CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED", 2)
CAPICOM_CERTIFICATE_FIND_SHA1_HASH = _const("CAPICOM_CERTIFICATE_FIND_SHA1_HASH", 0)
CADESCOM_CADES_BES = _const("CADESCOM_CADES_BES", 1)
CADESCOM_ENCODE_BASE64 = _const("CADESCOM_ENCODE_BASE64", 0)
CADESCOM_ENCODE_BINARY = _const("CADESCOM_ENCODE_BINARY", 1)
CADESCOM_BASE64_TO_BINARY = _const("CADESCOM_BASE64_TO_BINARY", 1)


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
    is_valid: Optional[bool]

    @property
    def common_name(self) -> str:
        """Выделяет CN из Subject, если возможно."""
        parts = [p.strip() for p in self.subject.split(",")]
        for part in parts:
            if part.upper().startswith("CN="):
                return part.split("=", 1)[1].strip()
        return self.subject


def _ensure_com_available():
    if sys.platform != "win32":
        raise SignerCadescomError(
            "Подпись через CAdESCOM доступна только в Windows."
        )
    if win32com is None:
        raise SignerCadescomError(
            "Не установлен модуль pywin32. Установите pywin32 и убедитесь в доступности CAdESCOM."
        )


def _dispatch(prog_id: str):
    try:
        obj = ensure_dispatch(prog_id)
        # Иногда EnsureDispatch возвращает IEventSource без нужных методов
        # (например, вместо CAdESCOM.Store), тогда пробуем обычный Dispatch.
        if prog_id == "CAdESCOM.Store" and not hasattr(obj, "Open"):
            logger.warning(
                "EnsureDispatch вернул %s без метода Open, пробуем Dispatch", obj
            )
            obj = win32com.client.Dispatch(prog_id)
        return obj
    except Exception as exc:  # pragma: no cover - зависит от окружения Windows
        message = str(exc)
        if "Class not registered" in message:
            raise SignerCadescomError(
                "CADESCOM не зарегистрирован/несовпадение разрядности"
            ) from exc
        raise


def _open_store(location: int, open_mode: int):
    store = _dispatch("CAdESCOM.Store")
    store.Open(location, CAPICOM_MY_STORE, open_mode)
    return store


def _safe_is_valid(cert) -> bool:
    """Проверяет валидность сертификата без падения на ошибках COM."""

    try:
        validation = cert.IsValid()
        return bool(getattr(validation, "Result", False))
    except Exception:  # pragma: no cover - зависит от данных сертификатов
        thumbprint = str(getattr(cert, "Thumbprint", ""))
        logger.warning(
            "Не удалось проверить валидность сертификата %s", thumbprint, exc_info=True
        )
        return False


def _log_com_error(exc, message: str):
    if pywintypes is not None and isinstance(exc, pywintypes.com_error):
        logger.error(
            "%s (hresult=%s excepinfo=%s)",
            message,
            hex(getattr(exc, "hresult", 0)),
            getattr(exc, "excepinfo", None),
            exc_info=True,
        )
    else:
        logger.exception(message)


def _collect_store(location: int) -> List[CertificateSummary]:
    certificates: List[CertificateSummary] = []
    try:
        store = _open_store(location, CAPICOM_STORE_OPEN_READ_ONLY)
    except SignerCadescomError:
        raise
    except Exception as exc:  # pragma: no cover - зависит от ОС
        logger.exception("Не удалось открыть хранилище сертификатов")
        raise SignerCadescomError(
            "Не удалось открыть хранилище сертификатов Windows"
        ) from exc

    try:
        certs = store.Certificates
        now = datetime.datetime.now(datetime.timezone.utc)
        for i in range(1, certs.Count + 1):
            cert = certs.Item(i)
            try:
                thumbprint = cert.Thumbprint
                has_private_key = bool(getattr(cert, "HasPrivateKey", False))
                not_before = cert.ValidFromDate
                not_after = cert.ValidToDate

                not_before_dt = (
                    not_before
                    if not_before.tzinfo is not None
                    else not_before.replace(tzinfo=datetime.timezone.utc)
                )
                not_after_dt = (
                    not_after
                    if not_after.tzinfo is not None
                    else not_after.replace(tzinfo=datetime.timezone.utc)
                )
                try:
                    is_valid = not_before_dt <= now <= not_after_dt
                except Exception:
                    is_valid = None
                certificates.append(
                    CertificateSummary(
                        subject=str(cert.SubjectName),
                        issuer=str(cert.IssuerName),
                        not_before=cert.ValidFromDate,
                        not_after=cert.ValidToDate,
                        thumbprint=str(thumbprint).replace(" ", ""),
                        has_private_key=has_private_key,
                        is_valid=is_valid,
                    )
                )
            except Exception:  # pragma: no cover - зависит от данных сертификатов
                logger.exception("Ошибка при разборе сертификата из хранилища")
                continue
    finally:
        try:
            store.Close()
        except Exception:
            pass
    return certificates


def list_certificates() -> List[CertificateSummary]:
    """Возвращает список сертификатов из хранилищ CurrentUser\My и LocalMachine\My."""

    _ensure_com_available()
    certificates: List[CertificateSummary] = []

    for location in (CAPICOM_CURRENT_USER_STORE, CAPICOM_LOCAL_MACHINE_STORE):
        try:
            certificates.extend(_collect_store(location))
        except SignerCadescomError:
            raise
        except Exception:
            logger.exception("Ошибка чтения хранилища %s", location)
            continue

    now = datetime.datetime.now(datetime.timezone.utc)
    for cert in certificates:
        if cert.not_before.tzinfo is None:
            cert.not_before = cert.not_before.replace(tzinfo=datetime.timezone.utc)
        if cert.not_after.tzinfo is None:
            cert.not_after = cert.not_after.replace(tzinfo=datetime.timezone.utc)

    certificates.sort(
        key=lambda c: (
            not c.has_private_key,
            not c.is_valid,
            c.not_after,
        )
    )
    return certificates


def _find_certificate(store, thumbprint: str):
    normalized = thumbprint.replace(" ", "").upper()
    found = store.Certificates.Find(
        CAPICOM_CERTIFICATE_FIND_SHA1_HASH, normalized
    )
    return found.Item(1) if found.Count > 0 else None


def _select_certificate(store, thumbprint: Optional[str]):
    if thumbprint:
        cert = _find_certificate(store, thumbprint)
        if not cert:
            raise SignerCadescomError(
                f"Сертификат с отпечатком {thumbprint} не найден в хранилище"
            )
        if not getattr(cert, "HasPrivateKey", False):
            raise SignerCadescomError("Сертификат без доступа к закрытому ключу")
        return cert

    certificates = list_certificates()
    for cert_summary in certificates:
        if cert_summary.has_private_key and cert_summary.is_valid:
            cert = _find_certificate(store, cert_summary.thumbprint)
            if cert is not None:
                return cert
    if certificates:
        cert = _find_certificate(store, certificates[0].thumbprint)
        if cert is not None:
            return cert
    raise SignerCadescomError(
        "Не найден подходящий сертификат в хранилище Windows"
    )


def _encode_signature(signature, encoding: str) -> bytes:
    if encoding == "base64":
        return signature.encode("utf-8") if isinstance(signature, str) else bytes(signature)
    return bytes(signature)


def sign_file(
    input_path: str,
    output_path: Optional[str] = None,
    thumbprint: Optional[str] = None,
    detached: bool = True,
    encoding: str = "base64",
) -> str:
    """Подписывает файл через COM-интерфейс CAdESCOM."""

    _ensure_com_available()

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Файл для подписи не найден: {input_path}")

    if encoding not in ("base64", "der"):
        raise ValueError("encoding должен быть 'base64' или 'der'")

    store = None
    try:
        store = _open_store(
            CAPICOM_CURRENT_USER_STORE, CAPICOM_STORE_OPEN_MAXIMUM_ALLOWED
        )
    except SignerCadescomError:
        raise
    except Exception as exc:  # pragma: no cover - зависит от ОС
        logger.exception("Ошибка инициализации CAdESCOM.Store")
        raise SignerCadescomError(
            "Не удалось открыть хранилище сертификатов Windows"
        ) from exc

    try:
        certificate = _select_certificate(store, thumbprint)

        if not getattr(certificate, "HasPrivateKey", False):
            raise SignerCadescomError("Сертификат без доступа к закрытому ключу")

        try:
            signer = _dispatch("CAdESCOM.CPSigner")
            signer.Certificate = certificate
        except SignerCadescomError:
            raise
        except Exception as exc:  # pragma: no cover - зависит от окружения
            logger.exception("Ошибка подготовки подписанта")
            raise SignerCadescomError(
                "Не удалось подготовить подписанта CAdESCOM"
            ) from exc

        try:
            signed_data = _dispatch("CAdESCOM.CadesSignedData")
        except SignerCadescomError:
            raise
        except Exception as exc:  # pragma: no cover
            logger.exception("Ошибка создания CadesSignedData")
            raise SignerCadescomError(
                "Не удалось создать объект CadesSignedData"
            ) from exc

        with open(input_path, "rb") as f:
            content_bytes = f.read()

        signed_data.ContentEncoding = CADESCOM_BASE64_TO_BINARY
        signed_data.Content = base64.b64encode(content_bytes).decode("ascii")

        encoding_type = (
            CADESCOM_ENCODE_BASE64
            if encoding == "base64"
            else CADESCOM_ENCODE_BINARY
        )

        try:
            raw_signature = signed_data.SignCades(
                signer, CADESCOM_CADES_BES, detached, encoding_type
            )
        except SignerCadescomError:
            raise
        except pywintypes.com_error as exc:  # pragma: no cover - зависит от токена/сертификата
            message = str(exc)
            _log_com_error(exc, "Ошибка подписи через CAdESCOM (COM)")
            if "0x80090016" in message or "NTE_BAD_KEYSET" in message:
                raise SignerCadescomError("Требуется подключить носитель ключа") from exc
            if "0x8009000D" in message:
                raise SignerCadescomError(
                    "Сертификат без доступа к закрытому ключу"
                ) from exc
            if "Class not registered" in message:
                raise SignerCadescomError(
                    "CADESCOM не зарегистрирован/несовпадение разрядности"
                ) from exc
            raise SignerCadescomError(
                "Не удалось создать подпись через CAdESCOM"
            ) from exc
        except Exception as exc:  # pragma: no cover - зависит от токена/сертификата
            message = str(exc)
            logger.exception("Ошибка подписи через CAdESCOM")
            if "0x80090016" in message or "NTE_BAD_KEYSET" in message:
                raise SignerCadescomError("Требуется подключить носитель ключа") from exc
            if "0x8009000D" in message:
                raise SignerCadescomError("Сертификат без доступа к закрытому ключу") from exc
            if "Class not registered" in message:
                raise SignerCadescomError(
                    "CADESCOM не зарегистрирован/несовпадение разрядности"
                ) from exc
            raise SignerCadescomError("Не удалось создать подпись через CAdESCOM") from exc

        signature_bytes = _encode_signature(raw_signature, encoding)

        if not output_path:
            output_path = f"{input_path}.p7s"

        with open(output_path, "wb") as f:
            f.write(signature_bytes)

        logger.info("Подпись создана через CAdESCOM: %s", output_path)
        return output_path
    finally:
        if store is not None:
            try:
                store.Close()
            except Exception:
                pass
