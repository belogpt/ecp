"""
Модуль для работы с P7S-подписью и сертификатом.
Использует asn1crypto для разбора структуры CMS/PKCS#7.
Поддерживает DER, PEM и "голый" base64 в .p7s.
"""

from __future__ import annotations

import datetime
import hashlib
import os
import logging
from dataclasses import dataclass
from typing import Optional, Tuple

try:
    from asn1crypto import cms, x509  # type: ignore
except ImportError as e:  # pragma: no cover - depends on external lib
    raise RuntimeError(
        "Модуль 'asn1crypto' не установлен. Установите его командой:\n"
        "    pip install asn1crypto"
    ) from e

logger = logging.getLogger(__name__)


@dataclass
class CertificateInfo:
    serial_number: str = "не удалось определить"
    subject: str = "не удалось определить"  # здесь будет Common Name (ФИО)
    issuer: str = "не удалось определить"
    valid_from: str = "не удалось определить"
    valid_to: str = "не удалось определить"
    signing_time: str = "не удалось определить"  # ДД.ММ.ГГГГ ЧЧ:ММ
    status: str = "не удалось определить"


def _format_serial(serial: int) -> str:
    hex_str = f"{serial:X}"
    if len(hex_str) % 2:
        hex_str = "0" + hex_str
    return " ".join(hex_str[i:i + 2] for i in range(0, len(hex_str), 2))


def _format_dt(dt: Optional[datetime.datetime]) -> str:
    """Формат даты/времени для времени подписи: ДД.ММ.ГГГГ ЧЧ:ММ."""
    if not dt:
        return "не удалось определить"
    return dt.strftime("%d.%m.%Y %H:%M")


def _format_date(dt: Optional[datetime.datetime]) -> str:
    """Формат даты: ДД.ММ.ГГГГ."""
    if not dt:
        return "не удалось определить"
    return dt.strftime("%d.%m.%Y")


def _normalize_to_utc(dt: Optional[datetime.datetime]) -> Optional[datetime.datetime]:
    """Переводит datetime в UTC и делает его aware. Если tzinfo отсутствует — считаем, что это UTC."""
    if dt is None:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=datetime.timezone.utc)
    return dt.astimezone(datetime.timezone.utc)


def _load_certificate_from_cer(path: str) -> Optional["x509.Certificate"]:
    if not path or not os.path.exists(path):
        logger.debug("CER-файл не указан или не существует: %s", path)
        return None
    logger.debug("Пробуем загрузить сертификат из CER-файла: %s", path)
    with open(path, "rb") as f:
        data = f.read()
    if b"-----BEGIN" in data:
        logger.debug("CER-файл в PEM-формате, декодируем base64")
        pem_lines = []
        in_block = False
        for line in data.splitlines():
            if b"-----BEGIN" in line:
                in_block = True
                continue
            if b"-----END" in line:
                break
            if in_block:
                pem_lines.append(line.strip())
        import base64

        der = base64.b64decode(b"".join(pem_lines))
    else:
        logger.debug("CER-файл в DER-формате")
        der = data
    cert = x509.Certificate.load(der)
    logger.debug("Сертификат из CER успешно загружен")
    return cert


def _load_cms(p7s_path: str) -> "cms.SignedData":
    """
    Загружает CMS из P7S-файла.

    Поддерживает:
    - чистый DER (первый байт 0x30);
    - PEM с заголовком -----BEGIN PKCS7----- / -----BEGIN SIGNED DATA-----;
    - "голый" base64 без заголовков.
    """
    logger.debug("Читаем P7S-файл: %s", p7s_path)
    with open(p7s_path, "rb") as f:
        raw = f.read()
    logger.debug(
        "Размер P7S: %d байт, первые 16 байт: %s",
        len(raw),
        raw[:16].hex(" "),
    )

    der = raw

    if raw and raw[0] == 0x30:
        logger.debug("P7S выглядит как DER (первый байт 0x30)")
    else:
        if raw.startswith(b"-----BEGIN"):
            logger.debug("P7S в PEM-формате с заголовком, извлекаем base64")
            import base64

            lines = []
            in_block = False
            for line in raw.splitlines():
                if line.startswith(b"-----BEGIN"):
                    in_block = True
                    continue
                if line.startswith(b"-----END"):
                    break
                if in_block:
                    lines.append(line.strip())
            der = base64.b64decode(b"".join(lines))
            logger.debug("После декодирования PEM: %d байт DER", len(der))
        else:
            logger.debug(
                "P7S не похож на DER и не содержит BEGIN/END — пробуем как текстовый base64"
            )
            import base64

            try:
                cleaned = b"".join(
                    line.strip()
                    for line in raw.splitlines()
                    if line.strip() and not line.strip().startswith(b"#")
                )
                der = base64.b64decode(cleaned)
                logger.debug("После base64-декодирования: %d байт DER", len(der))
            except Exception:
                logger.exception(
                    "Не удалось декодировать base64 из P7S, оставляем как есть и пробуем как DER"
                )
                der = raw

    try:
        content_info = cms.ContentInfo.load(der)
    except Exception:
        logger.exception("Не удалось распарсить asn1crypto.cms.ContentInfo")
        raise

    ct = content_info["content_type"].native
    logger.debug("ContentInfo.content_type: %s", ct)
    if ct != "signed_data":
        raise ValueError(f"Файл подписи не содержит signedData (тип: {ct})")
    signed_data: cms.SignedData = content_info["content"]
    return signed_data


def _get_signer_info(signed_data: "cms.SignedData") -> "cms.SignerInfo":
    signer_infos = signed_data["signer_infos"]
    if len(signer_infos) == 0:
        raise ValueError("В подписи отсутствуют сведения о подписанте")
    logger.debug("Количество SignerInfo в подписи: %d", len(signer_infos))
    return signer_infos[0]


def _get_signing_time(signer_info: "cms.SignerInfo") -> Optional[datetime.datetime]:
    attrs = signer_info["signed_attrs"]
    for attr in attrs:
        if attr["type"].native == "signing_time":
            st = attr["values"][0].native
            logger.debug("Найден signingTime: %s", st)
            return st
    logger.debug("signingTime в подписи не найден")
    return None


def _get_message_digest_and_alg(
    signer_info: "cms.SignerInfo",
) -> Tuple[Optional[bytes], Optional[str]]:
    attrs = signer_info["signed_attrs"]
    msg_digest = None
    for attr in attrs:
        if attr["type"].native == "message_digest":
            msg_digest = attr["values"][0].native
            break
    digest_alg = signer_info["digest_algorithm"]["algorithm"].native  # e.g. 'sha256' или OID
    logger.debug(
        "messageDigest найден: %s, digest_algorithm: %s",
        "да" if msg_digest else "нет",
        digest_alg,
    )
    return msg_digest, str(digest_alg)


def _compute_digest(pdf_path: str, alg_name: str) -> Optional[bytes]:
    """
    Считает хэш файла pdf_path для алгоритма alg_name.

    Поддерживаются:
      - стандартные SHA*/MD5 через hashlib;
      - ГОСТ 34.11-2012 (Streebog) через gostcrypto (если установлен).
    """
    alg_norm = alg_name.lower()

    alg_map = {
        "sha1": "sha1",
        "sha224": "sha224",
        "sha256": "sha256",
        "sha384": "sha384",
        "sha512": "sha512",
        "md5": "md5",
    }

    if alg_norm in alg_map:
        py_alg = alg_map[alg_norm]
        logger.debug("Считаем хэш PDF (%s): %s", py_alg, pdf_path)
        h = hashlib.new(py_alg)
        with open(pdf_path, "rb") as f:
            while True:
                chunk = f.read(8192)
                if not chunk:
                    break
                h.update(chunk)
        digest = h.digest()
        logger.debug("Хэш PDF (%s): %s", py_alg, digest.hex(" "))
        return digest

    gost_map = {
        "1.2.643.7.1.1.2.2": "streebog256",
        "1.2.643.7.1.1.2.3": "streebog512",
        "id-tc26-gost3411-12-256": "streebog256",
        "id-tc26-gost3411-12-512": "streebog512",
        "gost3411-2012-256": "streebog256",
        "gost3411-2012-512": "streebog512",
        "gost3411_2012_256": "streebog256",
        "gost3411_2012_512": "streebog512",
    }

    if alg_norm in gost_map or alg_norm.startswith("1.2.643.7.1.1.2."):
        try:
            import gostcrypto.gosthash as gosthash  # type: ignore
        except Exception:
            logger.warning(
                "Алгоритм %s похож на ГОСТ 34.11-2012, но пакет 'gostcrypto' не установлен",
                alg_name,
            )
            return None

        name = gost_map.get(alg_norm, "streebog256")
        logger.debug("Считаем GOST-хэш PDF (%s) через gostcrypto: %s", name, pdf_path)
        h = gosthash.new(name)
        with open(pdf_path, "rb") as f:
            while True:
                chunk = f.read(8192)
                if not chunk:
                    break
                h.update(chunk)
        digest = h.digest()
        logger.debug("GOST-хэш PDF (%s): %s", name, digest.hex(" "))
        return digest

    logger.warning("Неизвестный алгоритм хеширования: %s", alg_name)
    return None


def _pick_cert_from_signed_data(
    signed_data: "cms.SignedData",
) -> Optional["x509.Certificate"]:
    certs = signed_data["certificates"]
    if not certs:
        logger.warning("В подписи нет вложенных сертификатов")
        return None
    first = certs[0]
    if isinstance(first.chosen, x509.Certificate):
        logger.debug("Используем сертификат из подписи (первый)")
        return first.chosen
    logger.warning("CertificateChoice не является x509.Certificate")
    return None


def _extract_common_name(cert: "x509.Certificate") -> str:
    """
    Достаёт только Common Name (ФИО) из subject.
    Если не удалось — возвращает human_friendly.
    """
    try:
        cn = None
        # subject — это x509.Name
        for rdn in cert.subject.chosen:  # type: ignore[attr-defined]
            for type_val in rdn:
                if type_val["type"].native == "common_name":
                    cn = type_val["value"].native
                    break
            if cn:
                break
        if cn:
            return str(cn)
    except Exception:
        logger.exception("Не удалось извлечь Common Name из subject")

    try:
        return cert.subject.human_friendly
    except Exception:
        return "не удалось определить"


def get_certificate_info(
    pdf_path: str, p7s_path: str, cer_path: Optional[str] = None
) -> CertificateInfo:
    """
    Основная функция: проверяет подпись и возвращает CertificateInfo.
    Все ошибки логируются в ep_viewer.log.
    """
    logger.info(
        "Начинаем проверку подписи: pdf=%s, p7s=%s, cer=%s",
        pdf_path,
        p7s_path,
        cer_path,
    )

    info = CertificateInfo()
    try:
        signed_data = _load_cms(p7s_path)
        signer_info = _get_signer_info(signed_data)

        msg_digest_attr, digest_alg = _get_message_digest_and_alg(signer_info)
        pdf_digest = None
        matches_document = False

        if msg_digest_attr and digest_alg:
            pdf_digest = _compute_digest(pdf_path, digest_alg)
            if pdf_digest is not None:
                matches_document = pdf_digest == msg_digest_attr
                logger.debug(
                    "Сравнение messageDigest: %s",
                    "совпадает" if matches_document else "НЕ совпадает",
                )
        else:
            logger.warning(
                "В подписи отсутствуют messageDigest или digestAlgorithm. "
                "Проверка соответствия документу невозможна."
            )

        signing_dt = _get_signing_time(signer_info)
        info.signing_time = _format_dt(signing_dt)

        cert = None
        if cer_path:
            cert = _load_certificate_from_cer(cer_path)

        if cert is None:
            cert = _pick_cert_from_signed_data(signed_data)

        valid_from_dt = None
        valid_to_dt = None

        if cert is not None:
            tbs = cert["tbs_certificate"]
            serial = tbs["serial_number"].native
            info.serial_number = _format_serial(serial)

            try:
                info.subject = _extract_common_name(cert)
            except Exception:
                logger.exception("Не удалось получить subject из сертификата")
                info.subject = "не удалось определить"

            try:
                info.issuer = cert.issuer.human_friendly
            except Exception:
                logger.exception("Не удалось получить issuer из сертификата")
                info.issuer = "не удалось определить"

            try:
                validity = tbs["validity"]
                valid_from_dt = _normalize_to_utc(validity["not_before"].native)
                valid_to_dt = _normalize_to_utc(validity["not_after"].native)
                info.valid_from = _format_date(valid_from_dt)
                info.valid_to = _format_date(valid_to_dt)
                logger.debug(
                    "Срок действия сертификата (UTC): %s - %s",
                    valid_from_dt,
                    valid_to_dt,
                )
            except Exception:
                logger.exception("Не удалось получить срок действия сертификата")
        else:
            logger.warning("Сертификат ни в CER, ни в P7S не найден")

        now = datetime.datetime.now(datetime.timezone.utc)

        # Формируем статус
        if msg_digest_attr is None or not digest_alg:
            info.status = "не удалось проверить подпись (нет атрибута messageDigest)"
        elif pdf_digest is None:
            if "gost" in digest_alg.lower() or str(digest_alg).startswith("1.2.643.7.1.1.2."):
                info.status = (
                    "не удалось проверить подпись "
                    "(алгоритм ГОСТ не поддерживается — установите пакет 'gostcrypto')"
                )
            else:
                info.status = "не удалось проверить подпись (неизвестный алгоритм хеширования)"
        elif not matches_document:
            info.status = "не соответствует документу"
        else:
            if valid_from_dt and valid_to_dt and (now < valid_from_dt or now > valid_to_dt):
                info.status = "срок сертификата истёк"
            else:
                info.status = "действительна"

        logger.info("Результат проверки подписи: %s", info.status)
        return info

    except Exception as e:  # pragma: no cover - защитный код
        logger.exception("Общая ошибка при проверке подписи")
        msg = str(e)
        if (
            "ContentInfo" in msg
            and "universal" in msg
            and "application was found" in msg
        ):
            info.status = (
                "ошибка проверки: неподдерживаемый формат файла подписи "
                "(ожидается CMS/PKCS#7, см. ep_viewer.log)"
            )
        else:
            info.status = f"ошибка проверки: {e}"
        return info
