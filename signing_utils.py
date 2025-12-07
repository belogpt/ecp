"""
Утилиты для подписания PDF-файлов с помощью ЭЦП (PKCS#7 detached).
"""

from __future__ import annotations

import os
import logging
from typing import Optional

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization import pkcs7
from cryptography.hazmat.primitives.serialization import (
    load_pem_private_key,
    load_der_private_key,
)

logger = logging.getLogger(__name__)


class PKCS11PrivateKey:
    """Адаптер для использования ключа на токене PKCS#11 в PKCS7SignatureBuilder."""

    def __init__(self, key, mechanism):
        self._key = key
        self._mechanism = mechanism
        try:
            from pkcs11 import Attribute

            self.key_size = int(key[Attribute.MODULUS_BITS])
        except Exception:
            # Размер ключа нужен только для вспомогательной информации, поэтому по умолчанию 0.
            self.key_size = 0

    def sign(self, data, padding, algorithm):
        # padding и algorithm задаются PKCS7SignatureBuilder; для токена важен только механизм.
        return self._key.sign(data, mechanism=self._mechanism)


def _load_certificate(cert_path: str) -> x509.Certificate:
    """Загружает сертификат в формате PEM или DER."""
    with open(cert_path, "rb") as f:
        data = f.read()
    try:
        return x509.load_pem_x509_certificate(data)
    except ValueError:
        logger.debug("Сертификат не PEM, пробуем DER")
        return x509.load_der_x509_certificate(data)


def _load_private_key(key_path: str, password: Optional[str]):
    """Загружает закрытый ключ (PEM или DER)."""
    with open(key_path, "rb") as f:
        data = f.read()
    pwd_bytes = password.encode("utf-8") if password else None
    try:
        return load_pem_private_key(data, password=pwd_bytes)
    except ValueError:
        logger.debug("Ключ не PEM, пробуем DER")
        return load_der_private_key(data, password=pwd_bytes)


def _import_pkcs11():
    try:
        import pkcs11
        from pkcs11 import lib, Mechanism, ObjectClass, Attribute
    except Exception as exc:  # pragma: no cover - пакет может быть не установлен в окружении CI
        raise ImportError(
            "Для работы с токеном PKCS#11 установите пакет 'pkcs11'"
        ) from exc
    return pkcs11, lib, Mechanism, ObjectClass, Attribute


def _select_token(pkcs11_lib, token_label: Optional[str], slot: Optional[int]):
    if token_label:
        return pkcs11_lib.get_token(token_label=token_label)
    if slot is not None:
        return pkcs11_lib.get_token(slot=slot)

    tokens = pkcs11_lib.get_tokens()
    if not tokens:
        raise RuntimeError("Токены PKCS#11 не обнаружены")
    if len(tokens) > 1:
        raise RuntimeError(
            "Найдено несколько токенов. Уточните слот или метку токена в настройках."
        )
    return tokens[0]


def _load_cert_from_token(session, Attribute, ObjectClass) -> x509.Certificate:
    certs = list(
        session.get_objects({Attribute.CLASS: ObjectClass.CERTIFICATE})
    )
    if not certs:
        raise RuntimeError(
            "На токене не найден сертификат. Укажите путь к сертификату вручную."
        )

    cert_obj = certs[0]
    der_bytes = cert_obj[Attribute.VALUE]
    return x509.load_der_x509_certificate(der_bytes)


def _resolve_pkcs11_certificate(cert_path: Optional[str], session, Attribute, ObjectClass):
    if cert_path:
        return _load_certificate(cert_path)
    return _load_cert_from_token(session, Attribute, ObjectClass)


def _resolve_pkcs11_private_key(session, Attribute, ObjectClass, key_label: Optional[str]):
    query = {Attribute.CLASS: ObjectClass.PRIVATE_KEY}
    if key_label:
        query[Attribute.LABEL] = key_label

    keys = list(session.get_objects(query))
    if not keys:
        raise RuntimeError("На токене не найден закрытый ключ с указанными параметрами")
    return keys[0]


def sign_pdf(pdf_path: str, cert_path: str, key_path: str, password: Optional[str] = None,
             output_dir: Optional[str] = None) -> str:
    """
    Подписывает PDF-файл и создаёт отсоединённую подпись в формате PKCS#7.

    Возвращает путь к созданному файлу подписи (.p7s).
    """
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF не найден: {pdf_path}")
    if not os.path.exists(cert_path):
        raise FileNotFoundError(f"Сертификат не найден: {cert_path}")
    if not os.path.exists(key_path):
        raise FileNotFoundError(f"Закрытый ключ не найден: {key_path}")

    logger.info("Подпись PDF %s с использованием сертификата %s", pdf_path, cert_path)

    certificate = _load_certificate(cert_path)
    private_key = _load_private_key(key_path, password)

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    builder = pkcs7.PKCS7SignatureBuilder().set_data(pdf_bytes)
    builder = builder.add_signer(certificate, private_key, hashes.SHA256())
    signature = builder.sign(
        serialization.Encoding.DER,
        [pkcs7.PKCS7Options.DetachedSignature],
    )

    base_name, _ = os.path.splitext(os.path.basename(pdf_path))
    signature_name = f"{base_name}_Файл подписи.p7s"
    target_dir = output_dir or os.path.dirname(pdf_path) or os.getcwd()
    os.makedirs(target_dir, exist_ok=True)
    signature_path = os.path.join(target_dir, signature_name)

    with open(signature_path, "wb") as f:
        f.write(signature)

    logger.info("Файл подписи создан: %s", signature_path)
    return signature_path


def sign_pdf_with_pkcs11(
    pdf_path: str,
    pkcs11_lib_path: str,
    pin: str,
    token_label: Optional[str] = None,
    slot: Optional[int] = None,
    key_label: Optional[str] = None,
    cert_path: Optional[str] = None,
    output_dir: Optional[str] = None,
) -> str:
    """Подписывает PDF, используя закрытый ключ на токене PKCS#11."""

    _, lib, Mechanism, ObjectClass, Attribute = _import_pkcs11()

    if not os.path.exists(pdf_path):
        raise FileNotFoundError(f"PDF не найден: {pdf_path}")
    if not os.path.exists(pkcs11_lib_path):
        raise FileNotFoundError(
            f"Библиотека PKCS#11 не найдена: {pkcs11_lib_path}"
        )

    logger.info(
        "Подпись PDF %s через токен PKCS#11 (token_label=%s, slot=%s)",
        pdf_path,
        token_label,
        slot,
    )

    pkcs11_lib = lib(pkcs11_lib_path)
    token = _select_token(pkcs11_lib, token_label, slot)

    with token.open(user_pin=pin) as session:
        certificate = _resolve_pkcs11_certificate(
            cert_path, session, Attribute, ObjectClass
        )
        private_key_obj = _resolve_pkcs11_private_key(
            session, Attribute, ObjectClass, key_label
        )

        pkcs11_key = PKCS11PrivateKey(private_key_obj, Mechanism.SHA256_RSA_PKCS)

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        builder = pkcs7.PKCS7SignatureBuilder().set_data(pdf_bytes)
        builder = builder.add_signer(certificate, pkcs11_key, hashes.SHA256())
        signature = builder.sign(
            serialization.Encoding.DER,
            [pkcs7.PKCS7Options.DetachedSignature],
        )

    base_name, _ = os.path.splitext(os.path.basename(pdf_path))
    signature_name = f"{base_name}_Файл подписи.p7s"
    target_dir = output_dir or os.path.dirname(pdf_path) or os.getcwd()
    os.makedirs(target_dir, exist_ok=True)
    signature_path = os.path.join(target_dir, signature_name)

    with open(signature_path, "wb") as f:
        f.write(signature)

    logger.info("Файл подписи создан через PKCS#11: %s", signature_path)
    return signature_path
