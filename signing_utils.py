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
from cryptography.hazmat.primitives.serialization import load_pem_private_key, load_der_private_key

logger = logging.getLogger(__name__)


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
