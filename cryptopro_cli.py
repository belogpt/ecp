"""Интерфейс для подписи файлов через утилиты CryptoPro (cryptcp)."""
from __future__ import annotations

import logging
import os
import platform
import shutil
import subprocess
from pathlib import Path
from typing import List, Optional, Sequence

logger = logging.getLogger(__name__)

# Флаги cryptcp вынесены в одно место для удобства корректировки при смене версии CSP.
CRYPTCP_FLAGS = {
    "sign_detached": ["-sign", "-detached"],
    "sign_attached": ["-sign"],
    "verify_attached": ["-verify"],
    "verify_detached": ["-verify", "-detached"],
}


class CryptoProNotFoundError(FileNotFoundError):
    """Ошибка отсутствия утилит CryptoPro."""


class CertificateSelectorError(ValueError):
    """Ошибка выбора сертификата."""


def find_cryptopro_tools() -> str:
    """Возвращает путь к `cryptcp` или выбрасывает ошибку, если утилита не найдена."""

    path = shutil.which("cryptcp")
    if path:
        logger.debug("cryptcp найден в PATH: %s", path)
        return path

    system = platform.system().lower()
    if system.startswith("win"):
        candidates = [
            r"C:\\Program Files\\Crypto Pro\\CSP\\cryptcp.exe",
            r"C:\\Program Files (x86)\\Crypto Pro\\CSP\\cryptcp.exe",
            r"C:\\Program Files\\CryptoPro\\CSP\\cryptcp.exe",
            r"C:\\Program Files (x86)\\CryptoPro\\CSP\\cryptcp.exe",
        ]
    elif system == "linux":
        candidates = [
            "/opt/cprocsp/bin/amd64/cryptcp",
            "/opt/cprocsp/bin/ia32/cryptcp",
        ]
    elif system == "darwin":
        candidates = [
            "/Applications/CryptoPro/CSP/cryptcp",
        ]
    else:
        candidates = []

    for candidate in candidates:
        if os.path.isfile(candidate):
            logger.debug("cryptcp найден по умолчанию: %s", candidate)
            return candidate

    raise CryptoProNotFoundError(
        "Утилита cryptcp не найдена. Установите CryptoPro CSP и убедитесь, что cryptcp доступен."
    )


def _ensure_input_file(input_path: str) -> Path:
    normalized = Path(input_path).expanduser().resolve()
    if not normalized.exists():
        raise FileNotFoundError(f"Файл для подписи не найден: {normalized}")
    if not normalized.is_file():
        raise FileNotFoundError(f"Указан не файл: {normalized}")
    return normalized


def _resolve_output_path(default_suffix: str, explicit_output: Optional[str], input_path: Path) -> Path:
    if explicit_output:
        return Path(explicit_output).expanduser().resolve()
    return input_path.with_suffix(input_path.suffix + default_suffix)


def _build_selector_args(
    thumbprint: Optional[str],
    subject: Optional[str],
    container: Optional[str],
    *,
    choose_certificate: bool = False,
) -> List[str]:
    if choose_certificate:
        if thumbprint or subject or container:
            raise CertificateSelectorError(
                "Опция выбора сертификата (--choose) несовместима с явным указанием сертификата",
            )
        return ["-choose"]

    if thumbprint:
        return ["-thumbprint", thumbprint]
    if subject:
        return ["-subject", subject]
    if container:
        return ["-cont", container]

    raise CertificateSelectorError(
        "Не указан сертификат. Передайте отпечаток сертификата через --thumbprint или используйте --choose для выбора установленного сертификата."
    )


def _log_command(cmd: Sequence[str]) -> None:
    safe_cmd: List[str] = []
    redact_next = False
    for part in cmd:
        if redact_next:
            safe_cmd.append("<thumbprint>")
            redact_next = False
            continue
        if part.lower() == "-thumbprint":
            safe_cmd.append(part)
            redact_next = True
            continue
        safe_cmd.append(part)
    logger.info("Выполняется команда cryptcp: %s", " ".join(safe_cmd))


def _run(cmd: List[str], dry_run: bool = False) -> subprocess.CompletedProcess[str]:
    _log_command(cmd)
    if dry_run:
        return subprocess.CompletedProcess(cmd, returncode=0, stdout="", stderr="")

    try:
        result = subprocess.run(
            cmd,
            check=True,
            text=True,
            capture_output=True,
        )
        return result
    except FileNotFoundError as exc:  # pragma: no cover - зависит от окружения
        raise CryptoProNotFoundError("cryptcp не найден в системе") from exc
    except subprocess.CalledProcessError as exc:
        stderr = exc.stderr.strip() if exc.stderr else "неизвестная ошибка"
        raise RuntimeError(f"Ошибка выполнения cryptcp: {stderr}") from exc


def _prepare_command(
    base_flags: Sequence[str],
    input_path: Path,
    output_path: Optional[Path],
    selector_args: List[str],
    *,
    extra_flags: Optional[Sequence[str]] = None,
    tool_path: Optional[str] = None,
) -> List[str]:
    cmd = [tool_path or find_cryptopro_tools()]
    cmd.extend(base_flags)
    if extra_flags:
        cmd.extend(extra_flags)
    cmd.extend(selector_args)
    cmd.extend(["-in", str(input_path)])
    if output_path:
        cmd.extend(["-out", str(output_path)])
    return cmd


def sign_file_detached(
    input_path: str,
    output_sig_path: Optional[str] = None,
    thumbprint: Optional[str] = None,
    subject: Optional[str] = None,
    container: Optional[str] = None,
    choose: bool = False,
    *,
    dry_run: bool = False,
    tool_path: Optional[str] = None,
) -> Path:
    src = _ensure_input_file(input_path)
    dst = _resolve_output_path(".sig", output_sig_path, src)
    selector_args = _build_selector_args(
        thumbprint,
        subject,
        container,
        choose_certificate=choose,
    )
    cmd = _prepare_command(
        CRYPTCP_FLAGS["sign_detached"],
        src,
        dst,
        selector_args,
        tool_path=tool_path,
    )
    _run(cmd, dry_run=dry_run)
    return dst


def sign_file_attached(
    input_path: str,
    output_path: Optional[str] = None,
    thumbprint: Optional[str] = None,
    subject: Optional[str] = None,
    container: Optional[str] = None,
    choose: bool = False,
    *,
    dry_run: bool = False,
    tool_path: Optional[str] = None,
) -> Path:
    src = _ensure_input_file(input_path)
    dst = _resolve_output_path(".p7m", output_path, src)
    selector_args = _build_selector_args(
        thumbprint,
        subject,
        container,
        choose_certificate=choose,
    )
    cmd = _prepare_command(
        CRYPTCP_FLAGS["sign_attached"],
        src,
        dst,
        selector_args,
        tool_path=tool_path,
    )
    _run(cmd, dry_run=dry_run)
    return dst


def verify_signature(
    input_path: str,
    sig_path: Optional[str] = None,
    *,
    dry_run: bool = False,
    tool_path: Optional[str] = None,
) -> subprocess.CompletedProcess[str]:
    input_file = _ensure_input_file(input_path)
    if sig_path:
        signature = _ensure_input_file(sig_path)
        base_flags = CRYPTCP_FLAGS["verify_detached"]
        extra = ["-data", str(input_file)]
        cmd = _prepare_command(
            base_flags,
            signature,
            None,
            selector_args=[],
            extra_flags=extra,
            tool_path=tool_path,
        )
    else:
        base_flags = CRYPTCP_FLAGS["verify_attached"]
        cmd = _prepare_command(
            base_flags,
            input_file,
            None,
            selector_args=[],
            tool_path=tool_path,
        )
    return _run(cmd, dry_run=dry_run)

