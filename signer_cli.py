"""CLI-обёртка для подписи и проверки через CryptoPro cryptcp."""
from __future__ import annotations

import argparse
import logging
import sys

from cryptopro_cli import (
    CertificateSelectorError,
    CryptoProNotFoundError,
    find_cryptopro_tools,
    sign_file_attached,
    sign_file_detached,
    verify_signature,
)

logger = logging.getLogger(__name__)


def setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s: %(message)s")


def _add_certificate_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--thumbprint", help="Отпечаток сертификата")
    parser.add_argument("--subject", help="Часть subject сертификата (запасной вариант)")
    parser.add_argument("--container", help="Имя контейнера (fallback)")
    parser.add_argument(
        "--choose",
        action="store_true",
        help="Открыть встроенный выбор сертификатов и подписать выбранным сертификатом",
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Подпись и проверка файлов через утилиту CryptoPro cryptcp",
    )
    parser.add_argument("--dry-run", action="store_true", help="Показывать команду без выполнения")
    parser.add_argument("--verbose", action="store_true", help="Подробное логирование")

    subparsers = parser.add_subparsers(dest="command", required=True)

    sign_parser = subparsers.add_parser("sign", help="Подписать файл")
    sign_parser.add_argument("--file", required=True, help="Путь к файлу для подписи")
    sign_mode = sign_parser.add_mutually_exclusive_group()
    sign_mode.add_argument("--detached", action="store_true", help="Отсоединённая подпись (.sig)")
    sign_mode.add_argument("--attached", action="store_true", help="Присоединённая подпись (.p7m)")
    sign_parser.add_argument("--out", help="Путь к выходному файлу")
    _add_certificate_options(sign_parser)

    verify_parser = subparsers.add_parser("verify", help="Проверить подпись")
    verify_parser.add_argument("--file", required=True, help="Файл с подписью или исходный файл")
    verify_parser.add_argument("--sig", help="Отдельный файл подписи для detached режима")

    return parser


def _resolve_sign_mode(args: argparse.Namespace) -> str:
    if args.attached:
        return "attached"
    return "detached"


def _handle_sign(args: argparse.Namespace) -> int:
    mode = _resolve_sign_mode(args)
    try:
        if mode == "attached":
            out_path = sign_file_attached(
                args.file,
                output_path=args.out,
                thumbprint=args.thumbprint,
                subject=args.subject,
                container=args.container,
                choose=args.choose,
                dry_run=args.dry_run,
            )
        else:
            out_path = sign_file_detached(
                args.file,
                output_sig_path=args.out,
                thumbprint=args.thumbprint,
                subject=args.subject,
                container=args.container,
                choose=args.choose,
                dry_run=args.dry_run,
            )
    except (CertificateSelectorError, FileNotFoundError, CryptoProNotFoundError, RuntimeError) as exc:
        logger.error(str(exc))
        return 1

    logger.info("Файл подписи создан: %s", out_path)
    return 0


def _handle_verify(args: argparse.Namespace) -> int:
    try:
        result = verify_signature(args.file, sig_path=args.sig, dry_run=args.dry_run)
    except (FileNotFoundError, CryptoProNotFoundError, RuntimeError) as exc:
        logger.error(str(exc))
        return 1

    if args.dry_run:
        logger.info("Проверка (dry-run) завершена: %s", result.args)
        return 0

    if result.returncode == 0:
        logger.info("Подпись прошла проверку")
        return 0

    logger.error("Подпись не прошла проверку")
    return result.returncode


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    setup_logging(args.verbose)

    try:
        find_cryptopro_tools()
    except CryptoProNotFoundError as exc:
        logger.error(str(exc))
        return 1

    if args.command == "sign":
        return _handle_sign(args)
    if args.command == "verify":
        return _handle_verify(args)
    parser.error("Неизвестная команда")
    return 1


if __name__ == "__main__":
    sys.exit(main())

