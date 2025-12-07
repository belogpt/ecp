"""Build and obfuscation script for Windows releases.

This script obfuscates project sources with python-minifier and builds a
PyInstaller executable using the obfuscated sources.
"""
from __future__ import annotations

import platform
import shutil
import subprocess
import sys
from os import walk
from pathlib import Path

from python_minifier import minify

EXCLUDE_DIRS = {"obf_src", "build", "dist", "__pycache__", "venv", ".venv", ".git"}
EXCLUDE_FILES = {"build_release.py", "obfuscate.py"}
OBFUSCATED_ROOT = Path("obf_src")


class BuildError(Exception):
    """Custom exception for build-related errors."""


def ensure_windows() -> None:
    if platform.system() != "Windows":
        raise BuildError("Скрипт необходимо запускать на Windows.")


def ensure_required_files(root: Path) -> None:
    missing = [name for name in ("app.ico", "help.html") if not (root / name).exists()]
    if missing:
        raise BuildError(f"Отсутствуют необходимые файлы в корне проекта: {', '.join(missing)}")


def ensure_gostcrypto() -> None:
    try:
        import gostcrypto  # noqa: F401
    except ImportError:
        raise BuildError("Пакет 'gostcrypto' не установлен. Установите командой: python -m pip install gostcrypto")


def should_skip_dir(dirname: str) -> bool:
    return dirname in EXCLUDE_DIRS


def should_skip_file(path: Path) -> bool:
    return path.name in EXCLUDE_FILES


def os_walk_filtered(root: Path):
    for current_dir, dirnames, filenames in walk(root, topdown=True):
        dirnames[:] = [d for d in dirnames if not should_skip_dir(d)]
        yield Path(current_dir), dirnames, filenames


def obfuscate_project(root: Path = Path("."), obf_root: Path = OBFUSCATED_ROOT) -> None:
    obf_root.mkdir(exist_ok=True)

    for current_dir, dirnames, filenames in os_walk_filtered(root):
        rel_dir = current_dir.relative_to(root)
        target_dir = obf_root / rel_dir
        target_dir.mkdir(parents=True, exist_ok=True)

        for filename in filenames:
            source_path = current_dir / filename
            if source_path.suffix != ".py" or should_skip_file(source_path):
                continue

            target_path = target_dir / filename
            source_text = source_path.read_text(encoding="utf-8")
            minified = minify(
                source_text,
                remove_literal_statements=True,
                remove_annotations=True,
                hoist_literals=True,
                rename_locals=True,
                rename_globals=False,
            )
            target_path.write_text(minified, encoding="utf-8")


def copy_resources(root: Path = Path("."), obf_root: Path = OBFUSCATED_ROOT) -> None:
    for filename in ("app.ico", "help.html"):
        src = root / filename
        dest = obf_root / filename
        shutil.copy2(src, dest)


def clean_build_artifacts(obf_root: Path = OBFUSCATED_ROOT) -> None:
    for folder in (obf_root / "build", obf_root / "dist"):
        if folder.exists():
            shutil.rmtree(folder)

    for spec_file in obf_root.glob("*.spec"):
        spec_file.unlink()


def run_pyinstaller(obf_root: Path = OBFUSCATED_ROOT) -> None:
    main_path = obf_root / "main.py"
    if not main_path.exists():
        raise BuildError("Точка входа main.py не найдена после обфускации.")

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconsole",
        "--name",
        "MyApp",
        "--icon=app.ico",
        "--add-data",
        "app.ico;.",
        "--add-data",
        "help.html;.",
        "--hidden-import=gostcrypto",
        "--hidden-import=gostcrypto.gosthash",
        "--hidden-import=gostcrypto.gostcipher",
        "--hidden-import=gostcrypto.gostsignature",
        "--hidden-import=gostcrypto.gostrandom",
        "--hidden-import=gostcrypto.gosthmac",
        "--hidden-import=gostcrypto.gostpbkdf",
        "--hidden-import=gostcrypto.gostoid",
        "--hidden-import=fitz",
        "--clean",
        "--noconfirm",
        "main.py",
    ]
    subprocess.run(cmd, cwd=obf_root, check=True)


def main() -> None:
    root = Path(__file__).resolve().parent

    try:
        ensure_windows()
        ensure_required_files(root)
        ensure_gostcrypto()
        obfuscate_project(root=root, obf_root=OBFUSCATED_ROOT)
        copy_resources(root=root, obf_root=OBFUSCATED_ROOT)
        clean_build_artifacts(obf_root=OBFUSCATED_ROOT)
        run_pyinstaller(obf_root=OBFUSCATED_ROOT)
    except BuildError as exc:
        print(exc)
        sys.exit(1)
    except Exception as exc:  # noqa: BLE001
        print(f"Неожиданная ошибка: {exc}")
        sys.exit(1)

    print("Готовый exe: obf_src/dist/MyApp.exe")


if __name__ == "__main__":
    main()
