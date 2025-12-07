# ВАЖНО: явные импорты gostcrypto для корректной сборки PyInstaller

import logging
import os
import sys

import gostcrypto  # type: ignore
import gostcrypto.gosthash  # type: ignore
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from gui import MainWindow
from paths import get_resource_path, get_data_path


# Если нужно отключить лог-файл, установите значение False.
ENABLE_FILE_LOG = True


def setup_logging():
    """
    Настройка логирования:
    - в файл ep_viewer.log
    - дублирование в консоль
    """
    handlers = [logging.StreamHandler(sys.stdout)]
    if ENABLE_FILE_LOG:
        log_path = get_data_path("ep_viewer.log")
        handlers.append(logging.FileHandler(log_path, encoding="utf-8"))

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        handlers=handlers,
    )
    logging.getLogger(__name__).info("=== Приложение запущено ===")


def main():
    setup_logging()
    app = QApplication(sys.argv)
    icon_path = get_resource_path("app.ico")
    app.setWindowIcon(QIcon(icon_path))

    # Подчёркиваем путь к app.ico — он один и тот же для IDE и PyInstaller.
    # Именно эту иконку заберёт PyInstaller и покажет в итоговом .exe.
    window = MainWindow()
    window.setWindowIcon(QIcon(icon_path))
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
