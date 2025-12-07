import os
import json
import logging
from typing import Dict, Optional, List

import fitz  # PyMuPDF
from PySide6.QtCore import Qt, QRect, QBuffer, QIODevice
from PySide6.QtGui import (
    QImage,
    QPainter,
    QFont,
    QFontDatabase,
    QColor,
    QPen,
    QBrush,
    QFontMetrics,
)
from paths import get_resource_path, get_data_path

logger = logging.getLogger(__name__)

_STAMP_FONT_FAMILY: Optional[str] = None

_SETTINGS_PATH = get_data_path("settings.json")
_LEGACY_CONFIG_PATHS = [
    get_data_path("header_config.json"),
    get_data_path("stamp_header.json"),
]

_DEFAULT_STAMP_RECT = {
    "left": 0.2,
    "top": 0.6,
    "width": 0.6,
    "height": 0.2,
}

_DEFAULT_SETTINGS = {
    "image_path": "",
    "header_text": "",
    "show_sign_time": False,
    "default_output_dir": "",
    "stamp_rect": _DEFAULT_STAMP_RECT,
}


# ---------------------------------------------------------------------------
# Настройки шапки (логотип + текст + флаг времени подписи)
# ---------------------------------------------------------------------------


def _load_raw_settings() -> Dict[str, object]:
    if os.path.exists(_SETTINGS_PATH):
        path = _SETTINGS_PATH
    else:
        path = None
        for legacy in _LEGACY_CONFIG_PATHS:
            if os.path.exists(legacy):
                path = legacy
                break

    if not path:
        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        logger.exception("Не удалось прочитать файл настроек: %s", path)
        return {}


def _normalize_stamp_rect(value: object) -> Dict[str, float]:
    if isinstance(value, dict):
        try:
            left = float(value.get("left", _DEFAULT_STAMP_RECT["left"]))
            top = float(value.get("top", _DEFAULT_STAMP_RECT["top"]))
            width = float(value.get("width", _DEFAULT_STAMP_RECT["width"]))
            height = float(value.get("height", _DEFAULT_STAMP_RECT["height"]))
        except Exception:
            return dict(_DEFAULT_STAMP_RECT)

        def _clamp(v: float) -> float:
            return max(0.0, min(1.0, v))

        width = _clamp(width)
        height = _clamp(height)
        left = _clamp(left)
        top = _clamp(top)

        if left + width > 1.0:
            width = max(0.0, 1.0 - left)
        if top + height > 1.0:
            height = max(0.0, 1.0 - top)

        if width <= 0 or height <= 0:
            return dict(_DEFAULT_STAMP_RECT)

        return {
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }
    return dict(_DEFAULT_STAMP_RECT)


def load_header_config() -> Dict[str, object]:
    """
    Загружает настройки приложения из settings.json (или совместимых устаревших
    файлов) и гарантирует наличие всех ключей.
    """

    data = _load_raw_settings()

    image_path = str(data.get("image_path", "") or "")

    if "header_text" in data:
        header_text = str(data.get("header_text") or "")
    else:
        lines = data.get("lines", [])
        if isinstance(lines, str):
            header_text = lines
        elif isinstance(lines, list):
            header_text = " ".join(str(x) for x in lines if x)
        else:
            header_text = ""

    show_sign_time = bool(data.get("show_sign_time", False))
    default_output_dir = str(data.get("default_output_dir", "") or "")
    stamp_rect = _normalize_stamp_rect(data.get("stamp_rect"))

    merged = {
        "image_path": image_path,
        "header_text": header_text,
        "show_sign_time": show_sign_time,
        "default_output_dir": default_output_dir,
        "stamp_rect": stamp_rect,
    }

    if not os.path.exists(_SETTINGS_PATH):
        try:
            with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
                json.dump(merged, f, ensure_ascii=False, indent=2)
        except Exception:
            logger.exception("Не удалось создать файл настроек: %s", _SETTINGS_PATH)
        else:
            for legacy in _LEGACY_CONFIG_PATHS:
                if os.path.exists(legacy):
                    try:
                        os.remove(legacy)
                    except Exception:
                        logger.warning(
                            "Не удалось удалить устаревший конфиг: %s", legacy
                        )

    return merged


def save_header_config(partial_cfg: Dict[str, object]) -> None:
    """
    Сохраняет настройки приложения в settings.json, обновляя только переданные
    поля.
    """

    current = load_header_config()
    current.update(partial_cfg)

    data = {
        "image_path": str(current.get("image_path", "") or ""),
        "header_text": str(current.get("header_text", "") or ""),
        "show_sign_time": bool(current.get("show_sign_time", False)),
        "default_output_dir": str(current.get("default_output_dir", "") or ""),
        "stamp_rect": _normalize_stamp_rect(current.get("stamp_rect")),
    }

    try:
        with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.exception("Не удалось сохранить настройки: %s", _SETTINGS_PATH)
    else:
        for legacy in _LEGACY_CONFIG_PATHS:
            if legacy == _SETTINGS_PATH:
                continue
            if os.path.exists(legacy):
                try:
                    os.remove(legacy)
                except Exception:
                    logger.warning("Не удалось удалить устаревший конфиг: %s", legacy)


# ---------------------------------------------------------------------------
# Работа с PDF и отрисовкой штампа
# ---------------------------------------------------------------------------


def _get_stamp_font_family() -> str:
    """
    Загружаем Times New Roman (кириллический) из файла timesnrcyrmt.ttf
    в каталоге проекта, если он есть. Иначе используем системный Times New Roman.
    """
    global _STAMP_FONT_FAMILY
    if _STAMP_FONT_FAMILY is not None:
        return _STAMP_FONT_FAMILY

    ttf_path = get_resource_path("timesnrcyrmt.ttf")
    family: Optional[str] = None

    if os.path.exists(ttf_path):
        try:
            fid = QFontDatabase.addApplicationFont(ttf_path)
            if fid != -1:
                fams = QFontDatabase.applicationFontFamilies(fid)
                if fams:
                    family = fams[0]
                    logger.info("Загружен TTF-шрифт штампа: %s (%s)", family, ttf_path)
        except Exception:
            logger.exception("Не удалось загрузить шрифт штампа: %s", ttf_path)

    if not family:
        family = "Times New Roman"
        logger.info("Используем системный шрифт штампа: %s", family)

    _STAMP_FONT_FAMILY = family
    return family


def open_document(path: str) -> fitz.Document:
    logger.debug("Открываем PDF-документ: %s", path)
    return fitz.open(path)


def _wrap_text_to_lines(
    text: str,
    fm: QFontMetrics,
    max_width: int,
    max_lines: int = 3,
) -> List[str]:
    """
    Переносит текст по словам на строки, чтобы каждая строка
    помещалась в max_width. Возвращает не более max_lines строк.
    """
    text = (text or "").strip()
    if not text or max_width <= 0 or max_lines <= 0:
        return [""] * max_lines

    words = text.split()
    lines: List[str] = []
    cur = ""

    for word in words:
        if not cur:
            # начинаем новую строку
            if fm.horizontalAdvance(word) <= max_width:
                cur = word
                continue
            # если отдельное слово длиннее max_width — обрезаем его
            cut = word
            while len(cut) > 1 and fm.horizontalAdvance(cut) > max_width:
                cut = cut[:-1]
            cur = cut
            continue

        cand = cur + " " + word
        if fm.horizontalAdvance(cand) <= max_width:
            cur = cand
        else:
            lines.append(cur)
            if len(lines) >= max_lines:
                break
            cur = word

    if cur and len(lines) < max_lines:
        lines.append(cur)

    # заполняем пустыми строками до max_lines
    while len(lines) < max_lines:
        lines.append("")

    return lines[:max_lines]


def build_stamp_image(stamp_info: Dict[str, str], show_sign_time: bool = True) -> QImage:
    """
    Рисует штамп:

      • верхняя шапка (логотип + текст из настроек, максимум 3 строки);
      • чёрная полоса с надписью «ДОКУМЕНТ ПОДПИСАН ЭЛЕКТРОННОЙ ПОДПИСЬЮ»;
      • блок сведений о сертификате:
            Сертификат
            Владелец
            Действителен: ...
            (опционально) Время подписи  — последней строкой.

    Если логотип не указан — шапка целиком не выводится.
    """
    width_px, height_px = 1200, 600
    image = QImage(width_px, height_px, QImage.Format_ARGB32)
    image.fill(Qt.white)

    painter = QPainter(image)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setRenderHint(QPainter.TextAntialiasing, True)

    # ------------ Рамка ------------------------------------------------------
    pen_width = max(2, int(width_px * 0.01))  # ~1% ширины, не меньше 2 px
    pen = QPen(QColor(0, 0, 0))
    pen.setWidth(pen_width)
    painter.setPen(pen)
    painter.setBrush(QBrush(Qt.white))

    half = pen_width // 2
    total_rect = QRect(half, half, width_px - pen_width, height_px - pen_width)

    radius = int(min(width_px, height_px) * 0.10)
    painter.drawRoundedRect(total_rect, radius, radius)

    # ------------ Данные подписи --------------------------------------------
    family = _get_stamp_font_family()

    serial = stamp_info.get("serial_number") or "не удалось определить"
    subject = stamp_info.get("subject") or "не удалось определить"
    signing_time = stamp_info.get("signing_time") or "не удалось определить"
    valid_from = stamp_info.get("valid_from") or "не удалось определить"
    valid_to = stamp_info.get("valid_to") or "не удалось определить"

    band_text = "ДОКУМЕНТ ПОДПИСАН ЭЛЕКТРОННОЙ ПОДПИСЬЮ"

    line_cert = f"Сертификат: {serial}"
    line_owner = f"Владелец: {subject}"
    line_valid = f"Действителен: с {valid_from} по {valid_to}"
    line_sign = f"Время подписи: {signing_time}"

    body_lines = [line_cert, line_owner, line_valid]
    if show_sign_time:
        body_lines.append(line_sign)

    # ------------ Конфигурация шапки ----------------------------------------
    cfg = load_header_config()
    header_image_path = cfg.get("image_path", "") or ""
    header_text = cfg.get("header_text", "") or ""

    header_image: Optional[QImage] = None
    if header_image_path and os.path.exists(header_image_path):
        img = QImage(header_image_path)
        if not img.isNull():
            header_image = img

    has_header_image = header_image is not None
    has_header_text = bool(header_text.strip())
    has_header = has_header_image or has_header_text

    # ------------ Геометрия и подбор размеров -------------------------------
    margin = int(min(width_px, height_px) * 0.08)
    inner_width = width_px - 2 * margin
    inner_height = height_px - 2 * margin

    base_body_size = int(inner_height * 0.11)
    body_size = base_body_size

    # находим такой размер body_size, чтобы всё влезло
    while True:
        body_font = QFont(family)
        body_font.setPixelSize(body_size)
        fm_body = QFontMetrics(body_font)
        body_h = fm_body.height()

        # чёрная полоса
        band_font = QFont(family)
        band_font.setBold(True)
        band_font.setPixelSize(int(body_size * 1.05))
        fm_band = QFontMetrics(band_font)
        band_h = fm_band.height()
        band_block_height = int(band_h * 1.5)
        band_width = fm_band.horizontalAdvance(band_text)

        # блок сведений
        max_body_width = 0
        for line in body_lines:
            if not line:
                continue
            w = fm_body.horizontalAdvance(line)
            if w > max_body_width:
                max_body_width = w

        gap_band_to_body = int(body_h * 0.8)
        step_body = int(body_h * 1.4)
        details_block_height = len(body_lines) * step_body

        # шапка
        header_block_height = 0
        header_block_width = 0
        if has_header:
            top_font = QFont(family)
            top_font.setPixelSize(int(body_size * 0.95))
            fm_top = QFontMetrics(top_font)
            top_h = fm_top.height()

            top_step = int(top_h * 1.2)
            # высота блока текста: ровно 3 строки
            text_block_height = 3 * top_step

            gap_icon_text = int(body_h * 0.8) if has_header_image else 0
            icon_side = 3 * top_step if has_header_image else 0

            max_text_width_area = inner_width - icon_side - gap_icon_text
            if max_text_width_area < 10:
                header_text_width = 0
            else:
                header_lines_tmp = _wrap_text_to_lines(
                    header_text,
                    fm_top,
                    max_text_width_area,
                    max_lines=3,
                )
                header_text_width = 0
                for l in header_lines_tmp:
                    if l:
                        w = fm_top.horizontalAdvance(l)
                        if w > header_text_width:
                            header_text_width = w

            header_block_width = icon_side + gap_icon_text + header_text_width
            header_block_height = max(icon_side, text_block_height)
            gap_header_to_band = int(body_h * 0.8)
        else:
            gap_header_to_band = 0

        total_height = (
            header_block_height
            + gap_header_to_band
            + band_block_height
            + gap_band_to_body
            + details_block_height
        )

        widths_ok = (
            max_body_width <= inner_width
            and band_width <= inner_width
            and (not has_header or header_block_width <= inner_width)
        )
        heights_ok = total_height <= inner_height

        if widths_ok and heights_ok:
            break

        body_size = int(body_size * 0.95)
        if body_size <= 7:
            break

    # ==== Финальные метрики с найденным body_size ===========================
    painter.setPen(QPen(QColor(0, 0, 0)))

    body_font = QFont(family)
    body_font.setPixelSize(body_size)
    fm_body = QFontMetrics(body_font)
    body_h = fm_body.height()
    gap_band_to_body = int(body_h * 0.8)
    step_body = int(body_h * 1.4)

    band_font = QFont(family)
    band_font.setBold(True)
    band_font.setPixelSize(int(body_size * 1.05))
    fm_band = QFontMetrics(band_font)
    band_h = fm_band.height()
    band_block_height = int(band_h * 1.5)

    top_font = None
    fm_top = None
    top_h = 0
    top_step = 0
    icon_side = 0
    gap_icon_text = 0
    header_lines: List[str] = ["", "", ""]
    header_block_height = 0
    gap_header_to_band = 0

    if has_header:
        top_font = QFont(family)
        top_font.setPixelSize(int(body_size * 0.95))
        fm_top = QFontMetrics(top_font)
        top_h = fm_top.height()
        top_step = int(top_h * 1.2)

        if has_header_image:
            icon_side = 3 * top_step
            gap_icon_text = int(body_h * 0.8)

        max_text_width_area = inner_width - icon_side - gap_icon_text
        if max_text_width_area < 10:
            header_lines = ["", "", ""]
        else:
            header_lines = _wrap_text_to_lines(
                header_text,
                fm_top,
                max_text_width_area,
                max_lines=3,
            )

        header_block_height = max(icon_side, 3 * top_step)
        gap_header_to_band = int(body_h * 0.8)

    # ==== Рисование =========================================================
    y = margin

    # --- шапка --------------------------------------------------------------
    if has_header and top_font is not None and fm_top is not None:
        painter.setFont(top_font)

        x = margin

        if has_header_image and header_image is not None:
            # логотип
            logo_rect = QRect(x, y, icon_side, icon_side)
            scaled_logo = header_image.scaled(
                icon_side,
                icon_side,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation,
            )
            painter.drawImage(logo_rect, scaled_logo)
            x += icon_side + gap_icon_text

        # три строки текста
        text_y = y
        for line in header_lines:
            rect = QRect(x, text_y, inner_width - (x - margin), top_h)
            if line:
                painter.drawText(rect, Qt.AlignLeft | Qt.AlignVCenter, line)
            text_y += top_step

        y += header_block_height + gap_header_to_band

    # --- чёрная полоса ------------------------------------------------------
    painter.setFont(band_font)
    band_rect = QRect(margin, y, inner_width, band_block_height)

    painter.save()
    painter.setPen(Qt.NoPen)
    painter.setBrush(QBrush(Qt.black))
    painter.drawRect(band_rect)
    painter.restore()

    painter.setPen(QColor(255, 255, 255))
    painter.drawText(band_rect, Qt.AlignCenter, band_text)

    y += band_block_height + gap_band_to_body

    # --- сведения о сертификате --------------------------------------------
    painter.setPen(QColor(0, 0, 0))
    painter.setFont(body_font)

    for line in body_lines:
        rect = QRect(margin, y, inner_width, body_h)
        if line:
            painter.drawText(rect, Qt.AlignLeft | Qt.AlignVCenter, line)
        y += step_body

    painter.end()
    return image


def _image_to_png_bytes(img: QImage) -> bytes:
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    data = bytes(buf.data())
    buf.close()
    return data


def add_stamp_to_pdf(
    src_pdf_path: str,
    dst_pdf_path: str,
    page_index: int,
    rect_pdf: fitz.Rect,
    stamp_info: Dict[str, str],
    show_sign_time: bool = True,
) -> None:
    """
    Копирует исходный PDF и добавляет штамп на выбранной странице
    в прямоугольник rect_pdf.
    """
    logger.info(
        "Добавляем штамп в PDF: src=%s dst=%s page=%d rect=%s",
        src_pdf_path,
        dst_pdf_path,
        page_index,
        rect_pdf,
    )

    doc = fitz.open(src_pdf_path)

    img = build_stamp_image(stamp_info, show_sign_time=show_sign_time)
    img_bytes = _image_to_png_bytes(img)

    page = doc[page_index]
    page.insert_image(rect_pdf, stream=img_bytes)

    doc.save(dst_pdf_path)
    doc.close()
    logger.info("PDF со штампом сохранён: %s", dst_pdf_path)
