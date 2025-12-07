import os
import sys
import logging
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List, Dict

import fitz  # PyMuPDF
from PySide6.QtCore import Qt, QRectF, QSize, QPoint, QUrl, QTimer
from PySide6.QtGui import (
    QPixmap,
    QImage,
    QPainter,
    QWheelEvent,
    QMouseEvent,
    QIcon,
    QAction,
    QDesktopServices,
)
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QGraphicsView,
    QGraphicsScene,
    QGraphicsPixmapItem,
    QGraphicsRectItem,
    QFileDialog,
    QMessageBox,
    QPushButton,
    QFormLayout,
    QLineEdit,
    QPlainTextEdit,
    QSplitter,
    QCheckBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QTextBrowser,
    QFrame,
    QTabWidget,
    QSpinBox,
    QSpacerItem,
    QSizePolicy,
)

from pdf_utils import (
    open_document,
    build_stamp_image,
    add_stamp_to_pdf,
    load_header_config,
    save_header_config,
)
from signature_utils import get_certificate_info, CertificateInfo
from signing_utils import sign_pdf, sign_pdf_with_pkcs11
from browser_signing import BrowserSigningSession, BrowserSigningError
from paths import get_resource_path

logger = logging.getLogger(__name__)

PDF_RENDER_ZOOM = 2.0  # коэффициент рендеринга страниц в картинку

# Поддерживаемые расширения отсоединённых подписей PKCS#7/CMS.
# При необходимости можно добавить свои.
SIGNATURE_EXTS = (".p7s", ".sig", ".p7m", ".p7b", ".p7c")


@dataclass
class FileSession:
    """Состояние по одному PDF-файлу."""
    pdf_path: str
    p7s_path: Optional[str]
    doc: fitz.Document
    cert_info: CertificateInfo
    current_page_index: int = 0
    saved: bool = False
    output_dir: Optional[str] = None


class StampRectItem(QGraphicsRectItem):
    HANDLE_SIZE = 10.0

    def __init__(self, rect: QRectF, geometry_changed_callback=None):
        super().__init__(rect)
        self.setFlags(
            QGraphicsRectItem.ItemIsMovable
            | QGraphicsRectItem.ItemIsSelectable
            | QGraphicsRectItem.ItemSendsGeometryChanges
        )
        self._resizing = False
        self._geometry_changed_callback = geometry_changed_callback
        self._aspect_ratio = rect.width() / rect.height() if rect.height() > 0 else 1.0

    def paint(self, painter: QPainter, option, widget=None):
        rect = self.rect()
        handle_rect = QRectF(
            rect.right() - self.HANDLE_SIZE,
            rect.bottom() - self.HANDLE_SIZE,
            self.HANDLE_SIZE,
            self.HANDLE_SIZE,
        )
        painter.setBrush(Qt.blue)
        painter.setPen(Qt.blue)
        painter.drawRect(handle_rect)

    def _handle_rect_scene(self) -> QRectF:
        rect = self.rect()
        handle_rect = QRectF(
            rect.right() - self.HANDLE_SIZE,
            rect.bottom() - self.HANDLE_SIZE,
            self.HANDLE_SIZE,
            self.HANDLE_SIZE,
        )
        return self.mapRectToScene(handle_rect)

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.LeftButton:
            handle_rect = self._handle_rect_scene()
            if handle_rect.contains(event.scenePos()):
                self._resizing = True
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent):
        if self._resizing:
            scene_pos = event.scenePos()
            top_left = self.mapToScene(self.rect().topLeft())
            new_width = max(scene_pos.x() - top_left.x(), 20.0)
            new_height = new_width / self._aspect_ratio if self._aspect_ratio > 0 else new_width

            if new_height < 20.0:
                new_height = 20.0
                new_width = new_height * self._aspect_ratio if self._aspect_ratio > 0 else new_height

            self.prepareGeometryChange()
            self.setRect(QRectF(0.0, 0.0, new_width, new_height))
            self.setPos(top_left)
            self._notify_geometry_changed()
            event.accept()
            return

        super().mouseMoveEvent(event)
        self._notify_geometry_changed()

    def mouseReleaseEvent(self, event: QMouseEvent):
        if self._resizing and event.button() == Qt.LeftButton:
            self._resizing = False
            self._notify_geometry_changed()
            event.accept()
            return
        super().mouseReleaseEvent(event)
        self._notify_geometry_changed()

    def _notify_geometry_changed(self):
        if self._geometry_changed_callback is not None:
            self._geometry_changed_callback()

    def set_rect_and_update_aspect(self, rect: QRectF):
        self.prepareGeometryChange()
        self.setRect(QRectF(0.0, 0.0, rect.width(), rect.height()))
        self.setPos(rect.topLeft())
        if rect.height() > 0:
            self._aspect_ratio = rect.width() / rect.height()
        self._notify_geometry_changed()


class PDFPageView(QGraphicsView):
    """Виджет просмотра страницы с размещением штампа."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))

        self._pixmap_item: Optional[QGraphicsPixmapItem] = None
        self.stamp_item: Optional[StampRectItem] = None
        self.stamp_pixmap_item: Optional[QGraphicsPixmapItem] = None
        self.stamp_pixmap: Optional[QPixmap] = None
        self.zoom = 1.0

        self._space_pressed = False
        self._panning = False
        self._last_pan_pos: Optional[QPoint] = None

        self._external_geom_cb = None

        self.setRenderHint(QPainter.Antialiasing, True)
        self.setRenderHint(QPainter.SmoothPixmapTransform, True)
        self.setViewportUpdateMode(QGraphicsView.BoundingRectViewportUpdate)
        self.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)

    def set_external_geometry_callback(self, cb):
        self._external_geom_cb = cb

    def set_page(self, pixmap: QPixmap, pdf_zoom: float):
        self.scene().clear()
        self._pixmap_item = None
        self.stamp_item = None
        self.stamp_pixmap_item = None
        self.stamp_pixmap = None

        self._pixmap_item = self.scene().addPixmap(pixmap)
        self._pixmap_item.setZValue(0)
        self.zoom = pdf_zoom

        self._create_default_stamp()
        self._fit_page()

    def _create_default_stamp(self):
        if not self._pixmap_item:
            return
        page_rect = self._pixmap_item.boundingRect()
        width = page_rect.width() * 0.6
        height = page_rect.height() * 0.2
        x = page_rect.x() + (page_rect.width() - width) / 2
        y = page_rect.y() + page_rect.height() - height - page_rect.height() * 0.1
        rect = QRectF(x, y, width, height)
        self.stamp_item = StampRectItem(rect, geometry_changed_callback=self._update_pixmap_item)
        self.stamp_item.setZValue(10)
        self.scene().addItem(self.stamp_item)

    def hide_stamp(self):
        if self.stamp_pixmap_item:
            self.scene().removeItem(self.stamp_pixmap_item)
            self.stamp_pixmap_item = None
        self.stamp_pixmap = None
        if self.stamp_item:
            self.scene().removeItem(self.stamp_item)
            self.stamp_item = None

    def ensure_stamp_item(self):
        if self.stamp_item is None and self._pixmap_item is not None:
            self._create_default_stamp()

    def set_stamp_rect_normalized(self, norm_rect: QRectF):
        if not self._pixmap_item:
            return
        self.ensure_stamp_item()
        if not self.stamp_item:
            return
        page_rect = self._pixmap_item.sceneBoundingRect()
        w = page_rect.width() * norm_rect.width()
        h = page_rect.height() * norm_rect.height()
        x = page_rect.left() + page_rect.width() * norm_rect.left()
        y = page_rect.top() + page_rect.height() * norm_rect.top()
        rect = QRectF(x, y, w, h)
        self.stamp_item.set_rect_and_update_aspect(rect)
        self._update_pixmap_item()

    def get_stamp_rect_normalized(self) -> Optional[QRectF]:
        if not self._pixmap_item or not self.stamp_item:
            return None
        scene_rect = self.stamp_item.mapRectToScene(self.stamp_item.rect())
        page_rect = self._pixmap_item.sceneBoundingRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return None
        left = (scene_rect.left() - page_rect.left()) / page_rect.width()
        top = (scene_rect.top() - page_rect.top()) / page_rect.height()
        w = scene_rect.width() / page_rect.width()
        h = scene_rect.height() / page_rect.height()
        return QRectF(left, top, w, h)

    def set_stamp_pixmap(self, pixmap: QPixmap):
        self.stamp_pixmap = pixmap
        self.ensure_stamp_item()
        if self.stamp_pixmap_item is None:
            self.stamp_pixmap_item = self.scene().addPixmap(pixmap)
            self.stamp_pixmap_item.setZValue(5)
        self._update_pixmap_item()

    def _update_pixmap_item(self):
        if not self.stamp_item or not self.stamp_pixmap_item or not self.stamp_pixmap:
            if self._external_geom_cb:
                self._external_geom_cb()
            return
        rect = self.stamp_item.rect()
        if rect.width() <= 0 or rect.height() <= 0:
            if self._external_geom_cb:
                self._external_geom_cb()
            return
        target_size = QSize(int(rect.width()), int(rect.height()))
        scaled = self.stamp_pixmap.scaled(target_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.stamp_pixmap_item.setPixmap(scaled)
        scene_rect = self.stamp_item.mapRectToScene(self.stamp_item.rect())
        pix_rect = self.stamp_pixmap_item.boundingRect()
        offset_x = (scene_rect.width() - pix_rect.width()) / 2
        offset_y = (scene_rect.height() - pix_rect.height()) / 2
        self.stamp_pixmap_item.setPos(scene_rect.left() + offset_x, scene_rect.top() + offset_y)
        if self._external_geom_cb:
            self._external_geom_cb()

    def _fit_page(self):
        if self._pixmap_item:
            self.fitInView(self._pixmap_item, Qt.KeepAspectRatio)

    def wheelEvent(self, event: QWheelEvent):
        if event.modifiers() & Qt.ControlModifier:
            angle = event.angleDelta().y()
            factor = 1.25 if angle > 0 else 0.8
            self.scale(factor, factor)
        else:
            super().wheelEvent(event)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Space:
            self._space_pressed = True
            self.setCursor(Qt.OpenHandCursor)
        super().keyPressEvent(event)

    def keyReleaseEvent(self, event):
        if event.key() == Qt.Key_Space:
            self._space_pressed = False
            self._panning = False
            self.setCursor(Qt.ArrowCursor)
        super().keyReleaseEvent(event)

    def mousePressEvent(self, event: QMouseEvent):
        if self._space_pressed and event.button() == Qt.LeftButton:
            self._panning = True
            self._last_pan_pos = event.pos()
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event: QMouseEvent):
        if self._panning and self._last_pan_pos is not None:
            delta = event.pos() - self._last_pan_pos
            self._last_pan_pos = event.pos()
            self.horizontalScrollBar().setValue(self.horizontalScrollBar().value() - delta.x())
            self.verticalScrollBar().setValue(self.verticalScrollBar().value() - delta.y())
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event: QMouseEvent):
        if self._panning and event.button() == Qt.LeftButton:
            self._panning = False
            self.setCursor(Qt.OpenHandCursor if self._space_pressed else Qt.ArrowCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def get_stamp_rect_pdf_coords(self) -> Optional[fitz.Rect]:
        if not self._pixmap_item or not self.stamp_item:
            return None
        scene_rect = self.stamp_item.mapRectToScene(self.stamp_item.rect())
        page_rect = self._pixmap_item.sceneBoundingRect()
        if page_rect.width() <= 0 or page_rect.height() <= 0:
            return None
        x0 = (scene_rect.left() - page_rect.left()) / self.zoom
        y0 = (scene_rect.top() - page_rect.top()) / self.zoom
        x1 = x0 + scene_rect.width() / self.zoom
        y1 = y0 + scene_rect.height() / self.zoom
        return fitz.Rect(x0, y0, x1, y1)


class HeaderSettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки шапки штампа")

        cfg = load_header_config()
        image_path = cfg.get("image_path", "") or ""
        header_text = cfg.get("header_text", "") or ""

        self.edit_image = QLineEdit(image_path)
        self.btn_browse = QPushButton("…")
        self.btn_browse.setFixedWidth(32)
        self.btn_browse.clicked.connect(self.on_browse_clicked)

        self.edit_text = QPlainTextEdit(header_text)
        self.edit_text.setLineWrapMode(QPlainTextEdit.WidgetWidth)
        self.edit_text.setTabChangesFocus(True)
        self.edit_text.setPlaceholderText("Текст шапки (будет разбит максимум на 3 строки)")

        self.btn_clear = QPushButton("Удалить изображение и текст")
        self.btn_clear.clicked.connect(self.on_clear_clicked)

        form_layout = QGridLayout()
        row = 0

        form_layout.addWidget(QLabel("Картинка (PNG, квадратная):"), row, 0)
        img_layout = QHBoxLayout()
        img_layout.addWidget(self.edit_image, 1)
        img_layout.addWidget(self.btn_browse)
        form_layout.addLayout(img_layout, row, 1)
        row += 1

        form_layout.addWidget(QLabel("Текст шапки:"), row, 0)
        form_layout.addWidget(self.edit_text, row, 1)
        row += 1

        form_layout.addWidget(self.btn_clear, row, 1)
        row += 1

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        ok_btn = buttons.button(QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.Cancel)
        if ok_btn:
            ok_btn.setText("Ок")
        if cancel_btn:
            cancel_btn.setText("Отменить")

        main_layout = QVBoxLayout(self)
        main_layout.addLayout(form_layout)
        main_layout.addWidget(buttons)

    def on_browse_clicked(self):
        start_dir = ""
        if self.edit_image.text():
            start_dir = os.path.dirname(self.edit_image.text())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите картинку PNG",
            start_dir,
            "PNG файлы (*.png);;Все файлы (*.*)",
        )
        if path:
            self.edit_image.setText(path)

    def on_clear_clicked(self):
        self.edit_image.clear()
        self.edit_text.clear()

    def accept(self):
        cfg_old = load_header_config()
        cfg_old["image_path"] = self.edit_image.text().strip()
        cfg_old["header_text"] = self.edit_text.toPlainText().strip()
        save_header_config(cfg_old)
        super().accept()


class SignDialog(QDialog):
    """Диалог для выбора сертификата и ключа перед подписью PDF."""

    def __init__(self, pdf_path: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Подписать PDF ЭЦП")
        self.pdf_path = pdf_path
        self._result = None

        tabs = QTabWidget()
        tabs.addTab(self._build_files_tab(), "Файлы сертификата и ключа")
        tabs.addTab(self._build_pkcs11_tab(), "Токен PKCS#11")
        self.browser_tab_index = tabs.addTab(
            self._build_browser_tab(), "Через браузер (CryptoPro)"
        )
        self.tabs = tabs

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self._validate_and_accept)
        buttons.rejected.connect(self.reject)

        ok_btn = buttons.button(QDialogButtonBox.Ok)
        cancel_btn = buttons.button(QDialogButtonBox.Cancel)
        if ok_btn:
            ok_btn.setText("Подписать")
        if cancel_btn:
            cancel_btn.setText("Отмена")

        layout = QVBoxLayout(self)
        layout.addWidget(tabs)
        layout.addWidget(buttons)

    # --- вкладка файла ---
    def _build_files_tab(self):
        widget = QWidget()
        form = QFormLayout(widget)
        form.addRow("Файл:", QLabel(os.path.basename(self.pdf_path)))

        self.cert_edit = QLineEdit()
        self.cert_btn = QPushButton("…")
        self.cert_btn.setFixedWidth(32)
        self.cert_btn.clicked.connect(self._browse_cert)

        cert_layout = QHBoxLayout()
        cert_layout.addWidget(self.cert_edit, 1)
        cert_layout.addWidget(self.cert_btn)
        form.addRow("Сертификат (.cer/.pem):", cert_layout)

        self.key_edit = QLineEdit()
        self.key_btn = QPushButton("…")
        self.key_btn.setFixedWidth(32)
        self.key_btn.clicked.connect(self._browse_key)

        key_layout = QHBoxLayout()
        key_layout.addWidget(self.key_edit, 1)
        key_layout.addWidget(self.key_btn)
        form.addRow("Закрытый ключ (.key/.pem):", key_layout)

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        form.addRow("Пароль к ключу (если есть):", self.password_edit)

        form.addItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))
        return widget

    # --- вкладка браузера ---
    def _build_browser_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.addWidget(QLabel(f"Файл: <b>{os.path.basename(self.pdf_path)}</b>"))
        layout.addWidget(
            QLabel(
                "Откроется страница в браузере с плагином CryptoPro. "
                "Выберите сертификат, введите PIN/пароль по запросу и подтвердите подпись."
            )
        )
        layout.addWidget(
            QLabel(
                "Если плагин отсутствует или браузер не поддерживается, вернитесь в диалог и "
                "выберите другой способ подписи."
            )
        )
        layout.addStretch(1)
        return widget

    # --- вкладка токена ---
    def _build_pkcs11_tab(self):
        widget = QWidget()
        form = QFormLayout(widget)
        form.addRow("Файл:", QLabel(os.path.basename(self.pdf_path)))

        self.pkcs11_lib_edit = QLineEdit()
        self.pkcs11_lib_btn = QPushButton("…")
        self.pkcs11_lib_btn.setFixedWidth(32)
        self.pkcs11_lib_btn.clicked.connect(self._browse_pkcs11)
        lib_layout = QHBoxLayout()
        lib_layout.addWidget(self.pkcs11_lib_edit, 1)
        lib_layout.addWidget(self.pkcs11_lib_btn)
        form.addRow("Библиотека PKCS#11:", lib_layout)

        self.token_label_edit = QLineEdit()
        form.addRow("Метка токена (опц.):", self.token_label_edit)

        self.slot_spin = QSpinBox()
        self.slot_spin.setRange(-1, 1000)
        self.slot_spin.setValue(-1)
        self.slot_spin.setSpecialValueText("Авто")
        form.addRow("Слот (опц.):", self.slot_spin)

        self.key_label_edit = QLineEdit()
        form.addRow("Метка ключа (опц.):", self.key_label_edit)

        self.pin_edit = QLineEdit()
        self.pin_edit.setEchoMode(QLineEdit.Password)
        form.addRow("PIN токена:", self.pin_edit)

        self.token_cert_edit = QLineEdit()
        self.token_cert_btn = QPushButton("…")
        self.token_cert_btn.setFixedWidth(32)
        self.token_cert_btn.clicked.connect(self._browse_token_cert)
        token_cert_layout = QHBoxLayout()
        token_cert_layout.addWidget(self.token_cert_edit, 1)
        token_cert_layout.addWidget(self.token_cert_btn)
        form.addRow("Сертификат (если не на токене):", token_cert_layout)

        form.addItem(QSpacerItem(0, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))
        return widget

    def _browse_cert(self):
        start_dir = os.path.dirname(self.pdf_path) if os.path.isfile(self.pdf_path) else ""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл сертификата",
            start_dir,
            "Сертификаты (*.cer *.pem *.crt);;Все файлы (*.*)",
        )
        if path:
            self.cert_edit.setText(path)

    def _browse_key(self):
        start_dir = os.path.dirname(self.pdf_path) if os.path.isfile(self.pdf_path) else ""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите закрытый ключ",
            start_dir,
            "Ключи (*.key *.pem *.der);;Все файлы (*.*)",
        )
        if path:
            self.key_edit.setText(path)

    def _browse_pkcs11(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Укажите библиотеку PKCS#11",
            "",
            "*.dll *.so *.dylib ;;Все файлы (*.*)",
        )
        if path:
            self.pkcs11_lib_edit.setText(path)

    def _browse_token_cert(self):
        start_dir = os.path.dirname(self.pdf_path) if os.path.isfile(self.pdf_path) else ""
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл сертификата",
            start_dir,
            "Сертификаты (*.cer *.pem *.crt);;Все файлы (*.*)",
        )
        if path:
            self.token_cert_edit.setText(path)

    def _validate_and_accept(self):
        if self.tabs.currentIndex() == 0:
            cert_path = self.cert_edit.text().strip()
            key_path = self.key_edit.text().strip()
            if not cert_path or not os.path.exists(cert_path):
                QMessageBox.warning(self, "Нет сертификата", "Укажите путь к файлу сертификата.")
                return
            if not key_path or not os.path.exists(key_path):
                QMessageBox.warning(self, "Нет ключа", "Укажите путь к файлу закрытого ключа.")
                return
            self._result = {
                "mode": "files",
                "cert_path": cert_path,
                "key_path": key_path,
                "password": self.password_edit.text(),
            }
        elif self.tabs.currentIndex() == self.browser_tab_index:
            self._result = {"mode": "browser"}
        else:
            pkcs11_path = self.pkcs11_lib_edit.text().strip()
            pin = self.pin_edit.text()
            slot = self.slot_spin.value()
            slot_value = None if slot < 0 else slot
            if not pkcs11_path or not os.path.exists(pkcs11_path):
                QMessageBox.warning(
                    self, "Библиотека PKCS#11", "Укажите путь к библиотеке PKCS#11."
                )
                return
            if not pin:
                QMessageBox.warning(self, "PIN", "Введите PIN токена для подписи.")
                return
            self._result = {
                "mode": "pkcs11",
                "pkcs11_path": pkcs11_path,
                "token_label": self.token_label_edit.text().strip() or None,
                "slot": slot_value,
                "key_label": self.key_label_edit.text().strip() or None,
                "pin": pin,
                "cert_path": self.token_cert_edit.text().strip() or None,
            }
        super().accept()

    def get_result(self):
        return self._result


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Визуализация ЭЦП в PDF")
        self.resize(1400, 800)

        icon_path = get_resource_path("app.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.sessions: List[FileSession] = []
        self.current_session_index: int = -1

        self.pdf_path: Optional[str] = None
        self.p7s_path: Optional[str] = None
        self.doc: Optional[fitz.Document] = None
        self.current_page_index: int = 0
        self.cert_info: Optional[CertificateInfo] = None

        self.orphan_p7s: List[str] = []

        self.last_stamp_norm_rect: Optional[QRectF] = None

        self.default_output_dir: Optional[str] = None

        cfg = load_header_config()
        self.show_sign_time_on_stamp = bool(cfg.get("show_sign_time", False))
        self.default_output_dir = (cfg.get("default_output_dir") or "").strip() or None

        rect_cfg = cfg.get("stamp_rect", {})
        try:
            left = float(rect_cfg.get("left", 0.0))
            top = float(rect_cfg.get("top", 0.0))
            width = float(rect_cfg.get("width", 0.0))
            height = float(rect_cfg.get("height", 0.0))
            self.last_stamp_norm_rect = QRectF(left, top, width, height)
        except Exception:
            self.last_stamp_norm_rect = None

        self._setup_ui()
        logger.info("Главное окно создано")

    # ---------- утилита для стартовой директории ----------

    def _get_default_browse_dir(self) -> str:
        """
        Стартовая директория для диалогов выбора файлов/папок:
        1) если уже выбирали папку сохранения — она;
        2) иначе Desktop, если существует;
        3) иначе текущая рабочая директория.
        """
        if self.default_output_dir:
            return self.default_output_dir
        home = os.path.expanduser("~")
        desktop = os.path.join(home, "Desktop")
        if os.path.isdir(desktop):
            return desktop
        return os.getcwd()

    def _persist_settings(self):
        payload: Dict[str, object] = {
            "show_sign_time": self.show_sign_time_on_stamp,
            "default_output_dir": self.default_output_dir or "",
        }

        if self.last_stamp_norm_rect is not None:
            payload["stamp_rect"] = {
                "left": self.last_stamp_norm_rect.left(),
                "top": self.last_stamp_norm_rect.top(),
                "width": self.last_stamp_norm_rect.width(),
                "height": self.last_stamp_norm_rect.height(),
            }

        save_header_config(payload)

    # ---------- UI ----------

    def _setup_ui(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")
        act_exit = QAction("Выйти", self)
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

        settings_menu = menubar.addMenu("Настройки")
        act_header = QAction("Шапка штампа…", self)
        act_header.triggered.connect(self.on_edit_header_settings)
        settings_menu.addAction(act_header)

        help_menu = menubar.addMenu("Помощь")
        act_help = QAction("Инструкция", self)
        act_help.triggered.connect(self.show_help)
        help_menu.addAction(act_help)
        act_about = QAction("О программе", self)
        act_about.triggered.connect(self.show_about)
        help_menu.addAction(act_about)

        central = QWidget(self)
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(4, 4, 4, 4)

        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(6)
        main_layout.addWidget(splitter)

        # Левая панель: кнопки + список файлов
        files_widget = QWidget()
        files_layout = QVBoxLayout(files_widget)
        files_layout.setContentsMargins(4, 4, 4, 4)

        self.btn_add_files = QPushButton("Добавить файлы")
        self.btn_add_files.clicked.connect(self.on_add_files_clicked)
        files_layout.addWidget(self.btn_add_files)

        self.btn_add_folder = QPushButton("Добавить папку")
        self.btn_add_folder.clicked.connect(self.on_add_folder_clicked)
        files_layout.addWidget(self.btn_add_folder)

        self.btn_remove_file = QPushButton("Удалить выбранный")
        self.btn_remove_file.clicked.connect(self.on_remove_file_clicked)
        files_layout.addWidget(self.btn_remove_file)

        self.btn_remove_all = QPushButton("Удалить все")
        self.btn_remove_all.clicked.connect(self.on_remove_all_clicked)
        files_layout.addWidget(self.btn_remove_all)

        self.file_list = QListWidget()
        self.file_list.currentRowChanged.connect(self.on_file_selected)
        files_layout.addWidget(self.file_list, 1)

        files_widget.setMinimumWidth(200)
        splitter.addWidget(files_widget)

        # Панель миниатюр
        thumbs_widget = QWidget()
        thumbs_layout = QVBoxLayout(thumbs_widget)
        thumbs_layout.setContentsMargins(4, 4, 4, 4)

        self.thumb_list = QListWidget()
        self.thumb_list.setIconSize(QSize(80, 100))
        self.thumb_list.setFrameShape(QFrame.NoFrame)
        self.thumb_list.setStyleSheet("border: 1px solid #d0d0d0; border-radius: 4px;")
        self.thumb_list.currentRowChanged.connect(self.on_thumbnail_selected)
        thumbs_layout.addWidget(self.thumb_list, 1)

        thumbs_widget.setMinimumWidth(180)
        splitter.addWidget(thumbs_widget)
        self.thumbs_widget = thumbs_widget

        # Центральная область: страница
        self.page_view = PDFPageView()
        self.page_view.setFrameShape(QFrame.NoFrame)
        self.page_view.setStyleSheet("border: 1px solid #d0d0d0; border-radius: 4px;")
        self.page_view.set_external_geometry_callback(self.on_stamp_geometry_changed)
        splitter.addWidget(self.page_view)

        # Правая панель: инфо и сохранение
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(4, 4, 4, 4)

        info_label = QLabel("<b>Сведения о сертификате</b>")
        right_layout.addWidget(info_label)

        form = QFormLayout()
        self.lbl_serial = QLabel("—")
        self.lbl_owner = QLabel("—")
        self.lbl_sign_time = QLabel("—")
        self.lbl_valid_period = QLabel("—")
        self.lbl_status = QLabel("—")

        form.addRow("Сертификат:", self.lbl_serial)
        form.addRow("Владелец:", self.lbl_owner)
        form.addRow("Время подписи:", self.lbl_sign_time)
        form.addRow("Действителен:", self.lbl_valid_period)
        form.addRow("Статус подписи:", self.lbl_status)

        right_layout.addLayout(form)

        right_layout.addSpacing(10)
        right_layout.addWidget(QLabel("<b>Сохранение</b>"))

        out_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_btn = QPushButton("…")
        self.output_dir_btn.setFixedWidth(30)
        self.output_dir_btn.clicked.connect(self.on_browse_output_dir)
        if self.default_output_dir:
            self.output_dir_edit.setText(self.default_output_dir)
        out_layout.addWidget(self.output_dir_edit, 1)
        out_layout.addWidget(self.output_dir_btn)
        right_layout.addWidget(QLabel("Папка сохранения:"))
        right_layout.addLayout(out_layout)

        self.chk_save_to_source = QCheckBox("Сохранять файл в исходное месторасположение")
        self.chk_save_to_source.setChecked(True)
        self.chk_save_to_source.toggled.connect(self.on_save_to_source_toggled)
        right_layout.addWidget(self.chk_save_to_source)

        self.chk_show_sign_time = QCheckBox("Показывать время подписи на штампе")
        self.chk_show_sign_time.setChecked(self.show_sign_time_on_stamp)
        self.chk_show_sign_time.toggled.connect(self.on_show_sign_time_toggled)
        right_layout.addWidget(self.chk_show_sign_time)

        self.btn_sign_pdf = QPushButton("Подписать текущий PDF")
        self.btn_sign_pdf.setEnabled(False)
        self.btn_sign_pdf.clicked.connect(self.on_sign_pdf_clicked)
        right_layout.addWidget(self.btn_sign_pdf)

        self.btn_save = QPushButton("Сохранить файл с ЭЦП")
        self.btn_save.setEnabled(False)
        self.btn_save.clicked.connect(self.on_save_clicked)
        right_layout.addWidget(self.btn_save)

        right_layout.addStretch(1)

        right_widget.setMinimumWidth(260)
        right_widget.setMaximumWidth(420)
        splitter.addWidget(right_widget)

        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 0)
        splitter.setStretchFactor(2, 1)
        splitter.setStretchFactor(3, 0)
        splitter.setSizes([220, 220, 900, 320])

        self.on_save_to_source_toggled(self.chk_save_to_source.isChecked())

        self.setAcceptDrops(True)
        self.statusBar().showMessage("Перетащите файлы или нажмите «Добавить файлы»")

    # ---------- меню / помощь / настройки ----------

    def show_about(self):
        QMessageBox.about(
            self,
            "О программе",
            (
                "<b>Визуализация ЭЦП в PDF</b><br><br>"
                "Автор: Белобородов Алексей Александрович<br>"
                "Версия: 2.3.3<br>"
                "Дата создания: 30.11.2025"
            ),
        )

    def show_help(self):
        chm_path = get_resource_path("help.chm")
        html_path = get_resource_path("help.html")

        if os.path.exists(chm_path):
            try:
                os.startfile(chm_path)
                return
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "Инструкция",
                    f"Не удалось открыть help.chm:\n{e}",
                )

        if not os.path.exists(html_path):
            QMessageBox.warning(
                self,
                "Инструкция",
                "Файл справки 'help.html' не найден.\n"
                "Положите его в папку с программой.",
            )
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Инструкция")
        dlg.resize(900, 650)

        layout = QVBoxLayout(dlg)
        browser = QTextBrowser()
        browser.setOpenExternalLinks(True)
        browser.setSource(QUrl.fromLocalFile(html_path))
        layout.addWidget(browser)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.accepted.connect(dlg.accept)
        buttons.rejected.connect(dlg.reject)
        close_btn = buttons.button(QDialogButtonBox.Close)
        if close_btn:
            close_btn.setText("Закрыть")
        layout.addWidget(buttons)

        dlg.exec()

    def on_edit_header_settings(self):
        dlg = HeaderSettingsDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self.update_stamp_preview()

    def _create_browser_wait_dialog(self, browser_session: BrowserSigningSession, url: str) -> QDialog:
        dlg = QDialog(self)
        dlg.setWindowTitle("Подпись через браузер (CryptoPro)")
        layout = QVBoxLayout(dlg)
        layout.addWidget(
            QLabel(
                "В браузере открылось окно подписи через плагин CryptoPro. "
                "После выбора сертификата и ввода PIN вернитесь в приложение."
            )
        )
        link = QLabel(f"<a href=\"{url}\">Открыть страницу подписи ещё раз</a>")
        link.setOpenExternalLinks(True)
        layout.addWidget(link)
        status_label = QLabel("Ожидание ответа от браузера…")
        layout.addWidget(status_label)
        buttons = QDialogButtonBox(QDialogButtonBox.Cancel)
        buttons.rejected.connect(dlg.reject)
        layout.addWidget(buttons)

        timer = QTimer(dlg)

        def tick():
            if browser_session.is_finished():
                status_label.setText("Ответ получен, завершаем…")
                dlg.accept()

        timer.timeout.connect(tick)
        timer.start(300)
        dlg.finished.connect(timer.stop)
        return dlg

    def _sign_pdf_via_browser(self, pdf_path: str) -> str:
        with BrowserSigningSession(pdf_path) as browser_session:
            browser_session.start()
            url = browser_session.url()
            QDesktopServices.openUrl(QUrl(url))
            wait_dialog = self._create_browser_wait_dialog(browser_session, url)
            if wait_dialog.exec() != QDialog.Accepted:
                raise BrowserSigningError(
                    "Подпись через браузер отменена пользователем или прервана"
                )

            # Даем браузерному плагину достаточно времени для выбора сертификата
            # и создания подписи (по умолчанию 180 секунд в BrowserSigningSession).
            # Короткий таймаут приводил к преждевременному завершению сессии и
            # остановке встроенного HTTP-сервера, из-за чего на странице плагина
            # не появлялись логи и подпись не успевала формироваться.
            result = browser_session.wait()

        base_name, _ = os.path.splitext(os.path.basename(pdf_path))
        signature_name = f"{base_name}_Файл подписи.p7s"
        target_dir = os.path.dirname(pdf_path) or os.getcwd()
        os.makedirs(target_dir, exist_ok=True)
        signature_path = os.path.join(target_dir, signature_name)
        with open(signature_path, "wb") as f:
            f.write(result.signature)

        logger.info("Файл подписи создан через браузерный плагин: %s", signature_path)
        return signature_path

    def on_sign_pdf_clicked(self):
        if self.current_session_index < 0 or not self.doc:
            QMessageBox.warning(self, "Нет файла", "Сначала выберите PDF для подписи.")
            return

        session = self.sessions[self.current_session_index]
        dlg = SignDialog(session.pdf_path, self)
        if dlg.exec() != QDialog.Accepted:
            return

        result = dlg.get_result()
        if not result:
            return
        try:
            if result["mode"] == "files":
                signature_path = sign_pdf(
                    session.pdf_path,
                    result["cert_path"],
                    result["key_path"],
                    result["password"],
                )
            elif result["mode"] == "browser":
                signature_path = self._sign_pdf_via_browser(session.pdf_path)
            else:
                signature_path = sign_pdf_with_pkcs11(
                    session.pdf_path,
                    result["pkcs11_path"],
                    result["pin"],
                    token_label=result.get("token_label"),
                    slot=result.get("slot"),
                    key_label=result.get("key_label"),
                    cert_path=result.get("cert_path"),
                )
        except Exception as e:
            logger.exception("Ошибка создания подписи")
            QMessageBox.critical(
                self,
                "Ошибка подписи",
                f"Не удалось создать файл подписи:\n{e}",
            )
            return

        session.p7s_path = signature_path
        try:
            session.cert_info = get_certificate_info(session.pdf_path, signature_path, None)
            self.cert_info = session.cert_info
        except Exception as e:
            logger.exception("Ошибка чтения созданной подписи")
            session.cert_info = CertificateInfo(status=f"ошибка подписи: {e}")
            self.cert_info = session.cert_info
            QMessageBox.warning(
                self,
                "Предупреждение",
                "Подпись создана, но не удалось прочитать информацию о сертификате."
            )

        self.rebuild_file_list()
        for row in range(self.file_list.count()):
            it = self.file_list.item(row)
            if it and it.data(Qt.UserRole) == self.current_session_index:
                self.file_list.setCurrentRow(row)
                break
        self.switch_to_session(self.current_session_index)
        self.statusBar().showMessage(
            f"Создан файл подписи: {os.path.basename(signature_path)}",
            8000,
        )
        QMessageBox.information(
            self,
            "Подпись создана",
            f"Файл подписи создан рядом с PDF:\n{signature_path}\n\n"
            "Он автоматически добавлен в список для визуализации.",
        )

    # ---------- добавление / удаление файлов ----------

    def on_add_files_clicked(self):
        start_dir = self._get_default_browse_dir()
        sig_mask = " ".join(f"*{ext}" for ext in SIGNATURE_EXTS)
        filter_str = f"PDF и подписи (*.pdf {sig_mask});;Все файлы (*.*)"
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы PDF и подписи",
            start_dir,
            filter_str,
        )
        if not paths:
            return
        self.add_files_from_paths(paths)

    def on_add_folder_clicked(self):
        start_dir = self._get_default_browse_dir()
        directory = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку с PDF и подписями",
            start_dir,
        )

        if not directory:
            return

        paths: List[str] = []
        for root, _, files in os.walk(directory):
            for name in files:
                lower = name.lower()
                if lower.endswith(".pdf") or lower.endswith(SIGNATURE_EXTS):
                    paths.append(os.path.join(root, name))

        if not paths:
            QMessageBox.warning(self, "Файлы не найдены", "В выбранных папках нет PDF или файлов подписи.")
        else:
            self.add_files_from_paths(paths)

    def on_remove_file_clicked(self):
        row = self.file_list.currentRow()
        if row < 0:
            return
        item = self.file_list.item(row)
        if not item:
            return
        session_index = item.data(Qt.UserRole)
        if session_index is None:
            return

        session_index = int(session_index)
        if session_index < 0 or session_index >= len(self.sessions):
            return

        session = self.sessions.pop(session_index)
        try:
            session.doc.close()
        except Exception:
            pass

        if not self.sessions:
            self.current_session_index = -1
            self.file_list.clear()
            self.clear_current_view()
            return

        if self.current_session_index == session_index:
            self.current_session_index = -1

        self.rebuild_file_list()

    def on_remove_all_clicked(self):
        if not self.sessions:
            return
        ret = QMessageBox.question(
            self,
            "Удалить все",
            "Вы уверены, что хотите удалить все загруженные документы из списка?\n"
            "Исходные файлы на диске затронуты не будут.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if ret != QMessageBox.Yes:
            return

        for s in self.sessions:
            try:
                s.doc.close()
            except Exception:
                pass
        self.sessions.clear()
        self.orphan_p7s.clear()
        self.current_session_index = -1
        self.file_list.clear()
        self.clear_current_view()

    # drag & drop

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        paths = [u.toLocalFile() for u in event.mimeData().urls()]
        self.add_files_from_paths(paths)

    def add_files_from_paths(self, paths: List[str]):
        filtered: List[str] = []
        for p in paths:
            lower = p.lower()
            if lower.endswith(".pdf") or lower.endswith(SIGNATURE_EXTS):
                filtered.append(p)

        if not filtered:
            QMessageBox.warning(self, "Нет файлов", "Среди выбранных нет PDF или файлов подписи.")
            return

        total = len(filtered)
        pdf_count = 0
        p7s_count = 0
        first_pdf: Optional[str] = None

        for idx, p in enumerate(filtered, start=1):
            basename = os.path.basename(p)
            self.statusBar().showMessage(f"Обработка файла {idx} из {total}: {basename}")

            lower = p.lower()
            if lower.endswith(".pdf"):
                pdf_count += 1
                if first_pdf is None:
                    first_pdf = p
                self.load_single_pdf(p, rebuild_list=False)
            elif lower.endswith(SIGNATURE_EXTS):
                p7s_count += 1
                self._register_orphan_p7s(p)

        self.auto_match_signatures()

        if first_pdf and not self.default_output_dir:
            self.default_output_dir = os.path.dirname(first_pdf) or os.getcwd()
            if not self.chk_save_to_source.isChecked():
                self.output_dir_edit.setText(self.default_output_dir)
            self._persist_settings()

        if pdf_count and not p7s_count:
            self.statusBar().showMessage(
                "Добавлены PDF-файлы (для части может не быть подписи).",
                8000,
            )
        elif p7s_count and not pdf_count:
            self.statusBar().showMessage(
                "Добавлены файлы подписи, ожидаются соответствующие PDF.",
                8000,
            )
        else:
            self.statusBar().showMessage("Файлы добавлены.", 8000)

    def _register_orphan_p7s(self, p7s_path: str):
        abs_new = os.path.abspath(p7s_path)
        for s in self.sessions:
            if s.p7s_path and os.path.abspath(s.p7s_path) == abs_new:
                return
        for existing in self.orphan_p7s:
            if os.path.abspath(existing) == abs_new:
                return
        self.orphan_p7s.append(p7s_path)

    def load_single_pdf(self, pdf_path: str, rebuild_list: bool = True):
        logger.info("Добавление PDF: %s", pdf_path)
        for s in self.sessions:
            if os.path.abspath(s.pdf_path) == os.path.abspath(pdf_path):
                logger.info("PDF уже загружен, пропускаем: %s", pdf_path)
                return
        try:
            doc = open_document(pdf_path)
        except Exception:
            logger.exception("Не удалось открыть PDF: %s", pdf_path)
            QMessageBox.critical(self, "Ошибка PDF", f"Не удалось открыть PDF:\n{pdf_path}")
            return
        cert_info = CertificateInfo()
        cert_info.status = "нет файла подписи"
        session = FileSession(
            pdf_path=pdf_path,
            p7s_path=None,
            doc=doc,
            cert_info=cert_info,
            output_dir=os.path.dirname(pdf_path) or None,
        )
        self.sessions.append(session)

        if self.current_session_index == -1:
            self.current_session_index = len(self.sessions) - 1

        if rebuild_list:
            self.rebuild_file_list()

    # ---------- сопоставление подписей ----------

    @staticmethod
    def _is_signature_pair_match(info: CertificateInfo) -> bool:
        status = (info.status or "").lower()
        if not status:
            return False
        bad_words = [
            "не соответствует",
            "нет файла подписи",
            "не найдена подпись",
            "ошибка",
            "подпись отсутствует",
        ]
        return not any(w in status for w in bad_words)

    def _has_valid_signature(self, info: Optional[CertificateInfo]) -> bool:
        if not info:
            return False
        status = (info.status or "").lower()
        if not status.strip():
            return False
        bad_words = [
            "нет файла подписи",
            "не найдена подпись",
            "ошибка",
            "не соответствует",
            "не подтверждена",
        ]
        return not any(w in status for w in bad_words)

    def auto_match_signatures(self):
        if not self.orphan_p7s or not self.sessions:
            self.rebuild_file_list()
            return

        p7s_candidates = list(self.orphan_p7s)
        self.orphan_p7s = []
        total_signatures = len(p7s_candidates)
        processed_signatures = 0

        for idx, p7s_path in enumerate(p7s_candidates, start=1):
            self.statusBar().showMessage(
                f"Проверка подписи {idx} из {total_signatures}: {os.path.basename(p7s_path)}"
            )
            QApplication.processEvents()
            matched = False
            p7s_dir = os.path.dirname(p7s_path)

            same_dir = [s for s in self.sessions if s.p7s_path is None and os.path.dirname(s.pdf_path) == p7s_dir]
            other_dir = [s for s in self.sessions if s.p7s_path is None and os.path.dirname(s.pdf_path) != p7s_dir]

            for group in (same_dir, other_dir):
                if matched:
                    break
                for session in group:
                    try:
                        info = get_certificate_info(session.pdf_path, p7s_path, None)
                    except Exception:
                        logger.exception(
                            "Ошибка проверки пары PDF+подпись: %s + %s",
                            session.pdf_path,
                            p7s_path,
                        )
                        continue

                    if self._is_signature_pair_match(info):
                        logger.info("Подпись %s сопоставлена с PDF %s", p7s_path, session.pdf_path)
                        session.p7s_path = p7s_path
                        session.cert_info = info
                        matched = True
                        break

            if not matched:
                self.orphan_p7s.append(p7s_path)

            processed_signatures += 1
            self.statusBar().showMessage(
                f"Проверено подписей: {processed_signatures} из {total_signatures}"
            )

        self.rebuild_file_list()

        if 0 <= self.current_session_index < len(self.sessions):
            self.cert_info = self.sessions[self.current_session_index].cert_info
            self.update_cert_info_panel()
            self.update_stamp_preview()

    # ---------- список файлов (левая панель) ----------

    def _make_file_list_text(self, session_index: int) -> str:
        s = self.sessions[session_index]
        base = os.path.basename(s.pdf_path)
        prefix = "✔ " if s.saved else "  "
        return f"{prefix}{base}"

    def rebuild_file_list(self):
        self.file_list.blockSignals(True)
        self.file_list.clear()

        if not self.sessions:
            self.file_list.blockSignals(False)
            return

        def add_header(text: str):
            item = QListWidgetItem(text)
            font = item.font()
            font.setBold(True)
            item.setFont(font)
            flags = item.flags()
            flags &= ~Qt.ItemIsSelectable
            flags &= ~Qt.ItemIsDragEnabled
            item.setFlags(flags)
            item.setData(Qt.UserRole, None)
            self.file_list.addItem(item)

        current_index = self.current_session_index
        item_to_select = None

        has_with = any(s.p7s_path is not None for s in self.sessions)
        has_without = any(s.p7s_path is None for s in self.sessions)

        if has_with:
            add_header("Найдены подписи:")
            for idx, s in enumerate(self.sessions):
                if s.p7s_path is not None:
                    item = QListWidgetItem(self._make_file_list_text(idx))
                    item.setData(Qt.UserRole, idx)
                    self.file_list.addItem(item)
                    if idx == current_index:
                        item_to_select = item

        if has_without:
            add_header("Не найдены подписи:")
            for idx, s in enumerate(self.sessions):
                if s.p7s_path is None:
                    item = QListWidgetItem(self._make_file_list_text(idx))
                    item.setData(Qt.UserRole, idx)
                    self.file_list.addItem(item)
                    if idx == current_index:
                        item_to_select = item

        self.file_list.blockSignals(False)

        if item_to_select is not None:
            self.file_list.setCurrentItem(item_to_select)
        else:
            for row in range(self.file_list.count()):
                it = self.file_list.item(row)
                if it and it.data(Qt.UserRole) is not None:
                    self.file_list.setCurrentRow(row)
                    break

    def update_file_list_item_status(self, index: int):
        self.rebuild_file_list()

    def on_file_selected(self, row: int):
        if row < 0:
            return
        item = self.file_list.item(row)
        if not item:
            return
        session_index = item.data(Qt.UserRole)
        if session_index is None:
            for direction in (1, -1):
                r = row + direction
                while 0 <= r < self.file_list.count():
                    it = self.file_list.item(r)
                    if it and it.data(Qt.UserRole) is not None:
                        self.file_list.setCurrentRow(r)
                        return
                    r += direction
            return

        if session_index == self.current_session_index and self.doc is not None:
            return

        session_index = int(session_index)
        if 0 <= session_index < len(self.sessions):
            self.switch_to_session(session_index)

    def switch_to_session(self, index: int):
        self.current_session_index = index
        session = self.sessions[index]

        self.pdf_path = session.pdf_path
        self.p7s_path = session.p7s_path
        self.doc = session.doc

        # всегда показываем первую страницу при переключении файла
        self.current_page_index = 0
        session.current_page_index = 0

        self.cert_info = session.cert_info

        if not self.chk_save_to_source.isChecked():
            if self.default_output_dir:
                self.output_dir_edit.setText(self.default_output_dir)

        self.populate_thumbnails()
        self.show_page(0)
        self.update_cert_info_panel()

    # ---------- миниатюры и страницы ----------

    def populate_thumbnails(self):
        self.thumb_list.clear()
        if not self.doc:
            self.thumbs_widget.setVisible(False)
            return

        page_count = len(self.doc)
        self.thumbs_widget.setVisible(True)

        for page_index in range(page_count):
            page = self.doc[page_index]
            pix = page.get_pixmap(matrix=fitz.Matrix(0.3, 0.3))
            img = QImage(
                pix.samples,
                pix.width,
                pix.height,
                pix.stride,
                QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888,
            )
            qpix = QPixmap.fromImage(img)

            item = QListWidgetItem()
            item.setIcon(qpix)
            item.setText(str(page_index + 1))
            self.thumb_list.addItem(item)

        self.thumb_list.setCurrentRow(0)

    def on_thumbnail_selected(self, row: int):
        if row < 0 or not self.doc:
            return
        self.current_page_index = row
        if 0 <= self.current_session_index < len(self.sessions):
            self.sessions[self.current_session_index].current_page_index = row
        self.show_page(row)

    def show_page(self, index: int):
        if not self.doc:
            return
        if index < 0 or index >= len(self.doc):
            index = 0

        page = self.doc[index]
        matrix = fitz.Matrix(PDF_RENDER_ZOOM, PDF_RENDER_ZOOM)
        pix = page.get_pixmap(matrix=matrix)
        img = QImage(
            pix.samples,
            pix.width,
            pix.height,
            pix.stride,
            QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888,
        )
        qpix = QPixmap.fromImage(img)

        logger.debug("Отображаем страницу %d", index + 1)
        self.page_view.set_page(qpix, PDF_RENDER_ZOOM)

        if self.last_stamp_norm_rect is not None:
            self.page_view.set_stamp_rect_normalized(self.last_stamp_norm_rect)
        else:
            self.on_stamp_geometry_changed()

        self.update_stamp_preview()

    # ---------- сведения о сертификате и штамп ----------

    def update_cert_info_panel(self):
        info = self.cert_info or CertificateInfo()

        serial = info.serial_number or "не удалось определить"
        owner = info.subject or "не удалось определить"
        sign_time = info.signing_time or "не удалось определить"
        valid_from = info.valid_from or "не удалось определить"
        valid_to = info.valid_to or "не удалось определить"
        status = info.status or "не удалось определить"

        self.lbl_serial.setText(serial)
        self.lbl_owner.setText(owner)
        self.lbl_sign_time.setText(sign_time)
        self.lbl_valid_period.setText(f"с {valid_from} по {valid_to}")
        self.lbl_status.setText(status)

        status_lower = status.lower()
        if "действительн" in status_lower:
            self.lbl_status.setStyleSheet("color: green;")
        elif "не найдена подпись" in status_lower or "нет файла подписи" in status_lower:
            self.lbl_status.setStyleSheet("color: orange;")
        elif status_lower:
            self.lbl_status.setStyleSheet("color: red;")
        else:
            self.lbl_status.setStyleSheet("")

        has_sig = self._has_valid_signature(self.cert_info)
        self.btn_save.setEnabled(bool(self.doc) and has_sig)
        self.btn_sign_pdf.setEnabled(bool(self.doc))

    def _make_stamp_info_dict(self) -> Dict[str, str]:
        info = self.cert_info or CertificateInfo()
        return {
            "serial_number": info.serial_number or "не удалось определить",
            "subject": info.subject or "не удалось определить",
            "issuer": info.issuer or "не удалось определить",
            "valid_from": info.valid_from or "не удалось определить",
            "valid_to": info.valid_to or "не удалось определить",
            "signing_time": info.signing_time or "не удалось определить",
            "status": info.status or "не удалось определить",
        }

    def on_stamp_geometry_changed(self):
        norm = self.page_view.get_stamp_rect_normalized()
        if norm is not None:
            self.last_stamp_norm_rect = norm
            self._persist_settings()

    def update_stamp_preview(self):
        if not self.doc:
            return

        if not self._has_valid_signature(self.cert_info):
            self.page_view.hide_stamp()
            self.btn_save.setEnabled(False)
            return

        self.page_view.ensure_stamp_item()
        if self.last_stamp_norm_rect is not None:
            self.page_view.set_stamp_rect_normalized(self.last_stamp_norm_rect)

        stamp_info = self._make_stamp_info_dict()
        img = build_stamp_image(stamp_info, show_sign_time=self.show_sign_time_on_stamp)
        qpix = QPixmap.fromImage(img)
        self.page_view.set_stamp_pixmap(qpix)
        self.btn_save.setEnabled(True)

    def on_show_sign_time_toggled(self, checked: bool):
        self.show_sign_time_on_stamp = checked
        self._persist_settings()
        self.update_stamp_preview()

    # ---------- сохранение ----------

    def on_save_to_source_toggled(self, checked: bool):
        self.output_dir_edit.setEnabled(not checked)
        self.output_dir_btn.setEnabled(not checked)

        if not checked:
            if self.default_output_dir:
                self.output_dir_edit.setText(self.default_output_dir)
            elif self.current_session_index >= 0:
                cur_pdf = self.sessions[self.current_session_index].pdf_path
                self.output_dir_edit.setText(os.path.dirname(cur_pdf))

    def on_browse_output_dir(self):
        if self.chk_save_to_source.isChecked():
            return

        current = ""
        if self.output_dir_edit.text():
            current = self.output_dir_edit.text()
        elif self.default_output_dir:
            current = self.default_output_dir
        elif self.current_session_index >= 0:
            current = os.path.dirname(self.sessions[self.current_session_index].pdf_path)
        else:
            current = self._get_default_browse_dir()

        directory = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения",
            current,
        )
        if directory:
            self.output_dir_edit.setText(directory)
            self.default_output_dir = directory
            self._persist_settings()

    def on_save_clicked(self):
        if not self.doc or self.current_session_index < 0:
            return
        if not self._has_valid_signature(self.cert_info):
            QMessageBox.warning(self, "Подпись", "Для этого файла не найдена корректная ЭЦП.")
            return

        session = self.sessions[self.current_session_index]

        if self.chk_save_to_source.isChecked():
            output_dir = os.path.dirname(session.pdf_path) or os.getcwd()
        else:
            output_dir = self.output_dir_edit.text().strip()
            if not output_dir:
                QMessageBox.warning(
                    self,
                    "Папка сохранения",
                    "Выберите папку для сохранения или включите режим "
                    "«Сохранять файл в исходное месторасположение».",
                )
                self.on_browse_output_dir()
                output_dir = self.output_dir_edit.text().strip()
                if not output_dir:
                    return
            self.default_output_dir = output_dir

        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка папки", f"Не удалось использовать папку:\n{e}")
            return

        rect_pdf = self.page_view.get_stamp_rect_pdf_coords()
        if rect_pdf is None:
            QMessageBox.warning(self, "Штамп", "Не удалось определить положение штампа.")
            return

        norm_rect = self.page_view.get_stamp_rect_normalized()
        if norm_rect is not None:
            self.last_stamp_norm_rect = norm_rect

        base_name = os.path.splitext(os.path.basename(self.pdf_path or "document"))[0]
        timestamp = datetime.now().strftime("%d.%m.%Y_%H.%M.%S")
        out_name = f"{base_name}_ЭЦП_{timestamp}.pdf"
        out_path = os.path.join(output_dir, out_name)

        stamp_info = self._make_stamp_info_dict()
        try:
            add_stamp_to_pdf(
                self.pdf_path,
                out_path,
                self.current_page_index,
                rect_pdf,
                stamp_info,
                show_sign_time=self.show_sign_time_on_stamp,
            )
        except Exception:
            logger.exception("Ошибка при сохранении PDF со штампом")
            QMessageBox.critical(self, "Ошибка сохранения", "Не удалось сохранить PDF.")
            return

        self._persist_settings()

        session.saved = True
        self.update_file_list_item_status(self.current_session_index)

        msg = f"Сохранён файл: {out_path}"
        logger.info(msg)
        self.statusBar().showMessage(msg, 8000)

        # автоматически переключаемся на следующий файл с корректной подписью
        valid_indices = [
            i for i, s in enumerate(self.sessions)
            if self._has_valid_signature(s.cert_info)
        ]
        if self.current_session_index in valid_indices:
            pos = valid_indices.index(self.current_session_index)
            if pos + 1 < len(valid_indices):
                next_index = valid_indices[pos + 1]
                for row in range(self.file_list.count()):
                    it = self.file_list.item(row)
                    if it and it.data(Qt.UserRole) == next_index:
                        self.file_list.setCurrentRow(row)
                        break
            else:
                self.statusBar().showMessage(
                    "Все файлы с найденными подписями обработаны",
                    8000,
                )
        else:
            self.statusBar().showMessage(
                "Файл сохранён. Других файлов с корректной подписью не найдено.",
                8000,
            )

    def clear_current_view(self):
        self.doc = None
        self.pdf_path = None
        self.p7s_path = None
        self.current_page_index = 0
        self.cert_info = None
        self.page_view.scene().clear()
        self.thumb_list.clear()
        self.thumbs_widget.setVisible(False)
        self.btn_save.setEnabled(False)
        self.btn_sign_pdf.setEnabled(False)

        for lbl in [
            self.lbl_serial,
            self.lbl_owner,
            self.lbl_sign_time,
            self.lbl_valid_period,
            self.lbl_status,
        ]:
            lbl.setText("—")

        self.statusBar().showMessage("Перетащите файлы или нажмите «Добавить файлы»")


# -------- запуск приложения + настройка логирования --------

# Если нужно отключить лог-файл в проде — поставь здесь False.
ENABLE_FILE_LOG = True

if __name__ == "__main__":
    log_kwargs = {
        "level": logging.DEBUG,
        "format": "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    }
    if ENABLE_FILE_LOG:
        log_kwargs["filename"] = "ep_viewer.log"
        log_kwargs["filemode"] = "w"

    logging.basicConfig(**log_kwargs)

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
