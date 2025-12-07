# Визуализация электронной подписи на PDF

Приложение настольное (desktop) на Python + PySide6 для проверки электронной подписи под PDF-документом и визуального проставления штампа с краткой информацией о сертификате.

## 1. Назначение

Приложение позволяет:

1. Перетащить в окно:
   - **PDF** — исходный документ;
   - **P7S** — файл подписи;
   - опционально **CER/CRT** — файл сертификата (если он не встроен в подпись).
2. Проверить соответствие подписи PDF-файлу и статус сертификата.
3. Извлечь основные данные сертификата и время подписи.
4. Отобразить PDF с возможностью:
   - выбрать страницу;
   - поставить штамп с информацией о подписи;
   - перемещать и масштабировать штамп.
5. Сохранить новый PDF:
   - исходное содержимое не меняется;
   - на выбранной странице появляется штамп.

## 2. Структура проекта

Предполагаемая структура папки проекта:

```text
project_root/
├── main.py
├── gui.py
├── pdf_utils.py
├── signature_utils.py
├── timesnrcyrmt.ttf        # (опционально) шрифт Times New Roman (кириллица)
└── README.md               # этот файл
```

При необходимости можно добавить `requirements.txt`:

```text
PySide6
PyMuPDF
asn1crypto
cryptography
```

## 3. Требования

- Python **3.9+** (проверялось на 3.10).
- Библиотеки:
  - **PySide6** — GUI (Qt);
  - **PyMuPDF** (`fitz`) — работа с PDF:
    - рендер страницы в изображение,
    - вставка картинки штампа;
  - **asn1crypto**, **cryptography** — разбор CMS/PKCS#7 и сертификатов (ГОСТ-подписи и др.).
- ОС: Windows (но код практически кроссплатформенный — если есть Qt и PyMuPDF).

Установка зависимостей (внутри виртуального окружения):

```bash
pip install PySide6 PyMuPDF asn1crypto cryptography
```

Дополнительно можно положить рядом файл шрифта:

- `timesnrcyrmt.ttf` — кириллический Times New Roman (или совместимый).
  Если его нет, используется системный *Times New Roman*, а если и его нет — шрифт Qt по умолчанию.

### CryptoPro CSP и подпись через CLI

Для командной подписи без браузера установите **CryptoPro CSP** (версия с утилитой `cryptcp`).

Проверка наличия `cryptcp`:

```bash
cryptcp -help
```

Если утилита не найдена, убедитесь, что CryptoPro установлен и путь к `cryptcp` добавлен в переменную окружения `PATH` (или лежит в типовой папке `C:\Program Files\Crypto Pro\CSP\cryptcp.exe`).

Примеры команд (desktop/CLI режим):

- Отсоединённая подпись PDF:

  ```bash
  python signer_cli.py sign --file sample.pdf --detached --thumbprint "<отпечаток>"
  ```

- Присоединённая подпись:

  ```bash
  python signer_cli.py sign --file sample.pdf --attached --thumbprint "<отпечаток>"
  ```

- Проверка detached подписи:

  ```bash
  python signer_cli.py verify --file sample.pdf --sig sample.pdf.sig
  ```

Используйте флаг `--dry-run`, чтобы увидеть команду `cryptcp` без выполнения.

Типичные ошибки:

- `Утилита cryptcp не найдена` — установите CryptoPro CSP или поправьте `PATH`.
- `Не указан сертификат. Передайте отпечаток сертификата через --thumbprint` — укажите отпечаток (certmgr.msc → Сведения → Отпечаток).
- `Файл для подписи не найден` — проверьте путь к входному файлу или подписи.

## 3.1. Сборка в один файл через PyInstaller

Для корректного включения иконки и справки используйте ключи `--add-data` при сборке, например:

```bash
pyinstaller --onefile --noconsole \
  --add-data "help.html;." \
  --add-data "app.ico;." \
  main.py
```

Пути к ресурсам и файлам настроек в коде настроены так, чтобы работать как при запуске `python main.py`, так и в собранном `.exe`.

---

## 4. Общая логика работы

1. Пользователь запускает `main.py`.
2. Создаётся экземпляр `QApplication`, на котором поднимается окно `MainWindow` из `gui.py`.
3. Главное окно:
   - принимает файлы через **Drag & Drop**;
   - определяет среди них `*.pdf`, `*.p7s`, опционально `*.cer`;
   - открывает PDF через `pdf_utils.open_document`;
   - вызывает `signature_utils.get_certificate_info` для проверки подписи и получения сведений о сертификате;
   - отображает:
     - миниатюры страниц слева;
     - выбранную страницу по центру;
     - панель информации о сертификате справа;
     - поверх страницы — прямоугольник штампа (объект `StampRectItem` в `PDFPageView`).
4. Для штампа:
   - из `CertificateInfo` собирается словарь с полями:
     `serial_number`, `subject`, `valid_from`, `valid_to`, `signing_time`, `status`;
   - `pdf_utils.build_stamp_image(...)` создаёт **одну каноническую картинку штампа** (1200×600), где:
     - рисуется закруглённая рамка;
     - внутри — текст с отступами и автоматически подобранным размером шрифта.
   - Картинка передаётся в `PDFPageView` и масштабируется целиком при изменении размера прямоугольника штампа.
5. При сохранении:
   - пользователь выбирает путь;
   - `PDFPageView.get_stamp_rect_pdf_coords()` возвращает координаты прямоугольника штампа **в координатах PDF-страницы**;
   - `pdf_utils.add_stamp_to_pdf(...)` открывает исходный PDF, вставляет PNG-картинку штампа в указанный `fitz.Rect` и сохраняет новый файл.

Таким образом:
- **текст штампа никогда не перерисовывается в PDF** — всегда вставляется одинаковая картинка, просто с разным масштабом;
- расположение в интерфейсе всегда совпадает с расположением в итоговом документе.

---

## 5. Файл `main.py`

Входная точка приложения. Пример реализаций:

```python
import sys
import logging

from PySide6.QtWidgets import QApplication
from gui import MainWindow


def setup_logging():
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        filename="ep_viewer.log",
        filemode="w",
        encoding="utf-8",
    )
    # дублируем логи в консоль
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    logging.getLogger().addHandler(console)


if __name__ == "__main__":
    setup_logging()
    logging.getLogger(__name__).info("=== Приложение запущено ===")

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
```

Основные моменты:

- функция `setup_logging()` настраивает логирование в файл `ep_viewer.log` + в консоль;
- создаётся `QApplication` и главное окно `MainWindow` из `gui.py`.

---

## 6. Файл `gui.py` — интерфейс

Файл `gui.py` содержит всё, что касается GUI:

- классы:
  - `StampRectItem(QGraphicsRectItem)` — прямоугольник штампа;
  - `PDFPageView(QGraphicsView)` — центральный виджет, отображающий страницу и штамп;
  - `MainWindow(QMainWindow)` — главное окно.

### 6.1. Класс `StampRectItem`

```python
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
```

Назначение:

- Представляет **рамку штампа** на странице.
- Позволяет:
  - перетаскивать штамп;
  - изменять размер только за нижний правый угол (синий квадратик);
  - при изменении пропорционально масштабировать по ширине/высоте (фиксированное соотношение сторон).

Ключевые методы:

- `paint(...)` — переопределён так, чтобы:
  - **рамка не рисовалась** (видимый прямоугольник приходит из картинки);
  - рисовалась только синяя «ручка» (квадратик) для изменения размера.
- `mousePressEvent / mouseMoveEvent / mouseReleaseEvent` — логика определения, пользователь тянет ручку или двигает весь прямоугольник.
- `_notify_geometry_changed()` — вызывает коллбек `geometry_changed_callback` (обычно `PDFPageView._update_pixmap_item`), чтобы пересчитать размер картинки штампа под новый прямоугольник.

### 6.2. Класс `PDFPageView`

```python
class PDFPageView(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setScene(QGraphicsScene(self))
        self._pixmap_item = None               # изображение страницы
        self.stamp_item: Optional[StampRectItem] = None
        self.stamp_pixmap_item: Optional[QGraphicsPixmapItem] = None
        self.stamp_pixmap: Optional[QPixmap] = None
        self.zoom = 1.0                        # коэффициент рендеринга PDF -> изображение

        self._space_pressed = False            # для панорамирования
        self._panning = False
        self._last_pan_pos: Optional[QPoint] = None
```

Основные задачи:

1. Показывать изображение страницы (`_pixmap_item`).
2. Содержать объект `StampRectItem` (рамка) и `QGraphicsPixmapItem` с картинкой штампа.
3. Позволять:
   - масштабировать вид (Ctrl + колесико);
   - перемещать сцену (пробел + ЛКМ);
   - автоматически подгонять всю страницу под размер окна (`fitInView` при установке страницы).

Ключевые методы:

- `set_page(pixmap: QPixmap, pdf_zoom: float)`  
  Устанавливает изображение страницы и запоминает `pdf_zoom` — коэффициент, с которым страница была отрендерена из PDF. Этот коэффициент позже используется для пересчёта координат штампа в `get_stamp_rect_pdf_coords`.

- `_create_default_stamp()`  
  Создаёт рамку штампа:
  - ширина 60% от ширины страницы;
  - высота 20% от высоты;
  - позиция: по центру по горизонтали, немного выше нижнего края.

- `set_stamp_pixmap(pixmap: QPixmap)`  
  Принимает **каноническую картинку штампа** (1200×600) и запускает `_update_pixmap_item`.

- `_update_pixmap_item()`  
  Масштабирует картинку штампа под размеры рамки `StampRectItem` с сохранением пропорций (`Qt.KeepAspectRatio`) и центрирует её внутри рамки.

- `get_stamp_rect_pdf_coords() -> Optional[fitz.Rect]`  
  Самый важный метод: пересчитывает координаты рамки:

  1. Получает прямоугольник штампа в координатах сцены (`mapRectToScene`).
  2. Делит все координаты на `self.zoom` (коэффициент рендеринга страницы из PDF).
  3. Возвращает `fitz.Rect` в **координатах PDF-страницы**, которые используются в `add_stamp_to_pdf`.

  Масштаб вида (`scale()` при Ctrl+колёсико) не влияет на эти координаты, потому что он не меняет геометрию объектов в сцене — только/как они отображаются.

### 6.3. Класс `MainWindow`

`MainWindow` собирает всё вместе.

#### Основные поля

```python
self.pdf_path: Optional[str] = None
self.p7s_path: Optional[str] = None
self.cer_path: Optional[str] = None

self.doc: Optional[fitz.Document] = None
self.current_page_index: int = 0
self.cert_info: Optional[CertificateInfo] = None
```

#### UI-компоненты

- Слева:
  - `QListWidget` с миниатюрами страниц (`thumb_list`).
- Центр:
  - `PDFPageView` (`self.page_view`).
- Справа:
  - `QFormLayout` с полями:
    - серийный номер,
    - владелец,
    - издатель,
    - начало / окончание действия,
    - время подписи,
    - статус подписи.
  - кнопка **«Сохранить файл…»**.

#### Drag & Drop

- `dragEnterEvent()` — разрешает перенос, если есть URL-ы.
- `dropEvent()`:
  - отбирает файлы по расширению (`.pdf`, `.p7s`, `.cer`/`.crt`);
  - проверяет наличие PDF и P7S;
  - вызывает `load_files(pdf_path, p7s_path, cer_path)`.

#### `load_files(...)`

1. Открывает PDF через `open_document`.
2. Строит миниатюры страниц (`populate_thumbnails`).
3. Показывает первую страницу (`show_page(0)`).
4. Вызывает `get_certificate_info(pdf, p7s, cer)`:
   - результат сохраняется в `self.cert_info`;
   - обновляет панель справа `update_cert_info_panel()`.
5. Вызывает `update_stamp_preview()` — строит картинку штампа для текущей подписи и передаёт её в `PDFPageView`.
6. Разрешает кнопку сохранения.

#### `show_page(index)`

1. Получает страницу `page = self.doc[index]`.
2. Рендерит её в растровое изображение:

   ```python
   matrix = fitz.Matrix(PDF_RENDER_ZOOM, PDF_RENDER_ZOOM)
   pix = page.get_pixmap(matrix=matrix)
   ```

   где `PDF_RENDER_ZOOM = 2.0` — постоянный коэффициент для хорошего качества.

3. Конвертирует `pix` в `QPixmap` и передаёт в `self.page_view.set_page(qpix, PDF_RENDER_ZOOM)`.
4. Обновляет картинку штампа (`update_stamp_preview()`).

#### `update_cert_info_panel()`

Заполняет текстовые `QLabel` данными `CertificateInfo`. По статусу меняет цвет:

- «действительна» → зелёный;
- истёк / не соответствует / ошибка → красный.

#### `update_stamp_preview()`

1. Формирует словарь из `self.cert_info` с ключами:
   - `serial_number`, `subject`, `issuer`, `valid_from`, `valid_to`, `signing_time`, `status`.
2. Вызывает `pdf_utils.build_stamp_image(stamp_info)` → `QImage`.
3. Преобразует в `QPixmap` и передаёт в `self.page_view.set_stamp_pixmap(pix)`.

#### Сохранение `on_save_clicked()`

1. Получает `stamp_rect_pdf = self.page_view.get_stamp_rect_pdf_coords()`.
2. Даёт пользователю выбрать путь (с дефолтным именем `*_stamped.pdf`).
3. Проверяет существование файла и спрашивает о перезаписи.
4. Вызывает:

   ```python
   add_stamp_to_pdf(
       self.pdf_path,
       out_path,
       page_index,
       stamp_rect_pdf,
       stamp_info,
   )
   ```

где `stamp_info` — тот же словарь, что используется для превью. Новый файл сохраняется, исходный по умолчанию не перезаписывается.

---

## 7. Отладка плагина CryptoPro в браузере

Для диагностики проблем с браузерной подписью добавлен автономный тест:

- откройте файл `docs/browser_signing_debug.html` напрямую в браузере (двойной клик или через контекстное меню);
- выберите PDF для тестовой подписи или нажмите «Использовать встроенный пример»;
- нажмите «Подписать выбранный файл» и следите за журналом на странице — все сообщения остаются локально и помогут понять, загружается ли плагин и какие ошибки возникают.

Страница не отправляет данные на сервер и пригодна для проверки доступности плагина и выбора сертификата без запуска основного приложения.

## 7. Файл `pdf_utils.py` — работа с PDF и штампом

(Полный код описан в комментариях этого README — разделы про функции
`open_document`, `_get_stamp_font_family`, `_build_stamp_text`,
`_render_base_stamp_image`, `build_stamp_image`, `add_stamp_to_pdf`.)

---

## 8. Файл `signature_utils.py` — проверка подписи и сертификата

Реализация может отличаться, но логика примерно такая:

1. Прочитать `*.p7s`.
2. Попробовать распарсить как CMS/PKCS#7 (`asn1crypto.cms.ContentInfo` или через `cryptography`).
3. Извлечь:
   - подписанные данные (SignedData);
   - сертификат(ы);
   - атрибут `signingTime`.
4. Сравнить хэш PDF с тем, что подписан.
5. Определить статус:
   - «действительна» — если подпись корректна и срок сертификата не истёк;
   - «срок действия сертификата истёк» — если подпись корректна, но сейчас дата > NotAfter;
   - «не соответствует документу» — если хэш PDF не соответствует подписанному;
   - «ошибка проверки: ...» — если что-то пошло не так.

### Структура данных `CertificateInfo`

```python
class CertificateInfo:
    def __init__(self):
        self.serial_number: str = "не удалось определить"
        self.subject: str = "не удалось определить"
        self.issuer: str = "не удалось определить"
        self.valid_from: str = "не удалось определить"
        self.valid_to: str = "не удалось определить"
        self.signing_time: str = "не удалось определить"
        self.status: str = "не удалось определить"
```

Все поля — строки, уже подготовленные для отображения (ДД.ММ.ГГГГ ЧЧ:ММ и т.п.).

### `get_certificate_info(pdf_path, p7s_path, cer_path=None) -> CertificateInfo`

Основные шаги:

1. Прочитать P7S.
2. Попробовать извлечь из него сертификат; если не получилось и указан `cer_path` — прочесть CER.
3. Заполнить:
   - `serial_number` — серийный номер (обычно в hex);
   - `subject` — Common Name / ФИО + при необходимости организация/должность;
   - `issuer` — УЦ;
   - `valid_from / valid_to` — период действия (формат ДД.ММ.ГГГГ);
   - `signing_time` — время подписи (ДД.ММ.ГГГГ ЧЧ:ММ).
4. Проверить подпись по содержимому PDF.
5. Выставить `status` согласно результатам.
6. Логирование ошибок ведётся в `ep_viewer.log`, а пользователю отдаются «читаемые» статусы на русском.

---

## 9. Настройки приложения (`settings.json`)

- Файл `settings.json` создаётся автоматически в папке с программой (рядом с `main.py`).
- Все пользовательские настройки хранятся в одном JSON:

  ```json
  {
    "image_path": "C:/logos/company.png",
    "header_text": "ООО \"Компания\" Подпись документа",
    "show_sign_time": true,
    "default_output_dir": "D:/Signed",
    "stamp_rect": {
      "left": 0.2,
      "top": 0.6,
      "width": 0.6,
      "height": 0.2
    }
  }
  ```

- Поля:
  - `image_path` — PNG-логотип шапки штампа.
  - `header_text` — текст шапки (одной строкой; в штампе автоматически разбивается на 1–3 строки и обрезается по ширине).
  - `show_sign_time` — показывать ли время подписи на штампе.
  - `default_output_dir` — последняя выбранная папка сохранения (используется, когда снят флажок «Сохранять файл в исходное месторасположение»).
  - `stamp_rect` — нормированные координаты прямоугольника штампа (от 0 до 1 относительно страницы) для универсального позиционирования на всех файлах.

- Чтобы отключить запись логов в `ep_viewer.log`, откройте `main.py` и установите `ENABLE_FILE_LOG = False` перед запуском.


## 10. Горячие клавиши и управление

- **Drag & Drop**:
  - перетаскивание PDF + P7S (и опционально CER) в главное окно.
- **Колесико мыши + Ctrl** — зум страницы (масштабирование вида).
- **Пробел + ЛКМ** — панорамирование страницы (перемещение вида).
- **Штамп**:
  - LMB по рамке — перетаскивание;
  - LMB по синему квадратику в правом нижнем углу — пропорциональное изменение размера.

---

## 11. Как воспроизвести приложение с нуля

1. Создать виртуальное окружение и установить зависимости:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   pip install PySide6 PyMuPDF asn1crypto cryptography
   ```

2. Создать файлы:
   - `main.py` — как описано в разделе 5;
   - `gui.py` — как в разделе 6;
   - `pdf_utils.py` — как в разделе 7;
   - `signature_utils.py` — реализация проверки подписи (см. раздел 8, можно использовать свою/готовую);
   - (опционально) положить `timesnrcyrmt.ttf` рядом с `pdf_utils.py`.

3. Запустить:

   ```bash
   python main.py
   ```

4. В открывшееся окно перетащить:
   - PDF-файл;
   - соответствующий P7S.

5. При необходимости выбрать страницу, отрегулировать положение/размер штампа и нажать **«Сохранить файл…»**.

После прочтения этого README и кода можно полностью пересобрать приложение с нуля или доработать его под свои задачи.
