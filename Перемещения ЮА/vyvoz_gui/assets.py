"""Шрифты и значки: Inter и Material Symbols лежат в assets/fonts и встраиваются в exe (vyvoz_vne_matricy.spec)."""
import os
import sys

from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase, QIcon, QPainter, QPixmap

from . import tokens as T

ICONS = {               # Material Symbols Outlined, код символа (подмножество в MaterialSymbolsOutlined-subset.ttf)
    "folder": "", "folder_open": "", "refresh": "", "check_circle": "", "warning": "",
    "error": "", "info": "", "help": "", "visibility": "", "visibility_off": "",
    "select_all": "", "description": "", "local_shipping": "", "call_split": "",
    "cleaning_services": "", "inventory": "", "settings": "", "receipt_long": "",
    "play_arrow": "", "open_in_new": "", "light_mode": "", "dark_mode": "", "contrast": "",
    "close": "", "search": "", "content_copy": "", "delete": "", "task_alt": "",
    "schedule": "", "store": "", "table_view": "", "build": "", "remove_circle": "",
    "check": "", "expand_more": "", "arrow_forward": "", "playlist_remove": "",
    "fact_check": "",
}

FONT_FILES = ("Inter-Regular.ttf", "Inter-Medium.ttf", "Inter-SemiBold.ttf", "MaterialSymbolsOutlined-subset.ttf")


def base_dir():
    """Папка с assets: в exe - распакованные данные PyInstaller (_MEIPASS = vyvoz_lib), иначе папка «Перемещения ЮА»."""
    return getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def fonts_dir():
    return os.path.join(base_dir(), "assets", "fonts")


def load_fonts():
    """Регистрирует шрифты. -> список проблем (пусто - всё загружено). Не загрузился Inter - окно работает на Segoe UI."""
    probs = []
    for fn in FONT_FILES:
        p = os.path.join(fonts_dir(), fn)
        if not os.path.isfile(p):
            probs.append(u"нет файла шрифта %s" % p)
        elif QFontDatabase.addApplicationFont(p) < 0:
            probs.append(u"шрифт не загрузился: %s" % p)
    return probs


def ui_font(size_px=T.TYPO_BODY_MD[0], weight=400):
    f = QFont(T.FONT if T.FONT in QFontDatabase.families() else "Segoe UI")
    f.setPixelSize(size_px)
    f.setWeight(QFont.Weight(weight))
    return f


def has_icons():
    return T.FONT_ICONS in QFontDatabase.families()


def glyph_pixmap(name, color, size=20, dpr=1.0):
    """Значок Material Symbols в QPixmap нужного цвета (для кнопок, таблиц, чипов)."""
    px = QPixmap(QSize(int(size * dpr), int(size * dpr)))
    px.setDevicePixelRatio(dpr)
    px.fill(Qt.transparent)
    if not has_icons() or name not in ICONS:
        return px
    p = QPainter(px)
    p.setRenderHint(QPainter.Antialiasing)
    f = QFont(T.FONT_ICONS)
    f.setPixelSize(size)
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(QRectF(QPointF(0, 0), QSize(size, size)), Qt.AlignCenter, ICONS[name])
    p.end()
    return px


def icon(name, color, size=20):
    ic = QIcon()
    for dpr in (1.0, 1.25, 1.5, 2.0):            # чёткие значки на масштабе Windows 125-200 %
        ic.addPixmap(glyph_pixmap(name, color, size, dpr))
    return ic


def check_png(color):
    """Галочка для QCheckBox::indicator:checked: QSS берёт картинку только из файла - кладём во временную папку."""
    import tempfile
    if not has_icons():                      # без шрифта значков картинка была бы пустой и закэшировалась бы
        return u""
    p = os.path.join(tempfile.gettempdir(), "fm_vyvoz_check_%s.png" % color.lstrip("#"))
    if not os.path.isfile(p):
        glyph_pixmap("check", color, 16, 2.0).save(p, "PNG")
    return p.replace("\\", "/")


def app_icon():
    p = os.path.join(base_dir(), "vyvoz_fm.ico")
    return QIcon(p) if os.path.isfile(p) else QIcon()
