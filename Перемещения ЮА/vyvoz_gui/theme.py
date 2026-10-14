"""Светлая и тёмная тема (QSS из токенов) и выбор «Светлая / Тёмная / Как в Windows» с запоминанием (QSettings)."""
from PySide6.QtCore import QObject, QSettings, Qt, Signal
from PySide6.QtGui import QGuiApplication

from . import tokens as T

FONT_DEFAULT = "inter"

MODES = (("light", u"Светлая"), ("dark", u"Тёмная"), ("system", u"Как в Windows"))
DEFAULT_MODE = "light"          # решение владельца 14.10.2026: в магазинах привыкли к светлой, списки печатаются на светлом


def build_qss(c, check_png=u"", family=T.FONT, arrow_dn=u"", arrow_up=u""):
    """Полная таблица стилей для палитры c (tokens.LIGHT или tokens.DARK). check_png - картинка галочки (assets.check_png).
    family - шрифт из «Сервиса»; хинтинг QSS не задаёт: он берётся из шрифта приложения (assets.ui_font)."""
    b, s = T.TYPO_BODY_MD[0], T.TYPO_BODY_SM[0]
    return f"""
* {{ font-family: "{family}", "{T.FONT_FALLBACK}"; font-size: {b}px; }}
QWidget {{ color: {c['text']}; }}
QMainWindow, #content, #page, QScrollArea, #scrollBody {{ background: {c['bg']}; }}
QToolTip {{ background: {c['tooltip_bg']}; color: {c['tooltip_text']}; border: none; padding: {T.SPACING_SM}px; }}

/* ---------- боковая панель (тёмная в обеих темах) ---------- */
#sidebar {{ background: {c['side_bg']}; }}
#sidebar QLabel {{ background: transparent; }}
#brandName {{ color: #ffffff; font-size: 18px; font-weight: 600; }}
#brandSub {{ color: {c['side_text']}; font-size: {s}px; }}
#brandBar {{ background: {c['side_indicator']}; }}
#sidebarGroupLabel {{ color: {c['side_group']}; font-size: {T.TYPO_LABEL_SM[0]}px; font-weight: 600;
    padding: {T.SPACING_MD}px {T.SPACING_MD}px {T.SPACING_XS}px {T.SPACING_MD + 3}px; }}
#sidebarDivider {{ background: {c['side_divider']}; }}
#sidebarFooter {{ color: {c['side_group']}; font-size: {s}px; }}
#sidebarNavItem {{ background: transparent; border: none; border-left: 3px solid transparent; border-radius: 0;
    padding: {T.SPACING_SM}px {T.SPACING_MD}px; text-align: left; color: {c['side_text']}; font-weight: 500; min-height: 24px; }}
#sidebarNavItem:hover {{ background: {c['side_hover']}; color: {c['side_text_active']}; }}
#sidebarNavItem:focus {{ outline: none; border-left: 3px solid {c['side_text']}; background: {c['side_hover']}; }}
#sidebarNavItem[active="true"] {{ background: {c['side_active']}; color: {c['side_text_active']};
    border-left: 3px solid {c['side_indicator']}; font-weight: 600; }}

/* ---------- заголовки и тексты ---------- */
#pageTitle {{ font-size: {T.TYPO_HEADLINE_LG[0]}px; font-weight: 600; }}
#cardTitle {{ font-size: {T.TYPO_HEADLINE_SM[0]}px; font-weight: 600; }}
QLabel[muted="true"] {{ color: {c['text_muted']}; font-size: {s}px; }}
#card {{ background: {c['surface']}; border: 1px solid {c['outline_variant']}; border-radius: {T.RADIUS_LG}px; }}
#card QLabel, #card QCheckBox, #card QRadioButton {{ background: transparent; }}
#actionBar {{ background: {c['surface']}; border-top: 1px solid {c['outline_variant']}; }}
#actionBar QLabel {{ background: transparent; }}
#statusBar {{ background: {c['bg']}; border-top: 1px solid {c['outline_variant']}; }}
#statusBar QLabel {{ color: {c['text_muted']}; font-size: {s}px; background: transparent; }}

/* ---------- плашки результата ---------- */
#banner[level="OK"] {{ background: {c['success_bg']}; border-radius: {T.RADIUS_DEFAULT}px; }}
#banner[level="WARN"] {{ background: {c['warning_bg']}; border-radius: {T.RADIUS_DEFAULT}px; }}
#banner[level="ERR"] {{ background: {c['error_bg']}; border-radius: {T.RADIUS_DEFAULT}px; }}
#banner[level="INFO"] {{ background: {c['info_bg']}; border-radius: {T.RADIUS_DEFAULT}px; }}
#banner QLabel {{ background: transparent; }}
#banner[level="ERR"] QLabel {{ color: {c['on_error_bg']}; }}

/* ---------- кнопки ---------- */
QPushButton {{ background: {c['surface']}; color: {c['text']}; border: 1px solid {c['outline_variant']};
    border-radius: {T.RADIUS_DEFAULT}px; padding: {T.SPACING_SM - 2}px {T.SPACING_MD - 4}px; min-height: 22px; }}
QPushButton:hover {{ background: {c['surface_low']}; border-color: {c['outline']}; }}
QPushButton:pressed {{ background: {c['surface_mid']}; }}
QPushButton:focus {{ outline: none; border: 2px solid {c['focus']}; padding: {T.SPACING_SM - 3}px {T.SPACING_MD - 5}px; }}
QPushButton:disabled {{ color: {c['text_muted']}; background: {c['surface_low']}; border-color: {c['outline_variant']}; }}
QPushButton[kind="primary"] {{ background: {c['primary']}; color: {c['on_primary']}; border: none; font-weight: 600;
    padding: {T.SPACING_SM}px {T.SPACING_MD + 4}px; }}
QPushButton[kind="primary"]:hover {{ background: {c['primary_hover']}; }}
QPushButton[kind="primary"]:pressed {{ background: {c['primary_pressed']}; }}
QPushButton[kind="primary"]:focus {{ border: 2px solid {c['focus']}; padding: {T.SPACING_SM - 2}px {T.SPACING_MD + 2}px; }}
QPushButton[kind="primary"]:disabled {{ background: {c['primary_disabled_bg']}; color: {c['primary_disabled_text']}; }}
QPushButton[kind="ghost"] {{ background: transparent; border: none; padding: {T.SPACING_XS}px; }}
QPushButton[kind="ghost"]:hover {{ background: {c['surface_mid']}; }}
QPushButton[kind="ghost"]:focus {{ border: 2px solid {c['focus']}; padding: {T.SPACING_XS - 2}px; }}

/* ---------- поля и выбор ---------- */
QLineEdit, QComboBox, QSpinBox {{ background: {c['surface']}; border: 1px solid {c['outline']}; border-radius: {T.RADIUS_SM}px;
    padding: {T.SPACING_XS + 2}px {T.SPACING_SM}px; min-height: 22px; selection-background-color: {c['selection_bg']};
    selection-color: {c['selection_text']}; }}
QLineEdit:focus, QComboBox:focus, QSpinBox:focus {{ border: 2px solid {c['focus']}; padding: {T.SPACING_XS + 1}px {T.SPACING_SM - 1}px; }}
QComboBox::drop-down {{ border: none; width: 28px; subcontrol-origin: padding; subcontrol-position: center right; }}
QComboBox::down-arrow {{ image: url("{arrow_dn}"); width: 12px; height: 12px; }}
QComboBox::down-arrow:disabled, QSpinBox::up-arrow:disabled, QSpinBox::down-arrow:disabled {{ image: none; }}
QSpinBox {{ padding-right: 24px; }}
QSpinBox::up-button, QSpinBox::down-button {{ border: none; background: transparent; width: 22px; subcontrol-origin: border; }}
QSpinBox::up-button {{ subcontrol-position: top right; }}
QSpinBox::down-button {{ subcontrol-position: bottom right; }}
QSpinBox::up-button:hover, QSpinBox::down-button:hover {{ background: {c['surface_mid']}; }}
QSpinBox::up-arrow {{ image: url("{arrow_up}"); width: 10px; height: 10px; }}
QSpinBox::down-arrow {{ image: url("{arrow_dn}"); width: 10px; height: 10px; }}
QComboBox:disabled, QSpinBox:disabled, QLineEdit:disabled {{ color: {c['text_muted']}; background: {c['surface_low']};
    border-color: {c['outline_variant']}; }}
QComboBox QAbstractItemView {{ background: {c['surface']}; border: 1px solid {c['outline_variant']};
    selection-background-color: {c['selection_bg']}; selection-color: {c['selection_text']}; outline: none; }}
QCheckBox, QRadioButton {{ spacing: {T.SPACING_SM}px; }}
QCheckBox::indicator, QRadioButton::indicator {{ width: 16px; height: 16px; border: 2px solid {c['outline']};
    background: {c['surface']}; }}
QCheckBox::indicator {{ border-radius: {T.RADIUS_SM}px; }}
QRadioButton::indicator {{ border-radius: 10px; }}
QCheckBox::indicator:checked, QRadioButton::indicator:checked {{ background: {c['checked']}; border-color: {c['checked']}; }}
QCheckBox::indicator:checked {{ image: url("{check_png}"); }}
QRadioButton::indicator:checked {{ background: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5,
    stop:0 {c['checked']}, stop:0.5 {c['checked']}, stop:0.56 {c['surface']}, stop:1 {c['surface']}); }}
QCheckBox::indicator:hover, QRadioButton::indicator:hover {{ border-color: {c['focus']}; }}
QCheckBox:focus, QRadioButton:focus {{ outline: none; text-decoration: underline; }}
QCheckBox:disabled, QRadioButton:disabled {{ color: {c['text_muted']}; }}

/* ---------- таблицы ---------- */
QTableView {{ background: {c['surface']}; alternate-background-color: {c['surface_low']}; border: 1px solid {c['outline_variant']};
    border-radius: {T.RADIUS_DEFAULT}px; gridline-color: {c['outline_variant']}; selection-background-color: {c['selection_bg']};
    selection-color: {c['selection_text']}; outline: none; }}
QTableView::item {{ padding: 0 {T.SPACING_SM}px; border: none; }}
QTableView::item:selected {{ background: {c['selection_bg']}; color: {c['selection_text']}; }}
QTableView:focus {{ border: 2px solid {c['focus']}; }}
QHeaderView::section {{ background: {c['surface_low']}; color: {c['text_muted']}; border: none;
    border-bottom: 1px solid {c['outline_variant']}; padding: {T.SPACING_XS + 2}px {T.SPACING_SM}px;
    font-size: {T.TYPO_LABEL_MD[0]}px; font-weight: 600; }}
QPlainTextEdit, QTextEdit {{ background: {c['surface']}; border: 1px solid {c['outline_variant']}; border-radius: {T.RADIUS_DEFAULT}px;
    padding: {T.SPACING_SM}px; selection-background-color: {c['selection_bg']}; selection-color: {c['selection_text']}; }}

QScrollBar:vertical {{ background: transparent; width: 12px; margin: 0; }}
QScrollBar::handle:vertical {{ background: {c['outline_variant']}; border-radius: 4px; min-height: 32px; margin: 2px; }}
QScrollBar::handle:vertical:hover {{ background: {c['outline']}; }}
QScrollBar:horizontal {{ background: transparent; height: 12px; margin: 0; }}
QScrollBar::handle:horizontal {{ background: {c['outline_variant']}; border-radius: 4px; min-width: 32px; margin: 2px; }}
QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
QProgressBar {{ background: {c['surface_mid']}; border: none; border-radius: 2px; max-height: 4px; }}
QProgressBar::chunk {{ background: {c['checked']}; border-radius: 2px; }}
QDialog {{ background: {c['bg']}; }}
"""


class ThemeManager(QObject):
    """Режим light / dark / system; system следует теме Windows и меняется вместе с ней."""
    changed = Signal(dict)

    def __init__(self, app, settings=None):
        super().__init__()
        self.app = app
        self.settings = settings or QSettings("FamilyMarket", "VyvozVneMatricy")
        m = str(self.settings.value("theme", DEFAULT_MODE))
        self.mode = m if m in dict(MODES) else DEFAULT_MODE
        f = str(self.settings.value("font", FONT_DEFAULT))
        self.font_mode = f if f in dict(T.FONT_MODES) else FONT_DEFAULT
        try:
            QGuiApplication.styleHints().colorSchemeChanged.connect(lambda *_: self.mode == "system" and self.apply())
        except Exception:
            pass

    def is_dark(self):
        if self.mode == "system":
            try:
                return QGuiApplication.styleHints().colorScheme() == Qt.ColorScheme.Dark
            except Exception:
                return False
        return self.mode == "dark"

    def palette(self):
        return T.DARK if self.is_dark() else T.LIGHT

    def set_mode(self, mode):
        self.mode = mode if mode in dict(MODES) else DEFAULT_MODE
        self.settings.setValue("theme", self.mode)
        self.apply()

    def family(self):
        """Шрифт, который выбран в «Сервисе» (Inter / Segoe UI)."""
        return dict(T.FONT_MODES)[self.font_mode]

    def set_font(self, mode):
        self.font_mode = mode if mode in dict(T.FONT_MODES) else FONT_DEFAULT
        self.settings.setValue("font", self.font_mode)
        self.apply()

    def font_report(self):
        from .assets import font_report
        return font_report(self.app.font(), self.family())

    def apply(self):
        c = self.palette()
        from .assets import arrow_png, check_png, ui_font
        f = ui_font(family=self.family())
        self.app.setFont(f)
        self.app.setStyleSheet(build_qss(c, check_png(c["on_checked"]), f.family(), arrow_png(c["text_muted"]),
                                         arrow_png(c["text_muted"], up=True)))
        self.changed.emit(c)
