# QSS Patterns — PyQt6 Modern Enterprise Design System

Complete QSS templates and patterns. Always use token constants — no raw hex values
in QSS strings. Interpolate Python constants into QSS f-strings.

---

## Theme Architecture

Always define TWO complete QSS strings and apply via `app.setStyleSheet()` or
a window-level `setStyleSheet()`. Provide a `ThemeManager` helper:

```python
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import QObject

class ThemeManager(QObject):
    def __init__(self, app: QApplication):
        super().__init__()
        self.app = app
        self._dark = False

    @property
    def is_dark(self) -> bool:
        return self._dark

    def apply_light(self):
        self._dark = False
        self.app.setStyleSheet(LIGHT_QSS)

    def apply_dark(self):
        self._dark = True
        self.app.setStyleSheet(DARK_QSS)

    def toggle(self):
        self.apply_dark() if not self._dark else self.apply_light()
```

---

## Base QSS Template (Light)

```python
LIGHT_QSS = f"""
/* ─── Global Reset ─────────────────────────────── */
* {{
    font-family: '{FONT_BODY}';
    font-size: {TYPO_BODY_MD[0]}px;
    color: {COLOR_ON_SURFACE};
    outline: none;
    border: none;
    padding: 0;
    margin: 0;
    background: transparent;
    box-sizing: border-box;
}}

QMainWindow, QDialog {{
    background-color: {COLOR_BACKGROUND};
}}

QWidget {{
    background-color: transparent;
}}

/* ─── Scrollbars ────────────────────────────────── */
QScrollBar:vertical {{
    width: 6px;
    background: transparent;
    margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {COLOR_OUTLINE_VARIANT};
    border-radius: 3px;
    min-height: 24px;
}}
QScrollBar::handle:vertical:hover {{
    background: {COLOR_OUTLINE};
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {{
    background: transparent;
}}

/* ─── Buttons ───────────────────────────────────── */
QPushButton {{
    font-family: '{FONT_BODY}';
    font-size: {TYPO_LABEL_MD[0]}px;
    font-weight: {TYPO_LABEL_MD[1]};
    letter-spacing: 0.05em;
    padding: {SPACING_SM}px {SPACING_MD}px;
    border-radius: {RADIUS_DEFAULT}px;
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    background-color: {COLOR_SURFACE_LOWEST};
    color: {COLOR_ON_SURFACE};
}}
QPushButton:hover {{
    background-color: {COLOR_SURFACE_LOW};
    border-color: {COLOR_OUTLINE};
}}
QPushButton:pressed {{
    background-color: {COLOR_SURFACE};
}}
QPushButton:disabled {{
    color: {COLOR_OUTLINE_VARIANT};
    border-color: {COLOR_OUTLINE_VARIANT};
    background-color: {COLOR_SURFACE_LOW};
}}
QPushButton[class="primary"] {{
    background-color: {COLOR_PRIMARY};
    color: {COLOR_ON_PRIMARY};
    border: none;
}}
QPushButton[class="primary"]:hover {{
    background-color: {COLOR_PRIMARY_CONTAINER};
}}
QPushButton[class="primary"]:pressed {{
    background-color: {COLOR_PRIMARY};
    opacity: 0.85;
}}
QPushButton[class="primary"]:disabled {{
    background-color: {COLOR_OUTLINE_VARIANT};
    color: {COLOR_SURFACE_LOWEST};
}}
QPushButton[class="ghost"] {{
    background: transparent;
    border: none;
    color: {COLOR_ON_SURFACE_VARIANT};
    padding: {SPACING_SM}px;
}}
QPushButton[class="ghost"]:hover {{
    background-color: {COLOR_SURFACE_LOW};
    color: {COLOR_ON_SURFACE};
}}
QPushButton[class="danger"] {{
    background-color: {COLOR_ERROR};
    color: {COLOR_ON_ERROR};
    border: none;
}}
QPushButton[class="danger"]:hover {{
    background-color: {COLOR_ON_ERROR_CONTAINER};
}}

/* ─── Inputs ────────────────────────────────────── */
QLineEdit, QTextEdit, QPlainTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
    font-family: '{FONT_BODY}';
    font-size: {TYPO_BODY_MD[0]}px;
    background-color: {COLOR_SURFACE_LOWEST};
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: {RADIUS_SM}px;
    padding: {SPACING_SM}px {SPACING_MD}px;
    color: {COLOR_ON_SURFACE};
    selection-background-color: {COLOR_PRIMARY_FIXED};
    selection-color: {COLOR_PRIMARY};
}}
QLineEdit:focus, QTextEdit:focus, QPlainTextEdit:focus,
QSpinBox:focus, QDoubleSpinBox:focus {{
    border: 2px solid {COLOR_PRIMARY};
    padding: {SPACING_SM - 1}px {SPACING_MD - 1}px;
}}
QLineEdit:disabled, QTextEdit:disabled {{
    background-color: {COLOR_SURFACE_LOW};
    color: {COLOR_OUTLINE};
}}
QLineEdit[hasError="true"] {{
    border: 2px solid {COLOR_ERROR};
}}

/* ─── QComboBox ─────────────────────────────────── */
QComboBox::drop-down {{
    border: none;
    width: {SPACING_LG}px;
}}
QComboBox::down-arrow {{
    image: none;  /* use custom arrow or Material Symbols label overlay */
    width: 16px; height: 16px;
}}
QComboBox QAbstractItemView {{
    background-color: {COLOR_SURFACE_LOWEST};
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: {RADIUS_LG}px;
    padding: {SPACING_XS}px;
    selection-background-color: {COLOR_PRIMARY_FIXED};
    selection-color: {COLOR_PRIMARY};
    outline: none;
}}
QComboBox QAbstractItemView::item {{
    padding: {SPACING_SM}px {SPACING_MD}px;
    border-radius: {RADIUS_DEFAULT}px;
    min-height: 32px;
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {COLOR_SURFACE_LOW};
}}

/* ─── QLabel ────────────────────────────────────── */
QLabel[class="headline-lg"] {{
    font-family: '{FONT_HEADING}';
    font-size: {TYPO_HEADLINE_LG[0]}px;
    font-weight: {TYPO_HEADLINE_LG[1]};
    color: {COLOR_ON_SURFACE};
    letter-spacing: -0.01em;
}}
QLabel[class="headline-md"] {{
    font-family: '{FONT_HEADING}';
    font-size: {TYPO_HEADLINE_MD[0]}px;
    font-weight: {TYPO_HEADLINE_MD[1]};
    color: {COLOR_ON_SURFACE};
    letter-spacing: -0.01em;
}}
QLabel[class="headline-sm"] {{
    font-family: '{FONT_HEADING}';
    font-size: {TYPO_HEADLINE_SM[0]}px;
    font-weight: {TYPO_HEADLINE_SM[1]};
    color: {COLOR_ON_SURFACE};
}}
QLabel[class="label-md"] {{
    font-family: '{FONT_BODY}';
    font-size: {TYPO_LABEL_MD[0]}px;
    font-weight: {TYPO_LABEL_MD[1]};
    letter-spacing: 0.05em;
    color: {COLOR_ON_SURFACE_VARIANT};
}}
QLabel[class="muted"] {{
    color: {COLOR_ON_SURFACE_VARIANT};
}}
QLabel[class="icon"] {{
    font-family: 'Material Symbols Outlined';
    font-size: 20px;
    color: {COLOR_ON_SURFACE_VARIANT};
}}

/* ─── QTableWidget / QTableView ─────────────────── */
QTableWidget, QTableView {{
    background-color: {COLOR_SURFACE_LOWEST};
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: {RADIUS_LG}px;
    gridline-color: {COLOR_OUTLINE_VARIANT};
    alternate-background-color: {COLOR_SURFACE_LOW};
    selection-background-color: {COLOR_PRIMARY_FIXED};
    selection-color: {COLOR_ON_SURFACE};
    outline: none;
    font-size: {TYPO_BODY_SM[0]}px;
}}
QTableWidget::item, QTableView::item {{
    padding: 0 {SPACING_MD}px;
    min-height: 40px;
    border-bottom: 1px solid {COLOR_OUTLINE_VARIANT};
}}
QTableWidget::item:selected, QTableView::item:selected {{
    background-color: {COLOR_PRIMARY_FIXED};
    color: {COLOR_ON_SURFACE};
}}
QHeaderView {{
    background-color: {COLOR_SURFACE_LOW};
    border-bottom: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: 0;
}}
QHeaderView::section {{
    background-color: {COLOR_SURFACE_LOW};
    color: {COLOR_ON_SURFACE_VARIANT};
    font-family: '{FONT_BODY}';
    font-size: {TYPO_LABEL_MD[0]}px;
    font-weight: {TYPO_LABEL_MD[1]};
    letter-spacing: 0.05em;
    text-transform: uppercase;
    padding: 0 {SPACING_MD}px;
    height: 36px;
    border: none;
    border-right: 1px solid {COLOR_OUTLINE_VARIANT};
    border-bottom: 1px solid {COLOR_OUTLINE_VARIANT};
}}
QHeaderView::section:last {{
    border-right: none;
}}

/* ─── QListWidget ───────────────────────────────── */
QListWidget {{
    background-color: transparent;
    border: none;
    outline: none;
    padding: {SPACING_XS}px;
}}
QListWidget::item {{
    border-radius: {RADIUS_DEFAULT}px;
    padding: {SPACING_SM}px {SPACING_MD}px;
    color: {COLOR_ON_SURFACE_VARIANT};
    min-height: 36px;
}}
QListWidget::item:hover {{
    background-color: {COLOR_SURFACE_HIGH};
    color: {COLOR_ON_SURFACE};
}}
QListWidget::item:selected {{
    background-color: {COLOR_PRIMARY_FIXED};
    color: {COLOR_PRIMARY};
    font-weight: 600;
}}

/* ─── QFrame (Cards) ────────────────────────────── */
QFrame[class="card"] {{
    background-color: {COLOR_SURFACE_LOWEST};
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: {RADIUS_LG}px;
    padding: {SPACING_MD}px;
}}
QFrame[class="divider"] {{
    background-color: {COLOR_OUTLINE_VARIANT};
    max-height: 1px;
    border: none;
}}

/* ─── QTabWidget ────────────────────────────────── */
QTabWidget::pane {{
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: 0 {RADIUS_LG}px {RADIUS_LG}px {RADIUS_LG}px;
    background: {COLOR_SURFACE_LOWEST};
    top: -1px;
}}
QTabBar::tab {{
    padding: {SPACING_SM}px {SPACING_MD}px;
    font-size: {TYPO_LABEL_MD[0]}px;
    font-weight: {TYPO_LABEL_MD[1]};
    color: {COLOR_ON_SURFACE_VARIANT};
    border-bottom: 2px solid transparent;
    background: transparent;
    min-width: 80px;
}}
QTabBar::tab:selected {{
    color: {COLOR_PRIMARY};
    border-bottom: 2px solid {COLOR_PRIMARY};
    background: {COLOR_SURFACE_LOWEST};
}}
QTabBar::tab:hover:!selected {{
    color: {COLOR_ON_SURFACE};
    background: {COLOR_SURFACE_LOW};
}}

/* ─── QCheckBox / QRadioButton ──────────────────── */
QCheckBox, QRadioButton {{
    spacing: {SPACING_SM}px;
    color: {COLOR_ON_SURFACE};
    font-size: {TYPO_BODY_MD[0]}px;
}}
QCheckBox::indicator {{
    width: 18px; height: 18px;
    border: 2px solid {COLOR_OUTLINE};
    border-radius: {RADIUS_SM}px;
    background: {COLOR_SURFACE_LOWEST};
}}
QCheckBox::indicator:checked {{
    background-color: {COLOR_PRIMARY};
    border-color: {COLOR_PRIMARY};
    image: url(:/icons/check_white.svg);
}}
QCheckBox::indicator:hover {{
    border-color: {COLOR_PRIMARY};
}}
QRadioButton::indicator {{
    width: 18px; height: 18px;
    border: 2px solid {COLOR_OUTLINE};
    border-radius: 9px;
    background: {COLOR_SURFACE_LOWEST};
}}
QRadioButton::indicator:checked {{
    background-color: {COLOR_PRIMARY};
    border-color: {COLOR_PRIMARY};
}}

/* ─── QSlider ───────────────────────────────────── */
QSlider::groove:horizontal {{
    height: 4px;
    background: {COLOR_SURFACE_HIGH};
    border-radius: 2px;
}}
QSlider::handle:horizontal {{
    width: 16px; height: 16px;
    background: {COLOR_PRIMARY};
    border-radius: 8px;
    margin: -6px 0;
}}
QSlider::sub-page:horizontal {{
    background: {COLOR_PRIMARY};
    border-radius: 2px;
}}

/* ─── QMenuBar / QMenu ──────────────────────────── */
QMenuBar {{
    background-color: {COLOR_SURFACE};
    border-bottom: 1px solid {COLOR_OUTLINE_VARIANT};
    padding: 2px {SPACING_SM}px;
    font-size: {TYPO_BODY_MD[0]}px;
}}
QMenuBar::item {{
    padding: {SPACING_XS}px {SPACING_SM}px;
    border-radius: {RADIUS_DEFAULT}px;
    color: {COLOR_ON_SURFACE};
}}
QMenuBar::item:selected {{
    background-color: {COLOR_SURFACE_HIGH};
}}
QMenu {{
    background-color: {COLOR_SURFACE_LOWEST};
    border: 1px solid {COLOR_OUTLINE_VARIANT};
    border-radius: {RADIUS_LG}px;
    padding: {SPACING_XS}px;
}}
QMenu::item {{
    padding: {SPACING_SM}px {SPACING_LG}px {SPACING_SM}px {SPACING_MD}px;
    border-radius: {RADIUS_DEFAULT}px;
    font-size: {TYPO_BODY_MD[0]}px;
    color: {COLOR_ON_SURFACE};
    min-height: 32px;
}}
QMenu::item:selected {{
    background-color: {COLOR_SURFACE_LOW};
}}
QMenu::separator {{
    height: 1px;
    background: {COLOR_OUTLINE_VARIANT};
    margin: {SPACING_XS}px {SPACING_SM}px;
}}

/* ─── QToolTip ──────────────────────────────────── */
QToolTip {{
    background-color: {COLOR_INVERSE_SURFACE};
    color: {COLOR_INVERSE_ON_SURFACE};
    border: none;
    border-radius: {RADIUS_SM}px;
    padding: {SPACING_XS}px {SPACING_SM}px;
    font-size: {TYPO_LABEL_SM[0]}px;
}}

/* ─── QProgressBar ──────────────────────────────── */
QProgressBar {{
    background-color: {COLOR_SURFACE_HIGH};
    border-radius: 2px;
    border: none;
    height: 4px;
    text-align: center;
    font-size: 0px;
}}
QProgressBar::chunk {{
    background-color: {COLOR_PRIMARY};
    border-radius: 2px;
}}

/* ─── QStatusBar ────────────────────────────────── */
QStatusBar {{
    background-color: {COLOR_SURFACE};
    border-top: 1px solid {COLOR_OUTLINE_VARIANT};
    font-size: {TYPO_LABEL_SM[0]}px;
    color: {COLOR_ON_SURFACE_VARIANT};
    padding: 0 {SPACING_MD}px;
    min-height: 24px;
}}
"""
```

---

## Dark Theme Overrides

Rather than duplicating the full QSS, apply dark overrides that shadow light tokens:

```python
DARK_QSS = LIGHT_QSS.replace(
    COLOR_BACKGROUND,       DARK_BACKGROUND
).replace(
    COLOR_SURFACE_LOWEST,   DARK_SURFACE
).replace(
    COLOR_SURFACE_LOW,      DARK_SURFACE_CONTAINER
).replace(
    COLOR_SURFACE_HIGH,     DARK_SURFACE_HIGH
).replace(
    COLOR_OUTLINE_VARIANT,  DARK_BORDER
).replace(
    COLOR_ON_SURFACE,       DARK_ON_SURFACE
).replace(
    COLOR_ON_SURFACE_VARIANT, DARK_ON_SURFACE_VARIANT
).replace(
    COLOR_OUTLINE,          DARK_OUTLINE
)
# NOTE: Verify replacements don't cause false matches.
# For complex themes, prefer maintaining two full QSS strings.
```

**Better approach for production:** maintain two distinct QSS constants and
switch them wholesale. The replacement trick above works for prototypes.

---

## Sidebar QSS Pattern

```python
SIDEBAR_QSS = f"""
#sidebar {{
    background-color: {COLOR_SURFACE};
    border-right: 1px solid {COLOR_OUTLINE_VARIANT};
    min-width: 240px;
    max-width: 240px;
}}
#sidebar[collapsed="true"] {{
    min-width: 64px;
    max-width: 64px;
}}
#sidebarNavItem {{
    background: transparent;
    border: none;
    border-left: 3px solid transparent;
    border-radius: 0;
    padding: {SPACING_SM}px {SPACING_MD}px;
    text-align: left;
    color: {COLOR_ON_SURFACE_VARIANT};
    font-size: {TYPO_BODY_MD[0]}px;
    font-weight: 400;
    min-height: 40px;
}}
#sidebarNavItem:hover {{
    background-color: {COLOR_SURFACE_HIGH};
    color: {COLOR_ON_SURFACE};
    border-left-color: {COLOR_OUTLINE_VARIANT};
}}
#sidebarNavItem[active="true"] {{
    background-color: {COLOR_PRIMARY_FIXED};
    color: {COLOR_PRIMARY};
    border-left: 3px solid {COLOR_PRIMARY};
    font-weight: 600;
}}
#sidebarNavItem[active="true"]:hover {{
    background-color: {COLOR_PRIMARY_FIXED};
}}
"""
```

---

## Card Frame Pattern

```python
# In Python layout code:
def make_card(parent=None) -> QFrame:
    card = QFrame(parent)
    card.setObjectName("card")
    card.setProperty("class", "card")
    card.setFrameShape(QFrame.Shape.StyledPanel)
    layout = QVBoxLayout(card)
    layout.setContentsMargins(SPACING_MD, SPACING_MD, SPACING_MD, SPACING_MD)
    layout.setSpacing(SPACING_SM)
    return card

# Optional: Level-2 elevation shadow for modals/dropdowns
from PyQt6.QtWidgets import QGraphicsDropShadowEffect
from PyQt6.QtGui import QColor

def add_elevation_shadow(widget, level=2):
    shadow = QGraphicsDropShadowEffect()
    shadow.setBlurRadius(12)
    shadow.setOffset(0, 4)
    shadow.setColor(QColor(0, 0, 0, int(0.08 * 255)))  # rgba(0,0,0,0.08)
    widget.setGraphicsEffect(shadow)
```

---

## Font Loading Pattern

```python
from PyQt6.QtGui import QFontDatabase, QFont

def load_fonts():
    """
    Load custom fonts. Place .ttf/.otf files in assets/fonts/.
    Falls back to system fonts if not found.
    """
    fonts_dir = Path(__file__).parent / "assets" / "fonts"
    for font_file in fonts_dir.glob("*.ttf"):
        QFontDatabase.addApplicationFont(str(font_file))
    for font_file in fonts_dir.glob("*.otf"):
        QFontDatabase.addApplicationFont(str(font_file))

def make_font(family: str, size: int, weight: int = 400) -> QFont:
    font = QFont(family, size)
    font.setWeight(QFont.Weight(weight))
    return font

def apply_typography(widget, scale: tuple):
    """Apply a typography scale tuple (size, weight, line_height, letter_spacing)."""
    size, weight, *_ = scale
    font = widget.font()
    font.setPointSize(size)
    font.setWeight(QFont.Weight(weight))
    widget.setFont(font)
```

---

## Material Symbols Icon Pattern

```python
# Load Material Symbols Outlined font (include in assets/fonts/)
# Download from: https://github.com/google/material-design-icons

def make_icon_label(codepoint: str, size: int = 20,
                    color: str = COLOR_ON_SURFACE_VARIANT) -> QLabel:
    """
    Create a label that renders a Material Symbol.
    codepoint: unicode codepoint string e.g. "\\ue88a" for 'home'
    Common icons:
        dashboard     e871   search        e8b6
        settings      e8b8   notifications e7f4
        add           e145   download      f090
        inventory_2   e1bc   edit          e3c9
        delete        e872   close         e5cd
        check         e876   chevron_right e5cc
        person        e7fd   logout        e9ba
        dark_mode     e51c   light_mode    e518
        menu          e5d2   arrow_back    e5c4
        filter_list   ef4f   more_vert     e5d4
    """
    label = QLabel(codepoint)
    font = QFont("Material Symbols Outlined")
    font.setPixelSize(size)
    label.setFont(font)
    label.setStyleSheet(f"color: {color}; background: transparent;")
    return label
```

---

## Status Badge / Chip Pattern

```python
STATUS_QSS = f"""
QLabel[status="success"] {{
    background-color: {COLOR_SUCCESS_BG};
    color: {COLOR_SUCCESS};
    border-radius: {RADIUS_FULL}px;
    padding: 2px {SPACING_SM}px;
    font-size: {TYPO_LABEL_SM[0]}px;
    font-weight: 600;
}}
QLabel[status="warning"] {{
    background-color: {COLOR_WARNING_BG};
    color: {COLOR_WARNING};
    border-radius: {RADIUS_FULL}px;
    padding: 2px {SPACING_SM}px;
    font-size: {TYPO_LABEL_SM[0]}px;
    font-weight: 600;
}}
QLabel[status="danger"] {{
    background-color: {COLOR_ERROR_CONTAINER};
    color: {COLOR_ERROR};
    border-radius: {RADIUS_FULL}px;
    padding: 2px {SPACING_SM}px;
    font-size: {TYPO_LABEL_SM[0]}px;
    font-weight: 600;
}}
"""

def make_badge(text: str, status: str = "success") -> QLabel:
    """status: 'success' | 'warning' | 'danger'"""
    badge = QLabel(text)
    badge.setProperty("status", status)
    badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
    badge.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
    return badge
```

---

## Responsive Property Updates

After changing a widget property, call `style().polish()` to force QSS re-evaluation:

```python
widget.setProperty("active", True)
widget.style().unpolish(widget)
widget.style().polish(widget)
widget.update()
```

This is required when toggling dynamic QSS properties like `active`, `class`,
`collapsed`, `hasError`, `status`, etc.