# Component Library — PyQt6 Modern Enterprise Design System

Complete, copy-paste-ready PyQt6 component implementations.
Always read `design_tokens.md` and `qss_patterns.md` first.

---

## Table of Contents

1. [App Shell (Main Window)](#1-app-shell)
2. [Sidebar Navigation](#2-sidebar-navigation)
3. [Top App Bar](#3-top-app-bar)
4. [Stat / Metric Card](#4-stat-card)
5. [Data Table](#5-data-table)
6. [Search Input](#6-search-input)
7. [Primary / Secondary / Ghost Buttons](#7-buttons)
8. [Modal / Dialog](#8-modal)
9. [Settings Panel](#9-settings-panel)
10. [Toast Notification](#10-toast)

---

## 1. App Shell

```python
class AppShell(QMainWindow):
    """
    Fixed sidebar (240px) + fixed top bar (64px) + scrollable content area.
    Matches the layout defined in design_tokens.md Layout Structure.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Enterprise App")
        self.setMinimumSize(1024, 600)

        central = QWidget()
        self.setCentralWidget(central)
        root = QHBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # Sidebar
        self.sidebar = SidebarNav(self)
        root.addWidget(self.sidebar)

        # Right column: top bar + content
        right_col = QVBoxLayout()
        right_col.setContentsMargins(0, 0, 0, 0)
        right_col.setSpacing(0)

        self.top_bar = TopAppBar(self)
        right_col.addWidget(self.top_bar)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setObjectName("contentScroll")

        self.content = QWidget()
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(
            SPACING_LG, SPACING_LG, SPACING_LG, SPACING_LG
        )
        self.content_layout.setSpacing(SPACING_LG)
        self.content_layout.addStretch()

        self.scroll_area.setWidget(self.content)
        right_col.addWidget(self.scroll_area, 1)

        right_widget = QWidget()
        right_widget.setLayout(right_col)
        root.addWidget(right_widget, 1)
```

---

## 2. Sidebar Navigation

```python
class NavItem:
    def __init__(self, label: str, icon_codepoint: str, page_key: str):
        self.label = label
        self.icon = icon_codepoint
        self.page = page_key

NAV_ITEMS = [
    NavItem("Dashboard",  "\ue871", "dashboard"),
    NavItem("Inventory",  "\ue1bc", "inventory"),
    NavItem("Analytics",  "\ue6b1", "analytics"),
    NavItem("Settings",   "\ue8b8", "settings"),
]

class SidebarNav(QWidget):
    page_changed = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("sidebar")
        self.setFixedWidth(240)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, SPACING_LG, 0, SPACING_LG)
        layout.setSpacing(0)

        # Brand / logo area
        brand = QWidget()
        brand_layout = QVBoxLayout(brand)
        brand_layout.setContentsMargins(SPACING_MD, 0, SPACING_MD, 0)
        brand_layout.setSpacing(SPACING_XS)

        title = QLabel("Enterprise Core")
        title.setProperty("class", "headline-md")
        subtitle = QLabel("ADMIN CONSOLE")
        subtitle.setProperty("class", "label-md")

        brand_layout.addWidget(title)
        brand_layout.addWidget(subtitle)
        layout.addWidget(brand)

        # Spacer
        layout.addSpacing(SPACING_XL)

        # Nav items
        self._buttons: list[QPushButton] = []
        self._active_key: str = NAV_ITEMS[0].page

        for item in NAV_ITEMS:
            btn = self._make_nav_button(item)
            layout.addWidget(btn)
            self._buttons.append(btn)

        layout.addStretch()

        # User profile footer
        footer = self._make_user_footer()
        layout.addWidget(footer)

        self._set_active(self._active_key)

    def _make_nav_button(self, item: NavItem) -> QPushButton:
        btn = QPushButton()
        btn.setObjectName("sidebarNavItem")
        btn.setProperty("page", item.page)
        btn.setProperty("active", False)
        btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        btn.setFixedHeight(44)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)

        row = QHBoxLayout(btn)
        row.setContentsMargins(SPACING_MD, 0, SPACING_MD, 0)
        row.setSpacing(SPACING_SM + SPACING_XS)  # 12px

        icon = QLabel(item.icon)
        icon.setFont(QFont("Material Symbols Outlined", 20))
        icon.setFixedWidth(24)
        icon.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

        label = QLabel(item.label)
        label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)

        row.addWidget(icon)
        row.addWidget(label)
        row.addStretch()

        btn.clicked.connect(lambda checked, p=item.page: self._on_nav_click(p))
        return btn

    def _make_user_footer(self) -> QWidget:
        footer = QWidget()
        footer.setObjectName("sidebarFooter")
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(SPACING_MD, SPACING_MD, SPACING_MD, 0)
        footer_layout.setSpacing(SPACING_SM)

        # Avatar placeholder
        avatar = QLabel("AS")
        avatar.setFixedSize(32, 32)
        avatar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        avatar.setStyleSheet(f"""
            background-color: {COLOR_PRIMARY_CONTAINER};
            color: {COLOR_ON_PRIMARY};
            border-radius: 16px;
            font-weight: 600;
            font-size: 13px;
        """)

        name_col = QVBoxLayout()
        name_col.setSpacing(0)
        name = QLabel("Alex Sterling")
        name.setStyleSheet(f"font-weight: 600; font-size: {TYPO_LABEL_MD[0]}px;")
        role = QLabel("Lead Administrator")
        role.setProperty("class", "muted")

        name_col.addWidget(name)
        name_col.addWidget(role)

        footer_layout.addWidget(avatar)
        footer_layout.addLayout(name_col)
        footer_layout.addStretch()
        return footer

    def _on_nav_click(self, page_key: str):
        self._set_active(page_key)
        self.page_changed.emit(page_key)

    def _set_active(self, page_key: str):
        self._active_key = page_key
        for btn in self._buttons:
            is_active = btn.property("page") == page_key
            btn.setProperty("active", is_active)
            btn.style().unpolish(btn)
            btn.style().polish(btn)
```

---

## 3. Top App Bar

```python
class TopAppBar(QWidget):
    search_changed = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("topBar")
        self.setFixedHeight(64)
        self.setStyleSheet(f"""
            #topBar {{
                background-color: {COLOR_SURFACE};
                border-bottom: 1px solid {COLOR_OUTLINE_VARIANT};
            }}
        """)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(SPACING_LG, 0, SPACING_LG, 0)
        layout.setSpacing(SPACING_MD)

        # Search bar (left-aligned, grows)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search system resources...")
        self.search.setFixedHeight(36)
        self.search.setMaximumWidth(420)
        self.search.textChanged.connect(self.search_changed)
        layout.addWidget(self.search)
        layout.addStretch()

        # Action buttons (right)
        for codepoint, tooltip in [("\ue7f4", "Notifications"),
                                    ("\ue8fd", "Help")]:
            btn = QPushButton(codepoint)
            btn.setFont(QFont("Material Symbols Outlined", 20))
            btn.setFixedSize(36, 36)
            btn.setToolTip(tooltip)
            btn.setProperty("class", "ghost")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            layout.addWidget(btn)

        # Divider
        divider = QFrame()
        divider.setFrameShape(QFrame.Shape.VLine)
        divider.setFixedHeight(28)
        divider.setStyleSheet(f"background: {COLOR_OUTLINE_VARIANT};")
        layout.addWidget(divider)

        # App name / brand
        brand = QLabel("AdminPanel")
        brand.setProperty("class", "headline-sm")
        layout.addWidget(brand)
```

---

## 4. Stat Card

```python
class StatCard(QFrame):
    """
    A metric/KPI card: icon + label + large value + optional delta badge.
    """
    def __init__(self, title: str, value: str, delta: str = "",
                 icon: str = "\ue8b6", positive: bool = True, parent=None):
        super().__init__(parent)
        self.setProperty("class", "card")
        self.setFixedHeight(120)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(SPACING_MD, SPACING_MD, SPACING_MD, SPACING_MD)
        layout.setSpacing(SPACING_MD)

        # Icon container
        icon_bg = QLabel(icon)
        icon_bg.setFont(QFont("Material Symbols Outlined", 24))
        icon_bg.setFixedSize(48, 48)
        icon_bg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_bg.setStyleSheet(f"""
            background-color: {COLOR_PRIMARY_FIXED};
            color: {COLOR_PRIMARY};
            border-radius: {RADIUS_LG}px;
        """)
        layout.addWidget(icon_bg)

        # Text column
        text_col = QVBoxLayout()
        text_col.setSpacing(SPACING_XS)

        lbl = QLabel(title.upper())
        lbl.setProperty("class", "label-md")
        text_col.addWidget(lbl)

        val = QLabel(value)
        val.setProperty("class", "headline-md")
        text_col.addWidget(val)

        if delta:
            delta_color = COLOR_SUCCESS if positive else COLOR_DANGER
            delta_lbl = QLabel(delta)
            delta_lbl.setStyleSheet(f"""
                color: {delta_color};
                font-size: {TYPO_LABEL_SM[0]}px;
                font-weight: 600;
            """)
            text_col.addWidget(delta_lbl)

        text_col.addStretch()
        layout.addLayout(text_col)
        layout.addStretch()
```

---

## 5. Data Table

```python
class DataTable(QTableWidget):
    """
    A styled enterprise data table with alternating rows, compact density,
    and proper header typography.
    """
    def __init__(self, columns: list[str], parent=None):
        super().__init__(parent)
        self.setColumnCount(len(columns))
        self.setHorizontalHeaderLabels(columns)

        # Behavior
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.setShowGrid(False)
        self.setWordWrap(False)
        self.verticalHeader().setVisible(False)
        self.verticalHeader().setDefaultSectionSize(40)  # Standard density
        self.horizontalHeader().setStretchLastSection(True)
        self.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Interactive
        )
        self.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.setFrameShape(QFrame.Shape.NoFrame)

    def add_row(self, values: list[str],
                status_col: int = -1,
                status: str = ""):
        """Append a row. Optionally render a status badge in status_col."""
        row = self.rowCount()
        self.insertRow(row)
        for col, val in enumerate(values):
            if col == status_col and status:
                badge = make_badge(val, status)
                self.setCellWidget(row, col, badge)
            else:
                item = QTableWidgetItem(val)
                item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                self.setItem(row, col, item)

    def set_compact_density(self):
        self.verticalHeader().setDefaultSectionSize(32)

    def set_standard_density(self):
        self.verticalHeader().setDefaultSectionSize(40)
```

---

## 6. Search Input

```python
class SearchInput(QWidget):
    """Search bar with leading icon and clear button."""
    text_changed = pyqtSignal(str)

    def __init__(self, placeholder: str = "Search...", parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        container = QFrame()
        container.setStyleSheet(f"""
            QFrame {{
                background: {COLOR_SURFACE_LOW};
                border: 1px solid {COLOR_OUTLINE_VARIANT};
                border-radius: {RADIUS_DEFAULT}px;
            }}
        """)
        row = QHBoxLayout(container)
        row.setContentsMargins(SPACING_SM, 0, SPACING_SM, 0)
        row.setSpacing(SPACING_SM)

        icon = QLabel("\ue8b6")  # search
        icon.setFont(QFont("Material Symbols Outlined", 18))
        icon.setStyleSheet(f"color: {COLOR_ON_SURFACE_VARIANT};")
        row.addWidget(icon)

        self.input = QLineEdit()
        self.input.setPlaceholderText(placeholder)
        self.input.setFrame(False)
        self.input.setFixedHeight(36)
        self.input.setStyleSheet("background: transparent; border: none;")
        self.input.textChanged.connect(self.text_changed)
        row.addWidget(self.input, 1)

        layout.addWidget(container)
```

---

## 7. Buttons

```python
def make_primary_button(text: str, icon: str = "") -> QPushButton:
    btn = QPushButton(f"  {text}" if icon else text)
    btn.setProperty("class", "primary")
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFixedHeight(36)
    if icon:
        btn.setIcon(...)  # or use a label-based approach
    return btn

def make_secondary_button(text: str) -> QPushButton:
    btn = QPushButton(text)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFixedHeight(36)
    return btn

def make_ghost_button(icon_codepoint: str, tooltip: str = "") -> QPushButton:
    btn = QPushButton(icon_codepoint)
    btn.setFont(QFont("Material Symbols Outlined", 20))
    btn.setProperty("class", "ghost")
    btn.setFixedSize(36, 36)
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    if tooltip:
        btn.setToolTip(tooltip)
    return btn

def make_danger_button(text: str) -> QPushButton:
    btn = QPushButton(text)
    btn.setProperty("class", "danger")
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    btn.setFixedHeight(36)
    return btn
```

---

## 8. Modal / Dialog

```python
class ConfirmDialog(QDialog):
    """
    Standard confirmation modal with title, message, cancel + confirm.
    Level-2 elevation: uses drop shadow.
    """
    def __init__(self, title: str, message: str,
                 confirm_text: str = "Confirm",
                 danger: bool = False, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setModal(True)
        self.setFixedWidth(400)
        self.setWindowFlags(
            Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # Outer container with shadow
        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 16, 16, 16)

        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background: {COLOR_SURFACE_LOWEST};
                border-radius: {RADIUS_LG}px;
                border: 1px solid {COLOR_OUTLINE_VARIANT};
            }}
        """)
        add_elevation_shadow(card)

        inner = QVBoxLayout(card)
        inner.setContentsMargins(SPACING_LG, SPACING_LG, SPACING_LG, SPACING_LG)
        inner.setSpacing(SPACING_MD)

        # Title row
        title_row = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setProperty("class", "headline-sm")
        title_row.addWidget(title_label)
        title_row.addStretch()

        close_btn = make_ghost_button("\ue5cd")  # close
        close_btn.clicked.connect(self.reject)
        title_row.addWidget(close_btn)
        inner.addLayout(title_row)

        # Divider
        div = QFrame()
        div.setFrameShape(QFrame.Shape.HLine)
        div.setStyleSheet(f"background: {COLOR_OUTLINE_VARIANT}; max-height: 1px;")
        inner.addWidget(div)

        # Message
        msg = QLabel(message)
        msg.setWordWrap(True)
        msg.setProperty("class", "muted")
        inner.addWidget(msg)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel = make_secondary_button("Cancel")
        cancel.clicked.connect(self.reject)
        confirm = make_danger_button(confirm_text) if danger \
                  else make_primary_button(confirm_text)
        confirm.clicked.connect(self.accept)
        btn_row.addWidget(cancel)
        btn_row.addSpacing(SPACING_SM)
        btn_row.addWidget(confirm)
        inner.addLayout(btn_row)

        outer.addWidget(card)
```

---

## 9. Settings Panel

```python
class SettingsPanel(QWidget):
    """A settings page with grouped rows: label + control side by side."""
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(SPACING_LG)
        self._build(layout)

    def _build(self, layout):
        self._add_section(layout, "Appearance", [
            ("Theme", self._make_theme_toggle()),
            ("Font Size", self._make_combo(["Small", "Medium", "Large"])),
        ])
        self._add_section(layout, "Notifications", [
            ("Email Alerts", self._make_toggle()),
            ("Desktop Notifications", self._make_toggle(checked=True)),
        ])
        layout.addStretch()

    def _add_section(self, layout, title: str, rows: list):
        section = QFrame()
        section.setProperty("class", "card")
        sec_layout = QVBoxLayout(section)
        sec_layout.setContentsMargins(SPACING_MD, SPACING_MD, SPACING_MD, SPACING_MD)
        sec_layout.setSpacing(0)

        hdr = QLabel(title)
        hdr.setProperty("class", "headline-sm")
        sec_layout.addWidget(hdr)
        sec_layout.addSpacing(SPACING_SM)

        for i, (lbl_text, control) in enumerate(rows):
            row = QHBoxLayout()
            row.setContentsMargins(0, SPACING_SM, 0, SPACING_SM)
            lbl = QLabel(lbl_text)
            row.addWidget(lbl)
            row.addStretch()
            row.addWidget(control)
            sec_layout.addLayout(row)
            if i < len(rows) - 1:
                div = QFrame()
                div.setFrameShape(QFrame.Shape.HLine)
                div.setStyleSheet(
                    f"background: {COLOR_OUTLINE_VARIANT}; max-height: 1px;"
                )
                sec_layout.addWidget(div)

        layout.addWidget(section)

    def _make_combo(self, options: list[str]) -> QComboBox:
        combo = QComboBox()
        combo.addItems(options)
        combo.setFixedWidth(160)
        return combo

    def _make_toggle(self, checked: bool = False) -> QCheckBox:
        cb = QCheckBox()
        cb.setChecked(checked)
        return cb

    def _make_theme_toggle(self) -> QComboBox:
        return self._make_combo(["Light", "Dark", "System"])
```

---

## 10. Toast Notification

```python
class Toast(QFrame):
    """
    Ephemeral notification that slides in from bottom-right and auto-dismisses.
    """
    def __init__(self, message: str, level: str = "success",
                 duration_ms: int = 3000, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint |
            Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        color_map = {
            "success": (COLOR_SUCCESS, COLOR_SUCCESS_BG),
            "warning": (COLOR_WARNING, COLOR_WARNING_BG),
            "error":   (COLOR_DANGER,  COLOR_ERROR_CONTAINER),
        }
        icon_map = {
            "success": "\ue876",  # check_circle
            "warning": "\ue002",  # warning
            "error":   "\ue000",  # error
        }
        fg, bg = color_map.get(level, color_map["success"])

        self.setStyleSheet(f"""
            QFrame {{
                background: {bg};
                border: 1px solid {fg};
                border-radius: {RADIUS_SM}px;
                padding: {SPACING_SM}px;
            }}
        """)

        row = QHBoxLayout(self)
        row.setContentsMargins(SPACING_MD, SPACING_SM, SPACING_MD, SPACING_SM)
        row.setSpacing(SPACING_SM)

        icon = QLabel(icon_map.get(level, "\ue876"))
        icon.setFont(QFont("Material Symbols Outlined", 18))
        icon.setStyleSheet(f"color: {fg};")
        row.addWidget(icon)

        msg = QLabel(message)
        msg.setStyleSheet(f"color: {fg}; font-size: {TYPO_BODY_SM[0]}px;")
        row.addWidget(msg)
        row.addStretch()

        QTimer.singleShot(duration_ms, self.close)

    @classmethod
    def show_toast(cls, message: str, level: str = "success",
                   parent_widget=None, duration_ms: int = 3000):
        toast = cls(message, level, duration_ms, parent_widget)
        toast.adjustSize()
        if parent_widget:
            pos = parent_widget.rect().bottomRight()
            toast.move(
                parent_widget.mapToGlobal(pos) - QPoint(
                    toast.width() + SPACING_LG,
                    toast.height() + SPACING_LG
                )
            )
        toast.show()
        return toast
```