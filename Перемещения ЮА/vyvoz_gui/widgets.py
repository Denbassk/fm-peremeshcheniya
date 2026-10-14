"""Общие элементы окна: карточка, кнопки, «?»-подсказка, плашка, модель таблицы со статусом-значком."""
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSize, Qt, Signal
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (QAbstractItemView, QFrame, QHBoxLayout, QHeaderView, QLabel, QPushButton, QSizePolicy,
                               QTableView, QToolTip, QVBoxLayout, QWidget)

from . import assets as A
from . import tokens as T


class Theme(object):
    """Текущая палитра для того, что красится кодом (значки, цвета ячеек). Обновляется ThemeManager.changed."""
    c = T.LIGHT
    listeners = []

    @classmethod
    def set(cls, c):
        cls.c = c
        for fn in list(cls.listeners):
            try:
                fn()
            except RuntimeError:                     # виджет уже удалён
                cls.listeners.remove(fn)


def on_theme(fn):
    Theme.listeners.append(fn)
    fn()


def repolish(w):
    w.style().unpolish(w)
    w.style().polish(w)
    w.update()


def button(text, icon_name=None, kind=None, tip=None, on_click=None):
    b = QPushButton(text)
    b.setCursor(Qt.PointingHandCursor)
    if kind:
        b.setProperty("kind", kind)
    if tip:
        b.setToolTip(tip)
    if on_click:
        b.clicked.connect(on_click)
    if icon_name:
        def paint():
            col = Theme.c["on_primary"] if kind == "primary" else Theme.c["text"]
            b.setIcon(A.icon(icon_name, col, 18))
            b.setIconSize(QSize(18, 18))
        on_theme(paint)
    return b


def help_button(text, title=u""):
    """Кнопка «?»: длинное пояснение во всплывающей подсказке; по щелчку или Enter/пробелу - тоже (не только мышью)."""
    b = QPushButton()
    b.setProperty("kind", "ghost")
    b.setCursor(Qt.WhatsThisCursor)
    b.setToolTip(text)
    b.setAccessibleName(u"Пояснение" + (u": " + title if title else u""))
    b.setAccessibleDescription(text)
    b.setFixedSize(QSize(32, 32))
    on_theme(lambda: (b.setIcon(A.icon("help", Theme.c["text_muted"], 20)), b.setIconSize(QSize(20, 20))))
    b.clicked.connect(lambda: QToolTip.showText(b.mapToGlobal(b.rect().bottomLeft()), b.toolTip(), b))   # текст можно менять
    return b


def muted(text=u"", wrap=True):
    lb = QLabel(text)
    lb.setProperty("muted", "true")
    lb.setWordWrap(wrap)
    return lb


class Card(QFrame):
    """Карточка: заголовок (+ «?» и кнопки справа) и содержимое."""

    def __init__(self, title=None, help_text=None, parent=None):
        super().__init__(parent)
        self.setObjectName("card")
        self.lay = QVBoxLayout(self)
        self.lay.setContentsMargins(T.SPACING_MD, T.SPACING_MD, T.SPACING_MD, T.SPACING_MD)
        self.lay.setSpacing(T.SPACING_SM + 4)
        self.head = QHBoxLayout()
        self.head.setSpacing(T.SPACING_SM)
        if title:
            t = QLabel(title)
            t.setObjectName("cardTitle")
            self.head.addWidget(t)
        if help_text:
            self.head.addWidget(help_button(help_text, title or u""))
        self.head.addStretch(1)
        self.lay.addLayout(self.head)

    def add_head(self, w):
        self.head.addWidget(w)
        return w


class FitLabel(QLabel):
    """Текст с переносом, который сначала занимает ширину одной строки: QLabel с переносом сам просит узкий столбик,
    а этот - сколько нужно тексту в одну строку; не хватает места - переносится по ширине карточки."""

    def sizeHint(self):
        s = super().sizeHint()
        return QSize(self.fontMetrics().horizontalAdvance(self.text()) + 4, s.height())

    def minimumSizeHint(self):
        s = super().minimumSizeHint()
        return QSize(min(self.sizeHint().width(), 12 * T.SPACING_MD), s.height())


class StatusLine(QWidget):
    """Значок + текст: состояние читается без цвета (уровни OK / WARN / ERR / INFO). help - «?» сразу после текста
    (длинное пояснение во всплывающей подсказке, set_help меняет его)."""

    def __init__(self, text=u"", level="INFO", parent=None, help=None):
        super().__init__(parent)
        h = QHBoxLayout(self)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(T.SPACING_SM)
        self.ic = QLabel()
        self.ic.setFixedSize(20, 20)
        self.tx = FitLabel() if help is not None else QLabel()
        self.tx.setWordWrap(True)
        self.tx.setTextInteractionFlags(Qt.TextSelectableByMouse)
        h.addWidget(self.ic, 0, Qt.AlignTop)
        self.help = None
        if help is None:
            h.addWidget(self.tx, 1)
        else:
            self.help = help_button(help)
            h.addWidget(self.tx)
            h.addWidget(self.help, 0, Qt.AlignTop)
            h.addStretch(1)
        self.level = level
        on_theme(self._paint)
        self.set(text, level)

    def _paint(self):
        name, tok, _w = T.LEVELS.get(self.level, T.LEVELS["INFO"])
        self.ic.setPixmap(A.glyph_pixmap(name, Theme.c[tok], 20, self.devicePixelRatioF() or 1.0))

    def set(self, text, level="INFO"):
        self.level = level
        self.tx.setText(text)
        self.tx.updateGeometry()
        self._paint()
        self.setVisible(bool(text))

    def set_help(self, text):
        if self.help is not None:
            self.help.setToolTip(text)
            self.help.setAccessibleDescription(text)


class Banner(QFrame):
    """Плашка итога: фон по уровню, значок, текст и кнопка действия («Открыть журнал»)."""
    action = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("banner")
        h = QHBoxLayout(self)
        h.setContentsMargins(T.SPACING_MD - 4, T.SPACING_SM, T.SPACING_SM, T.SPACING_SM)
        self.line = StatusLine()
        h.addWidget(self.line, 1)
        self.btn = button(u"", kind=None, on_click=self.action.emit)
        h.addWidget(self.btn, 0, Qt.AlignVCenter)
        self.hide()

    def show_msg(self, text, level="OK", action_text=None):
        self.setProperty("level", "ERR" if level == "ERROR" else level)
        self.line.set(text, level)
        self.btn.setVisible(bool(action_text))
        self.btn.setText(action_text or u"")
        repolish(self)
        self.show()


class RowsModel(QAbstractTableModel):
    """Строки-словари для QTableView. cols: [(ключ, заголовок, выравнивание)]; level_key - уровень строки, status_col -
    колонка, где рисуется значок уровня; цвет текста не единственный признак: в колонке статуса есть и значок, и слово."""

    def __init__(self, cols, level_key=None, status_col=None, parent=None):
        super().__init__(parent)
        self.cols, self.level_key, self.status_col, self.rows = cols, level_key, status_col, []

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = list(rows)
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.cols)

    def headerData(self, section, orient, role=Qt.DisplayRole):
        if orient == Qt.Horizontal and role == Qt.DisplayRole:
            return self.cols[section][1]
        if orient == Qt.Horizontal and role == Qt.TextAlignmentRole:
            return int(self.cols[section][2] | Qt.AlignVCenter)
        return None

    def data(self, ix, role=Qt.DisplayRole):
        if not ix.isValid():
            return None
        r, key, al = self.rows[ix.row()], self.cols[ix.column()][0], self.cols[ix.column()][2]
        lv = r.get(self.level_key, "INFO") if self.level_key else None
        if role == Qt.ToolTipRole and r.get(key + "_tip"):
            return r[key + "_tip"]
        if role in (Qt.DisplayRole, Qt.ToolTipRole):
            v = r.get(key, u"")
            return u"" if v is None else str(v)
        if role == Qt.TextAlignmentRole:
            return int(al | Qt.AlignVCenter)
        if role == Qt.DecorationRole and ix.column() == self.status_col and lv:
            name, tok, _w = T.LEVELS.get(lv, T.LEVELS["INFO"])
            return A.icon(name, Theme.c[tok], 18)
        if role == Qt.ForegroundRole and ix.column() == self.status_col and lv in ("ERR", "ERROR", "WARN"):
            return QColor(Theme.c[T.LEVELS[lv][1]])
        return None

    def row(self, i):
        return self.rows[i]


def table(model, stretch_col=None, multi=False):
    v = QTableView()
    v.setModel(model)
    v.setAlternatingRowColors(True)
    v.setSelectionBehavior(QAbstractItemView.SelectRows)
    v.setSelectionMode(QAbstractItemView.ExtendedSelection if multi else QAbstractItemView.SingleSelection)
    v.setEditTriggers(QAbstractItemView.NoEditTriggers)
    v.setShowGrid(False)
    v.setWordWrap(False)
    v.setTextElideMode(Qt.ElideMiddle)
    v.verticalHeader().setVisible(False)
    v.verticalHeader().setDefaultSectionSize(T.ROW_H)
    hh = v.horizontalHeader()
    hh.setHighlightSections(False)
    hh.setSectionResizeMode(QHeaderView.ResizeToContents)
    if stretch_col is not None:
        hh.setSectionResizeMode(stretch_col, QHeaderView.Stretch)
    v.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    v.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
    v.setIconSize(QSize(18, 18))
    return v


def fit_rows(view, min_rows=2, max_rows=8):
    """Высота таблицы по числу строк: без пустого места и без второй прокрутки внутри страницы до max_rows строк."""
    n = max(min_rows, min(max_rows, view.model().rowCount()))
    h = view.horizontalHeader().sizeHint().height() + n * view.verticalHeader().defaultSectionSize() + 2 * view.frameWidth() + 4
    view.setFixedHeight(h)


class ResultCard(Card):
    """Итог действия: плашка (уровень + «Журнал»), таблица по строкам результата, кнопки открыть файл / папку.
    Скрыта до первого запуска. on_log - открыть «Журнал»."""

    def __init__(self, cols, status_col=None, stretch_col=None, on_log=None, title=u"Результат"):
        super().__init__(title)
        self.banner = Banner()
        if on_log:
            self.banner.action.connect(on_log)
        self.lay.addWidget(self.banner)
        self.model = RowsModel(cols, level_key="level" if status_col is not None else None, status_col=status_col)
        self.view = table(self.model, stretch_col=stretch_col)
        self.view.setAccessibleName(title)
        self.lay.addWidget(self.view)
        self.btns = QHBoxLayout()
        self.btns.addStretch(1)
        self.lay.addLayout(self.btns)
        self.hide()

    def add_button(self, b):
        self.btns.insertWidget(self.btns.count() - 1, b)
        return b

    def show_result(self, rows, text, level, max_rows=8):
        self.model.set_rows(rows)
        fit_rows(self.view, 1, max_rows)
        self.view.setVisible(bool(rows))
        self.banner.show_msg(text, level, u"Открыть журнал" if level in ("ERR", "WARN") else u"Журнал")
        self.show()


def pick_files(parent, title, init, multi=False, pattern=u"Excel (*.xlsx *.xlsm *.xls);;Все файлы (*.*)"):
    """Окно выбора файла(ов) -> список путей (пусто - отменили)."""
    import os
    from PySide6.QtWidgets import QFileDialog
    init = init if init and os.path.isdir(init) else u""
    if multi:
        fps, _f = QFileDialog.getOpenFileNames(parent, title, init, pattern)
    else:
        fp, _f = QFileDialog.getOpenFileName(parent, title, init, pattern)
        fps = [fp] if fp else []
    return [os.path.normpath(x) for x in fps]
