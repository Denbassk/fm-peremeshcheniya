"""Окно «Поставщики: кого не вывозить» (как open_suppliers_dialog): ГАЛОЧКА = НЕ вывозим, без галочки - вывозим.
«Сохранить» пишет Справочник\\vyvoz_suppliers.csv через core.save_supplier_choice."""
from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt
from PySide6.QtWidgets import QComboBox, QDialog, QHBoxLayout, QLabel, QLineEdit, QSizePolicy, QVBoxLayout

from . import tokens as T
from .widgets import button, help_button, muted, table

FILTERS = (u"Все", u"Только не вывозим", u"Только вывозим")


class SupModel(QAbstractTableModel):
    COLS = ((u"Поставщик (галочка - не вывозим)", Qt.AlignLeft), (u"Поз.", Qt.AlignRight), (u"Сумма, грн", Qt.AlignRight))

    def __init__(self, core, known, order):
        super().__init__()
        self.core, self.known, self.order = core, known, order
        self.on = {n: bool(known[n]["excluded"]) for n in order}

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.order)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else 3

    def headerData(self, s, o, role=Qt.DisplayRole):
        if o == Qt.Horizontal and role == Qt.DisplayRole:
            return self.COLS[s][0]
        if o == Qt.Horizontal and role == Qt.TextAlignmentRole:
            return int(self.COLS[s][1] | Qt.AlignVCenter)
        return None

    def data(self, ix, role=Qt.DisplayRole):
        n = self.order[ix.row()]
        d = self.known[n]
        if role == Qt.DisplayRole:
            if ix.column() == 0:
                return d["name"]
            if ix.column() == 1:
                return str(d["pos"]) if d["pos"] else u"-"
            return self.core.fmt_money(d["sum"]) if d["pos"] else u"-"
        if role == Qt.CheckStateRole and ix.column() == 0:
            return Qt.Checked if self.on[n] else Qt.Unchecked
        if role == Qt.TextAlignmentRole:
            return int(self.COLS[ix.column()][1] | Qt.AlignVCenter)
        if role == Qt.UserRole:
            return n
        return None

    def flags(self, ix):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        return f | Qt.ItemIsUserCheckable if ix.column() == 0 else f

    def setData(self, ix, value, role=Qt.EditRole):
        if role == Qt.CheckStateRole and ix.column() == 0:
            self.on[self.order[ix.row()]] = Qt.CheckState(value) == Qt.Checked
            self.dataChanged.emit(ix, ix, [Qt.CheckStateRole])
            return True
        return False

    def set_rows(self, names, val):
        for n in names:
            self.on[n] = (not self.on[n]) if val is None else val
        self.dataChanged.emit(self.index(0, 0), self.index(len(self.order) - 1, 0), [Qt.CheckStateRole])


class SupFilter(QSortFilterProxyModel):
    def __init__(self, core, model):
        super().__init__()
        self.core, self.q, self.mode = core, u"", FILTERS[0]
        self.setSourceModel(model)

    def filterAcceptsRow(self, r, parent):
        m = self.sourceModel()
        n = m.order[r]
        if self.q and self.q not in n:
            return False
        return self.mode == FILTERS[0] or (self.mode == FILTERS[1]) == m.on[n]

    def visible(self):
        return [self.sourceModel().order[self.mapToSource(self.index(i, 0)).row()] for i in range(self.rowCount())]


class SuppliersDialog(QDialog):
    def __init__(self, core, dirs, info=None, parent=None):
        super().__init__(parent)
        self.core, self.dirs, self.saved = core, dirs, 0
        self.setWindowTitle(u"Поставщики: кого не вывозить")
        self.setModal(True)
        self.resize(720, 640)
        known = core.known_suppliers(dirs, info)
        order = sorted(known, key=lambda n: (-(known[n]["pos"] or 0), known[n]["name"].lower()))   # как в прежнем окне
        self.model = SupModel(core, known, order)
        self.proxy = SupFilter(core, self.model)
        v = QVBoxLayout(self)
        v.setContentsMargins(T.SPACING_LG, T.SPACING_LG, T.SPACING_LG, T.SPACING_MD)
        v.setSpacing(T.SPACING_SM + 4)
        hd = QHBoxLayout()
        t = QLabel(u"Поставщики: кого не вывозить")
        t.setObjectName("cardTitle")
        hd.addWidget(t)
        hd.addWidget(help_button(u"ГАЛОЧКА СТОИТ - товар поставщика НЕ вывозим: он остаётся на полке и идёт в инвентаризацию ЮА. "
                                 u"Без галочки - вывозим (новые поставщики тоже). Список хранится в файле и действует на все "
                                 u"следующие расчёты, пока вы его не измените. Кнопки над таблицей действуют на видимые строки "
                                 u"(с учётом поиска и фильтра)."))
        hd.addStretch(1)
        v.addLayout(hd)
        n0, when0, path0 = core.suppliers_file_info(dirs)
        f = muted(u"Сейчас в файле: не вывозим %d (%s)" % (n0, (u"сохранён " + when0) if when0 else u"ещё не сохранялся"))
        f.setToolTip(path0)
        v.addWidget(f)
        r = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText(u"Поиск поставщика")
        self.search.setClearButtonEnabled(True)
        self.search.setAccessibleName(u"Поиск поставщика")
        self.search.textChanged.connect(self.refilter)
        r.addWidget(self.search, 1)
        self.mode = QComboBox()
        self.mode.addItems(FILTERS)
        self.mode.setAccessibleName(u"Фильтр")
        self.mode.currentTextChanged.connect(self.refilter)
        r.addWidget(self.mode)
        v.addLayout(r)
        r = QHBoxLayout()
        r.addWidget(button(u"Не вывозить видимых", on_click=lambda: self.set_vis(True)))
        r.addWidget(button(u"Вывозить видимых", on_click=lambda: self.set_vis(False)))
        r.addWidget(button(u"Инвертировать", on_click=lambda: self.set_vis(None)))
        r.addStretch(1)
        self.cnt = muted(wrap=False)
        r.addWidget(self.cnt)
        v.addLayout(r)
        self.view = table(self.proxy, stretch_col=0)
        self.view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.view.setMinimumHeight(240)
        self.view.setAccessibleName(u"Поставщики")
        self.view.activated.connect(self.toggle)
        v.addWidget(self.view, 1)
        r = QHBoxLayout()
        r.addStretch(1)
        r.addWidget(button(u"Отмена", on_click=self.reject))
        ok = button(u"Сохранить", "check", kind="primary", on_click=self.save)
        ok.setDefault(True)
        r.addWidget(ok)
        v.addLayout(r)
        self.model.dataChanged.connect(lambda *_: self.count())
        self.count()
        self.search.setFocus()

    def refilter(self, *_a):
        self.proxy.q = self.core.sup_norm(self.search.text())
        self.proxy.mode = self.mode.currentText()
        self.proxy.invalidateFilter()
        self.count()

    def count(self):
        self.cnt.setText(u"Не вывозим: %d из %d" % (sum(1 for x in self.model.on.values() if x), len(self.model.on)))

    def set_vis(self, val):
        self.model.set_rows(self.proxy.visible(), val)
        if self.proxy.mode != FILTERS[0]:
            self.proxy.invalidateFilter()

    def toggle(self, ix):
        """Enter или двойной щелчок по строке - поставить или снять галочку (не только мышью по квадратику)."""
        src = self.model.index(self.proxy.mapToSource(ix).row(), 0)
        self.model.setData(src, Qt.Unchecked if self.model.on[self.model.order[src.row()]] else Qt.Checked, Qt.CheckStateRole)

    def save(self):
        self.saved = self.core.save_supplier_choice(self.dirs, self.model.known, [n for n, x in self.model.on.items() if x])
        self.accept()
