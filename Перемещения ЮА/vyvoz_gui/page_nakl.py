"""Экран «Сверка накладной»: проведённая накладная перемещения -> галочка «Едет» и «Факт» по строке -> новый ТСД.

Решение 28.10.2026: что уехало - оставляем в накладной, что осталось на магазине - убираем (его переводят на ЮА); где
количество другое - правим «Факт». Чтение накладной и запись ТСД - read_waybill / write_waybill_tsd ядра.
"""
import os

from PySide6.QtCore import QAbstractTableModel, QModelIndex, QSortFilterProxyModel, Qt, QTimer
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import QAbstractItemView, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMessageBox, QSizePolicy

from .app import Page
from .widgets import Banner, Card, StatusLine, Theme, button, fit_rows, pick_files, table

FILE_HELP = (u"Накладная внутреннего перемещения из Торгсофта (.xls или .xlsx), как её сохраняет Торгсофт: шапка «Поставщик» / "
             u"«Получатель», колонки «Штрих код» и «Кол-во». После чтения программа сверяет строки с «Итого» накладной.")
ROWS_HELP = (u"Галочка «Едет» - товар уехал, остаётся в новой накладной. Снята - товар остался на магазине, в ТСД не попадёт.\n\n"
             u"«Факт» - сколько уехало на самом деле: двойной щелчок или F2, затем число (запятая или точка). 0 снимает галочку.\n\n"
             u"Пробел переключает галочку у выделенных строк. Поиск - по ШК или части названия; сканер ШК тоже работает: "
             u"Enter в поиске переключает галочку найденной строки.")
TSD_HELP = (u"Файл для ТСД: ШК;количество в каждой строке, как остальные файлы программы (cp1251). Лежит в "
            u"ВЫВОЗ\\<сегодня>\\<отправитель>\\Сверка_накладной_<время>\\ вместе с протоколом Excel: едет, убрано, изменено.")


def _num(s):
    try:
        return round(float(str(s).replace(u"\xa0", u"").replace(u" ", u"").replace(u",", u".")), 3)
    except ValueError:
        return None


def _q(v):
    return (u"%.3f" % v).rstrip(u"0").rstrip(u".").replace(u".", u",") if v is not None else u""


def _money(v):
    """Как в накладной Торгсофта: «19 563,08» - и в таблице, и в итоге под ней."""
    return u"{:,.2f}".format(float(v or 0)).replace(u",", u"\xa0").replace(u".", u",")


class WaybillModel(QAbstractTableModel):
    """Строки накладной: галочка «Едет» (по умолчанию стоит) и «Факт» (по умолчанию как в накладной)."""
    COL_GO, COL_BC, COL_NAME, COL_QTY, COL_FACT, COL_SUM = range(6)
    HEAD = (u"Едет", u"Штрих-код", u"Название товара", u"В накладной", u"Факт", u"Сумма факт, грн")

    def __init__(self, parent=None):
        super().__init__(parent)
        self.rows, self.go, self.fact = [], [], []

    def set_rows(self, rows):
        self.beginResetModel()
        self.rows = list(rows)
        self.go = [True] * len(self.rows)
        self.fact = [r["qty"] for r in self.rows]
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.HEAD)

    def headerData(self, section, orient, role=Qt.DisplayRole):
        if orient == Qt.Horizontal and role == Qt.DisplayRole:
            return self.HEAD[section]
        if orient == Qt.Horizontal and role == Qt.TextAlignmentRole:
            return int((Qt.AlignRight if section in (self.COL_QTY, self.COL_FACT, self.COL_SUM) else Qt.AlignLeft) | Qt.AlignVCenter)
        return None

    def flags(self, ix):
        f = Qt.ItemIsEnabled | Qt.ItemIsSelectable
        if ix.column() == self.COL_GO:
            f |= Qt.ItemIsUserCheckable
        if ix.column() == self.COL_FACT:
            f |= Qt.ItemIsEditable
        return f

    def changed(self, i):
        return self.go[i] and abs(self.fact[i] - self.rows[i]["qty"]) > 0.0005

    def data(self, ix, role=Qt.DisplayRole):
        if not ix.isValid():
            return None
        i, c, r = ix.row(), ix.column(), self.rows[ix.row()]
        if role == Qt.CheckStateRole and c == self.COL_GO:
            return Qt.Checked if self.go[i] else Qt.Unchecked
        if role == Qt.AccessibleTextRole and c == self.COL_GO:
            return u"едет" if self.go[i] else u"остаётся на магазине"
        if role in (Qt.DisplayRole, Qt.EditRole):
            if c == self.COL_GO:
                return None                                     # состояние - сама галочка, у снятой строки ещё «—» в «Факт»
            if c == self.COL_BC:
                return r["bc"]
            if c == self.COL_NAME:
                return r["name"]
            if c == self.COL_QTY:
                return _q(r["qty"])
            if c == self.COL_FACT:
                return _q(self.fact[i]) if self.go[i] else u"—"
            if c == self.COL_SUM:
                return _money(self.fact[i] * (r["price"] or 0.0)) if self.go[i] else u"—"
        if role == Qt.TextAlignmentRole:
            return int((Qt.AlignRight if c in (self.COL_QTY, self.COL_FACT, self.COL_SUM) else Qt.AlignLeft) | Qt.AlignVCenter)
        if role == Qt.ForegroundRole:
            if r.get("bad") and c == self.COL_BC:
                return QColor(Theme.c["error"])
            if not self.go[i]:
                return QColor(Theme.c["text_muted"])
            if c == self.COL_FACT and self.changed(i):
                return QColor(Theme.c["warning"])
        if role == Qt.FontRole and c == self.COL_FACT and self.changed(i):
            f = QFont()
            f.setBold(True)                                     # изменённый факт - жирным, не только цветом
            return f
        if role == Qt.ToolTipRole:
            if r.get("bad") and c == self.COL_BC:
                return u"ШК не из 6-14 цифр: в ТСД не загрузить. Снимите галочку и внесите вручную."
            if c == self.COL_FACT and self.changed(i):
                return u"Изменено: в накладной %s, факт %s" % (_q(r["qty"]), _q(self.fact[i]))
            if c == self.COL_NAME:
                return r["name"]
        return None

    def setData(self, ix, value, role=Qt.EditRole):
        if not ix.isValid():
            return False
        i, c = ix.row(), ix.column()
        if role == Qt.CheckStateRole and c == self.COL_GO:
            self._set_go(i, Qt.CheckState(value) == Qt.Checked if not isinstance(value, bool) else value)
            return True
        if role == Qt.EditRole and c == self.COL_FACT:
            q = _num(value)
            if q is None or q < 0:
                return False
            self.fact[i] = q
            self.go[i] = q > 0
            self._emit(i)
            return True
        return False

    def _set_go(self, i, on):
        self.go[i] = on
        if on and self.fact[i] <= 0:                            # вернули галочку у строки с нулём: факт как в накладной
            self.fact[i] = self.rows[i]["qty"]
        self._emit(i)

    def _emit(self, i):
        self.dataChanged.emit(self.index(i, 0), self.index(i, self.columnCount() - 1))

    def toggle(self, idxs):
        """Пробел по выделенным строкам: если хоть одна снята - отметить все, иначе снять все."""
        idxs = sorted(set(idxs))
        if not idxs:
            return
        on = not all(self.go[i] for i in idxs)
        for i in idxs:
            self._set_go(i, on)

    def set_all(self, on):
        for i in range(len(self.rows)):
            self._set_go(i, on)

    def decisions(self):
        return [(g, f) for g, f in zip(self.go, self.fact)]

    def counts(self):
        keep = [i for i in range(len(self.rows)) if self.go[i] and self.fact[i] > 0]
        return {"rows": len(self.rows), "keep": len(keep), "removed": len(self.rows) - len(keep),
                "changed": sum(1 for i in keep if self.changed(i)),
                "keep_sum": round(sum(self.fact[i] * (self.rows[i]["price"] or 0.0) for i in keep), 2)}


class _Filter(QSortFilterProxyModel):
    """Поиск по ШК (начало или конец) и по части названия."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.text = u""

    def set_text(self, t):
        self.text = t.strip().lower()
        self.invalidateFilter()

    def filterAcceptsRow(self, row, parent):
        if not self.text:
            return True
        r = self.sourceModel().rows[row]
        return self.text in r["bc"] or self.text in r["name"].lower()


class NaklPage(Page):
    def __init__(self, core, win):
        super().__init__("nakl", core, win)
        self.d = win.dirs
        self.wbi, self.out = None, None

        c = Card(u"Накладная", FILE_HELP)
        row = QHBoxLayout()
        self.load_btn = button(u"Загрузить накладную…", "folder_open", on_click=self.pick)
        row.addWidget(self.load_btn)
        self.file_line = StatusLine(u"Накладная не загружена", "INFO")
        row.addWidget(self.file_line, 1)
        c.lay.addLayout(row)
        self.body.addWidget(c)

        c = Card(u"Строки накладной", ROWS_HELP)
        row = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText(u"Поиск: ШК или часть названия")
        self.search.setAccessibleName(u"Поиск по накладной")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.on_search)
        self.search.returnPressed.connect(self.on_scan)
        row.addWidget(self.search, 1)
        self.b_all = button(u"Отметить все", "select_all", on_click=lambda: self.model.set_all(True))
        self.b_none = button(u"Снять все", "remove_circle", on_click=lambda: self.model.set_all(False))
        row.addWidget(self.b_all)
        row.addWidget(self.b_none)
        c.lay.addLayout(row)
        self.model = WaybillModel(self)
        self.proxy = _Filter(self)
        self.proxy.setSourceModel(self.model)
        self.view = table(self.proxy, stretch_col=WaybillModel.COL_NAME, multi=True)
        self.view.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed |
                                  QAbstractItemView.AnyKeyPressed)
        self.view.setTextElideMode(Qt.ElideRight)
        self.view.horizontalHeader().setSectionResizeMode(WaybillModel.COL_NAME, QHeaderView.Stretch)
        self.view.setAccessibleName(u"Строки накладной")
        self.view.installEventFilter(self)
        c.lay.addWidget(self.view)
        self.body.addWidget(c)
        self.model.dataChanged.connect(lambda *_: self.update_action())
        self.model.modelReset.connect(self.update_action)

        self.res = Card(u"Результат", TSD_HELP)
        self.banner = Banner()
        self.banner.action.connect(lambda: win.show_page("log"))
        self.res.lay.addWidget(self.banner)
        row = QHBoxLayout()
        row.addStretch(1)
        row.addWidget(button(u"Протокол (Excel)", "table_view", on_click=lambda: self.open_out("xlsx")))
        row.addWidget(button(u"Папка результата", "folder_open", on_click=lambda: self.open_out("out_dir")))
        self.res.lay.addLayout(row)
        self.res.hide()
        self.body.addWidget(self.res)
        self.finish()

        bl = self.action_bar()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        bl.addWidget(self.summary, 1)
        self.go_btn = button(u"Сформировать ТСД", "play_arrow", kind="primary", on_click=self.go)
        bl.addWidget(self.go_btn)
        win.busy_widgets += [self.load_btn, self.go_btn]
        self._fit()
        self.update_action()

    # ---------- таблица ----------
    def _fit(self):
        fit_rows(self.view, 4, 14)

    def eventFilter(self, obj, ev):
        from PySide6.QtCore import QEvent
        if obj is self.view and ev.type() == QEvent.KeyPress and ev.key() == Qt.Key_Space and self.view.state() != QAbstractItemView.EditingState:
            self.model.toggle(self.proxy.mapToSource(ix).row() for ix in self.view.selectionModel().selectedRows())
            return True
        return super().eventFilter(obj, ev)

    def on_search(self, t):
        self.proxy.set_text(t)

    def on_scan(self):
        """Enter в поиске (сканер ШК): одна найденная строка - переключить её галочку и показать."""
        if self.proxy.rowCount() != 1:
            self.win.status(u"Найдено строк: %d - уточните поиск" % self.proxy.rowCount(), "WARN" if self.proxy.rowCount() == 0 else "INFO")
            return
        src = self.proxy.mapToSource(self.proxy.index(0, 0)).row()
        self.model.toggle([src])
        r = self.model.rows[src]
        self.win.status(u"%s: %s" % (r["name"], u"едет" if self.model.go[src] else u"остаётся на магазине"), "OK")
        self.search.selectAll()

    def update_action(self):
        n = self.model.counts()
        if not n["rows"]:
            why = u"Загрузите накладную"
            self.summary.setText(why)
        else:
            why = u"Ни одна строка не отмечена «Едет»" if not n["keep"] else u""
            self.summary.setText(u"Едет %d из %d · %s грн · убрано %d · изменено количество %d"
                                 % (n["keep"], n["rows"], _money(n["keep_sum"]), n["removed"], n["changed"]))
        self.go_btn.setEnabled(not why and self.win.worker is None)
        self.go_btn.setToolTip(why or u"Записать файл для ТСД и протокол")
        for b in (self.b_all, self.b_none):
            b.setEnabled(bool(n["rows"]))

    # ---------- действия ----------
    def pick(self):
        if self.model.rows and self.model.counts()["removed"] + self.model.counts()["changed"] and \
                QMessageBox.question(self, u"Сверка накладной", u"Отметки по текущей накладной пропадут. Загрузить другую?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        fps = pick_files(self, u"Накладная перемещения", self.d.inp, pattern=u"Excel (*.xls *.xlsx);;Все файлы (*.*)")
        if not fps:
            return
        fp = fps[0]

        def done(wbi, out):
            self.wbi = wbi or {}
            self.res.hide()
            self.search.clear()
            self.model.set_rows(self.wbi.get("rows", []))
            self._fit()
            if self.wbi.get("ok"):
                bad = self.wbi.get("bad") or []
                self.file_line.set(u"%s → %s · %d строк · %s ед. · %s грн%s"
                                   % (self.wbi.get("sender") or u"?", self.wbi.get("receiver") or u"?", len(self.wbi["rows"]),
                                      _q(self.wbi["tot_qty"]), _money(self.wbi["tot_sum"]),
                                      (u" · ШК не загрузить: %d (красным)" % len(bad)) if bad else u""), "WARN" if bad else "OK")
                self.win.journal(u"Сверка накладной: загружена", u"%s\n%s -> %s, строк %d" % (fp, self.wbi.get("sender"),
                                                                                            self.wbi.get("receiver"), len(self.wbi["rows"])))
            else:
                self.file_line.set(u"Не прочитана: %s" % u"; ".join(self.wbi.get("problems", [])), "ERR")
                self.win.journal(u"Сверка накладной: ОШИБКА чтения", u"%s\n%s" % (fp, u"\n".join(self.wbi.get("problems", []))))
                self.model.set_rows([])                          # «Итого» не сошлось = часть строк не прочитана: ТСД не даём
            self.update_action()
        self.win.run_task(u"Читаю накладную", lambda: self.core.read_waybill(fp), done)

    def go(self):
        if not self.wbi or not self.model.rows:
            return
        n = self.model.counts()
        if QMessageBox.question(self, u"Сформировать ТСД",
                                u"В новый ТСД: %d строк из %d на %s грн.\nУбрано (остаётся на магазине): %d. Изменено количество: %d.\n\n"
                                u"Записать файл?" % (n["keep"], n["rows"], _money(n["keep_sum"]), n["removed"], n["changed"]),
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) != QMessageBox.Yes:
            return
        wbi, dec = self.wbi, self.model.decisions()

        def done(out, txt):
            self.out = out or {}
            self.win.journal(u"Сверка накладной: ТСД", u"\n".join([u"Файл: %s" % self.out.get("txt", u""),
                                                                   u"Протокол: %s" % self.out.get("xlsx", u"")]
                                                                  + [u"ОШИБКА: " + p for p in self.out.get("problems", [])]))
            if self.out.get("ok"):
                st = self.out["stats"]
                self.banner.show_msg(u"ТСД готов: %s (%d строк). Удалите старую накладную в Торгсофте, загрузите этот файл и "
                                     u"проведите новую." % (os.path.basename(self.out["txt"]), st["lines"]), "OK", u"Журнал")
                self.win.status(u"Сверка накладной: ТСД записан", "OK")
            else:
                self.banner.show_msg(u"; ".join(self.out.get("problems", [])) or u"Не получилось", "ERR", u"Открыть журнал")
                self.win.status(u"Сверка накладной: есть ошибки", "ERR")
            self.res.show()
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(u"Записываю ТСД", lambda: self.core.write_waybill_tsd(self.d, wbi, dec), done)

    def open_out(self, key):
        p = (self.out or {}).get(key)
        if p and os.path.exists(p):
            self.core.open_path(p)
        else:
            QMessageBox.information(self, u"Нет результата", u"Сначала сформируйте ТСД.")
