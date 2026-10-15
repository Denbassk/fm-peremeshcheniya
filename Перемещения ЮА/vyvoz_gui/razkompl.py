"""Окно «Раскомплектация: на точки» (решение 15.10.2026): поставщики, слова названия и ШК, чей товар «Распределение»
всегда везёт по точкам, а не на склад, при любом количестве. «Сохранить» пишет Справочник\\vyvoz_razkompl.csv."""
import os
from datetime import datetime

from PySide6.QtWidgets import (QAbstractItemView, QDialog, QHBoxLayout, QHeaderView, QLabel, QMessageBox, QTableWidget,
                               QTableWidgetItem, QVBoxLayout)

from . import tokens as T
from .widgets import button, help_button, muted

HELP = (u"Товар из этого списка «Распределение» везёт ПО ТОЧКАМ, а не на склад (Полевая-Склад), при любом количестве: "
        u"раскомплектованную упаковку склад обратно не собирает. Правило важнее справочника маршрутов и истории.\n\n"
        u"Что писать в строке:\n"
        u"• поставщика - как в приходах или в матрице (Джамп, Агропром ТОВ, Мак Кофе);\n"
        u"• слово из названия товара (Монжар, Петрівська Слобода): подходит, если оно есть в названии; пробелы и регистр "
        u"не важны, «МакКофе» = «Мак Кофе»;\n"
        u"• отдельный товар - его штрих-код (только цифры).\n\n"
        u"Список хранится в файле и действует на все следующие распределения. В «Распределение.xlsx» на листе «Маршрут» "
        u"у таких позиций основание «раскомплектация».")


class RazkomplDialog(QDialog):
    def __init__(self, core, dirs, parent=None):
        super().__init__(parent)
        self.core, self.dirs, self.saved = core, dirs, None
        self.setWindowTitle(u"Раскомплектация: на точки")
        self.setModal(True)
        self.resize(720, 520)
        v = QVBoxLayout(self)
        v.setContentsMargins(T.SPACING_LG, T.SPACING_LG, T.SPACING_LG, T.SPACING_MD)
        v.setSpacing(T.SPACING_SM + 4)
        hd = QHBoxLayout()
        t = QLabel(u"Раскомплектация: на точки, не на склад")
        t.setObjectName("cardTitle")
        hd.addWidget(t)
        hd.addWidget(help_button(HELP, u"Раскомплектация"))
        hd.addStretch(1)
        v.addLayout(hd)
        path = os.path.join(dirs.cache, core.RAZKOMPL_FILE)
        rows = core.load_razkompl(dirs)                      # файла нет - создаётся со списком по умолчанию
        when = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d.%m.%Y %H:%M") if os.path.isfile(path) else u""
        f = muted(u"Сейчас в файле: %d зап.%s" % (len(rows), (u", сохранён " + when) if when else u""))
        f.setToolTip(path)
        v.addWidget(f)

        self.tw = QTableWidget(0, 2)
        self.tw.setHorizontalHeaderLabels([u"Поставщик, слово из названия или ШК", u"Комментарий"])
        self.tw.setAccessibleName(u"Список раскомплектации")
        self.tw.setAlternatingRowColors(True)
        self.tw.setShowGrid(False)
        self.tw.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tw.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed | QAbstractItemView.AnyKeyPressed)
        self.tw.verticalHeader().setVisible(False)
        self.tw.verticalHeader().setDefaultSectionSize(T.ROW_H)
        hh = self.tw.horizontalHeader()
        hh.setHighlightSections(False)
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.Stretch)
        for e, c in rows:
            self._add(e, c)
        v.addWidget(self.tw, 1)

        r = QHBoxLayout()
        r.addWidget(button(u"Добавить строку", None, on_click=self.add_row))
        r.addWidget(button(u"Удалить выбранную", "delete", on_click=self.del_row))
        r.addStretch(1)
        v.addLayout(r)
        r = QHBoxLayout()
        r.addStretch(1)
        r.addWidget(button(u"Отмена", on_click=self.reject))
        ok = button(u"Сохранить", "check", kind="primary", on_click=self.save)
        ok.setDefault(True)
        r.addWidget(ok)
        v.addLayout(r)

    def _add(self, e=u"", c=u""):
        i = self.tw.rowCount()
        self.tw.insertRow(i)
        for j, t_ in enumerate((e, c)):
            it = QTableWidgetItem(t_)
            it.setToolTip(t_)                                # длинный комментарий виден целиком при наведении
            self.tw.setItem(i, j, it)
        return i

    def add_row(self):
        i = self._add()
        self.tw.setCurrentCell(i, 0)
        self.tw.editItem(self.tw.item(i, 0))

    def del_row(self):
        rs = sorted(set(ix.row() for ix in self.tw.selectedIndexes()), reverse=True)
        for i in rs:
            self.tw.removeRow(i)

    def rows(self):
        out = []
        for i in range(self.tw.rowCount()):
            e, c = self.tw.item(i, 0), self.tw.item(i, 1)
            e = (e.text() if e else u"").strip()
            if e:
                out.append((e, (c.text() if c else u"").strip()))
        return out

    def save(self):
        rs = self.rows()
        if not rs and QMessageBox.question(self, u"Раскомплектация", u"Список пустой: всё будет решаться по истории поставок. Сохранить?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        try:
            self.saved = self.core.save_razkompl(self.dirs, rs)
        except PermissionError:
            QMessageBox.warning(self, u"Раскомплектация", u"Файл %s открыт в Excel: закройте его и сохраните ещё раз." % self.core.RAZKOMPL_FILE)
            return
        self.accept()
