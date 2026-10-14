"""Экран «Инвентаризация ЮА» (финал): магазин + финальная выгрузка Family -> файл инвентаризации ЮА и сверка.

Поведение как в прежнем окне (run_app2, раздел «Инвентаризация ЮА»); расчёт - run_final_inventory ядра без изменений.
"""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox, QGridLayout, QLabel, QLineEdit, QMessageBox, QSizePolicy

from . import tokens as T
from .app import Page
from .widgets import Card, ResultCard, button, help_button, pick_files

FILE_HELP = (u"Файл можно не выбирать: тогда берётся свежая выгрузка магазина из ВХОД_ВЫВОЗ\\МАГАЗИНЫ. "
             u"Выгрузку снимайте ПОСЛЕ проведения вывоза.")


class FinPage(Page):
    def __init__(self, core, win):
        super().__init__("fin", core, win)
        self.d = win.dirs
        self.info = None

        c = Card(u"Финальная выгрузка Family")
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_SM)
        g.setVerticalSpacing(T.SPACING_SM + 4)
        g.setColumnStretch(1, 1)
        g.addWidget(QLabel(u"Магазин:"), 0, 0)
        self.shop_cb = QComboBox()
        self.shop_cb.setEditable(True)                       # как в прежнем окне: можно вписать магазин руками
        self.shop_cb.setAccessibleName(u"Магазин")
        self.shop_cb.lineEdit().setPlaceholderText(u"выберите магазин")
        self.shop_cb.currentTextChanged.connect(lambda *_: self.update_action())
        g.addWidget(self.shop_cb, 0, 1)
        g.addWidget(QLabel(u"Файл:"), 1, 0)
        self.file_ed = QLineEdit()
        self.file_ed.setAccessibleName(u"Финальная выгрузка")
        self.file_ed.setPlaceholderText(u"свежая выгрузка магазина из ВХОД_ВЫВОЗ\\МАГАЗИНЫ")
        self.file_ed.textChanged.connect(lambda *_: self.update_action())
        g.addWidget(self.file_ed, 1, 1)
        g.addWidget(button(u"Выбрать выгрузку…", "folder_open", on_click=self.pick), 1, 2)
        g.addWidget(help_button(FILE_HELP, u"Файл"), 1, 3)
        c.lay.addLayout(g)
        self.body.addWidget(c)

        self.res = ResultCard(((("part", u"Часть", Qt.AlignLeft), ("pos", u"Позиций", Qt.AlignRight),
                                ("qty", u"Единиц", Qt.AlignRight), ("sum", u"Сумма, грн", Qt.AlignRight))),
                              stretch_col=0, on_log=lambda: win.show_page("log"))
        self.res.add_button(button(u"Сверка (Excel)", "table_view", on_click=self.open_xlsx))
        self.res.add_button(button(u"Папка результата", "folder_open", on_click=self.open_dir))
        self.body.addWidget(self.res)
        self.finish()

        bl = self.action_bar()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        bl.addWidget(self.summary, 1)
        self.go_btn = button(u"Собрать файл инвентаризации", "inventory", kind="primary", on_click=self.go)
        bl.addWidget(self.go_btn)
        win.busy_widgets += [self.go_btn]
        self.update_action()

    def shown(self):
        """Магазины - те, что нашли «Списки вывоза» в ВХОД_ВЫВОЗ\\МАГАЗИНЫ (как в прежнем окне)."""
        found = sorted(getattr(self.win.pages.get("vyvoz"), "found", {}) or {})
        cur = self.shop_cb.currentText()
        if [self.shop_cb.itemText(i) for i in range(self.shop_cb.count())] != found:
            self.shop_cb.blockSignals(True)
            self.shop_cb.clear()
            self.shop_cb.addItems(found)
            self.shop_cb.setEditText(cur)
            self.shop_cb.blockSignals(False)
        self.update_action()

    def pick(self):
        fps = pick_files(self, u"Финальная выгрузка Family", self.d.stores if os.path.isdir(self.d.stores) else self.d.base)
        if fps:
            self.file_ed.setText(fps[0])

    def update_action(self):
        shop, fp = self.shop_cb.currentText().strip(), self.file_ed.text().strip()
        why = (u"Выберите магазин" if not shop else u"Файл не найден: проверьте путь" if fp and not os.path.isfile(fp) else u"")
        self.summary.setText(why or u"%s · %s" % (shop, os.path.basename(fp) if fp else u"свежая выгрузка из ВХОД_ВЫВОЗ\\МАГАЗИНЫ"))
        self.go_btn.setEnabled(not why and self.win.worker is None)
        self.go_btn.setToolTip(why or u"Файл инвентаризации ЮА и сверка")

    def go(self):
        core = self.core
        shop = self.shop_cb.currentText().strip()
        if not shop:
            QMessageBox.information(self, u"Инвентаризация ЮА", u"Выберите магазин.")
            return
        fp = self.file_ed.text().strip() or None
        if fp and not os.path.isfile(fp):
            QMessageBox.information(self, u"Инвентаризация ЮА", u"Файл не найден.")
            return
        if QMessageBox.question(self, u"Инвентаризация ЮА", u"Выгрузка Family снята ПОСЛЕ проведения вывоза?\n\n"
                                u"Если нет, в инвентаризацию попадёт и то, что ещё должно уехать.",
                                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return

        def done(info, out):
            self.info = info or {}
            rows = [{"part": a, "pos": n, "qty": core.fmt_qty(q), "sum": core.fmt_money(sm, 2)}
                    for a, n, q, sm in self.info.get("summary", [])]
            self.win.journal(u"Инвентаризация ЮА (финал)", core._final_text(self.info))
            if self.info.get("ok"):
                self.res.show_result(rows, u"Готово: файл инвентаризации ЮА для «%s» собран." % shop, "OK")
            else:
                self.res.show_result(rows, u"Есть ошибки - подробности в «Журнале».", "ERR")
                self.win.status(u"Инвентаризация ЮА: есть ошибки", "ERR")
            from PySide6.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(u"Финальная инвентаризация", lambda: core.run_final_inventory(self.d, shop, state_path=fp), done)

    def open_xlsx(self):
        fp = (self.info or {}).get("xlsx")
        self.core.open_path(fp) if fp and os.path.isfile(fp) else self.nores()

    def open_dir(self):
        d_ = (self.info or {}).get("out_dir")
        self.core.open_path(d_) if d_ and os.path.exists(d_) else self.nores()

    def nores(self):
        QMessageBox.information(self, u"Нет результата", u"Сначала соберите файл инвентаризации.")
