"""Экран «Инвентаризация ЮА» (финал): магазин + финальная выгрузка Family -> файл инвентаризации ЮА и сверка.

Поведение как в прежнем окне (run_app2, раздел «Инвентаризация ЮА»); расчёт - run_final_inventory ядра без изменений.
«Сверить с базой ЮА» - run_inv_check ядра: дошёл ли txt инвентаризации до магазина в базе ЮА, доинвентаризация.
"""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QComboBox, QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QSizePolicy

from . import tokens as T
from .app import Page
from .widgets import Card, ResultCard, button, help_button, pick_files

FILE_HELP = (u"Файл можно не выбирать: тогда берётся свежая выгрузка магазина из ВХОД_ВЫВОЗ\\МАГАЗИНЫ. "
             u"Выгрузку снимайте ПОСЛЕ проведения вывоза.")
CHECK_HELP = (u"ЮА провела txt инвентаризации - проверьте, что дошло. Снимите выгрузку базы ЮА по всем точкам в "
              u"ВХОД_ВЫВОЗ\\БАЗА_ЮА (или выгрузку магазина в МАГАЗИНЫ_ЮА) и свежую выгрузку Family (поле «Файл» выше).\n\n"
              u"По каждой строке txt: дошло / количество другое / не дошло (карточка на ЮА есть или нет) / у Фемели уже 0; "
              u"плюс остаток Фемели вне txt и где ещё товар есть на ЮА.\n\n"
              u"Результат: Excel и два txt - доинвентаризация (карточка на ЮА есть) и «нет карточки» (сначала завести "
              u"карточки). Количество в них - нынешний остаток Фемели. «Количество другое» в txt не идёт - решать вручную.")
TXT_HELP = u"Не выбран - берётся последний txt инвентаризации магазина из папки ВЫВОЗ. Можно выбрать несколько: более поздний важнее."
UA_HELP = u"Не выбрана - свежая выгрузка из ВХОД_ВЫВОЗ\\БАЗА_ЮА или МАГАЗИНЫ_ЮА, где есть этот магазин."


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

        c = Card(u"Сверка с базой ЮА", CHECK_HELP)
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_SM)
        g.setVerticalSpacing(T.SPACING_SM + 4)
        g.setColumnStretch(1, 1)
        self.txt_ed, self.ua_ed = QLineEdit(), QLineEdit()
        for i, (lb, ed, ph, an, fn, hp) in enumerate(((u"txt:", self.txt_ed, u"последний txt инвентаризации магазина",
                                                       u"txt инвентаризации", self.pick_txt, TXT_HELP),
                                                      (u"База ЮА:", self.ua_ed, u"свежая выгрузка из БАЗА_ЮА / МАГАЗИНЫ_ЮА",
                                                       u"Выгрузка ЮА", self.pick_ua, UA_HELP))):
            g.addWidget(QLabel(lb), i, 0)
            ed.setAccessibleName(an)
            ed.setPlaceholderText(ph)
            ed.textChanged.connect(lambda *_: self.update_action())
            g.addWidget(ed, i, 1)
            g.addWidget(button(u"Выбрать…", "folder_open", on_click=fn), i, 2)
            g.addWidget(help_button(hp, an), i, 3)
        c.lay.addLayout(g)
        row = QHBoxLayout()
        row.addStretch(1)
        self.chk_btn = button(u"Сверить с базой ЮА", "fact_check", on_click=self.check)
        row.addWidget(self.chk_btn)
        c.lay.addLayout(row)
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
        win.busy_widgets += [self.go_btn, self.chk_btn]
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
        bad = [p for p in self.txt_paths() + [self.ua_ed.text().strip()] if p and not os.path.isfile(p)]
        why2 = (u"Выберите магазин" if not shop else u"Файл не найден: %s" % os.path.basename(bad[0]) if bad
                else u"Файл не найден: проверьте путь" if fp and not os.path.isfile(fp) else u"")
        self.chk_btn.setEnabled(not why2 and self.win.worker is None)
        self.chk_btn.setToolTip(why2 or u"Дошёл ли txt инвентаризации до магазина в базе ЮА; доинвентаризация")

    def txt_paths(self):
        return [p.strip() for p in self.txt_ed.text().split(";") if p.strip()]

    def pick_txt(self):
        fps = pick_files(self, u"txt инвентаризации ЮА", self.d.out, multi=True, pattern=u"txt (*.txt);;Все файлы (*.*)")
        if fps:
            self.txt_ed.setText(u";".join(fps))

    def pick_ua(self):
        fps = pick_files(self, u"Выгрузка базы ЮА", self.d.ua_base if os.path.isdir(self.d.ua_base) else self.d.base)
        if fps:
            self.ua_ed.setText(fps[0])

    def check(self):
        core = self.core
        shop = self.shop_cb.currentText().strip()
        if not shop:
            QMessageBox.information(self, u"Сверка с базой ЮА", u"Выберите магазин.")
            return
        txts, ua, fam = self.txt_paths(), self.ua_ed.text().strip() or None, self.file_ed.text().strip() or None

        def done(info, out):
            self.info = info or {}
            st = self.info.get("stats", {})
            rows = [{"part": a, "pos": n, "qty": core.fmt_qty(q), "sum": core.fmt_money(sm, 2)}
                    for a, n, q, sm in self.info.get("summary", [])]
            self.win.journal(u"Сверка инвентаризации с базой ЮА", core._final_text(self.info))
            if self.info.get("ok"):
                self.res.show_result(rows, u"Сверка «%s»: не дошло строк txt - %d; в доинвентаризацию %d, без карточки на ЮА %d."
                                     % (shop, st.get(core.CHK_MISS, 0) + st.get(core.CHK_NOCARD, 0), st.get("add_lines", 0),
                                        st.get("nocard_lines", 0)), "OK", max_rows=10)
            else:
                self.res.show_result(rows, u"Есть ошибки - подробности в «Журнале».", "ERR")
                self.win.status(u"Сверка с базой ЮА: есть ошибки", "ERR")
            from PySide6.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(u"Сверка с базой ЮА", lambda: core.run_inv_check(self.d, shop, txt_paths=txts, ua_path=ua, fam_path=fam),
                          done)

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
