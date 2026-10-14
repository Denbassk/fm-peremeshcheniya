"""Экран «Распределение»: шаг 1 - лист из Корректировка_ЮА (или автоматический «без правок»), шаг 2 - один выбор
«Как делить» (по N точкам / подобрать по сумме / всё на один магазин); «Распределить» закреплена внизу.

Поведение как в прежнем окне (run_app2, раздел «Распределение»); расчёт - run_split ядра без изменений.
"""
import os
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QButtonGroup, QCheckBox, QComboBox, QGridLayout, QHBoxLayout, QLabel, QMessageBox,
                               QRadioButton, QSizePolicy, QSpinBox)

from . import tokens as T
from .app import Page
from .widgets import Card, ResultCard, StatusLine, button, help_button, muted

LIST_HELP = (u"Распределение берёт листы ТОЛЬКО из папки «%s»: исправили список вывоза в Excel - сохраните его туда. "
             u"Новые листы - сверху. Файлы «проверка» и временные ~$ не показываются.")
AUTO_HELP = (u"«Взять без правок» кладёт в «%s» копию листа «Вывезти на склад» (7 колонок, как в исправленных листах); "
             u"сам автоматический список не меняется.")
HOW_HELP = (u"Товар РЦ всегда едет на %s (если галочка ниже включена), на точки он не делится. Прямые поставки делятся "
            u"по торговым точкам с самым большим спросом на товары этого списка, не меньше 5 тыс. грн на точку.")
ONE_HELP = (u"Всё, что обычно делится по точкам (прямые поставки и кеги), едет на выбранный магазин; товар РЦ - как обычно на "
            u"%s. В Распределение.xlsx - оценка продаж за %d дн. на этом магазине и сравнение магазинов. В списке только "
            u"магазины, которые распределение примет: без Полевой, без магазинов ЮА и без магазина-источника.")
HOW = (("top", u"По точкам"), ("auto", u"Подобрать число точек по сумме"), ("one", u"Всё на один магазин"))


class SplitPage(Page):
    def __init__(self, core, win):
        super().__init__("split", core, win)
        self.d = win.dirs
        self.corr, self.auto, self.info, self.chk = [], [], None, None
        corr = core.CORR_DIR

        # ---------- шаг 1: лист ----------
        c = Card(u"Шаг 1 · Какой список распределяем", LIST_HELP % corr)
        c.add_head(button(u"Перечитать", "refresh", tip=u"Найти листы заново", on_click=lambda: self.refresh_lists()))
        c.add_head(button(u"Открыть папку", "folder_open", tip=u"Папка «%s»" % corr, on_click=self.open_corr))
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_SM)
        g.setVerticalSpacing(T.SPACING_SM + 4)
        g.setColumnStretch(1, 1)
        g.addWidget(QLabel(u"Лист:"), 0, 0)
        self.sp_cb = QComboBox()
        self.sp_cb.setAccessibleName(u"Лист из папки %s" % corr)
        self.sp_cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.sp_cb.currentIndexChanged.connect(lambda *_: self.on_list())
        g.addWidget(self.sp_cb, 0, 1)
        self.sp_info = StatusLine()
        g.addWidget(self.sp_info, 1, 1)
        g.addWidget(QLabel(u"Без правок:"), 2, 0)
        self.auto_cb = QComboBox()
        self.auto_cb.setAccessibleName(u"Автоматический список вывоза")
        self.auto_cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        g.addWidget(self.auto_cb, 2, 1)
        self.b_take = button(u"Взять без правок", "content_copy", on_click=self.take_as_is)
        g.addWidget(self.b_take, 2, 2)
        g.addWidget(help_button(AUTO_HELP % corr, u"Взять без правок"), 2, 3)
        c.lay.addLayout(g)
        self.body.addWidget(c)

        # ---------- шаг 2: как делить (один выбор) ----------
        c = Card(u"Шаг 2 · Как делить", HOW_HELP % core.SPLIT_RC_NAME)
        self.how = QButtonGroup(self)
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_MD)
        g.setVerticalSpacing(T.SPACING_SM + 4)
        g.setColumnStretch(1, 1)
        self.top_sp = QSpinBox()
        self.top_sp.setRange(1, core.SPLIT_TOP_N)
        self.top_sp.setValue(core.SPLIT_TOP_N)
        self.top_sp.setFixedWidth(5 * T.SPACING_MD)
        self.top_sp.setAccessibleName(u"Число точек")
        self.top_sp.valueChanged.connect(lambda *_: self.update_action())
        self.one_cb = QComboBox()
        self.one_cb.setAccessibleName(u"Магазин, на который едет весь вывоз")
        self.one_cb.setMinimumWidth(16 * T.SPACING_MD)
        self.one_cb.currentIndexChanged.connect(lambda *_: self.update_action())
        self.one_line = StatusLine(u"", "INFO", help=ONE_HELP % (core.SPLIT_RC_NAME, core.EFF_DAYS))
        rows = {"top": ([self.top_sp], u"из %d самых сильных по спросу" % core.SPLIT_TOP_N,
                        u"Прямые поставки делятся на столько точек с самым большим спросом на товары этого списка."),
                "auto": ([], u"5-7 тыс. грн прямых поставок на точку",
                         u"Число точек подбирается по сумме прямых поставок: 5-7 тыс. грн на точку, не больше %d точек."
                         % core.SPLIT_TOP_N),
                "one": ([self.one_cb, self.one_line], None, None)}
        for i, (val, text) in enumerate(HOW):
            rb = QRadioButton(text)
            rb.setProperty("how", val)
            rb.setChecked(val == "top")                        # как в прежнем окне: по умолчанию 10 точек
            self.how.addButton(rb)
            g.addWidget(rb, i, 0)
            h = QHBoxLayout()
            h.setSpacing(T.SPACING_SM)
            ws, hint, tip = rows[val]
            for w in ws:
                h.addWidget(w)
            if hint:                                           # пояснение - одной строкой, подробно - под «?» сразу после неё
                h.addWidget(muted(hint, wrap=False))
                h.addWidget(help_button(tip, text))
            h.addStretch(1)
            g.addLayout(h, i, 1)
        c.lay.addLayout(g)
        self.rc_cb = QCheckBox(u"Сделать файл на %s: туда идёт товар РЦ, на точки он не делится" % core.SPLIT_RC_NAME)
        self.rc_cb.setChecked(True)
        c.lay.addWidget(self.rc_cb)
        self.how.buttonClicked.connect(lambda *_: self.on_how())
        self.body.addWidget(c)

        # ---------- результат ----------
        self.res = ResultCard(((("addr", u"Куда (адрес)", Qt.AlignLeft), ("pos", u"Позиций", Qt.AlignRight),
                                ("qty", u"Единиц", Qt.AlignRight), ("sum", u"Сумма, грн", Qt.AlignRight))),
                              stretch_col=0, on_log=lambda: win.show_page("log"))
        self.res.add_button(button(u"Распределение (Excel)", "table_view", on_click=lambda: self.open_info("xlsx")))
        self.res.add_button(button(u"Папка файлов ТСД", "folder_open", on_click=lambda: self.open_info("out_dir")))
        self.body.addWidget(self.res)
        self.finish()

        # ---------- панель действия ----------
        bl = self.action_bar()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        bl.addWidget(self.summary, 1)
        self.go_btn = button(u"Распределить", "call_split", kind="primary", on_click=self.go)
        bl.addWidget(self.go_btn)
        win.busy_widgets += [self.go_btn]
        self.on_how()

    # ---------------- шаг 1 ----------------
    def shown(self):
        self.refresh_lists()

    def refresh_lists(self, select=None):
        core = self.core
        cur = select or self.list_path()
        self.corr = core.list_corrected(self.d)
        self.sp_cb.blockSignals(True)
        self.sp_cb.clear()
        for p in self.corr:
            self.sp_cb.addItem(u"%s    (%s)" % (os.path.basename(p), datetime.fromtimestamp(core._mtime(p)).strftime("%d.%m %H:%M")))
            self.sp_cb.setItemData(self.sp_cb.count() - 1, p, Qt.ToolTipRole)
        want = os.path.normcase(cur or u"")
        ix = next((i for i, p in enumerate(self.corr) if os.path.normcase(p) == want), 0 if self.corr else -1)
        self.sp_cb.setCurrentIndex(ix)
        self.sp_cb.blockSignals(False)
        items = core.auto_lists(self.d)
        self.auto = [fp for _l, fp in items]
        a_cur = self.auto_cb.currentIndex()
        self.auto_cb.clear()
        self.auto_cb.addItems([l_ for l_, _fp in items])
        self.auto_cb.setCurrentIndex(a_cur if 0 <= a_cur < len(items) else (0 if items else -1))
        self.auto_cb.setEnabled(bool(items))
        self.b_take.setEnabled(bool(items))
        self.b_take.setToolTip(u"" if items else u"Автоматических списков пока нет: сформируйте список в «Списки вывоза»")
        self.on_list()

    def list_path(self):
        i = self.sp_cb.currentIndex()
        return self.corr[i] if 0 <= i < len(self.corr) else u""

    def on_list(self):
        fp = self.list_path()
        if not fp:
            self.chk = None
            self.sp_info.set(u"В «%s» нет листов: сохраните туда исправленный список или возьмите автоматический ниже."
                             % self.core.CORR_DIR, "WARN")
            self.set_targets(None)
        else:
            self.chk = self.core.corrected_check(fp)
            self.sp_info.set(self.chk["text"], "OK" if self.chk["ok"] else "ERR")
            self.set_targets(self.chk["shop"] if self.chk["ok"] else None)
        self.update_action()

    def set_targets(self, source):
        """«Всё на один магазин» - только магазины, которые run_split примет (split_targets: без источника, Полевой и ЮА)."""
        cur = self.one_cb.currentText()
        vals = self.core.split_targets(self.d, source)
        self.one_cb.blockSignals(True)
        self.one_cb.clear()
        self.one_cb.addItems(vals)
        self.one_cb.setCurrentIndex(vals.index(cur) if cur in vals else -1)
        self.one_cb.setPlaceholderText(u"выберите магазин" if vals else u"нет подходящих магазинов")
        self.one_cb.blockSignals(False)

    def open_corr(self):
        d_ = self.core.corr_dir(self.d)
        os.makedirs(d_, exist_ok=True)
        self.core.open_path(d_)

    def take_as_is(self):
        core = self.core
        i = self.auto_cb.currentIndex()
        if not (0 <= i < len(self.auto)):
            QMessageBox.information(self, u"Взять без правок", u"Автоматических списков пока нет: сформируйте список в «Списки вывоза».")
            return
        src = self.auto[i]
        if core.has_manual_note(src) and QMessageBox.question(
                self, u"Взять без правок", u"Это список РУЧНОГО ВЫБОРА: в нём сигареты и кеги, которые обычно не вывозят.\nЕсли вы "
                u"ничего не удаляли, они поедут в распределение (кеги - на точки, где они продаются).\n\nВзять список как есть?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        try:
            dst = core.copy_list_as_is(self.d, src)
        except Exception as e:
            QMessageBox.critical(self, u"Взять без правок", u"Не получилось: %s" % e)
            return
        self.win.journal(u"Список вывоза взят без правок", u"%s\n->  %s" % (src, dst))
        self.refresh_lists(select=dst)
        self.win.status(u"Список взят без правок: %s" % os.path.basename(dst), "OK")

    # ---------------- шаг 2 ----------------
    def how_val(self):
        b = self.how.checkedButton()
        return b.property("how") if b else "top"

    def on_how(self):
        h = self.how_val()
        self.top_sp.setEnabled(h == "top")
        self.one_cb.setEnabled(h == "one")
        self.update_action()

    # ---------------- действие ----------------
    def update_action(self):
        h, n_t = self.how_val(), self.one_cb.count()
        if h == "one" and not n_t:
            self.one_line.set(u"Нет магазинов, которые примет распределение", "WARN")
        elif h == "one" and not self.one_cb.currentText():
            self.one_line.set(u"Выберите магазин", "WARN")
        else:
            self.one_line.set(u"Товар РЦ - на %s" % self.core.SPLIT_RC_NAME, "INFO")
        why = u""
        if not (self.chk and self.chk.get("ok")):
            why = u"Сначала выберите годный лист в шаге 1"
        elif h == "one" and not self.one_cb.currentText():
            why = u"Выберите магазин в шаге 2"
        if why:
            self.summary.setText(why)
        else:
            how_t = {"top": u"на %d точек" % self.top_sp.value(), "auto": u"число точек по сумме",
                     "one": u"всё на %s" % self.one_cb.currentText()}[h]
            self.summary.setText(u"%s · %d поз. · %s" % (self.chk["shop"], self.chk["n"], how_t))
        self.go_btn.setEnabled(not why and self.win.worker is None)
        self.go_btn.setToolTip(why or u"Файлы ТСД и Распределение.xlsx - в «ТСД_по_адресам» рядом со списком")

    def go(self):
        core, fp = self.core, self.list_path()
        v_ = core.validate_split_input(self.d, fp)
        if not v_["ok"]:
            (QMessageBox.information if v_["level"] == "info" else QMessageBox.warning)(self, u"Распределение", v_["text"])
            return
        chk, h = v_["chk"], self.how_val()
        n = core._clamp_top(self.top_sp.value())
        auto, skip = h == "auto", not self.rc_cb.isChecked()
        one = self.one_cb.currentText().strip() if h == "one" else u""
        if h == "one" and not one:
            return
        if one and QMessageBox.question(
                self, u"Распределение", u"Всё, что обычно делится по точкам (в т.ч. кеги), - на один магазин: %s?\nТовар РЦ - как "
                u"обычно на %s (лист: %d поз.)." % (one, core.SPLIT_RC_NAME, chk["n"])) != QMessageBox.Yes:
            return
        label = (u"Весь вывоз на %s" % one) if one else (u"Распределение (число точек по сумме)" if auto else u"Распределение на %d точек" % n)

        def done(info, out):
            self.info = info or {}
            rows = [{"addr": a, "pos": kk, "qty": core.fmt_qty(q), "sum": core.fmt_money(sm, 2)}
                    for a, kk, q, sm in self.info.get("summary", [])]
            self.win.journal(label, core._split_text(self.info))
            if self.info.get("ok"):
                tot = sum(float(r[3] or 0) for r in self.info.get("summary", []))
                self.res.show_result(rows, u"Готово: файлов ТСД %d на %s грн. Файлы - в «ТСД_по_адресам» рядом со списком."
                                     % (len(self.info.get("files", [])), core.fmt_money(tot)), "OK")
            else:
                self.res.show_result(rows, u"Есть ошибки - подробности в «Журнале».", "ERR")
                self.win.status(u"Распределение: есть ошибки", "ERR")
            from PySide6.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(label, lambda: core.run_split(self.d, fp, top_n=n, skip_rc=skip, auto_points=auto, target=one or None), done)

    def open_info(self, key):
        v = (self.info or {}).get(key)
        if v and os.path.exists(v):
            self.core.open_path(v)
        else:
            QMessageBox.information(self, u"Нет результата", u"Сначала нажмите «Распределить».")
