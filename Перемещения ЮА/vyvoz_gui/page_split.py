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
HOW_HELP = (u"Товар РЦ едет на %s (если не включена галочка «весь товар на магазины»), на точки он не делится. Прямые "
            u"поставки делятся по торговым точкам с самым большим спросом на товары этого списка, не меньше 5 тыс. грн на точку.")
ONE_HELP = (u"Всё, что обычно делится по точкам (прямые поставки и кеги), едет на выбранный магазин; товар РЦ - как обычно на "
            u"%s, а с галочкой «весь товар на магазины» тоже на этот магазин. В Распределение.xlsx - оценка продаж за %d дн. на "
            u"этом магазине и сравнение магазинов. В списке только магазины, которые распределение примет: без Полевой, без "
            u"магазинов ЮА и без магазина-источника.")
BASH_HELP = (u"Заморозка Бащинского (ШК из ВХОД_ВЫВОЗ\\ДОВЫВОЗ\\заморозка_бащинский.txt) едет одной партией на ОДНУ точку с его "
             u"холодильником, без порога суммы и не на склад. Холодильники есть не во всех магазинах: «Найти по продажам» берёт "
             u"точки, где за 60 дней продано не меньше 5 позиций этой заморозки не меньше чем в 6 неделях (BigQuery), лучшая - первая. "
             u"Можно выбрать точку из списка. Выбор запоминается. Остальной товар распределяется как обычно.")
RCS_HELP = (u"Весь товар едет на магазины, включая тот, что обычно идёт через РЦ: он делится по точкам по спросу, файл на %s не "
            u"создаётся. Нужно, когда склад не принимает (например, магазин закрывается и вывоз идёт сразу на точки).")
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
        self.rcs_cb = QCheckBox(u"Весь товар на магазины, включая товар РЦ (на склад ничего не едет)")
        self.rcs_cb.setToolTip(RCS_HELP % core.SPLIT_RC_NAME)
        self.rcs_cb.setAccessibleName(u"Весь товар на магазины, включая товар РЦ")
        self.rcs_cb.toggled.connect(lambda *_: self.on_rcs())
        c.lay.addWidget(self.rcs_cb)
        self.bash_chk = QCheckBox(u"Заморозка Бащинского - одной партией на точку с холодильником")
        self.bash_chk.setToolTip(BASH_HELP)
        self.bash_chk.setAccessibleName(u"Заморозка Бащинского одной партией на точку")
        self.bash_chk.setChecked(self._sget("split_bash") == "1")
        self.bash_chk.toggled.connect(lambda *_: self.on_bash())
        c.lay.addWidget(self.bash_chk)
        brow = QHBoxLayout()
        brow.setSpacing(T.SPACING_SM)
        brow.setContentsMargins(T.SPACING_LG + T.SPACING_SM, 0, 0, 0)
        self.bash_cb = QComboBox()
        self.bash_cb.setAccessibleName(u"Куда везти заморозку Бащинского")
        self.bash_cb.setMinimumWidth(14 * T.SPACING_MD)
        self.bash_cb.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
        self.bash_cb.currentIndexChanged.connect(lambda *_: self.on_bash())
        brow.addWidget(self.bash_cb)
        self.bash_btn = button(u"Найти по продажам", "refresh", tip=u"Точки с холодильником Бащинского - из BigQuery",
                               on_click=self.load_bash)
        brow.addWidget(self.bash_btn)
        brow.addWidget(help_button(BASH_HELP, u"Заморозка Бащинского"))
        brow.addStretch(1)
        c.lay.addLayout(brow)
        self.bash_line = StatusLine()
        c.lay.addWidget(self.bash_line)
        self.pinfo = None
        self._fill_bash()
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
        win.busy_widgets += [self.go_btn, self.bash_btn]
        self.on_how()
        self.on_bash()

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
        kind = core.manual_note_kind(src)
        if kind == "all":
            ask = (u"Это список «ВЫВОЗ ВСЕГО»: в нём весь остаток магазина, в т.ч. то, что обычно остаётся (кеги, сигареты, расходники, "
                   u"кулинария, матрица и ЮА).\nЕсли вы ничего не удаляли, поедет ВСЁ.\n\nВзять список как есть?")
        else:
            ask = (u"Это список РУЧНОГО ВЫБОРА: в нём сигареты и кеги, которые обычно не вывозят.\nЕсли вы "
                   u"ничего не удаляли, они поедут в распределение (кеги - на точки, где они продаются).\n\nВзять список как есть?")
        if kind and QMessageBox.question(self, u"Взять без правок", ask, QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.No) != QMessageBox.Yes:
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

    # ---------------- заморозка Бащинского ----------------
    def _sget(self, key):
        try:
            s = getattr(getattr(self.win, "themes", None), "settings", None)
            v = s.value(key) if s is not None else None
            return None if v is None else str(v)
        except Exception:
            return None

    def _sset(self, key, val):
        try:
            s = getattr(getattr(self.win, "themes", None), "settings", None)
            if s is not None:
                s.setValue(key, val)
        except Exception:
            pass

    def _fill_bash(self):
        fr = (self.pinfo or {}).get("bash") or []
        saved = self._sget("split_bash_point")
        self.bash_cb.blockSignals(True)
        self.bash_cb.clear()
        if self.pinfo is None:
            self.bash_cb.addItem(u"лучшая по продажам", None)
        elif not fr:
            self.bash_cb.addItem(u"не найдены - заморозка поделится по спросу", None)
        else:
            for i, (s, q) in enumerate(fr):
                t = s + (u"  · лучшая по продажам" if i == 0 else u"")
                if q:
                    t += u"  · %.0f шт за %d дн." % (q, self.core.DV_BASH_DAYS)
                self.bash_cb.addItem(t, s)
            names = [s for s, _q in fr]
            self.bash_cb.setCurrentIndex(names.index(saved) if saved in names else 0)
        self.bash_cb.blockSignals(False)

    def load_bash(self):
        if self.win.worker is not None:
            return
        core, src = self.core, (self.chk or {}).get("shop") or u""
        self.win.run_task(u"Точки с холодильником Бащинского", lambda: core.dovyvoz_points(self.d, [src] if src else []), self._bash_done)

    def _bash_done(self, info, out):
        self.pinfo = info if isinstance(info, dict) else {"bash": [], "problem": u"нет ответа"}
        self._fill_bash()
        self.on_bash()

    def bash_point(self):
        """Выбранная точка (None - пусть расчёт найдёт лучшую сам) или None, если галочка выключена."""
        return self.bash_cb.currentData() if self.bash_chk.isChecked() else None

    def on_bash(self):
        on = self.bash_chk.isChecked()
        self.bash_cb.setEnabled(on)
        self.bash_btn.setEnabled(on and self.win.worker is None)
        v = self.bash_cb.currentData()
        if v:
            self._sset("split_bash_point", v)
        self._sset("split_bash", "1" if on else "0")
        if not on:
            self.bash_line.set(u"", "INFO")
        elif self.pinfo is None:
            self.bash_line.set(u"Точка не выбрана: расчёт возьмёт лучшую по продажам; «Найти по продажам» - выбрать из списка", "INFO")
        elif v:
            self.bash_line.set(u"Вся заморозка Бащинского - на %s" % v, "OK")
        else:
            self.bash_line.set(u"Холодильники Бащинского не определены%s: заморозка поделится по спросу - проверьте"
                               % ((u" (" + self.pinfo["problem"] + u")") if self.pinfo.get("problem") else u""), "WARN")
        self.update_action()

    def on_rcs(self):
        on = self.rcs_cb.isChecked()
        self.rc_cb.setEnabled(not on)
        self.rc_cb.setToolTip(u"Не действует: товар РЦ тоже едет на магазины" if on else u"")
        self.update_action()

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
            self.one_line.set(u"Товар РЦ - тоже на магазин" if self.rcs_cb.isChecked() else u"Товар РЦ - на %s" % self.core.SPLIT_RC_NAME, "INFO")
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
            self.summary.setText(u"%s · %d поз. · %s%s" % (self.chk["shop"], self.chk["n"], how_t,
                                                           (u" · РЦ-товар на магазины" if self.rcs_cb.isChecked() else u"")
                                                           + (u" · заморозка Бащинского одной партией" if self.bash_chk.isChecked() else u"")))
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
        rcs = self.rcs_cb.isChecked()
        auto, skip = h == "auto", (not self.rc_cb.isChecked()) or rcs
        one = self.one_cb.currentText().strip() if h == "one" else u""
        if h == "one" and not one:
            return
        if one and QMessageBox.question(
                self, u"Распределение", u"Всё, что обычно делится по точкам (в т.ч. кеги), - на один магазин: %s?\nТовар РЦ - %s "
                u"(лист: %d поз.)." % (one, u"тоже на этот магазин" if rcs else u"как обычно на " + core.SPLIT_RC_NAME, chk["n"])) != QMessageBox.Yes:
            return
        label = (u"Весь вывоз на %s" % one) if one else (u"Распределение (число точек по сумме)" if auto else u"Распределение на %d точек" % n)

        def done(info, out):
            self.info = info or {}
            rows = [{"addr": a, "pos": kk, "qty": core.fmt_qty(q), "sum": core.fmt_money(sm, 2)}
                    for a, kk, q, sm in self.info.get("summary", [])]
            self.win.journal(label, core._split_text(self.info))
            if self.info.get("ok"):
                tot = sum(float(r[3] or 0) for r in self.info.get("summary", []))
                warns = list(dict.fromkeys(self.info.get("dv_warn") or []))
                self.res.show_result(rows, u"Готово: файлов ТСД %d на %s грн. Файлы - в «ТСД_по_адресам» рядом со списком."
                                     % (len(self.info.get("files", [])), core.fmt_money(tot))
                                     + ((u" ВНИМАНИЕ: " + u"; ".join(warns)) if warns else u""), "WARN" if warns else "OK")
            else:
                self.res.show_result(rows, u"Есть ошибки - подробности в «Журнале».", "ERR")
                self.win.status(u"Распределение: есть ошибки", "ERR")
            from PySide6.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(label, lambda: core.run_split(self.d, fp, top_n=n, skip_rc=skip, auto_points=auto, target=one or None,
                                                rc_to_shops=rcs, bash_freeze=self.bash_chk.isChecked(),
                                                bash_point=self.bash_point()), done)

    def open_info(self, key):
        v = (self.info or {}).get(key)
        if v and os.path.exists(v):
            self.core.open_path(v)
        else:
            QMessageBox.information(self, u"Нет результата", u"Сначала нажмите «Распределить».")
