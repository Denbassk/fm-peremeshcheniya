"""Экран «Довывоз»: остаток магазинов после перехода на ЮА вывозится полностью. Что остаётся в магазине - галочки по
категориям; один магазин или несколько сразу; куда - сильнейшие точки, своя сумма на точку или точки вручную.
Расчёт - core.run_dovyvoz (зачистка с выбором категорий + распределение с приоритетом точек)."""
import os

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QComboBox, QGridLayout, QHBoxLayout, QLabel,
                               QListWidget, QListWidgetItem, QMessageBox, QRadioButton, QSizePolicy, QSpinBox)

from . import tokens as T
from .app import Page
from .widgets import Card, ResultCard, StatusLine, button, help_button, muted

TITLE = u"Довывоз после перехода"
SHOPS_HELP = (u"Выгрузки «Состояние склада» из базы FAMILY - в папку ВХОД_ВЫВОЗ\\%s, один файл на магазин. «Один магазин» - "
              u"расчёт для выбранного; «Несколько сразу» - один запуск на отмеченные. У каждого магазина свои файлы ТСД. "
              u"Магазины из этой папки получателями не бывают и спросом не считаются.")
KEEP_HELP = (u"Отмеченное остаётся в магазине, всё остальное вывозится. Товар с некорректным ШК и тех. товары остаются "
             u"всегда (в ТСД не загрузить). Категории заморозки и скоропорта берутся из Справочник\\vyvoz_perishable.csv "
             u"(по комментарию строки) и по названию. Выбор запоминается.")
WHERE_HELP = (u"Прямые поставки, раскомплектация и кеги - на точки. Товар РЦ - тоже на точки, если его продают хоть на одной; "
              u"на %s - только товар РЦ, который нигде не продаётся и есть в матрице. Остатки получателей не учитываются.")
WHERE = (("top", u"Сильнейшие точки"), ("sum", u"По сумме на точку"), ("manual", u"Выбрать точки вручную"))


def _check_list():
    w = QListWidget()
    w.setSelectionMode(QAbstractItemView.NoSelection)
    w.setAlternatingRowColors(True)
    return w


def _fit(w, max_rows=8):
    n = max(2, min(max_rows, w.count()))
    w.setFixedHeight(n * T.ROW_H + 2 * w.frameWidth() + 4)


def _checked(w):
    return [w.item(i).text() for i in range(w.count()) if w.item(i).checkState() == Qt.Checked]


class DovyvozPage(Page):
    def __init__(self, core, win):
        core.UI_PAGES.setdefault("dv", (TITLE, u"Довывоз остатка после перехода на ЮА."))
        super().__init__("dv", core, win)
        self.d = win.dirs
        self.found, self.sig, self.info = [], None, None

        # ---------- шаг 1: магазины ----------
        c = Card(u"Шаг 1 · Какие магазины", SHOPS_HELP % core.DOVYVOZ_IN_DIR)
        c.add_head(button(u"Перечитать", "refresh", tip=u"Найти выгрузки заново", on_click=lambda: self.refresh(force=True)))
        c.add_head(button(u"Открыть папку", "folder_open", tip=u"ВХОД_ВЫВОЗ\\%s" % core.DOVYVOZ_IN_DIR, on_click=self.open_in))
        row = QHBoxLayout()
        self.mode = QButtonGroup(self)
        self.rb_one, self.rb_many = QRadioButton(u"Один магазин"), QRadioButton(u"Несколько сразу")
        self.rb_one.setChecked(True)
        for rb in (self.rb_one, self.rb_many):
            self.mode.addButton(rb)
            row.addWidget(rb)
        row.addStretch(1)
        c.lay.addLayout(row)
        self.one_cb = QComboBox()
        self.one_cb.setAccessibleName(u"Магазин для довывоза")
        self.one_cb.setMinimumWidth(16 * T.SPACING_MD)
        self.one_cb.currentIndexChanged.connect(lambda *_: self.update_action())
        c.lay.addWidget(self.one_cb)
        self.many = _check_list()
        self.many.setAccessibleName(u"Магазины для довывоза")
        self.many.itemChanged.connect(lambda *_: self.update_action())
        c.lay.addWidget(self.many)
        self.in_line = StatusLine()
        c.lay.addWidget(self.in_line)
        self.mode.buttonClicked.connect(lambda *_: self.on_mode())
        self.body.addWidget(c)

        # ---------- шаг 2: что оставляем ----------
        c = Card(u"Шаг 2 · Что оставляем в магазине", KEEP_HELP)
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_LG)
        g.setVerticalSpacing(T.SPACING_SM)
        saved = self._load_keep()
        self.keep = {}
        for i, (k, label, dflt) in enumerate(core.DOVYVOZ_CATS):
            cb = QCheckBox(label)
            cb.setChecked((k in saved) if saved is not None else dflt)
            if k == "keg":
                cb.setToolTip(u"Не отмечено - кеги едут целыми на точки, где их продают. На склад кеги не едут никогда.")
            cb.toggled.connect(lambda *_: (self._save_keep(), self.update_action()))
            self.keep[k] = cb
            g.addWidget(cb, i // 2, i % 2)
        c.lay.addLayout(g)
        row = QHBoxLayout()
        row.addWidget(button(u"Как по умолчанию", "refresh", on_click=lambda: self.set_keep(core.DOVYVOZ_KEEP_DEFAULT)))
        row.addWidget(button(u"Ничего не оставлять", "playlist_remove", on_click=lambda: self.set_keep(set())))
        row.addStretch(1)
        c.lay.addLayout(row)
        self.body.addWidget(c)

        # ---------- шаг 3: куда ----------
        c = Card(u"Шаг 3 · Куда везти", WHERE_HELP % core.SPLIT_RC_NAME)
        self.where = QButtonGroup(self)
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
        self.sum_sp = QSpinBox()
        self.sum_sp.setRange(100, 200000)
        self.sum_sp.setSingleStep(500)
        self.sum_sp.setValue(3000)
        self.sum_sp.setSuffix(u" грн")
        self.sum_sp.setFixedWidth(8 * T.SPACING_MD)
        self.sum_sp.setAccessibleName(u"Сумма на точку")
        self.sum_sp.valueChanged.connect(lambda *_: self.update_action())
        rows = {"top": (self.top_sp, u"по спросу на товары списка, без порога суммы"),
                "sum": (self.sum_sp, u"точка, которой досталось меньше, снимается"),
                "manual": (None, u"отметьте точки в списке ниже, порога нет")}
        for i, (val, text) in enumerate(WHERE):
            rb = QRadioButton(text)
            rb.setProperty("where", val)
            rb.setChecked(val == "top")
            self.where.addButton(rb)
            g.addWidget(rb, i, 0)
            h = QHBoxLayout()
            h.setSpacing(T.SPACING_SM)
            w_, hint = rows[val]
            if w_ is not None:
                h.addWidget(w_)
            h.addWidget(muted(hint, wrap=False))
            h.addStretch(1)
            g.addLayout(h, i, 1)
        c.lay.addLayout(g)
        self.pts = _check_list()
        self.pts.setAccessibleName(u"Точки-получатели")
        self.pts.itemChanged.connect(lambda *_: self.update_action())
        c.lay.addWidget(self.pts)
        self.where.buttonClicked.connect(lambda *_: self.on_where())
        self.body.addWidget(c)

        # ---------- результат ----------
        self.res = ResultCard(((("addr", u"Откуда → куда", Qt.AlignLeft), ("pos", u"Позиций", Qt.AlignRight),
                                ("qty", u"Единиц", Qt.AlignRight), ("sum", u"Сумма, грн", Qt.AlignRight))),
                              stretch_col=0, on_log=lambda: win.show_page("log"))
        self.res.add_button(button(u"Папка результата", "folder_open", on_click=self.open_out))
        self.body.addWidget(self.res)
        self.cats = ResultCard(((("cat", u"Категория", Qt.AlignLeft), ("st", u"Итог", Qt.AlignLeft),
                                 ("pos", u"Позиций", Qt.AlignRight), ("sum", u"Сумма, грн", Qt.AlignRight))),
                               stretch_col=0, on_log=lambda: win.show_page("log"), title=u"По категориям (все магазины)")
        self.body.addWidget(self.cats)
        self.finish()

        # ---------- панель действия ----------
        bl = self.action_bar()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        bl.addWidget(self.summary, 1)
        self.go_btn = button(u"Сформировать довывоз", "local_shipping", kind="primary", on_click=self.go)
        bl.addWidget(self.go_btn)
        win.busy_widgets += [self.go_btn]
        self.on_mode()
        self.on_where()

    # ---------------- настройки ----------------
    def _settings(self):
        return getattr(getattr(self.win, "themes", None), "settings", None)

    def _load_keep(self):
        try:
            s = self._settings()
            v = s.value("dovyvoz_keep") if s is not None else None
        except Exception:
            v = None
        if v is None:
            return None
        if isinstance(v, (list, tuple)):
            v = u",".join(str(x) for x in v)
        v = str(v)
        return set() if v == u"-" else set(x for x in v.split(u",") if x)

    def _save_keep(self):
        try:
            s = self._settings()
            if s is not None:
                s.setValue("dovyvoz_keep", u",".join(k for k, cb in self.keep.items() if cb.isChecked()) or u"-")
        except Exception:
            pass

    def set_keep(self, keys):
        for k, cb in self.keep.items():
            cb.setChecked(k in keys)

    # ---------------- шаг 1 ----------------
    def shown(self):
        self.refresh()

    def refresh(self, force=False):
        core = self.core
        files = core.list_xlsx(core.dovyvoz_in_dir(self.d))
        sig = tuple((p, core._mtime(p)) for p in files)
        if not force and sig == self.sig:
            return
        self.sig = sig
        self.found = core.dovyvoz_shops(self.d)
        good = [x for x in self.found if not x[2]]
        cur_one, cur_many = self.one_cb.currentText(), _checked(self.many)
        names = [x[0] for x in good]
        self.one_cb.blockSignals(True)
        self.one_cb.clear()
        self.one_cb.addItems(names)
        self.one_cb.setCurrentIndex(names.index(cur_one) if cur_one in names else (0 if names else -1))
        self.one_cb.setPlaceholderText(u"нет выгрузок" if not names else u"выберите магазин")
        self.one_cb.blockSignals(False)
        self.many.blockSignals(True)
        self.many.clear()
        for nm, p, _e in good:
            it = QListWidgetItem(nm)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if (not cur_many or nm in cur_many) else Qt.Unchecked)
            it.setToolTip(p)
            self.many.addItem(it)
        self.many.blockSignals(False)
        _fit(self.many)
        bad = [x for x in self.found if x[2]]
        if not self.found:
            self.in_line.set(u"В ВХОД_ВЫВОЗ\\%s нет выгрузок «Состояние склада» (Family)" % core.DOVYVOZ_IN_DIR, "WARN")
        elif bad:
            self.in_line.set(u"Магазинов: %d; не прочитаны: %s" % (len(good), u"; ".join(os.path.basename(x[1]) for x in bad)), "WARN")
        else:
            self.in_line.set(u"Магазинов в папке: %d" % len(good), "OK")
        self.fill_points()
        self.update_action()

    def fill_points(self):
        core = self.core
        mine = set(core._skey(x[0]) for x in self.found)
        vals = [s for s in core.split_targets(self.d, None) if core._skey(s) not in mine]
        cur = set(_checked(self.pts))
        self.pts.blockSignals(True)
        self.pts.clear()
        for s in vals:
            it = QListWidgetItem(s)
            it.setFlags(it.flags() | Qt.ItemIsUserCheckable)
            it.setCheckState(Qt.Checked if s in cur else Qt.Unchecked)
            self.pts.addItem(it)
        self.pts.blockSignals(False)
        _fit(self.pts, 10)

    def on_mode(self):
        one = self.rb_one.isChecked()
        self.one_cb.setVisible(one)
        self.many.setVisible(not one)
        self.update_action()

    def selected_shops(self):
        if self.rb_one.isChecked():
            t = self.one_cb.currentText().strip()
            return [t] if t else []
        return _checked(self.many)

    # ---------------- шаг 3 ----------------
    def where_val(self):
        b = self.where.checkedButton()
        return b.property("where") if b else "top"

    def on_where(self):
        h = self.where_val()
        self.top_sp.setEnabled(h == "top")
        self.sum_sp.setEnabled(h == "sum")
        self.pts.setVisible(h == "manual")
        self.update_action()

    # ---------------- действие ----------------
    def update_action(self):
        if not hasattr(self, "go_btn"):
            return
        shops, h, pts = self.selected_shops(), self.where_val(), _checked(self.pts)
        why = u""
        if not self.found:
            why = u"Положите выгрузки в ВХОД_ВЫВОЗ\\%s" % self.core.DOVYVOZ_IN_DIR
        elif not shops:
            why = u"Выберите магазин в шаге 1"
        elif h == "manual" and not pts:
            why = u"Отметьте хотя бы одну точку в шаге 3"
        if why:
            self.summary.setText(why)
        else:
            where_t = {"top": u"на %d сильнейших точек" % self.top_sp.value(),
                       "sum": u"не меньше %d грн на точку" % self.sum_sp.value(),
                       "manual": u"на точки: %s" % u", ".join(pts)}[h]
            n_keep = sum(1 for cb in self.keep.values() if cb.isChecked())
            self.summary.setText(u"%s · оставляем категорий: %d · %s" % (u", ".join(shops), n_keep, where_t))
        self.go_btn.setEnabled(not why and self.win.worker is None)
        self.go_btn.setToolTip(why or u"Файлы ТСД - в ВЫВОЗ\\<дата>_довывоз\\<магазин>\\ТСД_по_адресам")

    def go(self):
        core = self.core
        shops, h = self.selected_shops(), self.where_val()
        if not shops:
            return
        keep = [k for k, cb in self.keep.items() if cb.isChecked()]
        pts = _checked(self.pts) if h == "manual" else None
        top = self.top_sp.value() if h == "top" else None
        msum = float(self.sum_sp.value()) if h == "sum" else None
        keep_t = u", ".join(cb.text() for cb in self.keep.values() if cb.isChecked()) or u"ничего"
        if QMessageBox.question(
                self, TITLE, u"Довывоз: %s\n\nОстаётся в магазине: %s\n\nВсё остальное уедет: на точки; на %s - только товар РЦ, "
                u"который нигде не продаётся и есть в матрице.\n\nСформировать?" % (u", ".join(shops), keep_t, core.SPLIT_RC_NAME),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) != QMessageBox.Yes:
            return
        label = u"Довывоз: %s" % (shops[0] if len(shops) == 1 else u"%d магазинов" % len(shops))

        def done(info, out):
            self.info = info or {}
            splits = self.info.get("splits", [])
            rows = [{"addr": u"%s → %s" % (nm, a), "pos": kk, "qty": core.fmt_qty(q), "sum": core.fmt_money(sm, 2)}
                    for nm, sp in splits for a, kk, q, sm in (sp or {}).get("summary", [])]
            self.win.journal(label, core._dovyvoz_text(self.info))
            crow = [{"cat": c_, "st": s_, "pos": n_, "sum": core.fmt_money(sm_, 2)}
                    for c_, s_, n_, sm_ in core.dovyvoz_stat_rows(self.info.get("stats"))]
            if self.info.get("ok"):
                tot = sum(float(r_[3] or 0) for _n, sp in splits for r_ in (sp or {}).get("summary", []))
                nf = sum(len((sp or {}).get("files", [])) for _n, sp in splits)
                self.res.show_result(rows, u"Готово: файлов ТСД %d на %s грн. В ТСД грузить только из папок «ТСД_по_адресам»."
                                     % (nf, core.fmt_money(tot)), "OK", max_rows=12)
            else:
                self.res.show_result(rows, self.info.get("problem") or u"Есть ошибки - подробности в «Журнале».", "ERR")
                self.win.status(u"Довывоз: есть ошибки", "ERR")
            self.cats.show_result(crow, u"Что осталось в магазинах и что уехало", "INFO", max_rows=12)
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(label, lambda: core.run_dovyvoz(self.d, shops=shops, top_n=top, points=pts, min_sum=msum, keep=keep), done)

    def open_in(self):
        d_ = self.core.dovyvoz_in_dir(self.d)
        os.makedirs(d_, exist_ok=True)
        self.core.open_path(d_)

    def open_out(self):
        for _nm, sp in (self.info or {}).get("splits", []):
            od = (sp or {}).get("out_dir")
            if od and os.path.isdir(od):
                self.core.open_path(os.path.dirname(os.path.dirname(od)))
                return
        QMessageBox.information(self, TITLE, u"Сначала нажмите «Сформировать довывоз».")