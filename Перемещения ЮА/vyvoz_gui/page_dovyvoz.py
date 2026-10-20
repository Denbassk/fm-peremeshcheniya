"""Экран «Довывоз»: остаток магазинов после перехода на ЮА вывозится полностью. Что остаётся в магазине - галочки по
категориям; один магазин или несколько сразу; куда - сильнейшие точки, своя сумма на точку или точки вручную.
Заморозка Бащинского - одной партией на одну точку с его холодильником (выбор в шаге 3), стаканы и кофе - на точку с кофеаппаратом.
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
              u"на %s - только товар РЦ, который нигде не продаётся и есть в матрице. Стаканы, крышки, кофе и сырьё "
              u"кофеаппарата - одной партией на точку с кофеаппаратом (если она среди отмеченных вручную - на неё). Пакеты - по "
              u"продажам. Заморозка Бащинского - на точку, выбранную ниже. Всё это на склад не едет. Остатки получателей не учитываются.")
BASH_HELP = (u"Холодильники Бащинского есть не во всех магазинах. Точки с холодильником - где за 60 дней продано не меньше 5 позиций его заморозки (ВХОД_ВЫВОЗ\\ДОВЫВОЗ\\заморозка_бащинский.txt) не меньше чем в 6 неделях "
             u"(BigQuery), или список из файла ВХОД_ВЫВОЗ\\ДОВЫВОЗ\\холодильники_бащинский.txt (по магазину в строке). "
             u"Вся заморозка Бащинского из всех магазинов довывоза едет одной партией на выбранную точку, без порога суммы. "
             u"Выбор запоминается: при поочерёдном расчёте магазинов точка та же. Хладик - как обычный товар, его холодильники везде.")
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
        self.pinfo, self._bash_warn = None, u""

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
            elif k in ("cup", "coffee"):
                cb.setToolTip(u"Не отмечено - одной партией на точку с кофеаппаратом. На склад не едут никогда.")
            elif k == "bag":
                cb.setToolTip(u"Не отмечено - по точкам по продажам. На склад не едут никогда.")
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
        brow = QHBoxLayout()
        brow.setSpacing(T.SPACING_SM)
        brow.addWidget(QLabel(u"Заморозка Бащинского →"))
        self.bash_cb = QComboBox()
        self.bash_cb.setAccessibleName(u"Куда везти заморозку Бащинского")
        self.bash_cb.setMinimumWidth(18 * T.SPACING_MD)
        self.bash_cb.currentIndexChanged.connect(lambda *_: (self._save_bash(), self.update_action()))
        brow.addWidget(self.bash_cb)
        self.bash_btn = button(u"Найти по продажам", "refresh",
                               tip=u"Точки с холодильником Бащинского и с кофеаппаратом - из BigQuery", on_click=self.load_points)
        brow.addWidget(self.bash_btn)
        brow.addWidget(help_button(BASH_HELP, u"Заморозка Бащинского"))
        brow.addStretch(1)
        c.lay.addLayout(brow)
        self.bash_line = StatusLine()
        c.lay.addWidget(self.bash_line)
        self.coffee_line = StatusLine()
        c.lay.addWidget(self.coffee_line)
        self._fill_bash()
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
        win.busy_widgets += [self.go_btn, self.bash_btn]
        self.on_mode()
        self.on_where()

    # ---------------- настройки ----------------
    def _settings(self):
        return getattr(getattr(self.win, "themes", None), "settings", None)

    def _sget(self, key):
        try:
            s = self._settings()
            v = s.value(key) if s is not None else None
            return None if v is None else str(v)
        except Exception:
            return None

    def _sset(self, key, val):
        try:
            s = self._settings()
            if s is not None:
                s.setValue(key, val)
        except Exception:
            pass

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
        self._sset("dovyvoz_keep", u",".join(k for k, cb in self.keep.items() if cb.isChecked()) or u"-")

    def _save_bash(self):
        v = self.bash_cb.currentData()
        if v:
            self._sset("dovyvoz_bash_point", v)

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
        self._mark_points()

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

    # ---------------- точки Бащинского и кофеаппарата ----------------
    def load_points(self):
        if self.win.worker is not None:
            return
        core, mine = self.core, [x[0] for x in self.found]
        self.win.run_task(u"Довывоз: точки Бащинского и кофеаппарата", lambda: core.dovyvoz_points(self.d, mine), self._points_done)

    def _points_done(self, info, out):
        self.pinfo = info if isinstance(info, dict) else {"bash": [], "coffee": [], "problem": u"нет ответа"}
        self._fill_bash()
        self._mark_points()
        self.update_action()

    def _fill_bash(self):
        p = self.pinfo
        fr = (p or {}).get("bash") or []
        saved = self._sget("dovyvoz_bash_point")
        self.bash_cb.blockSignals(True)
        self.bash_cb.clear()
        if p is None:
            self.bash_cb.addItem(u"нажмите «Найти по продажам» (или расчёт найдёт сам)", None)
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

    def _mark_points(self):
        p = self.pinfo or {}
        sk = self.core._skey
        fr = set(sk(s) for s, _q in p.get("bash") or [])
        cf = set(sk(s) for s, _q in p.get("coffee") or [])
        for i in range(self.pts.count()):
            it = self.pts.item(i)
            tags = ([u"есть холодильник Бащинского"] if sk(it.text()) in fr else []) + ([u"кофеаппарат"] if sk(it.text()) in cf else [])
            it.setToolTip(u", ".join(tags))

    def _point_lines(self, h, pts):
        core, p = self.core, self.pinfo
        self._bash_warn = u""
        if p is None:
            self.bash_line.set(u"Точки с холодильником Бащинского ещё не искались: «Найти по продажам» - или расчёт возьмёт лучшую сам", "OK")
            self.coffee_line.setVisible(False)
            return
        sk = core._skey
        fr = [s for s, _q in p.get("bash") or []]
        bp = self.bash_cb.currentData()
        if not fr:
            self.bash_line.set(u"Холодильники Бащинского не определены%s: его заморозка поделится по спросу - проверьте"
                               % ((u" (" + p["problem"] + u")") if p.get("problem") else u""), "WARN")
        elif h == "manual" and pts and not any(sk(s) in set(sk(x) for x in fr) for s in pts):
            self._bash_warn = u"Среди выбранных нет холодильника Бащинского, заморозка поедет на %s" % bp
            self.bash_line.set(self._bash_warn, "WARN")
        else:
            self.bash_line.set(u"Холодильник Бащинского (%d): %s. Вся его заморозка - на %s" % (len(fr), u", ".join(fr), bp), "OK")
        cl = [s for s, _q in p.get("coffee") or []]
        self.coffee_line.setVisible(True)
        if not cl:
            self.coffee_line.set(u"Точки с кофеаппаратом не определены: стаканы и кофе поделятся по спросу", "WARN")
        else:
            sel = set(sk(x) for x in pts) if h == "manual" else set()
            cp = next((s for s in cl if sk(s) in sel), cl[0])
            self.coffee_line.set(u"Стаканы, кофе, сырьё кофеаппарата - одной партией на %s (кофеаппарат на %d точках)" % (cp, len(cl)), "OK")

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
        self._point_lines(h, pts)
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
        bp = self.bash_cb.currentData()
        keep_t = u", ".join(cb.text() for cb in self.keep.values() if cb.isChecked()) or u"ничего"
        bash_t = (u"Заморозка Бащинского: вся на %s" % bp) if bp else u"Заморозка Бащинского: на лучшую точку с холодильником (найдёт расчёт)"
        if self._bash_warn:
            bash_t += u"\nВНИМАНИЕ: " + self._bash_warn
        if QMessageBox.question(
                self, TITLE, u"Довывоз: %s\n\nОстаётся в магазине: %s\n\nВсё остальное уедет: на точки; на %s - только товар РЦ, "
                u"который нигде не продаётся и есть в матрице.\n\n%s\n\nСформировать?"
                % (u", ".join(shops), keep_t, core.SPLIT_RC_NAME, bash_t),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) != QMessageBox.Yes:
            return
        label = u"Довывоз: %s" % (shops[0] if len(shops) == 1 else u"%d магазинов" % len(shops))

        def done(info, out):
            self.info = info or {}
            if isinstance(self.info.get("points"), dict):
                self.pinfo = self.info["points"]
                if self.info.get("bash_point"):
                    self._sset("dovyvoz_bash_point", self.info["bash_point"])
                self._fill_bash()
                self._mark_points()
                self.update_action()
            splits = self.info.get("splits", [])
            rows = [{"addr": u"%s → %s" % (nm, a), "pos": kk, "qty": core.fmt_qty(q), "sum": core.fmt_money(sm, 2)}
                    for nm, sp in splits for a, kk, q, sm in (sp or {}).get("summary", [])]
            self.win.journal(label, core._dovyvoz_text(self.info))
            crow = [{"cat": c_, "st": s_, "pos": n_, "sum": core.fmt_money(sm_, 2)}
                    for c_, s_, n_, sm_ in core.dovyvoz_stat_rows(self.info.get("stats"))]
            warns = list(dict.fromkeys(w_ for _n, sp in splits for w_ in ((sp or {}).get("dv_warn") or [])))
            if self.info.get("ok"):
                tot = sum(float(r_[3] or 0) for _n, sp in splits for r_ in (sp or {}).get("summary", []))
                nf = sum(len((sp or {}).get("files", [])) for _n, sp in splits)
                msg = u"Готово: файлов ТСД %d на %s грн. В ТСД грузить только из папок «ТСД_по_адресам»." % (nf, core.fmt_money(tot))
                if warns:
                    msg += u" ВНИМАНИЕ: " + u"; ".join(warns)
                self.res.show_result(rows, msg, "WARN" if warns else "OK", max_rows=12)
            else:
                self.res.show_result(rows, self.info.get("problem") or u"Есть ошибки - подробности в «Журнале».", "ERR")
                self.win.status(u"Довывоз: есть ошибки", "ERR")
            self.cats.show_result(crow, u"Что осталось в магазинах и что уехало", "INFO", max_rows=12)
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(label, lambda: core.run_dovyvoz(self.d, shops=shops, top_n=top, points=pts, min_sum=msum,
                                                          keep=keep, bash_point=bp), done)

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