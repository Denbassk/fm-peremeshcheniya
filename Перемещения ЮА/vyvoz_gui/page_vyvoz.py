"""Экран «Списки вывоза»: шаг 1 выгрузки, шаг 2 магазин, шаг 3 режим; «Сформировать список» закреплена внизу.

Поведение как в прежнем окне (run_app2, раздел «Списки вывоза»); логика - функции ядра (scan_inputs, mode_description, ...).
"""
import os

from PySide6.QtCore import QItemSelectionModel, Qt
from PySide6.QtWidgets import QButtonGroup, QCheckBox, QHBoxLayout, QLabel, QMessageBox, QRadioButton, QWidget

from . import tokens as T
from .app import Page
from .suppliers import SuppliersDialog
from .widgets import Banner, Card, RowsModel, StatusLine, button, fit_rows, muted, table

INPUT_HELP = (u"Нужен свежий файл в каждой обязательной папке. Склад ЮА - со ВСЕМИ строками, и нулевыми; приходы ЮА - "
              u"за весь период. Если в папке два файла (×2), берётся самый свежий. Старый или отсутствующий файл выгрузите "
              u"заново и нажмите «Перечитать». Двойной щелчок по строке открывает папку.")
STORE_HELP = (u"По умолчанию выбран магазин с самой свежей выгрузкой. Нужен другой - щёлкните по нему; несколько - "
              u"Ctrl или Shift. «Убрать из списка» убирает магазин только из таблицы, файл выгрузки остаётся в папке; "
              u"«Вернуть убранные» показывает их снова.")
MODE_SHORT = {                  # одна строка на экране; полное описание режима - под «?» (mode_description ядра)
    "normal": u"Автоматические исключения: сигареты, кеги, кулинария, расходники, поставщики и ШК из настроек.",
    "manual": u"Сигареты и кеги тоже попадут в список - лишнее удалите в Excel.",
}
MODE_HELP = (u"Обычный - автоматические исключения (сигареты, кеги, расходники, кулинария, список ШК, поставщики). "
             u"Ручной выбор - сигареты и кеги тоже попадают в список с причиной в последней колонке, лишнее удаляете в Excel. "
             u"Галочка «Вывезти ВСЁ» отменяет все правила: остаётся только то, что стоит на приходах или остатке этого магазина "
             u"на ЮА. Списки поставщиков и ШК «не вывозить» хранятся в папке Справочник и действуют на все следующие расчёты.")


class VyvozPage(Page):
    def __init__(self, core, win):
        super().__init__("vyvoz", core, win)
        self.d = win.dirs
        self.found, self.info = {}, None

        # ---------- шаг 1: выгрузки ----------
        c = Card(u"Шаг 1 · Проверьте выгрузки", INPUT_HELP)
        self.b_reread = c.add_head(button(u"Перечитать", "refresh", tip=u"Найти выгрузки заново (F5)", on_click=self.refresh))
        c.add_head(button(u"Открыть папку", "folder_open", tip=u"Папка ВХОД_ВЫВОЗ или папка выделенной строки",
                          on_click=self.open_input_folder))
        self.in_model = RowsModel((("fold", u"Папка", Qt.AlignLeft), ("what_need", u"Что положить", Qt.AlignLeft),
                                   ("file", u"Файл", Qt.AlignLeft), ("text", u"Состояние", Qt.AlignLeft)),
                                  level_key="level", status_col=3)
        self.in_view = table(self.in_model, stretch_col=1)      # имя файла целиком; сжимается пояснение (полное - в подсказке)
        self.in_view.setTextElideMode(Qt.ElideRight)
        self.in_view.setAccessibleName(u"Входные выгрузки")
        self.in_view.doubleClicked.connect(lambda *_: self.open_input_folder())
        c.lay.addWidget(self.in_view)
        self.in_line = StatusLine()
        c.lay.addWidget(self.in_line)
        self.body.addWidget(c)

        # ---------- шаг 2: магазин ----------
        c = Card(u"Шаг 2 · Выберите магазин", STORE_HELP)
        c.add_head(button(u"Папка МАГАЗИНЫ", "folder_open", on_click=lambda: core.open_path(self.d.stores)))
        self.b_all = button(u"Выбрать все", "select_all", on_click=lambda: self.st_view.selectAll())
        self.b_hide = button(u"Убрать из списка", "visibility_off", tip=u"Убрать выделенные магазины из таблицы (файлы остаются)",
                             on_click=self.hide_selected)
        self.b_show = button(u"Вернуть убранные", "visibility", on_click=self.restore_hidden)
        self.st_model = RowsModel((("shop", u"Магазин", Qt.AlignLeft), ("file", u"Файл выгрузки", Qt.AlignLeft),
                                   ("time", u"Выгружено", Qt.AlignLeft), ("age", u"Свежесть", Qt.AlignLeft)),
                                  level_key="level", status_col=3)
        self.st_view = table(self.st_model, stretch_col=1, multi=True)
        self.st_view.setAccessibleName(u"Магазины")
        self.st_view.selectionModel().selectionChanged.connect(self.on_pick)
        c.lay.addWidget(self.st_view)
        row = QHBoxLayout()
        for b in (self.b_all, self.b_hide, self.b_show):
            row.addWidget(b)
        row.addStretch(1)
        c.lay.addLayout(row)
        self.pick_line = StatusLine()                        # на всю ширину карточки, а не столбиком рядом с кнопками
        c.lay.addWidget(self.pick_line)
        self.hid_lbl = muted()
        c.lay.addWidget(self.hid_lbl)
        self.body.addWidget(c)

        # ---------- шаг 3: режим (настройки видны сразу, без «Настройки ▾») ----------
        c = Card(u"Шаг 3 · Режим", MODE_HELP)
        row = QHBoxLayout()
        row.setSpacing(T.SPACING_LG)
        self.mode_group = QButtonGroup(self)
        for val, text in core.VYVOZ_MODES:
            if val == "sweep":                               # зачистка в окне больше не нужна: её заменила галочка «Вывезти ВСЁ»
                continue
            rb = QRadioButton(text)
            rb.setProperty("mode", val)
            rb.setChecked(val == "manual")                   # как в прежнем окне: по умолчанию ручной выбор
            self.mode_group.addButton(rb)
            row.addWidget(rb)
        row.addStretch(1)
        c.lay.addLayout(row)
        self.mode_line = StatusLine(help=u"")                  # «?» сразу после строки: полное описание режима и поставщики
        c.lay.addWidget(self.mode_line)
        self.rm_cb = QCheckBox(u"Вывозить и товар из матрицы, который на ЮА ещё не завозился")
        self.rm_cb.setChecked(True)
        c.lay.addWidget(self.rm_cb)
        self.all_cb = QCheckBox(u"Вывезти ВСЁ, кроме того, что уже стоит у этого магазина на ЮА")
        self.all_cb.setToolTip(core.TAKE_ALL_TEXT)
        self.all_cb.setAccessibleName(u"Вывезти всё, отменив все правила")
        c.lay.addWidget(self.all_cb)
        # свёрнутый блок «Настройки вывоза»: списки «не вывозить» меняют редко
        self.opts_btn = button(u"Настройки вывоза", "expand_more", kind="ghost")
        self.opts_btn.setCheckable(True)
        self.opts_btn.setToolTip(u"Поставщики и ШК, которые не вывозим. Действуют на все следующие расчёты")
        self.opts_btn.setAccessibleName(u"Настройки вывоза: развернуть или свернуть")
        row = QHBoxLayout()
        row.addWidget(self.opts_btn)
        row.addStretch(1)
        c.lay.addLayout(row)
        self.opts = QWidget()
        orow = QHBoxLayout(self.opts)
        orow.setContentsMargins(T.SPACING_MD, 0, 0, 0)
        self.b_sup = button(u"Поставщики", "store", on_click=self.open_suppliers)
        self.b_keep = button(u"Исключённые ШК", "playlist_remove", tip=u"Список ШК «не вывозить» (Excel)", on_click=self.open_keep)
        orow.addWidget(self.b_sup)
        orow.addWidget(self.b_keep)
        orow.addStretch(1)
        self.opts.hide()
        c.lay.addWidget(self.opts)
        self.opts_btn.toggled.connect(self.opts.setVisible)
        self.body.addWidget(c)
        self.mode_group.buttonClicked.connect(lambda *_: self.mode_refresh())
        self.rm_cb.toggled.connect(lambda *_: self.mode_refresh())
        self.all_cb.toggled.connect(lambda *_: self.mode_refresh())

        # ---------- результат: появляется после расчёта ----------
        self.res_card = Card(u"Результат")
        self.banner = Banner()
        self.banner.action.connect(lambda: win.show_page("log"))
        self.res_card.lay.addWidget(self.banner)
        self.res_model = RowsModel((("shop", u"Магазин", Qt.AlignLeft), ("pos", u"Вывезти, поз.", Qt.AlignRight),
                                    ("sum", u"Сумма, грн", Qt.AlignRight), ("c1", u"Проверить 1", Qt.AlignRight),
                                    ("c2", u"Проверить 2", Qt.AlignRight), ("status", u"Статус", Qt.AlignLeft)),
                                   level_key="level", status_col=5)
        self.res_view = table(self.res_model, stretch_col=5)
        self.res_view.setAccessibleName(u"Результат по магазинам")
        self.res_card.lay.addWidget(self.res_view)
        row = QHBoxLayout()
        self.b_list = button(u"Открыть лист (Excel)", "description", on_click=self.open_list)
        self.b_dir = button(u"Папка результата", "folder_open", on_click=lambda: self.open_info("day_dir"))
        self.b_sum = button(u"Сводка (Excel)", "table_view", on_click=lambda: self.open_info("summary"))
        for b in (self.b_list, self.b_dir, self.b_sum):
            row.addWidget(b)
        row.addStretch(1)
        self.res_card.lay.addLayout(row)
        self.next_line = StatusLine()
        self.res_card.lay.addWidget(self.next_line)
        self.res_card.hide()
        self.body.addWidget(self.res_card)
        self.finish()

        # ---------- панель действия ----------
        bl = self.action_bar()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        bl.addWidget(self.summary, 1)
        self.go_btn = button(u"Сформировать список", "play_arrow", kind="primary", on_click=self.go)
        self.go_btn.setDefault(True)
        bl.addWidget(self.go_btn)
        win.busy_widgets += [self.b_reread, self.go_btn]
        from PySide6.QtGui import QKeySequence, QShortcut
        QShortcut(QKeySequence("F5"), self, activated=self.refresh)
        self.render_inputs([])
        self.render_stores()
        self.mode_refresh()

    # ---------------- шаг 1 ----------------
    def refresh(self):
        def done(res, out):
            res = res or {}
            self.found = res.get("found") or {}
            self.render_inputs(res.get("inputs") or [])
            self.render_stores(keep=self.selected())
            errs = res.get("errors") or []
            self.win.status(u"Магазинов найдено: %d" % len(self.found) if self.found else u"В ВХОД_ВЫВОЗ\\МАГАЗИНЫ нет выгрузок",
                            "OK" if self.found else "WARN")
            if errs:
                self.win.journal(u"Поиск выгрузок", u"\n".join(errs))
        self.win.run_task(u"Поиск выгрузок магазинов", lambda: self.core.scan_inputs(self.d), done)

    def render_inputs(self, rows):
        rows = [dict(x, file=x.get("file", u"").replace(u"Состояние склада ", u"", 1),
                     what_need=x.get("what", u"")) for x in rows]    # «обязательно / желательно» видно в состоянии
        for x in rows:                                       # время выгрузки - в подсказке, ширина остаётся пояснению
            x["text_tip"] = (u"Выгружено %s · %s" % (x["when"], x.get("text", u""))) if x.get("when") else x.get("text", u"")
            x["what_need_tip"] = u"%s · %s" % (x.get("what", u""), x.get("need", u""))
        self.in_model.set_rows(rows)
        fit_rows(self.in_view, 3, 8)
        bad = [x for x in rows if x.get("level") in ("ERR", "WARN")]
        if not rows:
            self.in_line.set(u"Нажмите «Перечитать», чтобы найти выгрузки.", "INFO")
        elif bad:
            self.in_line.set(u"Нужно обновить: %s. Выгрузите заново и нажмите «Перечитать»." % u"; ".join(x["fold"] for x in bad),
                             "ERR" if any(x["level"] == "ERR" for x in bad) else "WARN")
        else:
            self.in_line.set(u"Все выгрузки на месте и свежие.", "OK")

    def open_input_folder(self):
        ix = self.in_view.currentIndex()
        fp = self.in_model.row(ix.row()).get("folder") if ix.isValid() else None
        self.core.open_path(fp if fp and os.path.isdir(fp) else self.d.inp)

    # ---------------- шаг 2 ----------------
    def selected(self):
        return [self.st_model.row(ix.row())["shop"] for ix in self.st_view.selectionModel().selectedRows()]

    def render_stores(self, keep=None):
        core = self.core
        hidden = core.load_hidden_stores(self.d)
        shown = core.visible_stores(self.found, hidden)
        rows = []
        for x, v in shown.items():
            fp = v[0] if isinstance(v, (tuple, list)) else u""
            try:
                stamp, _age, lv, txt = core.file_state(fp)
                tm = stamp.strftime("%d.%m %H:%M")
            except Exception:
                tm, lv, txt = u"", "INFO", u""
            rows.append({"shop": x, "file": os.path.basename(fp), "time": tm, "age": txt, "level": lv, "path": fp})
        self.st_model.set_rows(rows)
        fit_rows(self.st_view, 2, 8)
        pick = [x for x in (keep or []) if x in shown] or ([core.pick_default_store(shown)] if shown else [])
        sm = self.st_view.selectionModel()
        sm.clearSelection()
        for i, r in enumerate(rows):
            if r["shop"] in pick:
                sm.select(self.st_model.index(i, 0), QItemSelectionModel.Select | QItemSelectionModel.Rows)
        n_h = len(self.found) - len(shown)
        self.b_show.setEnabled(bool(n_h))
        self.b_show.setText(u"Вернуть убранные (%d)" % n_h if n_h else u"Вернуть убранные")
        self.hid_lbl.setText((u"Убраны из списка: %s" % u"; ".join(sorted(x for x in self.found if x not in shown))) if n_h else u"")
        self.hid_lbl.setVisible(bool(n_h))
        self.b_all.setEnabled(len(rows) > 1)
        self.b_hide.setEnabled(bool(rows))
        self.on_pick()

    def on_pick(self, *_a):
        sel = self.selected()
        if not self.st_model.rows:
            self.pick_line.set(u"Магазинов нет: положите выгрузку «Состояние склада» в ВХОД_ВЫВОЗ\\МАГАЗИНЫ и нажмите «Перечитать».",
                               "WARN")
        elif not sel:
            self.pick_line.set(u"Магазин не выбран - выделите строку в таблице.", "WARN")
        elif len(sel) == 1:
            self.pick_line.set(u"", "INFO")
        else:
            self.pick_line.set(u"Будут сформированы списки для %d магазинов: %s" % (len(sel), u"; ".join(sel)), "WARN")
        self.upd_store_row(sel)
        self.update_action()

    def upd_store_row(self, sel):
        """Строка «склад магазина» в шаге 1 показывает выгрузку выбранного магазина (из нескольких - самую старую)."""
        rows = self.in_model.rows
        i = next((k for k, x in enumerate(rows) if x.get("key") == "stores"), None)
        paths = [self.found[x][0] for x in sel if x in self.found]
        if i is None or not paths:
            return
        worst = min(paths, key=self.core._file_stamp)
        stamp, _age, lv, txt = self.core.file_state(worst)
        rows[i] = dict(rows[i], file=os.path.basename(worst).replace(u"Состояние склада ", u"", 1),
                       when=stamp.strftime("%d.%m %H:%M"), text=txt, level=lv)
        rows[i]["text_tip"] = u"Выгружено %s · %s" % (rows[i]["when"], txt)
        self.in_model.set_rows(rows)

    def hide_selected(self):
        sel = self.selected()
        if not sel:
            QMessageBox.information(self, u"Убрать из списка", u"Выделите магазин, который нужно убрать из таблицы.")
            return
        if QMessageBox.question(self, u"Убрать из списка", u"Убрать из таблицы: %s?\n\nФайл выгрузки остаётся в папке, ничего не "
                                u"удаляется. Вернуть - кнопка «Вернуть убранные»." % u"; ".join(sel)) != QMessageBox.Yes:
            return
        try:
            self.core.save_hidden_stores(self.d, self.core.load_hidden_stores(self.d) | set(sel))
        except Exception as e:
            QMessageBox.critical(self, u"Убрать из списка", u"Не удалось сохранить: %s" % e)
            return
        self.render_stores()

    def restore_hidden(self):
        try:
            self.core.save_hidden_stores(self.d, set())
        except Exception as e:
            QMessageBox.critical(self, u"Вернуть убранные", u"Не удалось сохранить: %s" % e)
            return
        self.render_stores(keep=self.selected())

    # ---------------- шаг 3 ----------------
    def mode(self):
        b = self.mode_group.checkedButton()
        return b.property("mode") if b else "manual"

    def mode_refresh(self):
        n, when, _p = self.core.suppliers_file_info(self.d)
        m = self.mode()
        allout = self.all_cb.isChecked()
        for b in self.mode_group.buttons():
            b.setEnabled(not allout)
        self.rm_cb.setEnabled(not allout)
        self.rm_cb.setToolTip(u"Не действует: вывозится всё" if allout else u"")
        t, lv = self.core.mode_description(m, n, self.rm_cb.isChecked())
        note = self.core.suppliers_note(m, n, when)
        if allout:
            t, lv, note = self.core.TAKE_ALL_TEXT, "WARN", u"При «Вывезти ВСЁ» список поставщиков и ШК «не вывозить» не действуют."
            self.mode_line.set(self.core.TAKE_ALL_SHORT, "WARN")
        else:
            self.mode_line.set(MODE_SHORT.get(m, u""), "WARN" if lv == "WARN" else "INFO")
        self.mode_line.set_help(u"%s\n\n%s" % (t, note))
        self.b_sup.setToolTip(note)
        self.update_action()

    def open_suppliers(self):
        dlg = SuppliersDialog(self.core, self.d, self.info, self)
        if dlg.exec():
            self.mode_refresh()
            self.win.status(u"Поставщики сохранены: не вывозим %d" % dlg.saved, "OK")

    def open_keep(self):
        self.core.load_keep_shk(self.d)
        self.core.open_path(os.path.join(self.d.cache, self.core.KEEP_FILE))

    # ---------------- действие ----------------
    def update_action(self):
        sel = self.selected()
        mode_t = u"ВЫВЕЗТИ ВСЁ" if self.all_cb.isChecked() else dict(self.core.VYVOZ_MODES).get(self.mode(), u"")
        if not sel:
            self.summary.setText(u"Выберите магазин в шаге 2")
            self.go_btn.setEnabled(False)
            self.go_btn.setToolTip(u"Сначала выберите магазин в шаге 2")
        else:
            self.summary.setText(u"%s · %s" % (sel[0] if len(sel) == 1 else u"%d магазина(ов)" % len(sel), mode_t))
            self.go_btn.setEnabled(self.win.worker is None)
            self.go_btn.setToolTip(u"Excel-лист и файл для ТСД - в папке ВЫВОЗ\\<дата>\\<магазин>")

    def shown(self):
        self.update_action()

    def go(self):
        shops = self.selected()
        if not shops:
            return
        if len(shops) > 1 and QMessageBox.question(
                self, u"Списки вывоза", u"Выбрано магазинов: %d (%s).\n\nФормировать списки по всем?" % (len(shops), u"; ".join(shops)),
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        fl = self.core.vyvoz_mode_flags(self.mode())
        rm = self.rm_cb.isChecked()
        take_all = self.all_cb.isChecked()
        if take_all and QMessageBox.question(
                self, u"Вывезти ВСЁ", u"Включено «Вывезти ВСЁ»: уедет весь остаток (%s), кроме того, что стоит на приходах или остатке этого магазина на ЮА. "
                u"В список попадёт и то, что обычно остаётся: матрица и ЮА вообще, сигареты, кеги, расходники, кулинария, "
                u"овощи, поставщики с галочкой.\n\nЧто оставить на месте - удалите строки в Excel. Сформировать?" % u"; ".join(shops), QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No) != QMessageBox.Yes:
            return
        core, d = self.core, self.d

        def done(info, out):
            self.info = info or {}
            rows = []
            for sh in self.info.get("shops", []):
                ok = sh.get("status") == "OK"
                rows.append({"shop": sh.get("shop", u""), "pos": sh.get("vyvoz_pos", u""), "sum": core.fmt_money(sh.get("vyvoz_sum", 0)),
                             "c1": sh.get("check1", u""), "c2": sh.get("check2", u""), "status": sh.get("status", u""),
                             "level": "OK" if ok else "ERR"})
            self.res_model.set_rows(rows)
            fit_rows(self.res_view, 1, 8)
            self.res_view.setVisible(bool(rows))
            self.win.journal(u"Списки вывоза", core._result_text(self.info))
            ne = sum(1 for a, _b, _c in core.R.checks if a == "ERROR")
            okn = sum(1 for r in rows if r["level"] == "OK")
            pos = sum(int(sh.get("vyvoz_pos") or 0) for sh in self.info.get("shops", []) if sh.get("status") == "OK")
            sm = sum(float(sh.get("vyvoz_sum") or 0) for sh in self.info.get("shops", []) if sh.get("status") == "OK")
            if ne:
                self.banner.show_msg(u"Списков готово: %d из %d. Ошибок: %d - подробности в «Журнале»." % (okn, len(rows), ne), "ERR",
                                     u"Открыть журнал")
            elif self.info.get("take_all"):
                ok_sh = [sh for sh in self.info.get("shops", []) if sh.get("status") == "OK"]
                st_pos = sum(int(sh.get("stock_pos") or 0) for sh in ok_sh)
                st_sum = sum(float(sh.get("stock_sum") or 0) for sh in ok_sh)
                bad = sum(int(sh.get("bad_pos") or 0) for sh in ok_sh)
                stay = sum(int(sh.get("stay_pos") or 0) for sh in ok_sh)
                txt = (u"Уедет %d поз. на %s грн из %d на остатке (%s грн); остаётся на месте %d поз. (стоят на приходах магазина на ЮА)."
                       % (pos, core.fmt_money(sm), st_pos, core.fmt_money(st_sum), stay))
                if bad:
                    self.banner.show_msg(txt + u" НЕ загрузить в ТСД (ШК не 6-14 цифр): %d поз. - вывозить вручную, список в «Журнале»." % bad,
                                         "WARN", u"Журнал")
                else:
                    self.banner.show_msg(txt + u" Лишнее удалите в Excel.", "OK", u"Журнал")
            else:
                self.banner.show_msg(u"Готово списков: %d · вывезти %d поз. на %s грн." % (okn, pos, core.fmt_money(sm)), "OK", u"Журнал")
            self.has_list = bool(self.first_list())
            for b in (self.b_list,):
                b.setEnabled(self.has_list)
            self.b_dir.setEnabled(bool(self.info.get("day_dir")))
            self.b_sum.setEnabled(bool(self.info.get("summary")))
            self.next_line.set(u"Дальше: откройте лист, удалите лишнее и сохраните в «%s», затем раздел «Распределение»." % core.CORR_DIR
                               if self.has_list else u"", "INFO")
            self.res_card.show()
            from PySide6.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(u"Формирование списков вывоза",
                          lambda: core.run_vyvoz(d, shops=shops, remove_not_on_ua=rm, manual=fl["manual"],
                                                 take_all=take_all), done)

    def first_list(self):
        return next((x["list"] for x in (self.info or {}).get("shops", []) if x.get("list") and os.path.isfile(x["list"])), None)

    def open_list(self):
        fp = self.first_list()
        if fp:
            self.core.open_path(fp)

    def open_info(self, key):
        v = (self.info or {}).get(key)
        if v and os.path.exists(v):
            self.core.open_path(v)
        else:
            QMessageBox.information(self, u"Нет результата", u"Файл не найден: сформируйте список ещё раз.")
