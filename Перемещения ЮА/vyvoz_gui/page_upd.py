"""Экран «Подчистка списка» (на следующий день): исправленный лист + выгрузки ЮА -> очищенные копии в Обновление_<дата>.

Поведение как в прежнем окне (run_app2, раздел «Подчистка списка»); расчёт - run_update ядра без изменений.
"""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QGridLayout, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QSizePolicy

from . import tokens as T
from .app import Page
from .widgets import Card, ResultCard, button, help_button, pick_files

LIST_HELP = (u"Ваш исправленный лист вывоза (обычно из «%s»). Он и готовые файлы ТСД не меняются: очищенные копии "
             u"кладутся в папку Обновление_<дата> - в ТСД берите файлы оттуда.")
WH_HELP = u"Удаляем позиции с любым маршрутом, если ШК есть в выгрузке - с любым остатком: плюс, ноль или минус."
SHOP_HELP = u"Только для прямой доставки: удаляем, если ШК есть в «Состоянии склада» этого магазина в базе ЮА."


class UpdPage(Page):
    def __init__(self, core, win):
        super().__init__("upd", core, win)
        self.d = win.dirs
        self.wh, self.shp, self.info = [], [], None

        # ---------- шаг 1: лист ----------
        c = Card(u"Шаг 1 · Исправленный лист вывоза", LIST_HELP % core.CORR_DIR)
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_SM)
        g.setColumnStretch(0, 1)
        self.up_ed = QLineEdit()
        self.up_ed.setAccessibleName(u"Исправленный лист вывоза")
        self.up_ed.setPlaceholderText(u"путь к листу .xlsx")
        fs = core.list_corrected(self.d)
        self.up_ed.setText(os.path.normpath(fs[0]) if fs else u"")         # как в прежнем окне: самый свежий из Корректировка_ЮА
        self.up_ed.textChanged.connect(lambda *_: self.update_action())
        g.addWidget(self.up_ed, 0, 0)
        g.addWidget(button(u"Выбрать лист…", "folder_open", on_click=self.pick_list), 0, 1)
        c.lay.addLayout(g)
        self.body.addWidget(c)

        # ---------- шаг 2: выгрузки ЮА ----------
        c = Card(u"Шаг 2 · Выгрузки ЮА")
        g = QGridLayout()
        g.setHorizontalSpacing(T.SPACING_SM)
        g.setVerticalSpacing(T.SPACING_SM + 4)
        g.setColumnStretch(1, 1)
        self.wh_ed, self.shp_ed = QLineEdit(), QLineEdit()
        for i, (ed, lbl, hlp, ph) in enumerate(((self.wh_ed, u"Склад ЮА", WH_HELP, u"обязательно: «Состояние склада» Склад ЮА"),
                                                (self.shp_ed, u"Магазин в базе ЮА", SHOP_HELP, u"по желанию"))):
            h = QHBoxLayout()
            h.setSpacing(T.SPACING_XS)
            h.addWidget(QLabel(lbl))
            h.addWidget(help_button(hlp, lbl))                 # «?» сразу после подписи, а не через всю строку
            h.addStretch(1)
            g.addLayout(h, i, 0)
            ed.setReadOnly(True)
            ed.setPlaceholderText(ph)
            ed.setAccessibleName(lbl)
            g.addWidget(ed, i, 1)
        g.addWidget(button(u"Выбрать файл(ы)…", "folder_open", on_click=lambda: self.pick_into("wh")), 0, 2)
        g.addWidget(button(u"Выбрать файл(ы)…", "folder_open", on_click=lambda: self.pick_into("shp")), 1, 2)
        g.addWidget(button(u"Очистить", "close", tip=u"Не учитывать магазин в базе ЮА", on_click=lambda: self.set_files("shp", [])), 1, 3)
        c.lay.addLayout(g)
        self.body.addWidget(c)

        # ---------- результат ----------
        self.res = ResultCard(((("bc", u"ШК", Qt.AlignLeft), ("name", u"Название", Qt.AlignLeft), ("qty", u"Кол-во", Qt.AlignRight),
                                ("why", u"Причина", Qt.AlignLeft), ("where", u"Куда должно было ехать", Qt.AlignLeft))),
                              stretch_col=1, on_log=lambda: win.show_page("log"), title=u"Удалено из списка")
        self.res.add_button(button(u"Список удалённого (Excel)", "table_view", on_click=self.open_removed))
        self.res.add_button(button(u"Папка обновления", "folder_open", on_click=self.open_dir))
        self.body.addWidget(self.res)
        self.finish()

        bl = self.action_bar()
        self.summary = QLabel()
        self.summary.setWordWrap(True)
        self.summary.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        bl.addWidget(self.summary, 1)
        self.go_btn = button(u"Подчистить список", "cleaning_services", kind="primary", on_click=self.go)
        bl.addWidget(self.go_btn)
        win.busy_widgets += [self.go_btn]
        self.update_action()

    def pick_list(self):
        init = self.core.corr_dir(self.d)
        fps = pick_files(self, u"Исправленный лист вывоза", init if os.path.isdir(init) else self.d.out,
                         pattern=u"Excel (*.xlsx *.xlsm);;Все файлы (*.*)")
        if fps:
            self.up_ed.setText(fps[0])

    def pick_into(self, key):
        init = getattr(self.d, "ua_wh", self.d.inp) if key == "wh" else self.d.inp
        fps = pick_files(self, u"Склад ЮА" if key == "wh" else u"Магазин в базе ЮА", init, multi=True)
        if fps:
            self.set_files(key, fps)

    def set_files(self, key, fps):
        setattr(self, key, list(fps))
        ed = self.wh_ed if key == "wh" else self.shp_ed
        ed.setText(u"; ".join(os.path.basename(x) for x in fps))
        ed.setCursorPosition(0)                              # видно начало имени, а не хвост
        ed.setToolTip(u"\n".join(fps))
        self.update_action()

    def update_action(self):
        fp = self.up_ed.text().strip()
        why = (u"Выберите исправленный лист в шаге 1" if not fp or not os.path.isfile(fp) else
               u"Выберите «Состояние склада» Склад ЮА в шаге 2" if not self.wh else u"")
        self.summary.setText(why or u"%s · склад ЮА: %d файл(ов)%s" % (
            os.path.basename(fp), len(self.wh), u" · магазин ЮА: %d" % len(self.shp) if self.shp else u""))
        self.go_btn.setEnabled(not why and self.win.worker is None)
        self.go_btn.setToolTip(why or u"Очищенные копии - в папке Обновление_<дата>")

    def shown(self):
        if not self.up_ed.text().strip():                    # лист мог появиться после запуска («Взять без правок»)
            fs = self.core.list_corrected(self.d)
            if fs:
                self.up_ed.setText(os.path.normpath(fs[0]))
                self.up_ed.setCursorPosition(0)
        self.update_action()

    def go(self):
        core, fp = self.core, self.up_ed.text().strip()
        if not fp or not os.path.isfile(fp):
            QMessageBox.information(self, u"Подчистка списка", u"Выберите исправленный лист вывоза (xlsx).")
            return
        if core.is_auto_list(self.d, fp) and QMessageBox.question(
                self, u"Подчистка списка", u"Это автоматический лист из папки ВЫВОЗ, а не исправленный вами (он лежит в «%s»).\n\n"
                u"Продолжить именно с ним?" % core.CORR_DIR, QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes:
            return
        if not self.wh:
            QMessageBox.information(self, u"Подчистка списка", u"Выберите «Состояние склада» Склад ЮА.")
            return
        wh, shp = list(self.wh), list(self.shp)

        def done(info, out):
            self.info = info or {}
            rem = self.info.get("removed", [])
            rows = [{"bc": x["bc"], "name": x["name"], "qty": x["qty"], "why": x["why"], "where": x["where"]} for x in rem]
            self.win.journal(u"Подчистка списка", core._update_text(self.info))
            ok = self.info.get("ok")
            warn = any(x.startswith(u"ВНИМАНИЕ") for x in self.info.get("notes", []))
            if not ok:
                text, lv = u"Есть ошибки - подробности в «Журнале».", "ERR"
                self.win.status(u"Подчистка списка: есть ошибки", "ERR")
            else:
                text = (u"Удалено позиций: %d. Очищенные файлы - в папке обновления." % len(rem)) if rem else \
                    u"Ничего не стало ЮА-шным: список не изменился."
                text, lv = (text + u" Есть предупреждения - см. «Журнал».", "WARN") if warn else (text, "OK")
            self.res.show_result(rows, text, lv)
            from PySide6.QtCore import QTimer
            QTimer.singleShot(50, lambda: self.scroll.verticalScrollBar().setValue(self.scroll.verticalScrollBar().maximum()))
        self.win.run_task(u"Подчистка списка", lambda: core.run_update(self.d, fp, wh, shop_paths=shp), done)

    def open_dir(self):
        d_ = (self.info or {}).get("out_dir")
        self.core.open_path(d_) if d_ and os.path.exists(d_) else self.nores()

    def open_removed(self):
        d_ = (self.info or {}).get("out_dir")
        fp = os.path.join(d_, self.core.UPD_REMOVED_FILE) if d_ else u""
        self.core.open_path(fp) if fp and os.path.isfile(fp) else self.nores()

    def nores(self):
        QMessageBox.information(self, u"Нет результата", u"Сначала нажмите «Подчистить список».")
