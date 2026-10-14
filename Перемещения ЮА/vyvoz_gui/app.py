"""Главное окно: тёмная боковая панель с группами, экраны, строка состояния, фоновые задачи, журнал."""
import os
import sys
from datetime import datetime

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QFont, QKeySequence, QShortcut
from PySide6.QtWidgets import (QApplication, QButtonGroup, QFrame, QHBoxLayout, QLabel, QMainWindow, QMessageBox,
                               QPlainTextEdit, QProgressBar, QPushButton, QRadioButton, QScrollArea, QStackedWidget,
                               QVBoxLayout, QWidget)

from . import assets as A
from . import tokens as T
from . import theme as TH
from .widgets import Card, Theme, button, help_button, muted, on_theme, repolish
from .worker import Worker

# Меню по группам (решение владельца 14.10.2026): номеров нет, порядок - внутри группы «Сегодня»
NAV = (
    (u"СЕГОДНЯ", (("vyvoz", u"Списки вывоза", "local_shipping"), ("split", u"Распределение", "call_split"))),
    (u"НА СЛЕДУЮЩИЙ ДЕНЬ", (("upd", u"Подчистка списка", "cleaning_services"),)),
    (u"ОТДЕЛЬНО", (("fin", u"Инвентаризация ЮА", "inventory"),)),
    (None, (("serv", u"Сервис", "settings"), ("log", u"Журнал", "receipt_long"))),
)


class Page(QWidget):
    """Экран: заголовок с «?», прокручиваемое содержимое и закреплённая внизу панель действий."""

    def __init__(self, key, core, win):
        super().__init__()
        self.setObjectName("page")
        self.key, self.core, self.win = key, core, win
        title, help_text = core.UI_PAGES[key]
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        hd = QHBoxLayout()
        hd.setContentsMargins(T.SPACING_LG, T.SPACING_LG, T.SPACING_LG, T.SPACING_MD)
        hd.setSpacing(T.SPACING_SM)
        t = QLabel(title)
        t.setObjectName("pageTitle")
        t.setAccessibleName(title)
        hd.addWidget(t)
        hd.addWidget(help_button(help_text, title))
        hd.addStretch(1)
        v.addLayout(hd)
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)     # содержимое подстраивается под ширину окна
        inner = QWidget()
        inner.setObjectName("scrollBody")
        self.body = QVBoxLayout(inner)
        self.body.setContentsMargins(T.SPACING_LG, 0, T.SPACING_LG, T.SPACING_LG)
        self.body.setSpacing(T.SPACING_MD)
        self.scroll.setWidget(inner)
        v.addWidget(self.scroll, 1)
        self.bar = None

    def action_bar(self):
        """Панель внизу экрана: главное действие всегда видно без прокрутки."""
        if self.bar is None:
            self.bar = QFrame()
            self.bar.setObjectName("actionBar")
            self.bar_lay = QHBoxLayout(self.bar)
            self.bar_lay.setContentsMargins(T.SPACING_LG, T.SPACING_SM + 4, T.SPACING_LG, T.SPACING_SM + 4)
            self.bar_lay.setSpacing(T.SPACING_SM)
            self.layout().addWidget(self.bar)
        return self.bar_lay

    def finish(self):
        self.body.addStretch(1)

    def shown(self):
        pass


class Placeholder(Page):
    """Экран, который ещё не перенесён: работает в прежнем окне."""

    def __init__(self, key, core, win):
        super().__init__(key, core, win)
        c = Card(u"Этот раздел пока в прежнем окне")
        c.lay.addWidget(muted(u"Перенос идёт по одному экрану. До переноса откройте программу без ключа --qt "
                              u"(обычный ярлык): там раздел работает как раньше."))
        self.body.addWidget(c)
        self.finish()


class JournalPage(Page):
    def __init__(self, core, win):
        super().__init__("log", core, win)
        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        mono = QFont("Consolas")
        mono.setStyleHint(QFont.Monospace)
        mono.setPixelSize(T.TYPO_BODY_SM[0])
        self.text.setFont(mono)
        self.text.setAccessibleName(u"Журнал сеанса")
        self.body.addWidget(self.text, 1)
        bl = self.action_bar()
        bl.addWidget(button(u"Скопировать", "content_copy", on_click=self.copy))
        bl.addWidget(button(u"Очистить", "delete", on_click=self.text.clear))
        bl.addStretch(1)

    def add(self, title, text):
        self.text.appendPlainText(u"%s  %s\n%s\n" % (datetime.now().strftime("%H:%M:%S"), title, (text or u"").rstrip()))

    def copy(self):
        QApplication.clipboard().setText(self.text.toPlainText())
        self.win.status(u"Журнал скопирован", "OK")


class ServicePage(Page):
    def __init__(self, core, win):
        super().__init__("serv", core, win)
        d = core.DIRS if win.dirs is None else win.dirs
        c = Card(u"Оформление", u"Тема запоминается и действует при следующих запусках. «Как в Windows» следует настройке "
                                u"«Режим приложения» в параметрах Windows.")
        row = QHBoxLayout()
        self.theme_group = QButtonGroup(self)
        for mode, text in TH.MODES:
            rb = QRadioButton(text)
            rb.setProperty("mode", mode)
            rb.setChecked(win.themes.mode == mode)
            self.theme_group.addButton(rb)
            row.addWidget(rb)
        row.addStretch(1)
        self.theme_group.buttonClicked.connect(lambda b: win.themes.set_mode(b.property("mode")))
        c.lay.addLayout(row)
        self.body.addWidget(c)

        c = Card(u"Проверка программы", u"Самотест проверяет расчёты на тестовых данных; подробности - в «Журнале».")
        row = QHBoxLayout()
        b = button(u"Запустить самотест", "fact_check",
                   on_click=lambda: win.run_task(u"Самотест", core.self_test, lambda ok, out: (
                       win.journal(u"Самотест: %s" % (u"OK" if ok else u"ПРОВАЛ"), out), win.show_page("log"))))
        win.busy_widgets.append(b)
        row.addWidget(b)
        row.addStretch(1)
        c.lay.addLayout(row)
        self.body.addWidget(c)

        c = Card(u"Таблица соответствия ШК Family - ЮА")
        row = QHBoxLayout()
        b1 = button(u"Составить список кандидатов (Excel)", "table_view",
                    on_click=lambda: win.run_task(u"Список соответствия ШК", lambda: core.run_map_list(d), self._map_done))
        b2 = button(u"Загрузить мои решения из Excel…", "open_in_new", on_click=self._map_import)
        b3 = button(u"Открыть таблицу пар (Excel)", "open_in_new",
                    on_click=lambda: (core.load_ua_map(d), core.open_path(os.path.join(d.cache, core.UA_MAP_FILE))))
        win.busy_widgets += [b1, b2]
        for b in (b1, b2, b3):
            row.addWidget(b)
        row.addStretch(1)
        c.lay.addLayout(row)
        self.body.addWidget(c)

        c = Card(u"Папки")
        row = QHBoxLayout()
        for text, pth in ((u"Данные", d.base), (u"Входные выгрузки", d.inp), (u"Результаты", d.out),
                          (u"Справочник", d.cache), (u"Логи", d.logs)):
            row.addWidget(button(text, "folder", on_click=(lambda q=pth: core.open_path(q))))
        row.addStretch(1)
        c.lay.addLayout(row)
        self.body.addWidget(c)
        self.finish()
        self.d = d

    def _map_done(self, res, out):
        self.win.journal(u"Список соответствия ШК", self.core._map_text(res or {}))
        if (res or {}).get("path"):
            self.core.open_path(res["path"])
        else:
            QMessageBox.warning(self, u"Список соответствия", self.core._map_text(res or {}) or u"Не получилось")

    def _map_import(self):
        from PySide6.QtWidgets import QFileDialog
        md = os.path.join(self.d.base, self.core.MAP_DIR)
        fp, _f = QFileDialog.getOpenFileName(self, u"Список соответствия с вашими решениями", md if os.path.isdir(md) else self.d.base,
                                             u"Excel (*.xlsx *.xlsm);;Все файлы (*.*)")
        if not fp:
            return

        def done(res, out):
            self.win.journal(u"Загрузка решений по соответствию ШК", self.core._map_text(res or {}))
            QMessageBox.information(self, u"Соответствие ШК", self.core._map_text(res or {}))
        self.win.run_task(u"Загружаю решения", lambda: self.core.run_map_import(self.d, os.path.normpath(fp)), done)


class Sidebar(QFrame):
    def __init__(self, win, data_path):
        super().__init__()
        self.setObjectName("sidebar")
        self.setFixedWidth(T.SIDEBAR_W)
        v = QVBoxLayout(self)
        v.setContentsMargins(0, T.SPACING_LG, 0, T.SPACING_MD)
        v.setSpacing(0)
        brand = QHBoxLayout()
        brand.setContentsMargins(T.SPACING_MD, 0, T.SPACING_MD, 0)
        brand.setSpacing(T.SPACING_SM + 4)
        logo = QLabel()
        logo.setPixmap(A.app_icon().pixmap(QSize(40, 40)))
        logo.setFixedSize(40, 40)
        brand.addWidget(logo)
        name = QLabel(u"Фемелі Маркет")
        name.setObjectName("brandName")
        brand.addWidget(name, 1)
        v.addLayout(brand)
        bar = QFrame()
        bar.setObjectName("brandBar")
        bar.setFixedHeight(2)
        v.addSpacing(T.SPACING_SM + 4)
        bl = QHBoxLayout()
        bl.setContentsMargins(T.SPACING_MD, 0, T.SPACING_MD, 0)
        bl.addWidget(bar)
        v.addLayout(bl)
        sub = QLabel(u"Вывоз вне матрицы")
        sub.setObjectName("brandSub")
        sub.setContentsMargins(T.SPACING_MD, T.SPACING_SM, T.SPACING_MD, T.SPACING_SM)
        v.addWidget(sub)
        self.items = {}
        for group, items in NAV:
            if group:
                g = QLabel(group)
                g.setObjectName("sidebarGroupLabel")
                v.addWidget(g)
            else:
                v.addSpacing(T.SPACING_SM)
                dv = QFrame()
                dv.setObjectName("sidebarDivider")
                dv.setFixedHeight(1)
                dl = QHBoxLayout()
                dl.setContentsMargins(T.SPACING_MD, T.SPACING_SM, T.SPACING_MD, T.SPACING_SM)
                dl.addWidget(dv)
                v.addLayout(dl)
            for key, text, ic in items:
                b = QPushButton(text)
                b.setObjectName("sidebarNavItem")
                b.setCursor(Qt.PointingHandCursor)
                b.setAccessibleName(text)
                b.clicked.connect(lambda _c=False, k=key: win.show_page(k))
                self.items[key] = (b, ic)
                v.addWidget(b)
        v.addStretch(1)
        foot = QLabel(data_path)
        foot.setObjectName("sidebarFooter")
        foot.setToolTip(u"Папка с данными: " + data_path)
        foot.setWordWrap(False)
        foot.setContentsMargins(T.SPACING_MD, 0, T.SPACING_MD, 0)
        fm = foot.fontMetrics()
        foot.setText(fm.elidedText(u"Данные: " + data_path, Qt.ElideMiddle, T.SIDEBAR_W - 2 * T.SPACING_MD))
        v.addWidget(foot)
        on_theme(self.paint)

    def paint(self, active=None):
        for key, (b, ic) in self.items.items():
            on = b.property("active") == "true"
            b.setIcon(A.icon(ic, Theme.c["side_text_active"] if on else Theme.c["side_text"], 20))
            b.setIconSize(QSize(20, 20))

    def set_active(self, key):
        for k, (b, _ic) in self.items.items():
            b.setProperty("active", "true" if k == key else "false")
            b.setAccessibleDescription(u"открыт" if k == key else u"")
            repolish(b)
        self.paint()

    def keyPressEvent(self, e):
        """Стрелки вверх / вниз переходят между пунктами меню (Tab - тоже)."""
        keys = list(self.items)
        cur = next((i for i, k in enumerate(keys) if self.items[k][0].hasFocus()), None)
        if cur is not None and e.key() in (Qt.Key_Down, Qt.Key_Up):
            self.items[keys[(cur + (1 if e.key() == Qt.Key_Down else -1)) % len(keys)]][0].setFocus()
            return
        super().keyPressEvent(e)


class MainWindow(QMainWindow):
    def __init__(self, core, dirs, themes):
        super().__init__()
        self.core, self.dirs, self.themes = core, dirs, themes
        self.busy_widgets, self.worker = [], None
        self.setWindowTitle(u"Фемелі Маркет · Вывоз вне матрицы")
        self.setWindowIcon(A.app_icon())
        self.setMinimumSize(960, 600)
        central = QWidget()
        central.setObjectName("content")
        self.setCentralWidget(central)
        h = QHBoxLayout(central)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        self.side = Sidebar(self, dirs.base)
        h.addWidget(self.side)
        right = QVBoxLayout()
        right.setContentsMargins(0, 0, 0, 0)
        right.setSpacing(0)
        self.stack = QStackedWidget()
        right.addWidget(self.stack, 1)
        sb = QFrame()
        sb.setObjectName("statusBar")
        sl = QHBoxLayout(sb)
        sl.setContentsMargins(T.SPACING_LG, T.SPACING_XS + 2, T.SPACING_LG, T.SPACING_XS + 2)
        from .widgets import StatusLine
        self.stat = StatusLine(u"Готово к работе", "OK")
        sl.addWidget(self.stat, 1)
        self.prog = QProgressBar()
        self.prog.setRange(0, 0)
        self.prog.setFixedWidth(160)
        self.prog.setTextVisible(False)
        self.prog.hide()
        sl.addWidget(self.prog)
        right.addWidget(sb)
        h.addLayout(right, 1)

        from .page_vyvoz import VyvozPage
        self.log = JournalPage(core, self)
        self.pages = {"vyvoz": VyvozPage(core, self), "split": Placeholder("split", core, self),
                      "upd": Placeholder("upd", core, self), "fin": Placeholder("fin", core, self),
                      "serv": ServicePage(core, self), "log": self.log}
        for p in self.pages.values():
            self.stack.addWidget(p)
        for i, key in enumerate(("vyvoz", "split", "upd", "fin", "serv", "log"), 1):
            QShortcut(QKeySequence("Ctrl+%d" % i), self, activated=lambda k=key: self.show_page(k))
        self.show_page("vyvoz")

    def show_page(self, key):
        self.stack.setCurrentWidget(self.pages[key])
        self.side.set_active(key)
        self.pages[key].shown()

    def status(self, text, level="INFO"):
        self.stat.set(text, level)

    def journal(self, title, text):
        self.log.add(title, text)

    def run_task(self, label, fn, done, failed=None):
        """Одна фоновая задача за раз: кнопки запуска выключаются, внизу полоса прогресса и что сейчас делается."""
        if self.worker is not None:
            return False
        for w in self.busy_widgets:
            w.setEnabled(False)
        self.status(label + u"…", "INFO")
        self.prog.show()
        QApplication.setOverrideCursor(Qt.BusyCursor)
        self.worker = Worker(self.core, fn, self)

        def finish():
            self.worker = None
            for w in self.busy_widgets:
                w.setEnabled(True)
            self.prog.hide()
            QApplication.restoreOverrideCursor()
            for p in self.pages.values():
                p.shown() if p is self.stack.currentWidget() else None

        def ok(res, out):
            finish()
            self.status(u"Готово: %s" % label, "OK")
            done(res, out)

        def bad(err, out):
            finish()
            self.status(u"Ошибка: %s" % err, "ERR")
            self.journal(label + u" - ОШИБКА", err + u"\n" + out)
            if failed:
                failed(err, out)
            else:
                QMessageBox.critical(self, label, err)
        self.worker.done.connect(ok)
        self.worker.failed.connect(bad)
        self.worker.start()
        return True


def make_app(argv=None):
    app = QApplication.instance() or QApplication(argv or sys.argv[:1])
    app.setApplicationName(u"Вывоз вне матрицы")
    app.setOrganizationName("FamilyMarket")
    app.setWindowIcon(A.app_icon())
    probs = A.load_fonts()
    app.setFont(A.ui_font())
    return app, probs


def run(core, dirs=None, test_mode=False, settings=None):
    """Открыть окно. test_mode: собрать все экраны, показать каждый и закрыть -> (ключи экранов, окно не показывается)."""
    dirs = dirs or core.DIRS
    dirs.ensure()
    app, probs = make_app()
    themes = TH.ThemeManager(app, settings)
    themes.changed.connect(Theme.set)
    themes.apply()
    win = MainWindow(core, dirs, themes)
    for p_ in probs:
        win.journal(u"Шрифты", p_)
    if test_mode:
        for k in win.pages:
            win.show_page(k)
        keys = sorted(win.pages)
        win.close()
        win.deleteLater()
        return keys
    win.resize(1180, 800)
    win.show()
    win.pages["vyvoz"].refresh()
    return app.exec()
