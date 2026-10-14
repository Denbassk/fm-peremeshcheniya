"""Долгие операции (чтение Excel, BigQuery, формирование списков) - в QThread, окно не зависает."""
from PySide6.QtCore import QThread, Signal


class Worker(QThread):
    """fn выполняется в потоке через core.run_captured: в журнал задачи попадает только её вывод.
    done(результат, вывод) или failed(ошибка, вывод) приходят в поток окна."""
    done = Signal(object, str)
    failed = Signal(str, str)

    def __init__(self, core, fn, parent=None):
        super().__init__(parent)
        self.core, self.fn = core, fn

    def run(self):
        box = self.core.run_captured(self.fn)
        if "err" in box:
            self.failed.emit(box["err"], box.get("out", u""))
        else:
            self.done.emit(box.get("res"), box.get("out", u""))
