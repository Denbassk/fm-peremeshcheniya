# -*- mode: python ; coding: utf-8 -*-
r"""
Сборка vyvoz_vne_matricy (вывоз вне матрицы при переходе на ЮА Маркет), раскладка onedir:
    vyvoz_vne_matricy.exe (около 20 МБ) + папка vyvoz_lib\ (около 220 МБ). Держать ВМЕСТЕ.

Запуск: build_vyvoz.bat (самотест -> сборка в _build -> install_vyvoz.bat) или вручную:
    python -m PyInstaller --noconfirm --clean --workpath _build\work_vyvoz --distpath _build\dist_vyvoz vyvoz_vne_matricy.spec

Скрипт лежит в «Перемещения ЮА» и импортирует raschet_zakaza из соседней папки «РАСЧЕТ ЗАКАЗА» (pathex ниже).
Абсолютных путей нет: программа работает из своей папки (ВХОД_ВЫВОЗ, ВЫВОЗ, Справочник), ключ BigQuery берёт
из «РАСЧЕТ ЗАКАЗА\credentials». Переопределить: FM_VYVOZ_DIR, GOOGLE_APPLICATION_CREDENTIALS.
"""
import os

from PyInstaller.utils.hooks import collect_all

RZ = os.path.join("..", "РАСЧЕТ ЗАКАЗА")
datas, binaries, hiddenimports = [], [], []

# Пакеты, которые PyInstaller сам не дотягивает целиком. Нет в системе - пропускаем: скрипт умеет работать
# от кэша Справочник\vyvoz_*.csv без BigQuery (первый запуск без BigQuery и без кэша остановится: матрицы нет).
for _pkg in (
    "google.cloud.bigquery",
    "google.api_core",
    "google.auth",
    "google.oauth2",
    "google.resumable_media",
    "google.cloud.core",
    "db_dtypes",
):
    try:
        _d, _b, _h = collect_all(_pkg)
        datas += _d
        binaries += _b
        hiddenimports += _h
    except Exception:
        pass

# Окно PySide6 (--qt): шрифты Inter и значки Material Symbols, иконка окна. Ищутся в sys._MEIPASS (= vyvoz_lib).
datas += [(os.path.join("assets", "fonts"), os.path.join("assets", "fonts")), ("vyvoz_fm.ico", ".")]

if os.path.isfile(os.path.join(RZ, "raschet.ico")):
    datas += [(os.path.join(RZ, "raschet.ico"), ".")]      # raschet_zakaza ищет свою иконку рядом

hiddenimports += [
    "raschet_zakaza",
    "tkinter",
    "tkinter.ttk",
    "vyvoz_gui", "vyvoz_gui.app", "vyvoz_gui.page_vyvoz", "vyvoz_gui.page_split", "vyvoz_gui.page_upd", "vyvoz_gui.page_fin", "vyvoz_gui.suppliers", "vyvoz_gui.razkompl", "vyvoz_gui.widgets",
    "vyvoz_gui.theme", "vyvoz_gui.tokens", "vyvoz_gui.assets", "vyvoz_gui.worker",
    "openpyxl.cell._writer",
    "pandas._libs.tslibs.base",
    "google.cloud.bigquery_storage",
]

a = Analysis(
    ["vyvoz_vne_matricy.py"],
    pathex=[".", RZ],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "matplotlib", "scipy", "PyQt5", "PySide2",
        "IPython", "notebook", "jupyter", "pytest", "sphinx",
        "PIL", "sqlalchemy",
        # части Qt, которые окну не нужны (иначе сборка тяжелее на сотни МБ)
        "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick", "PySide6.QtWebChannel",
        "PySide6.QtQuick", "PySide6.QtQuick3D", "PySide6.QtQml", "PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets",
        "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtPdf",
        "PySide6.QtPdfWidgets", "PySide6.QtBluetooth", "PySide6.QtPositioning", "PySide6.QtLocation", "PySide6.QtSensors",
        "PySide6.QtSerialPort", "PySide6.QtSql", "PySide6.QtTest", "PySide6.QtDesigner", "PySide6.QtHelp",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="vyvoz_vne_matricy",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["vyvoz_fm.ico"],
    contents_directory="vyvoz_lib",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="vyvoz_vne_matricy",
)
