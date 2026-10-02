# -*- mode: python ; coding: utf-8 -*-
"""
Сборка raschet_zakaza.exe.

Пересборка после правки скрипта — одна команда:
    build_raschet.bat
или вручную:
    python -m PyInstaller --noconfirm --clean raschet_zakaza.spec

Абсолютных путей внутри нет: exe берёт рабочей папкой ту, в которой лежит сам,
поэтому запускается на любом компьютере. Переопределить можно переменной
окружения FM_BASE_DIR.
"""

from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []

# Пакеты, которые PyInstaller сам не дотягивает целиком.
# Если какого-то нет в системе — просто пропускаем, скрипт умеет работать
# от кэша Справочник\transfer_pack.csv без обращения к BigQuery.
for _pkg in (
    "google.cloud.bigquery",
    "google.api_core",
    "google.auth",
    "google.oauth2",
    "google.resumable_media",
    "google.cloud.core",
    "db_dtypes",
    "sv_ttk",          # тема Windows 11 для окна
):
    try:
        _d, _b, _h = collect_all(_pkg)
        datas += _d
        binaries += _b
        hiddenimports += _h
    except Exception:
        pass

datas += [("raschet.ico", ".")]        # иконка окна

hiddenimports += [
    "openpyxl.cell._writer",
    "pandas._libs.tslibs.base",
    "google.cloud.bigquery_storage",
]

a = Analysis(
    ["raschet_zakaza.py"],
    pathex=[],
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
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="raschet_zakaza",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["raschet.ico"],
)
