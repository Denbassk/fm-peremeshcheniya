@echo off
setlocal
cd /d "%~dp0"
title Build vyvoz_vne_matricy

echo ============================================================
echo   BUILD vyvoz_vne_matricy  (exe + folder vyvoz_lib)
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found in PATH.
    pause
    exit /b 1
)
if not exist "vyvoz_vne_matricy.py"   goto nofile
if not exist "vyvoz_vne_matricy.spec" goto nospec

echo [1/4] Installing build dependencies...
python -m pip install --upgrade --quiet pip pyinstaller
if errorlevel 1 goto fail
python -m pip install --quiet pandas openpyxl google-cloud-bigquery db-dtypes
if errorlevel 1 goto fail

echo [2/4] Self-test (no BigQuery needed)...
python vyvoz_vne_matricy.py --selftest
if errorlevel 1 goto fail

echo [3/4] Building into _build (takes a few minutes)...
python -m PyInstaller --noconfirm --clean --workpath "_build\work_vyvoz" --distpath "_build\dist_vyvoz" vyvoz_vne_matricy.spec
if errorlevel 1 goto fail

echo [4/4] Installing next to the script...
call install_vyvoz.bat
exit /b %errorlevel%

:nofile
echo [ERROR] vyvoz_vne_matricy.py not found next to this file.
pause
exit /b 1

:nospec
echo [ERROR] vyvoz_vne_matricy.spec not found next to this file.
pause
exit /b 1

:fail
echo.
echo [ERROR] Build failed. See messages above.
pause
exit /b 1
