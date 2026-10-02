@echo off
setlocal
cd /d "%~dp0"
title Build raschet_zakaza.exe

echo ============================================================
echo   BUILD raschet_zakaza.exe
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found in PATH.
    echo         Install Python 3.11+ with "Add Python to PATH".
    pause
    exit /b 1
)

if not exist "raschet_zakaza.py"   goto nofile
if not exist "raschet_zakaza.spec" goto nospec

echo [1/4] Installing build dependencies...
python -m pip install --upgrade --quiet pip pyinstaller
if errorlevel 1 goto fail
python -m pip install --quiet pandas openpyxl google-cloud-bigquery db-dtypes
if errorlevel 1 goto fail
python -m pip install --quiet sv-ttk
if errorlevel 1 goto fail

echo [2/4] Building (takes a few minutes)...
python -m PyInstaller --noconfirm --clean ^
    --workpath "_build\work" --distpath "_build\dist" ^
    raschet_zakaza.spec
if errorlevel 1 goto fail

echo [3/4] Placing exe next to the working folders...
copy /y "_build\dist\raschet_zakaza.exe" "raschet_zakaza.exe" >nul
if errorlevel 1 goto fail

echo [4/4] Cleaning up...
if exist "__pycache__" rd /s /q "__pycache__"

echo.
echo ============================================================
echo   DONE:  raschet_zakaza.exe   (in this folder)
echo.
echo   To hand the tool to another laptop, copy THIS whole folder.
echo   Everything is relative - no absolute paths inside.
echo   Temporary build files live in _build and can be deleted.
echo ============================================================
pause
exit /b 0

:nofile
echo [ERROR] raschet_zakaza.py not found next to this file.
pause
exit /b 1

:nospec
echo [ERROR] raschet_zakaza.spec not found next to this file.
pause
exit /b 1

:fail
echo.
echo [ERROR] Build failed. See messages above.
pause
exit /b 1
