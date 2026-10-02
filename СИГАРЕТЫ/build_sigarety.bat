@echo off
setlocal
cd /d "%~dp0"
title Build sigarety_tsd.exe

echo ============================================================
echo   BUILD sigarety_tsd.exe
echo ============================================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found in PATH.
    echo         Install Python 3.11+ with "Add Python to PATH".
    pause
    exit /b 1
)

if not exist "sigarety_tsd.py"   goto nofile
if not exist "sigarety_tsd.spec" goto nospec

echo [1/4] Installing build dependencies...
python -m pip install --upgrade --quiet pip pyinstaller
if errorlevel 1 goto fail
python -m pip install --quiet openpyxl
if errorlevel 1 goto fail

echo [2/4] Building...
python -m PyInstaller --noconfirm --clean ^
    --workpath "_build\work" --distpath "_build\dist" ^
    sigarety_tsd.spec
if errorlevel 1 goto fail

echo [3/4] Placing exe next to the data folders...
copy /y "_build\dist\sigarety_tsd.exe" "sigarety_tsd.exe" >nul
if errorlevel 1 goto fail

echo [4/4] Cleaning up...
if exist "__pycache__" rd /s /q "__pycache__"

echo.
echo ============================================================
echo   DONE:  sigarety_tsd.exe   (in this folder)
echo.
echo   The tool asks for the input file and the output folder
echo   through Explorer dialogs - no paths are hardcoded.
echo   Temporary build files live in _build and can be deleted.
echo ============================================================
pause
exit /b 0

:nofile
echo [ERROR] sigarety_tsd.py not found next to this file.
pause
exit /b 1

:nospec
echo [ERROR] sigarety_tsd.spec not found next to this file.
pause
exit /b 1

:fail
echo.
echo [ERROR] Build failed. See messages above.
pause
exit /b 1
