@echo off
setlocal
cd /d "%~dp0"
title Install vyvoz_vne_matricy

if not exist "_build\dist_vyvoz\vyvoz_vne_matricy\vyvoz_vne_matricy.exe" (
    echo [ERROR] No fresh build in _build\dist_vyvoz. Run build_vyvoz.bat first.
    pause
    exit /b 1
)

tasklist /FI "IMAGENAME eq vyvoz_vne_matricy.exe" 2>nul | find /I "vyvoz_vne_matricy.exe" >nul
if not errorlevel 1 (
    echo [STOP] vyvoz_vne_matricy.exe is running. Close the program and run install_vyvoz.bat again.
    echo        The fresh build is kept in _build\dist_vyvoz, nothing was replaced.
    pause
    exit /b 1
)

echo Installing: folder vyvoz_lib and vyvoz_vne_matricy.exe next to this file...
robocopy "_build\dist_vyvoz\vyvoz_vne_matricy\vyvoz_lib" "vyvoz_lib" /MIR /NFL /NDL /NJH /NJS /NP >nul
if errorlevel 8 goto fail
copy /y "_build\dist_vyvoz\vyvoz_vne_matricy\vyvoz_vne_matricy.exe" "vyvoz_vne_matricy.exe" >nul
if errorlevel 1 goto fail

echo.
echo ============================================================
echo   DONE: vyvoz_vne_matricy.exe + vyvoz_lib (keep them together)
echo   Data are read from this folder: VHOD_VYVOZ, VYVOZ, Spravochnik.
echo   raschet_zakaza.py and credentials stay in the sibling folder.
echo ============================================================
pause
exit /b 0

:fail
echo.
echo [ERROR] Install failed. Close the program and try again.
pause
exit /b 1
