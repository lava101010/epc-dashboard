@echo off
rem ============================================================
rem  EPC Production Dashboard - Daily Build (double-click me)
rem  Picks the NEWEST export from incoming/ by itself,
rem  shows everything on screen and writes last_build.txt
rem ============================================================
cd /d "%~dp0"

set "PY=%~dp0tools\python\python.exe"
if exist "%PY%" goto py_ok
where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PY=python"
    goto py_ok
)
echo [ERROR] No Python found (portable or system).
pause
exit /b 1

:py_ok

set "SRC="
if exist incoming\ (
  for /f "delims=" %%f in ('dir /b /o-d "incoming\0040_qa*.xls" "incoming\0040_qa*.xlsx" 2^>nul') do if not defined SRC set "SRC=incoming\%%f"
  for /f "delims=" %%f in ('dir /b /o-d "incoming\*.xls" "incoming\*.xlsx" 2^>nul') do if not defined SRC set "SRC=incoming\%%f"
)
if defined SRC (
    echo [build] Using export file: %SRC%
    "%PY%" tools\update_dashboard.py "%SRC%" > last_build.txt 2>&1
) else (
    echo [build] incoming/ is empty - letting the engine search other places...
    "%PY%" tools\update_dashboard.py > last_build.txt 2>&1
)
type last_build.txt
if errorlevel 1 (
    echo.
    echo [FAILED] See above / send last_build.txt for support.
    pause
    exit /b 1
)

echo.
echo Opening dashboard (index.html)...
start "" "%~dp0dist\index.html"
ping 127.0.0.1 -n 4 >nul
