@echo off
rem ============================================================
rem  EPC Dashboard - DIAGNOSTIC (double-click, then photograph
rem  the window or copy its text and send it for support)
rem ============================================================
cd /d "%~dp0"

echo ================================================
echo  1) Portable Python inside the project:
echo ================================================
if exist "tools\python\python.exe" (echo    FOUND: tools\python\python.exe) else (echo    MISSING! tools\python\python.exe not found)
"tools\python\python.exe" --version 2>&1
echo.

echo ================================================
echo  2) incoming/ folder contents (newest first):
echo ================================================
dir /b /o-d "incoming" 2>&1
echo.

echo ================================================
echo  3) Running the build NOW (live output):
echo ================================================
set "SRC="
for /f "delims=" %%f in ('dir /b /o-d "incoming\0040_qa*.xls" "incoming\0040_qa*.xlsx" 2^>nul') do if not defined SRC set "SRC=incoming\%%f"
for /f "delims=" %%f in ('dir /b /o-d "incoming\*.xls" "incoming\*.xlsx" 2^>nul') do if not defined SRC set "SRC=incoming\%%f"
if defined SRC (
    echo    Using: %SRC%
    "tools\python\python.exe" tools\update_dashboard.py "%SRC%"
) else (
    echo    incoming/ empty - letting the engine search...
    "tools\python\python.exe" tools\update_dashboard.py
)
echo.

echo ================================================
echo  4) Result on disk:
echo ================================================
dir /t:w "dist\index.html" | findstr /i "index"
dir /t:w "dist\standalone.html" | findstr /i "standalone"
"tools\python\python.exe" -c "import re,io;s=io.open('dist/js/data.js',encoding='utf-8').read();m=re.search(r'generated\D+([0-9-]+\s+[0-9:]+)',s);f=re.search(r'fingerprint\D+([a-f0-9]+)',s);print('    index (data.js) built at:',m.group(1) if m else '?',' | fingerprint:',f.group(1) if f else '?')" 2>&1
echo.
echo ================================================
echo  Done. Send a photo of this window (or copy the text).
echo ================================================
pause
