@echo off
setlocal
cd /d "%~dp0"

python sync_from_excel.py
if %ERRORLEVEL% NEQ 0 (
    py sync_from_excel.py
)

echo.
pause
