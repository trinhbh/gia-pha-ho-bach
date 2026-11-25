@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo   KHOI DONG WEB QUAN TRI GIA PHA HO BACH
echo   Truy cap trinh duyet tai: http://127.0.0.1:8000
echo ============================================================

start "" http://127.0.0.1:8000
python -m uvicorn app:app --host 127.0.0.1 --port 8000
if %ERRORLEVEL% NEQ 0 (
    py -m uvicorn app:app --host 127.0.0.1 --port 8000
)

pause
