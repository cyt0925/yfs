@echo off
REM ASCII-only on purpose: Chinese characters break cmd.exe batch parsing.
REM See README.md for the Chinese instructions.
cd /d "%~dp0"
title Coupang OMS

echo ============================================================
echo   Coupang OMS  /  Coupang Order Management System
echo ============================================================
echo.
echo   Starting... your browser will open in a few seconds.
echo.
echo   [!] Do NOT close this window - the system stops if you do.
echo       To stop: press Ctrl-C, or just close this window.
echo.
echo ============================================================
echo.

start "" /min cmd /c "timeout /t 4 >nul && start "" http://127.0.0.1:5000"

python app.py

echo.
echo System stopped. Press any key to close.
pause >nul
