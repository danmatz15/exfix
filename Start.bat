@echo off
cd /d "%~dp0"
start "EXfix" powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Serve-EXfix.ps1"
timeout /t 2 /nobreak >nul
start http://127.0.0.1:8765/
