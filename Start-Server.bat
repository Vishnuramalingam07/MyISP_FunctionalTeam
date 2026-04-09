@echo off
title MyISP Internal Tools Server
color 0A
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║          MyISP Internal Tools - Starting Server...           ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Server will start on: http://192.168.1.2:8000
echo.
echo ✓ Flask server with report generation support
echo ✓ Team can access from any browser
echo ✓ Regression reports can be generated
echo.
echo ⚠️  IMPORTANT: Keep this window open while team is using the site!
echo.
echo ───────────────────────────────────────────────────────────────
echo.
echo Installing Flask if needed...
python -m pip install flask --quiet 2>nul
echo.

cd /d "%~dp0"
python app.py

echo.
echo Server stopped.
pause
