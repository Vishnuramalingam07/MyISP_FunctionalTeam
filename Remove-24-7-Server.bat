@echo off
title Stop 24/7 Server - MyISP Tools
color 0C
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║        Removing 24/7 Auto-Start Server - MyISP Tools        ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Delete scheduled task
schtasks /delete /tn "MyISP_Tools_Server" /f 2>nul

:: Kill any running Python server processes
taskkill /F /FI "WINDOWTITLE eq MyISP*" 2>nul
taskkill /F /IM python.exe /FI "MEMUSAGE gt 30000" 2>nul

:: Clean up helper files
if exist "%~dp0\RunServer-Hidden.vbs" del "%~dp0\RunServer-Hidden.vbs"
if exist "%~dp0\Start-Server-Background.bat" del "%~dp0\Start-Server-Background.bat"

echo.
echo ✓ 24/7 server auto-start removed
echo ✓ Server process stopped
echo ✓ Helper files cleaned up
echo.
echo The server will no longer start automatically.
echo.
pause
