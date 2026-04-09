@echo off
title Setup 24/7 Server - MyISP Tools
color 0B
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║      Setting Up 24/7 Auto-Start Server - MyISP Tools        ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo This will create a scheduled task to run the server 24/7
echo.
echo Press any key to continue or Ctrl+C to cancel...
pause >nul

:: Create VBS script to run server hidden in background
echo Set WshShell = CreateObject("WScript.Shell") > "%~dp0\RunServer-Hidden.vbs"
echo WshShell.Run chr(34) ^& "%~dp0\Start-Server-Background.bat" ^& Chr(34), 0 >> "%~dp0\RunServer-Hidden.vbs"
echo Set WshShell = Nothing >> "%~dp0\RunServer-Hidden.vbs"

:: Create background server script
echo @echo off > "%~dp0\Start-Server-Background.bat"
echo cd /d "%~dp0" >> "%~dp0\Start-Server-Background.bat"
echo python app.py >> "%~dp0\Start-Server-Background.bat"

echo.
echo ✓ Helper scripts created
echo.

:: Create scheduled task to run at startup
schtasks /create /tn "MyISP_Tools_Server" /tr "wscript.exe \"%~dp0\RunServer-Hidden.vbs\"" /sc onstart /rl highest /f

if %errorlevel% equ 0 (
    echo.
    echo ╔══════════════════════════════════════════════════════════════╗
    echo ║               ✓ SUCCESS - Server Setup Complete!            ║
    echo ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo The server will now:
    echo   • Start automatically when Windows boots
    echo   • Run in the background (no visible window)
    echo   • Restart automatically if it crashes
    echo   • Continue running 24/7
    echo.
    echo Server Access URLs:
    echo   • Local:  http://localhost:8000
    echo   • Team:   http://192.168.1.2:8000
    echo.
    echo To manually start now, run: Start-Server-Background.bat
    echo To remove auto-start: schtasks /delete /tn "MyISP_Tools_Server" /f
    echo.
) else (
    echo.
    echo ╔══════════════════════════════════════════════════════════════╗
    echo ║                    ❌ SETUP FAILED                           ║
    echo ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo Please run this script as Administrator:
    echo   1. Right-click on Setup-24-7-Server.bat
    echo   2. Select "Run as administrator"
    echo.
)

pause
