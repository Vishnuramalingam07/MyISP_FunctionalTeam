@echo off
title Temporarily Disable Firewall for Testing
color 0C
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   ⚠️  WARNING: TEMPORARY FIREWALL DISABLE FOR TESTING       ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo This will TEMPORARILY disable Windows Firewall to test if
echo that's what's blocking access to the website.
echo.
echo ⚠️  SECURITY WARNING: Only use this for testing!
echo    Remember to ENABLE IT AGAIN after testing.
echo.
pause
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c cd /d %~dp0 && %~nx0' -Verb RunAs"
    exit /b
)

echo ✓ Running as Administrator
echo.
echo Disabling Windows Firewall for Private network...
netsh advfirewall set privateprofile state off

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║  Firewall DISABLED for Private Network                      ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Now try accessing from another computer:
echo    http://192.168.1.2:8000
echo.
echo If it works now, the firewall was the problem.
echo.
echo ⚠️  IMPORTANT: Run "Enable-Firewall.bat" when done testing!
echo.
pause
