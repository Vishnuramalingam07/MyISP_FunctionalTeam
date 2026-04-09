@echo off
title Re-Enable Windows Firewall
color 0A
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║         Re-Enabling Windows Firewall                         ║
echo ╚══════════════════════════════════════════════════════════════╝
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
echo Enabling Windows Firewall for all profiles...
netsh advfirewall set allprofiles state on

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║  ✓ Firewall RE-ENABLED                                       ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Windows Firewall is now back on and protecting your computer.
echo.
pause
