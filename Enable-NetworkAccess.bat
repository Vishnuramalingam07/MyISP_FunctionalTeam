@echo off
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   MyISP Internal Tools - Quick Firewall Setup               ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo This will add a firewall rule to allow network access.
echo.
echo Requesting Administrator privileges...
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% == 0 (
    echo ✓ Running as Administrator
    echo.
    goto :runCommand
) else (
    echo Elevating privileges...
    powershell -Command "Start-Process cmd -ArgumentList '/c cd /d %~dp0 && %~nx0' -Verb RunAs"
    exit /b
)

:runCommand
echo Adding firewall rule for port 8080...
netsh advfirewall firewall add rule name="MyISP Internal Tools Server" dir=in action=allow protocol=TCP localport=8080 profile=any description="Allows access to MyISP Internal Tools on port 8080"

if %errorLevel% == 0 (
    echo.
    echo ╔══════════════════════════════════════════════════════════════╗
    echo ║             ✓ FIREWALL CONFIGURED SUCCESSFULLY!             ║
    echo ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo Your team can now access the website at:
    echo    http://192.168.1.2:8080
    echo.
) else (
    echo.
    echo ❌ Failed to add firewall rule
    echo Please run this file as Administrator
    echo.
)

pause
