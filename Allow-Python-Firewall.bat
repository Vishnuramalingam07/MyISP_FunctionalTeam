@echo off
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   Allow Python Through Firewall - One-Time Setup            ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo This will allow Python to accept network connections.
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
echo Adding Python to Windows Firewall allowed programs...
echo.

:: Add Python 3.12 to firewall
netsh advfirewall firewall add rule name="Python 3.12 - MyISP Tools" dir=in action=allow program="C:\Users\vishnu.ramalingam\AppData\Local\Programs\Python\Python312\python.exe" profile=any description="Allow Python web server for MyISP Internal Tools"

if %errorLevel% == 0 (
    echo ✓ Python 3.12 added to firewall
) else (
    echo ⚠ Could not add Python 3.12
)

:: Add Python 3.13 to firewall
netsh advfirewall firewall add rule name="Python 3.13 - MyISP Tools" dir=in action=allow program="C:\Users\vishnu.ramalingam\AppData\Local\Programs\Python\Python313\python.exe" profile=any description="Allow Python web server for MyISP Internal Tools"

if %errorLevel% == 0 (
    echo ✓ Python 3.13 added to firewall
) else (
    echo ⚠ Could not add Python 3.13
)

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║             ✓ FIREWALL CONFIGURATION COMPLETE!              ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Python is now allowed through Windows Firewall.
echo Your team can access the website at:
echo.
echo    http://192.168.1.2:8000
echo.
echo IMPORTANT: The server must be running (Start-Server.bat)
echo.
pause
