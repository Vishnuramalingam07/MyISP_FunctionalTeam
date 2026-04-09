@echo off
setlocal enabledelayedexpansion
title MyISP Tools - Create Network Share
color 0B
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║   MyISP Tools - Network Share Setup (NO FIREWALL NEEDED!)   ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo This method allows team access WITHOUT needing a web server
echo or dealing with firewall issues!
echo.
echo What this does:
echo  1. Shares the MyISP_Tools folder on the network
echo  2. Team members access files directly via network path
echo  3. No firewall, no ports, no configuration needed!
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

:: Start file sharing service
echo Starting File Sharing service...
net start LanmanServer >nul 2>&1
if %errorLevel% == 0 (
    echo ✓ File sharing service started
) else (
    echo ✓ File sharing service already running
)
echo.

:: Get computer name
for /f "tokens=*" %%i in ('hostname') do set COMPUTER_NAME=%%i

:: Share the folder
echo Creating network share...
net share MyISP_Tools="C:\Users\vishnu.ramalingam\MyISP_Tools" /GRANT:Everyone,READ >nul 2>&1

if %errorLevel% == 0 (
    echo ✓ Share created successfully!
) else (
    echo Share already exists or creating new share...
    net share MyISP_Tools /DELETE >nul 2>&1
    net share MyISP_Tools="C:\Users\vishnu.ramalingam\MyISP_Tools" /GRANT:Everyone,READ
)

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║             ✓ NETWORK SHARE CREATED!                        ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Share this path with your team:
echo.
echo    \\%COMPUTER_NAME%\MyISP_Tools\index.html
echo.
echo Team members should:
echo  1. Open File Explorer (Windows + E)
echo  2. Type in address bar: \\%COMPUTER_NAME%\MyISP_Tools\index.html
echo  3. Press Enter
echo  4. The website will open in their browser!
echo.
echo ✓ No firewall issues
echo ✓ No web server needed
echo ✓ Works immediately
echo.
echo Creating shortcut file for easy access...

:: Create a URL shortcut
echo [InternetShortcut] > "MyISP_Tools_Shortcut.url"
echo URL=file://\\%COMPUTER_NAME%\MyISP_Tools\index.html >> "MyISP_Tools_Shortcut.url"

echo ✓ Created MyISP_Tools_Shortcut.url
echo.
echo You can email this .url file to team members!
echo They just double-click it to access the tools.
echo.
pause
