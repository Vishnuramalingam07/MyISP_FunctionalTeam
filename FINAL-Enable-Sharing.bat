@echo off
title Enable File Sharing - Run as Administrator
color 0B
echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║        Enable Windows File Sharing for Network Access       ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo This script MUST be run as Administrator!
    echo.
    echo Right-click this file and select "Run as Administrator"
    echo.
    pause
    exit /b 1
)

echo ✓ Running as Administrator
echo.

echo Step 1: Starting File Sharing Service...
sc config LanmanServer start= auto
net start LanmanServer

if %errorLevel% == 0 (
    echo ✓ File sharing service started successfully
) else (
    echo ℹ File sharing service already running
)
echo.

echo Step 2: Creating Network Share...
:: Delete existing share if present
net share MyISP_Tools /DELETE >nul 2>&1

:: Create new share
net share MyISP_Tools="C:\Users\vishnu.ramalingam\MyISP_Tools" /GRANT:Everyone,READ

if %errorLevel% == 0 (
    echo ✓ Network share created successfully!
    echo.
    for /f "tokens=*" %%i in ('hostname') do set COMPUTER_NAME=%%i
    echo ╔══════════════════════════════════════════════════════════════╗
    echo ║                  ✓ SETUP COMPLETE!                           ║
    echo ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo Share this path with your team:
    echo.
    echo    \\!COMPUTER_NAME!\MyISP_Tools\index.html
    echo.
    echo They can access it by:
    echo  1. Press Windows + R
    echo  2. Type that path
    echo  3. Press Enter
    echo.
) else (
    echo ✗ Error creating share
    echo.
    echo This might be because:
    echo  - File sharing is disabled by group policy
    echo  - Network discovery is off
    echo  - Antivirus is blocking it
    echo.
)

echo.
pause
