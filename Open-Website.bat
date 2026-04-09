@echo off
setlocal enabledelayedexpansion
title MyISP Internal Tools - Open Website
color 0B

:: Get current user's home directory automatically
set "TOOLS_PATH=%USERPROFILE%\MyISP_Tools\index.html"

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║          MyISP Internal Tools - Opening Website...           ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Current User: %USERNAME%
echo Tools Path: %TOOLS_PATH%
echo.
echo Opening website...

:: Open the file in default browser
start "" "%TOOLS_PATH%"

if %errorLevel% == 0 (
    echo ✓ Website opened successfully!
) else (
    echo ✗ Error opening file
    echo.
    echo Make sure the MyISP_Tools folder exists at:
    echo %USERPROFILE%\MyISP_Tools
)

echo.
timeout /t 3
