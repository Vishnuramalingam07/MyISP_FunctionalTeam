@echo off
setlocal enabledelayedexpansion
title MyISP Internal Tools - Create Desktop Shortcut
color 0B

:: Get current user automatically
set "TOOLS_PATH=%USERPROFILE%\MyISP_Tools\index.html"
set "SHORTCUT_PATH=%USERPROFILE%\Desktop\MyISP Internal Tools.url"

echo.
echo ╔══════════════════════════════════════════════════════════════╗
echo ║      MyISP Internal Tools - Create Desktop Shortcut         ║
echo ╚══════════════════════════════════════════════════════════════╝
echo.
echo Current User: %USERNAME%
echo Creating shortcut to: %TOOLS_PATH%
echo.

:: Create URL shortcut on desktop
(
echo [InternetShortcut]
echo URL=file:///%TOOLS_PATH:\=/%
echo IconIndex=0
) > "%SHORTCUT_PATH%"

if exist "%SHORTCUT_PATH%" (
    echo ╔══════════════════════════════════════════════════════════════╗
    echo ║              ✓ SHORTCUT CREATED SUCCESSFULLY!               ║
    echo ╚══════════════════════════════════════════════════════════════╝
    echo.
    echo Location: %USERPROFILE%\Desktop
    echo Name: MyISP Internal Tools.url
    echo.
    echo ✓ Double-click the shortcut on your desktop to open the tools!
) else (
    echo ✗ Error creating shortcut
)

echo.
pause
