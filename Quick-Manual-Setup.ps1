# Quick Manual 24/7 Server Setup - Run as Administrator
# Right-click this file and select "Run with PowerShell"

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║      Quick Manual 24/7 Server Setup - MyISP Tools           ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

$scriptPath = $PSScriptRoot
$pythonPath = "$scriptPath\.venv\Scripts\python.exe"
$appPath = "$scriptPath\app.py"

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "❌ This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "`nPlease:" -ForegroundColor Yellow
    Write-Host "  1. Right-click on Quick-Manual-Setup.ps1" -ForegroundColor Yellow
    Write-Host "  2. Select 'Run with PowerShell'" -ForegroundColor Yellow
    Write-Host "  3. Click 'Yes' on the UAC prompt`n" -ForegroundColor Yellow
    Pause
    exit
}

Write-Host "✓ Running as Administrator" -ForegroundColor Green

# Create VBS helper to run server hidden
$vbsContent = @"
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "$scriptPath\Start-Server-Background.bat" & Chr(34), 0
Set WshShell = Nothing
"@

$vbsPath = "$scriptPath\RunServer-Hidden.vbs"
$vbsContent | Out-File -FilePath $vbsPath -Encoding ASCII -Force
Write-Host "✓ Created VBS helper script" -ForegroundColor Green

# Create background batch file
$batContent = @"
@echo off
cd /d "$scriptPath"
"$pythonPath" app.py
"@

$batPath = "$scriptPath\Start-Server-Background.bat"
$batContent | Out-File -FilePath $batPath -Encoding ASCII -Force
Write-Host "✓ Created background server script" -ForegroundColor Green

# Create scheduled task
$taskName = "MyISP_Tools_Server"
$taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

if ($taskExists) {
    Write-Host "`n⚠️  Task already exists. Removing old task..." -ForegroundColor Yellow
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"$vbsPath`""
$trigger = New-ScheduledTaskTrigger -AtStartup
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "MyISP Internal Tools Flask Server - 24/7 Auto-Start" | Out-Null

Write-Host "✓ Scheduled task created successfully" -ForegroundColor Green

# Start the task immediately
Write-Host "`n⏳ Starting server now..." -ForegroundColor Yellow
Start-ScheduledTask -TaskName $taskName
Start-Sleep -Seconds 3

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║               ✓ SUCCESS - Server Setup Complete!            ║" -ForegroundColor Green
Write-Host "╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Green

Write-Host "The server is now running 24/7!" -ForegroundColor Cyan
Write-Host ""
Write-Host "✓ Starts automatically when Windows boots" -ForegroundColor White
Write-Host "✓ Runs in the background (no visible window)" -ForegroundColor White
Write-Host "✓ Restarts automatically if it crashes" -ForegroundColor White
Write-Host "✓ Continues running 24/7" -ForegroundColor White
Write-Host ""
Write-Host "Server Access URLs:" -ForegroundColor Cyan
Write-Host "  • Local:  http://localhost:8000" -ForegroundColor White
Write-Host "  • Team:   http://192.168.1.2:8000" -ForegroundColor White
Write-Host ""
Write-Host "To check if server is running:" -ForegroundColor Yellow
Write-Host "  curl http://localhost:8000" -ForegroundColor Gray
Write-Host ""
Write-Host "To stop the server:" -ForegroundColor Yellow
Write-Host "  Run Remove-24-7-Server.bat as Administrator" -ForegroundColor Gray
Write-Host ""

# Test server
Write-Host "Testing server..." -ForegroundColor Yellow
Start-Sleep -Seconds 2
try {
    $response = Invoke-WebRequest -Uri "http://localhost:8000" -TimeoutSec 5 -UseBasicParsing
    if ($response.StatusCode -eq 200) {
        Write-Host "✅ Server is responding! Setup successful!`n" -ForegroundColor Green
    }
} catch {
    Write-Host "⚠️  Server might still be starting up. Please wait a moment and try:" -ForegroundColor Yellow
    Write-Host "    curl http://localhost:8000`n" -ForegroundColor Gray
}

Pause
