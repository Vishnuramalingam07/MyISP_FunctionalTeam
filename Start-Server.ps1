# MyISP Internal Tools - PowerShell Server Launcher
# Double-click this file to start the web server

Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║          MyISP Internal Tools - Starting Server...           ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "Server starting on: http://192.168.1.2:8000" -ForegroundColor Green
Write-Host ""
Write-Host "✓ Flask server with report generation support" -ForegroundColor Green
Write-Host "✓ Team can access from any browser" -ForegroundColor Green
Write-Host "✓ Regression reports can be generated" -ForegroundColor Green
Write-Host ""
Write-Host "⚠️  IMPORTANT: Keep this window open while team is using the site!" -ForegroundColor Yellow
Write-Host ""
Write-Host "───────────────────────────────────────────────────────────────" -ForegroundColor Gray
Write-Host ""
Write-Host "Installing Flask if needed..." -ForegroundColor Yellow
python -m pip install flask --quiet 2>$null
Write-Host ""
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Gray
Write-Host ""

Set-Location $PSScriptRoot
python app.py
