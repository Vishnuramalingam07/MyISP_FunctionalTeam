# Start PostgreSQL and MyISP Tools Application
# Run this script to start everything

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "🚀 Starting MyISP Tools with PostgreSQL" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""

# Start PostgreSQL
Write-Host "1. Starting PostgreSQL database..." -ForegroundColor Yellow
$pgCtl = "C:\Program Files\PostgreSQL\17\bin\pg_ctl.exe"
$pgData = "C:\Program Files\PostgreSQL\17\data"

if (Test-Path $pgCtl) {
    & $pgCtl -D $pgData status | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "   ✓ PostgreSQL is already running" -ForegroundColor Green
    } else {
        & $pgCtl -D $pgData -l "$pgData\logfile.log" start | Out-Null
        Start-Sleep -Seconds 2
        Write-Host "   ✓ PostgreSQL started successfully" -ForegroundColor Green
    }
} else {
    Write-Host "   ✗ PostgreSQL not found at: $pgCtl" -ForegroundColor Red
    Write-Host "   Please check your PostgreSQL installation" -ForegroundColor Red
    pause
    exit 1
}

Write-Host ""  
Write-Host "2. Starting Flask application..." -ForegroundColor Yellow
Write-Host ""

# Start the Flask app
.\.venv\Scripts\python.exe app.py
