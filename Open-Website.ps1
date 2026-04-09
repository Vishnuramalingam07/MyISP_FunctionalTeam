# MyISP Internal Tools - Dynamic Path Opener
# Works for any user on any computer

$currentUser = $env:USERNAME
$toolsPath = Join-Path $env:USERPROFILE "MyISP_Tools\index.html"
$computerName = $env:COMPUTERNAME

Write-Host "`n╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║          MyISP Internal Tools - Quick Access                 ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

Write-Host "Current User:    " -NoNewline -ForegroundColor White
Write-Host $currentUser -ForegroundColor Green

Write-Host "Computer Name:   " -NoNewline -ForegroundColor White
Write-Host $computerName -ForegroundColor Green

Write-Host "Tools Path:      " -NoNewline -ForegroundColor White
Write-Host $toolsPath -ForegroundColor Yellow

Write-Host "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n" -ForegroundColor Gray

# Check if file exists
if (Test-Path $toolsPath) {
    Write-Host "✓ File found! Opening website..." -ForegroundColor Green
    Start-Process $toolsPath
    Write-Host "`n✓ Website opened in your default browser!`n" -ForegroundColor Green
} else {
    Write-Host "✗ Error: File not found!" -ForegroundColor Red
    Write-Host "`nExpected location: $toolsPath" -ForegroundColor Yellow
    Write-Host "`nMake sure MyISP_Tools folder exists in your user directory.`n" -ForegroundColor Yellow
}

Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
