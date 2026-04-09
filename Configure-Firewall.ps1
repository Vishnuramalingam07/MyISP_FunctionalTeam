# Configure Windows Firewall for PostgreSQL Network Access
# RIGHT-CLICK this file and select "Run with PowerShell as Administrator"

Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "🔥 Configuring Windows Firewall for PostgreSQL" -ForegroundColor Cyan
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "❌ ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Right-click this file and select 'Run with PowerShell as Administrator'" -ForegroundColor Yellow
    Write-Host ""
    pause
    exit 1
}

Write-Host "✓ Running with Administrator privileges" -ForegroundColor Green
Write-Host ""

# Remove existing rule if it exists
Write-Host "Checking for existing PostgreSQL firewall rules..." -ForegroundColor Yellow
$existingRule = Get-NetFirewallRule -DisplayName "PostgreSQL Database Server" -ErrorAction SilentlyContinue

if ($existingRule) {
    Write-Host "  Removing existing rule..." -ForegroundColor Yellow
    Remove-NetFirewallRule -DisplayName "PostgreSQL Database Server"
    Write-Host "  ✓ Existing rule removed" -ForegroundColor Green
}

# Create new firewall rule
Write-Host ""
Write-Host "Creating new firewall rule..." -ForegroundColor Yellow
try {
    New-NetFirewallRule `
        -DisplayName "PostgreSQL Database Server" `
        -Direction Inbound `
        -Protocol TCP `
        -LocalPort 5432 `
        -Action Allow `
        -Profile Private,Domain `
        -Description "Allow PostgreSQL connections from local network for MyISP Tools" | Out-Null
    
    Write-Host "  ✓ Firewall rule created successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "PostgreSQL port 5432 is now open for network connections" -ForegroundColor Green
    
} catch {
    Write-Host "  ❌ Failed to create firewall rule: $_" -ForegroundColor Red
    pause
    exit 1
}

Write-Host ""
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host "✅ Firewall configuration complete!" -ForegroundColor Green
Write-Host "================================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Your team can now connect to PostgreSQL at:" -ForegroundColor White
Write-Host "  Host: $env:COMPUTERNAME or 192.168.1.2" -ForegroundColor Cyan
Write-Host "  Port: 5432" -ForegroundColor Cyan
Write-Host ""
pause
