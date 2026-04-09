# MyISP Internal Tools - Firewall Configuration Script
# Run this script as Administrator to allow network access

Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║   MyISP Internal Tools - Firewall Configuration             ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "❌ ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host ""
    Write-Host "To run as Administrator:" -ForegroundColor Yellow
    Write-Host "1. Right-click on PowerShell" -ForegroundColor Yellow
    Write-Host "2. Select 'Run as Administrator'" -ForegroundColor Yellow
    Write-Host "3. Navigate to: C:\Users\vishnu.ramalingam\MyISP_Tools" -ForegroundColor Yellow
    Write-Host "4. Run: .\Enable-NetworkAccess.ps1" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit
}

Write-Host "✓ Running with Administrator privileges" -ForegroundColor Green
Write-Host ""
Write-Host "Adding firewall rule for port 8080..." -ForegroundColor Yellow

try {
    # Remove existing rule if it exists
    $existingRule = Get-NetFirewallRule -DisplayName "MyISP Internal Tools Server" -ErrorAction SilentlyContinue
    if ($existingRule) {
        Write-Host "Removing existing firewall rule..." -ForegroundColor Yellow
        Remove-NetFirewallRule -DisplayName "MyISP Internal Tools Server"
    }

    # Create new firewall rule
    New-NetFirewallRule -DisplayName "MyISP Internal Tools Server" `
                        -Direction Inbound `
                        -LocalPort 8080 `
                        -Protocol TCP `
                        -Action Allow `
                        -Profile Domain,Private,Public `
                        -Description "Allows inbound HTTP traffic for MyISP Internal Tools website on port 8080" | Out-Null

    Write-Host "✓ Firewall rule created successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "╔══════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║             CONFIGURATION COMPLETE!                          ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host ""
    Write-Host "Your team can now access the website at:" -ForegroundColor Cyan
    Write-Host "http://192.168.1.2:8080" -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host ""
    
    # Verify the rule
    $rule = Get-NetFirewallRule -DisplayName "MyISP Internal Tools Server"
    Write-Host "Firewall Rule Details:" -ForegroundColor Yellow
    Write-Host "  Name: $($rule.DisplayName)" -ForegroundColor White
    Write-Host "  Status: $($rule.Enabled)" -ForegroundColor White
    Write-Host "  Direction: $($rule.Direction)" -ForegroundColor White
    Write-Host "  Port: 8080" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host "❌ ERROR: Failed to create firewall rule" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
}

Write-Host "Press Enter to exit..." -ForegroundColor Gray
Read-Host
