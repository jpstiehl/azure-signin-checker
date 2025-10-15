# Azure AD Sign-in Checker Launcher
# This script ensures proper execution environment and launches the GUI

param(
    [switch]$Force,
    [switch]$Help
)

if ($Help) {
    Write-Host @"
Azure AD Sign-in Checker Launcher

Usage: .\Launch-SignInChecker.ps1 [-Force] [-Help]

Options:
  -Force    Skip execution policy checks
  -Help     Show this help message

This launcher script will:
1. Check execution policy and adjust if needed
2. Verify required modules are available
3. Launch the GUI application
4. Handle any startup errors gracefully
"@ -ForegroundColor Green
    exit 0
}

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$GuiScript = Join-Path $ScriptDir "Check-UserSignIns-GUI.ps1"

# Verify GUI script exists
if (-not (Test-Path $GuiScript)) {
    Write-Error "GUI script not found: $GuiScript"
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Azure AD Sign-in Checker Launcher" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Check execution policy
$currentPolicy = Get-ExecutionPolicy
if ($currentPolicy -eq 'Restricted' -and -not $Force) {
    Write-Host "Current execution policy: $currentPolicy" -ForegroundColor Yellow
    Write-Host "This may prevent the script from running properly." -ForegroundColor Yellow
    Write-Host ""
    $response = Read-Host "Would you like to temporarily set execution policy to RemoteSigned? (Y/N)"
    
    if ($response -match '^[Yy]') {
        try {
            Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force
            Write-Host "✅ Execution policy set to RemoteSigned for this session" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to set execution policy. Running with -ExecutionPolicy Bypass instead."
            & powershell.exe -ExecutionPolicy Bypass -File $GuiScript
            exit $LASTEXITCODE
        }
    }
}

Write-Host "Launching Azure AD Sign-in Checker GUI..." -ForegroundColor Green
Write-Host "Script location: $GuiScript" -ForegroundColor Gray
Write-Host ""

try {
    # Launch the GUI script
    & $GuiScript
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Warning "Script exited with code: $LASTEXITCODE"
        Read-Host "Press Enter to close"
    }
}
catch {
    Write-Host ""
    Write-Error "Error launching script: $($_.Exception.Message)"
    Write-Host "Full error details:" -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor Red
    Read-Host "Press Enter to close"
}