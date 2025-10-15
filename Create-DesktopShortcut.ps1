# Create Desktop Shortcut for Azure AD Sign-in Checker
# Run this script once to create a desktop shortcut

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatchFile = Join-Path $ScriptDir "Run-SignInChecker.bat"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $DesktopPath "Azure AD Sign-in Checker.lnk"

if (-not (Test-Path $BatchFile)) {
    Write-Error "Batch launcher not found: $BatchFile"
    Write-Host "Please ensure Run-SignInChecker.bat exists in the same directory."
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    # Create shortcut
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($ShortcutPath)
    $Shortcut.TargetPath = $BatchFile
    $Shortcut.WorkingDirectory = $ScriptDir
    $Shortcut.Description = "Azure AD Sign-in Checker - Check user sign-in activity"
    $Shortcut.IconLocation = "powershell.exe,0"
    $Shortcut.Save()
    
    Write-Host "✅ Desktop shortcut created successfully!" -ForegroundColor Green
    Write-Host "Shortcut location: $ShortcutPath" -ForegroundColor Gray
    Write-Host ""
    Write-Host "You can now double-click the 'Azure AD Sign-in Checker' shortcut on your desktop to run the application." -ForegroundColor Cyan
}
catch {
    Write-Error "Failed to create desktop shortcut: $($_.Exception.Message)"
}

Read-Host "Press Enter to close"