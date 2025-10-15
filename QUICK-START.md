# 🚀 How to Run Azure AD Sign-in Checker

## Easiest Methods (Pick One):

### Method 1: Double-Click the Batch File ⭐ **RECOMMENDED**
1. Double-click `Run-SignInChecker.bat`
2. The application will start automatically

### Method 2: Use Desktop Shortcut
1. First, run `Create-DesktopShortcut.ps1` (right-click → "Run with PowerShell")
2. Double-click the new "Azure AD Sign-in Checker" shortcut on your desktop

### Method 3: Right-Click the Main Script
1. Right-click `Check-UserSignIns-GUI.ps1`
2. Select "Run with PowerShell"

### Method 4: Use the Launcher Script
1. Right-click `Launch-SignInChecker.ps1` 
2. Select "Run with PowerShell"

## If You Get "Execution Policy" Errors:

The script will automatically detect and offer to fix execution policy issues. If you prefer to fix it manually:

**Run PowerShell as Administrator and execute:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
```

## Prerequisites:
- ✅ Windows PowerShell 5.0 or later
- ✅ Azure AD Premium P1 or P2 license  
- ✅ Reports Reader role (or higher) in Azure AD
- ✅ Internet connection for module installation

## First-Time Setup:
The script will automatically:
1. Check for required PowerShell modules
2. Prompt to install missing modules
3. Guide you through Azure AD authentication
4. Launch the GUI interface

## Files in this Package:
- `Check-UserSignIns-GUI.ps1` - Main application
- `Run-SignInChecker.bat` - Easy launcher (recommended)
- `Launch-SignInChecker.ps1` - PowerShell launcher with policy handling
- `Create-DesktopShortcut.ps1` - Creates desktop shortcut
- `Test-GraphPermissions.ps1` - Permission testing utility
- `sample-input.csv` - Example CSV format
- `README.md` - This instructions file

## Need Help?
- Check the Details column in output CSV for error information
- Ensure you have Reports Reader role in Azure AD
- Verify Azure AD Premium license is assigned
- Contact your IT administrator for permission issues

## Security Note:
This script requires privileged permissions to read user sign-in data. It uses Microsoft Graph API with standard Microsoft authentication.