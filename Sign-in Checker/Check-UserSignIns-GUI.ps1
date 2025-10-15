# Azure AD Sign-in Checker Script with GUI
# This script provides a GUI to either read email addresses from a CSV file or from a Microsoft 365 Group,
# checks their last sign-in time in Azure AD, and exports the results to a new CSV file with customizable day analysis.
#
# QUICK START:
# 1. Right-click this file and select "Run with PowerShell" OR
# 2. Double-click "Run-SignInChecker.bat" for easier execution OR  
# 3. Use the desktop shortcut (run Create-DesktopShortcut.ps1 first)
#
# Required: Azure AD Premium P1/P2 license and Reports Reader role or higher

# Check if running with restricted execution policy and offer to fix it
$currentPolicy = Get-ExecutionPolicy
if ($currentPolicy -eq 'Restricted') {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $policyMessage = @"
PowerShell Execution Policy Restriction Detected

Current policy: $currentPolicy

This script cannot run with the current execution policy. 

Options:
• Click 'Yes' to run with bypassed policy (recommended)
• Click 'No' to exit and manually change policy
• Click 'Cancel' for more information

To permanently fix this, run as Administrator:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
"@
    
    $result = [System.Windows.Forms.MessageBox]::Show($policyMessage, "Execution Policy Issue", [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Restart with bypass policy
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process powershell.exe -ArgumentList "-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"" -WindowStyle Hidden
        exit 0
    }
    elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        $infoMessage = @"
PowerShell Execution Policy Information

The execution policy is a security feature that controls script execution.

To fix this permanently (run as Administrator):
• Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine

Alternative ways to run this script:
• Use the Run-SignInChecker.bat file
• Run: powershell.exe -ExecutionPolicy Bypass -File "Check-UserSignIns-GUI.ps1"
• Right-click the .ps1 file and select "Run with PowerShell"

For more info: Get-Help about_Execution_Policies
"@
        [System.Windows.Forms.MessageBox]::Show($infoMessage, "Execution Policy Help", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    
    exit 0
}

# Set error handling preferences
$ErrorActionPreference = "Continue"  # GUI scripts need different error handling
Set-StrictMode -Version Latest

# Clean up any existing Microsoft Graph modules to prevent version conflicts
Write-Host "Cleaning existing Graph modules..." -ForegroundColor Yellow
$existingGraphModules = Get-Module -Name "Microsoft.Graph.*" -ErrorAction SilentlyContinue
if ($existingGraphModules) {
    $existingGraphModules | Remove-Module -Force -ErrorAction SilentlyContinue
    Write-Host "Removed $($existingGraphModules.Count) existing Graph modules" -ForegroundColor Green
}

# Global error handler for GUI
trap {
    $errorMessage = "Critical Error: $($_.Exception.Message)`n`nStack Trace: $($_.ScriptStackTrace)"
    [System.Windows.Forms.MessageBox]::Show($errorMessage, "Critical Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    
    # Cleanup
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    } catch { }
    
    exit 1
}

# Check if running in appropriate environment
if ([Environment]::UserInteractive -eq $false) {
    Write-Error "This script requires an interactive user session to display the GUI."
    exit 1
}

# Load Windows Forms assemblies with error handling
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    Write-Host "✅ Windows Forms assemblies loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "❌ Failed to load Windows Forms assemblies. This script requires a full Windows environment."
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Check PowerShell version compatibility
if ($PSVersionTable.PSVersion.Major -lt 5) {
    [System.Windows.Forms.MessageBox]::Show("This script requires PowerShell 5.0 or later. Current version: $($PSVersionTable.PSVersion)", "Version Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit 1
}

# Function to install and import required modules with GUI-friendly error handling
function Install-RequiredModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Reports',
        'Microsoft.Graph.Groups'
    )
    
    # Check execution policy
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq 'Restricted') {
        $policyMessage = "PowerShell execution policy is set to 'Restricted'. This may prevent module installation.`n`nWould you like to continue anyway?"
        $result = [System.Windows.Forms.MessageBox]::Show($policyMessage, "Execution Policy Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            exit 0
        }
    }
    
    foreach ($module in $requiredModules) {
        # Check if module is available
        $moduleAvailable = Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue
        
        if (-not $moduleAvailable) {
            # Show installation dialog
            $installMessage = "Module '$module' is not installed. This module is required to run the script.`n`nWould you like to install it now? (This may take a few minutes)"
            $result = [System.Windows.Forms.MessageBox]::Show($installMessage, "Module Installation Required", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
            
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                [System.Windows.Forms.MessageBox]::Show("Cannot continue without required modules. Please install manually:`nInstall-Module Microsoft.Graph -Scope CurrentUser", "Installation Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                exit 0
            }
            
            # Show progress during installation
            $progressForm = New-Object System.Windows.Forms.Form
            $progressForm.Text = "Installing Modules"
            $progressForm.Size = New-Object System.Drawing.Size(400, 120)
            $progressForm.StartPosition = "CenterScreen"
            $progressForm.FormBorderStyle = "FixedDialog"
            $progressForm.MaximizeBox = $false
            $progressForm.MinimizeBox = $false
            
            $progressLabel = New-Object System.Windows.Forms.Label
            $progressLabel.Location = New-Object System.Drawing.Point(10, 20)
            $progressLabel.Size = New-Object System.Drawing.Size(370, 40)
            $progressLabel.Text = "Installing $module..."
            $progressForm.Controls.Add($progressLabel)
            
            $progressBar = New-Object System.Windows.Forms.ProgressBar
            $progressBar.Location = New-Object System.Drawing.Point(10, 50)
            $progressBar.Size = New-Object System.Drawing.Size(370, 20)
            $progressBar.Style = "Marquee"
            $progressForm.Controls.Add($progressBar)
            
            $progressForm.Show()
            [System.Windows.Forms.Application]::DoEvents()
            
            try {
                # Test internet connectivity
                $testConnection = Test-NetConnection -ComputerName "www.powershellgallery.com" -Port 443 -InformationLevel Quiet -ErrorAction SilentlyContinue
                if (-not $testConnection) {
                    throw "No internet connection to PowerShell Gallery"
                }
                
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                $progressLabel.Text = "Successfully installed $module"
                Start-Sleep -Seconds 1
            }
            catch {
                $progressForm.Close()
                $errorMessage = "Failed to install module '$module'.`n`nError: $($_.Exception.Message)`n`nPlease install manually using:`nInstall-Module Microsoft.Graph -Scope CurrentUser"
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Installation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                exit 1
            }
            finally {
                $progressForm.Close()
            }
        }
        
        # Import the module with conflict resolution
        try {
            # Remove any existing loaded version to prevent conflicts
            $existingModule = Get-Module -Name $module -ErrorAction SilentlyContinue
            if ($existingModule) {
                Remove-Module -Name $module -Force -ErrorAction SilentlyContinue
            }
            
            # Import with the latest available version
            $availableVersions = Get-Module -ListAvailable -Name $module | Sort-Object Version -Descending
            if ($availableVersions) {
                $latestVersion = $availableVersions[0]
                Import-Module -Name $module -RequiredVersion $latestVersion.Version -Force -ErrorAction Stop
                Write-Verbose "✅ Loaded $module version $($latestVersion.Version)" -Verbose
            }
            else {
                Import-Module $module -Force -ErrorAction Stop
                Write-Verbose "✅ Loaded $module (no version specified)" -Verbose
            }
        }
        catch {
            # If versioned import fails, try simple import
            try {
                Import-Module $module -Force -ErrorAction Stop
                Write-Verbose "✅ Loaded $module (fallback method)" -Verbose
            }
            catch {
                $errorMessage = "Failed to import module '$module'.`n`nError: $($_.Exception.Message)`n`nThis may be due to module version conflicts. Try restarting PowerShell and running the script again."
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Module Import Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                exit 1
            }
        }
    }
    
    # Verify all modules are loaded
    foreach ($module in $requiredModules) {
        $loadedModule = Get-Module -Name $module -ErrorAction SilentlyContinue
        if (-not $loadedModule) {
            [System.Windows.Forms.MessageBox]::Show("Module '$module' failed to load properly.", "Module Verification Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            exit 1
        }
    }
}

# Function to check if a user signed in within the specified number of days
function Test-SignInWithinDays {
    param(
        [datetime]$LastSignInDate,
        [int]$Days
    )
    
    if ($null -eq $LastSignInDate -or $LastSignInDate -eq [datetime]::MinValue) {
        return $false
    }
    
    $cutoffDate = (Get-Date).AddDays(-$Days)
    return $LastSignInDate -gt $cutoffDate
}

# Function to get user sign-in information
function Get-UserSignInInfo {
    param(
        [string]$UserEmail,
        [int]$Days,
        [string]$FirstName = $null,
        [string]$LastName = $null
    )
    
    try {
        # Add debugging information for API calls
        Write-Verbose "Attempting to get user info for: $UserEmail" -Verbose
        
        # First try to find the user and check if account is enabled to avoid unnecessary API calls
        # Use Filter parameter for email addresses instead of UserId which expects GUID
        $userBasic = Get-MgUser -Filter "userPrincipalName eq '$UserEmail'" -Property "displayName,userPrincipalName,id,accountEnabled" -ErrorAction Stop
        
        if (-not $userBasic) {
            throw "User not found: $UserEmail"
        }
        
        Write-Verbose "Successfully found user: $($userBasic.DisplayName) (ID: $($userBasic.Id)) - Account Enabled: $($userBasic.AccountEnabled)" -Verbose
        
        # Check if account is disabled - skip expensive sign-in activity lookup for disabled accounts
        if ($userBasic.AccountEnabled -eq $false) {
            Write-Verbose "Account is disabled - skipping sign-in activity lookup for performance" -Verbose
            
            # Parse first and last name from display name
            $displayName = $userBasic.DisplayName
            $nameParts = if ($displayName) { $displayName.Split(' ', [StringSplitOptions]::RemoveEmptyEntries) } else { @() }
            $firstName = if ($nameParts.Count -gt 0) { $nameParts[0] } else { "Unknown" }
            $lastName = if ($nameParts.Count -gt 1) { $nameParts[-1] } else { "Unknown" }
            
            # Create result object for disabled account
            $resultObject = New-Object PSObject
            $resultObject | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $firstName
            $resultObject | Add-Member -MemberType NoteProperty -Name "LastName" -Value $lastName
            $resultObject | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $userBasic.UserPrincipalName
            $resultObject | Add-Member -MemberType NoteProperty -Name "LastSignInDateTime" -Value "Account Disabled"
            $resultObject | Add-Member -MemberType NoteProperty -Name "Status" -Value "Success"
            $resultObject | Add-Member -MemberType NoteProperty -Name "Error" -Value $null
            
            # Add disabled account details
            $resultObject | Add-Member -MemberType NoteProperty -Name "Details" -Value "Account is disabled - sign-in activity not applicable"
            
            # Add the dynamic days property - disabled accounts always show "No"
            $daysPropertyName = "SignedInLast${Days}Days"
            $resultObject | Add-Member -MemberType NoteProperty -Name $daysPropertyName -Value "No"
            
            return $resultObject
        }
        
        # Account is enabled - proceed with sign-in activity lookup using the user's GUID
        $user = Get-MgUser -UserId $userBasic.Id -Property "displayName,userPrincipalName,signInActivity" -ErrorAction Stop
        
        $lastSignInDateTime = $null
        
        if ($user.SignInActivity) {
            if ($user.SignInActivity.LastSignInDateTime) {
                $lastSignInDateTime = [datetime]$user.SignInActivity.LastSignInDateTime
            }
            elseif ($user.SignInActivity.LastNonInteractiveSignInDateTime) {
                $lastSignInDateTime = [datetime]$user.SignInActivity.LastNonInteractiveSignInDateTime
            }
        }
        
        $withinSpecifiedDays = if ($lastSignInDateTime) {
            Test-SignInWithinDays -LastSignInDate $lastSignInDateTime -Days $Days
        } else {
            $false
        }
        
        # Parse first and last name from display name
        $displayName = $user.DisplayName
        $nameParts = if ($displayName) { $displayName.Split(' ', [StringSplitOptions]::RemoveEmptyEntries) } else { @() }
        $firstName = if ($nameParts.Count -gt 0) { $nameParts[0] } else { "Unknown" }
        $lastName = if ($nameParts.Count -gt 1) { $nameParts[-1] } else { "Unknown" }
        
        # Create result object using New-Object to avoid property conflicts
        $resultObject = New-Object PSObject
        $resultObject | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $firstName
        $resultObject | Add-Member -MemberType NoteProperty -Name "LastName" -Value $lastName
        $resultObject | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $user.UserPrincipalName
        $resultObject | Add-Member -MemberType NoteProperty -Name "LastSignInDateTime" -Value $(if ($lastSignInDateTime) { $lastSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" })
        $resultObject | Add-Member -MemberType NoteProperty -Name "Status" -Value "Success"
        $resultObject | Add-Member -MemberType NoteProperty -Name "Error" -Value $null
        
        # Add success details
        $successDetails = if ($lastSignInDateTime) {
            "Successfully retrieved sign-in data - Last sign-in: $($lastSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss"))"
        } else {
            "Successfully retrieved user data - No sign-in activity found"
        }
        $resultObject | Add-Member -MemberType NoteProperty -Name "Details" -Value $successDetails
        
        # Add the dynamic days property safely
        $daysPropertyName = "SignedInLast${Days}Days"
        $resultObject | Add-Member -MemberType NoteProperty -Name $daysPropertyName -Value $(if ($withinSpecifiedDays) { "Yes" } else { "No" })
        
        return $resultObject
    }
    catch {
        # Use provided names from CSV if available, otherwise "Unknown"
        $errorFirstName = if (-not [string]::IsNullOrWhiteSpace($FirstName)) { $FirstName } else { "Unknown" }
        $errorLastName = if (-not [string]::IsNullOrWhiteSpace($LastName)) { $LastName } else { "Unknown" }
        
        # Create error result object using New-Object to avoid property conflicts
        $errorObject = New-Object PSObject
        $errorObject | Add-Member -MemberType NoteProperty -Name "FirstName" -Value $errorFirstName
        $errorObject | Add-Member -MemberType NoteProperty -Name "LastName" -Value $errorLastName
        $errorObject | Add-Member -MemberType NoteProperty -Name "EmailAddress" -Value $UserEmail
        $errorObject | Add-Member -MemberType NoteProperty -Name "LastSignInDateTime" -Value "Error"
        $errorObject | Add-Member -MemberType NoteProperty -Name "Status" -Value "Error"
        # Create detailed error message with specific diagnostics
        $errorType = "Unknown"
        $specificError = $_.Exception.Message
        
        if ($specificError -match "Resource '.*' does not exist|Request_ResourceNotFound|User not found") {
            $errorType = "UserNotFound"
            $specificError = "User not found in Azure AD"
        }
        elseif ($specificError -match "Insufficient privileges|Forbidden|403") {
            $errorType = "InsufficientPermissions" 
            $specificError = "Insufficient permissions to read user or sign-in data"
        }
        elseif ($specificError -match "signInActivity|AuditLog") {
            $errorType = "SignInPermissions"
            $specificError = "Missing AuditLog.Read.All permission for sign-in activity"
        }
        elseif ($specificError -match "Unauthorized|401") {
            $errorType = "Unauthorized"
            $specificError = "Authentication token invalid or expired"
        }
        
        $detailedError = "[$errorType] $UserEmail`: $specificError"
        Write-Verbose "Error details: $detailedError" -Verbose
        $errorObject | Add-Member -MemberType NoteProperty -Name "Error" -Value $detailedError
        
        # Add error details to Details column
        $errorObject | Add-Member -MemberType NoteProperty -Name "Details" -Value "Error details: $detailedError"
        
        # Add the dynamic days property safely
        $daysPropertyName = "SignedInLast${Days}Days"
        $errorObject | Add-Member -MemberType NoteProperty -Name $daysPropertyName -Value "No"
        
        return $errorObject
    }
}

# Function to get users from Microsoft 365 Group with enhanced error handling
function Get-M365GroupMembers {
    param([string]$GroupId)
    
    if ([string]::IsNullOrWhiteSpace($GroupId)) {
        throw "Group ID cannot be null or empty"
    }
    
    try {
        # Get group members with retry logic
        $maxRetries = 3
        $members = $null
        
        for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
            try {
                $members = Get-MgGroupMember -GroupId $GroupId -All -ErrorAction Stop
                break
            }
            catch {
                if ($attempt -eq $maxRetries) {
                    throw
                }
                Start-Sleep -Seconds 2
            }
        }
        
        if (-not $members) {
            throw "No members found in group or group is empty"
        }
        
        $userEmails = @()
        $memberCount = 0
        
        foreach ($member in $members) {
            $memberCount++
            
            try {
                # Check if member is a user (not a group or other object type)
                if ($member.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                    $user = Get-MgUser -UserId $member.Id -Property "userPrincipalName,accountEnabled" -ErrorAction Stop
                    
                    if ($user -and $user.UserPrincipalName) {
                        # Only include enabled users
                        if ($user.AccountEnabled -eq $true) {
                            $userEmails += $user.UserPrincipalName
                        }
                    }
                }
            }
            catch {
                Write-Warning "Failed to get details for member $($member.Id): $($_.Exception.Message)"
                continue
            }
        }
        
        if ($userEmails.Count -eq 0) {
            throw "No valid user email addresses found in the group. Total members processed: $memberCount"
        }
        
        return $userEmails
    }
    catch {
        $errorMsg = $_.Exception.Message
        if ($errorMsg -match "not found|404") {
            throw "Microsoft 365 Group not found. Please verify the group name or email address."
        }
        elseif ($errorMsg -match "forbidden|unauthorized|403|401") {
            throw "Access denied to the group. Please ensure you have permissions to read group membership."
        }
        else {
            throw "Failed to get group members: $errorMsg"
        }
    }
}

# Function to process users and generate report
function Process-Users {
    param(
        [array]$UserData,  # Changed from UserEmails to UserData
        [string]$OutputPath,
        [int]$Days,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatusLabel
    )
    
    $results = @()
    $totalUsers = $UserData.Count
    $currentUser = 0
    
    foreach ($user in $UserData) {
        $currentUser++
        
        # Handle both old format (just email string) and new format (object with email/names)
        if ($user -is [string]) {
            $email = $user
            $firstName = $null
            $lastName = $null
        } else {
            $email = $user.Email
            $firstName = $user.FirstName
            $lastName = $user.LastName
        }
        
        if ([string]::IsNullOrWhiteSpace($email)) {
            continue
        }
        
        # Update progress
        $percentage = [math]::Round(($currentUser / $totalUsers) * 100)
        $ProgressBar.Value = $percentage
        $StatusLabel.Text = "Processing: $email ($currentUser of $totalUsers)"
        [System.Windows.Forms.Application]::DoEvents()
        
        $userInfo = Get-UserSignInInfo -UserEmail $email -Days $Days -FirstName $firstName -LastName $lastName
        $results += $userInfo
        
        Start-Sleep -Milliseconds 100
    }
    
    # Export results - dynamically select the correct column name
    $signInColumnName = "SignedInLast${Days}Days"
    $results | Select-Object FirstName, LastName, EmailAddress, LastSignInDateTime, $signInColumnName, Details | Export-Csv -Path $OutputPath -NoTypeInformation
    
    return $results
}

# Install required modules
Install-RequiredModules

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Azure AD Sign-in Checker"
$form.Size = New-Object System.Drawing.Size(500, 520)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Title label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(10, 10)
$titleLabel.Size = New-Object System.Drawing.Size(460, 30)
$titleLabel.Text = "Azure AD Sign-in Checker - Customizable Days"
$titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
$titleLabel.TextAlign = "MiddleCenter"
$form.Controls.Add($titleLabel)

# Days selection group box
$daysGroupBox = New-Object System.Windows.Forms.GroupBox
$daysGroupBox.Location = New-Object System.Drawing.Point(10, 50)
$daysGroupBox.Size = New-Object System.Drawing.Size(460, 70)
$daysGroupBox.Text = "Select Time Period"
$form.Controls.Add($daysGroupBox)

# Days label
$daysLabel = New-Object System.Windows.Forms.Label
$daysLabel.Location = New-Object System.Drawing.Point(15, 25)
$daysLabel.Size = New-Object System.Drawing.Size(200, 20)
$daysLabel.Text = "Check sign-ins within last:"
$daysGroupBox.Controls.Add($daysLabel)

# Days numeric up/down with validation
$daysNumericUpDown = New-Object System.Windows.Forms.NumericUpDown
$daysNumericUpDown.Location = New-Object System.Drawing.Point(220, 23)
$daysNumericUpDown.Size = New-Object System.Drawing.Size(80, 25)
$daysNumericUpDown.Minimum = 1
$daysNumericUpDown.Maximum = 90
$daysNumericUpDown.Value = 90
$daysNumericUpDown.DecimalPlaces = 0
$daysNumericUpDown.Increment = 1

# Add validation event handler
$daysNumericUpDown.Add_ValueChanged({
    if ($daysNumericUpDown.Value -lt 1) {
        $daysNumericUpDown.Value = 1
    }
    elseif ($daysNumericUpDown.Value -gt 90) {
        $daysNumericUpDown.Value = 90
    }
})

$daysGroupBox.Controls.Add($daysNumericUpDown)

# Days suffix label
$daysSuffixLabel = New-Object System.Windows.Forms.Label
$daysSuffixLabel.Location = New-Object System.Drawing.Point(310, 25)
$daysSuffixLabel.Size = New-Object System.Drawing.Size(50, 20)
$daysSuffixLabel.Text = "days"
$daysGroupBox.Controls.Add($daysSuffixLabel)

# Input method group box
$inputGroupBox = New-Object System.Windows.Forms.GroupBox
$inputGroupBox.Location = New-Object System.Drawing.Point(10, 130)
$inputGroupBox.Size = New-Object System.Drawing.Size(460, 150)
$inputGroupBox.Text = "Select Input Method"
$form.Controls.Add($inputGroupBox)

# CSV file radio button
$csvRadio = New-Object System.Windows.Forms.RadioButton
$csvRadio.Location = New-Object System.Drawing.Point(15, 25)
$csvRadio.Size = New-Object System.Drawing.Size(200, 20)
$csvRadio.Text = "CSV File with Email Addresses"
$csvRadio.Checked = $true
$inputGroupBox.Controls.Add($csvRadio)

# CSV file path textbox
$csvPathTextBox = New-Object System.Windows.Forms.TextBox
$csvPathTextBox.Location = New-Object System.Drawing.Point(15, 50)
$csvPathTextBox.Size = New-Object System.Drawing.Size(350, 23)
$csvPathTextBox.ReadOnly = $true
$inputGroupBox.Controls.Add($csvPathTextBox)

# CSV browse button
$csvBrowseButton = New-Object System.Windows.Forms.Button
$csvBrowseButton.Location = New-Object System.Drawing.Point(370, 49)
$csvBrowseButton.Size = New-Object System.Drawing.Size(75, 25)
$csvBrowseButton.Text = "Browse..."
$inputGroupBox.Controls.Add($csvBrowseButton)

# Group radio button
$groupRadio = New-Object System.Windows.Forms.RadioButton
$groupRadio.Location = New-Object System.Drawing.Point(15, 85)
$groupRadio.Size = New-Object System.Drawing.Size(200, 20)
$groupRadio.Text = "Microsoft 365 Group"
$inputGroupBox.Controls.Add($groupRadio)

# Group email textbox
$groupEmailTextBox = New-Object System.Windows.Forms.TextBox
$groupEmailTextBox.Location = New-Object System.Drawing.Point(15, 110)
$groupEmailTextBox.Size = New-Object System.Drawing.Size(430, 23)
$groupEmailTextBox.Text = "Enter Microsoft 365 Group email address (e.g., group@contoso.com)"
$groupEmailTextBox.ForeColor = [System.Drawing.Color]::Gray
$groupEmailTextBox.Enabled = $false
$inputGroupBox.Controls.Add($groupEmailTextBox)

# Output section
$outputGroupBox = New-Object System.Windows.Forms.GroupBox
$outputGroupBox.Location = New-Object System.Drawing.Point(10, 290)
$outputGroupBox.Size = New-Object System.Drawing.Size(460, 80)
$outputGroupBox.Text = "Output Location"
$form.Controls.Add($outputGroupBox)

# Output path textbox
$outputPathTextBox = New-Object System.Windows.Forms.TextBox
$outputPathTextBox.Location = New-Object System.Drawing.Point(15, 25)
$outputPathTextBox.Size = New-Object System.Drawing.Size(350, 23)
$outputPathTextBox.Text = ".\SignInResults.csv"
$outputGroupBox.Controls.Add($outputPathTextBox)

# Output browse button
$outputBrowseButton = New-Object System.Windows.Forms.Button
$outputBrowseButton.Location = New-Object System.Drawing.Point(370, 24)
$outputBrowseButton.Size = New-Object System.Drawing.Size(75, 25)
$outputBrowseButton.Text = "Browse..."
$outputGroupBox.Controls.Add($outputBrowseButton)

# Progress bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 380)
$progressBar.Size = New-Object System.Drawing.Size(460, 23)
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Status label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 410)
$statusLabel.Size = New-Object System.Drawing.Size(460, 20)
$statusLabel.Text = "Ready to process..."
$statusLabel.Visible = $false
$form.Controls.Add($statusLabel)

# Run button
$runButton = New-Object System.Windows.Forms.Button
$runButton.Location = New-Object System.Drawing.Point(300, 440)
$runButton.Size = New-Object System.Drawing.Size(80, 30)
$runButton.Text = "Run Check"
$runButton.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($runButton)

# Exit button
$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Location = New-Object System.Drawing.Point(390, 440)
$exitButton.Size = New-Object System.Drawing.Size(80, 30)
$exitButton.Text = "Exit"
$exitButton.BackColor = [System.Drawing.Color]::LightCoral
$form.Controls.Add($exitButton)

# Event handlers
$csvRadio.Add_CheckedChanged({
    $csvPathTextBox.Enabled = $csvRadio.Checked
    $csvBrowseButton.Enabled = $csvRadio.Checked
    $groupEmailTextBox.Enabled = $groupRadio.Checked
})

$groupRadio.Add_CheckedChanged({
    $csvPathTextBox.Enabled = $csvRadio.Checked
    $csvBrowseButton.Enabled = $csvRadio.Checked
    $groupEmailTextBox.Enabled = $groupRadio.Checked
})

$csvBrowseButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV File"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPathTextBox.Text = $openFileDialog.FileName
    }
})

$outputBrowseButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveFileDialog.Title = "Save Results As"
    $saveFileDialog.FileName = "SignInResults.csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputPathTextBox.Text = $saveFileDialog.FileName
    }
})

# Placeholder text functionality for group email textbox
$placeholderText = "Enter Microsoft 365 Group email address (e.g., group@contoso.com)"

$groupEmailTextBox.Add_GotFocus({
    if ($groupEmailTextBox.Text -eq $placeholderText) {
        $groupEmailTextBox.Text = ""
        $groupEmailTextBox.ForeColor = [System.Drawing.Color]::Black
    }
})

$groupEmailTextBox.Add_LostFocus({
    if ([string]::IsNullOrWhiteSpace($groupEmailTextBox.Text)) {
        $groupEmailTextBox.Text = $placeholderText
        $groupEmailTextBox.ForeColor = [System.Drawing.Color]::Gray
    }
})

$runButton.Add_Click({
    # Comprehensive input validation
    try {
        # Validate days selection
        $selectedDays = [int]$daysNumericUpDown.Value
        if ($selectedDays -lt 1 -or $selectedDays -gt 90) {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid number of days (1-90).", "Invalid Days Selection", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Validate input method selection
        if ($csvRadio.Checked) {
            if ([string]::IsNullOrWhiteSpace($csvPathTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("Please select a CSV file.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Validate CSV file exists and is accessible
            if (-not (Test-Path $csvPathTextBox.Text)) {
                [System.Windows.Forms.MessageBox]::Show("The selected CSV file does not exist or cannot be accessed.", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
            
            # Check file size
            try {
                $fileInfo = Get-Item $csvPathTextBox.Text
                if ($fileInfo.Length -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("The selected CSV file is empty.", "Empty File", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    return
                }
                elseif ($fileInfo.Length -gt 50MB) {
                    $result = [System.Windows.Forms.MessageBox]::Show("The selected CSV file is very large ($([math]::Round($fileInfo.Length/1MB, 2)) MB). This may take a long time to process. Continue?", "Large File Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                        return
                    }
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Cannot access the selected CSV file: $($_.Exception.Message)", "File Access Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        elseif ($groupRadio.Checked) {
            if ([string]::IsNullOrWhiteSpace($groupEmailTextBox.Text) -or $groupEmailTextBox.Text -eq $placeholderText) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a Microsoft 365 Group email address or name.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
            
            # Basic email format validation if it looks like an email
            $groupInput = $groupEmailTextBox.Text.Trim()
            if ($groupInput -match "@" -and $groupInput -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$') {
                [System.Windows.Forms.MessageBox]::Show("The group email address format appears to be invalid.", "Invalid Email Format", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select an input method (CSV file or Microsoft 365 Group).", "Input Method Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Validate output path
        if ([string]::IsNullOrWhiteSpace($outputPathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please specify an output file path.", "Output Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        # Validate output directory is writable
        $outputDir = Split-Path $outputPathTextBox.Text -Parent
        if ([string]::IsNullOrEmpty($outputDir)) {
            $outputDir = [System.IO.Directory]::GetCurrentDirectory()
        }
        
        if (-not (Test-Path $outputDir)) {
            try {
                New-Item -Path $outputDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Cannot create output directory: $($_.Exception.Message)", "Directory Creation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }
        
        # Test write access to output location
        try {
            $testFile = Join-Path $outputDir "write_test_$(Get-Random).tmp"
            "test" | Out-File -FilePath $testFile -ErrorAction Stop
            Remove-Item $testFile -ErrorAction SilentlyContinue
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Cannot write to the output location. Please check permissions.", "Write Access Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Validation error: $($_.Exception.Message)", "Input Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Disable controls during processing
    $runButton.Enabled = $false
    $progressBar.Visible = $true
    $statusLabel.Visible = $true
    $progressBar.Value = 0
    $statusLabel.Text = "Connecting to Microsoft Graph..."
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Enhanced Microsoft Graph connection with MFA support
        $statusLabel.Text = "Connecting to Microsoft Graph (MFA may be required)..."
        [System.Windows.Forms.Application]::DoEvents()
        
        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context -and $context.Account) {
            $statusLabel.Text = "Already connected to Microsoft Graph"
            [System.Windows.Forms.Application]::DoEvents()
        }
        else {
            # Show MFA preparation message
            $mfaMessage = "Authentication Required`n`n" +
                         "• A browser window will open for sign-in`n" +
                         "• Complete any multi-factor authentication prompts`n" +
                         "• This process may take up to 2 minutes`n" +
                         "• Do not close this application during authentication"
            
            $mfaResult = [System.Windows.Forms.MessageBox]::Show($mfaMessage, "Authentication Required", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            if ($mfaResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
                return
            }
            
            $statusLabel.Text = "Opening browser for authentication..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Prepare connection parameters
            $connectParams = @{
                Scopes = @("User.Read.All", "AuditLog.Read.All", "Directory.Read.All", "Group.Read.All")
                ErrorAction = "Stop"
                NoWelcome = $true
            }
            
            # Try interactive authentication first
            try {
                Connect-MgGraph @connectParams
            }
            catch {
                $authError = $_.Exception.Message
                
                # Handle specific authentication scenarios
                if ($authError -match "browser|interactive|AADSTS50058|AADSTS50079") {
                    $statusLabel.Text = "Trying alternative authentication method..."
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    $deviceCodeMessage = "Browser authentication failed. Would you like to try device code authentication?`n`n" +
                                       "Device code authentication will provide a code that you enter in your browser."
                    
                    $deviceResult = [System.Windows.Forms.MessageBox]::Show($deviceCodeMessage, "Alternative Authentication", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
                    
                    if ($deviceResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                        try {
                            Connect-MgGraph @connectParams -UseDeviceAuthentication
                        }
                        catch {
                            throw
                        }
                    }
                    else {
                        throw "Authentication cancelled by user"
                    }
                }
                else {
                    throw
                }
            }
            
            # Verify connection with progress updates
            $statusLabel.Text = "Verifying authentication..."
            [System.Windows.Forms.Application]::DoEvents()
            
            $verificationTimeout = 30
            $verificationStart = Get-Date
            $context = $null
            
            while (-not $context -and ((Get-Date) - $verificationStart).TotalSeconds -lt $verificationTimeout) {
                Start-Sleep -Seconds 1
                $context = Get-MgContext -ErrorAction SilentlyContinue
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            if (-not ($context -and $context.Account)) {
                throw "Authentication completed but no valid context was established"
            }
            
            $statusLabel.Text = "Authentication successful - verifying permissions..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Show current context information
            Write-Verbose "Connected to Microsoft Graph successfully." -Verbose
            Write-Verbose "Account: $($context.Account)" -Verbose
            Write-Verbose "Environment: $($context.Environment)" -Verbose
            
            # Show current permissions
            try {
                $currentPermissions = $context.Scopes
                if ($currentPermissions) {
                    Write-Verbose "Current scopes: $($currentPermissions -join ', ')" -Verbose
                }
                else {
                    Write-Verbose "No scopes information available from context" -Verbose
                }
            }
            catch {
                Write-Verbose "Could not retrieve current permissions: $($_.Exception.Message)" -Verbose
            }
            
            # Skip detailed permission testing to avoid GUI hanging
            $statusLabel.Text = "Skipping permission verification to prevent GUI hang..."
            [System.Windows.Forms.Application]::DoEvents()
            
            Write-Verbose "⚠️ Skipping detailed permission verification to prevent GUI hang" -Verbose
            Write-Verbose "✅ Basic authentication successful - proceeding with processing" -Verbose
            
            # Set empty permission issues array since we're skipping verification
            $permissionIssues = @()
            
            Start-Sleep -Milliseconds 500  # Brief pause for user feedback
            
            if ($permissionIssues.Count -gt 0) {
                $permissionWarning = "Permission verification found issues:`n`n" +
                                   "$($permissionIssues -join "`n")`n`n" +
                                   "This will likely cause processing failures. Required permissions:`n" +
                                   "• User.Read.All (to read user profiles)`n" +
                                   "• AuditLog.Read.All (to read sign-in activity)`n" +
                                   "• Directory.Read.All (for directory access)`n" +
                                   "• Group.Read.All (for group membership)`n`n" +
                                   "Continue anyway? (Users will show 'Error' for sign-in data)"
                
                $permResult = [System.Windows.Forms.MessageBox]::Show($permissionWarning, "Permission Issues Detected", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                
                if ($permResult -eq [System.Windows.Forms.DialogResult]::No) {
                    Disconnect-MgGraph -ErrorAction SilentlyContinue
                    return
                }
            }
            else {
                Write-Verbose "✅ All required permissions verified successfully" -Verbose
            }
        }
        
        $userEmails = @()
        
        if ($csvRadio.Checked) {
            # Read from CSV file
            $statusLabel.Text = "Reading CSV file..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if (-not (Test-Path $csvPathTextBox.Text)) {
                throw "CSV file not found: $($csvPathTextBox.Text)"
            }
            
            # Read CSV file content directly to handle duplicate column names
            $csvContent = Get-Content $csvPathTextBox.Text
            if ($csvContent.Count -lt 2) {
                throw "CSV file appears to be empty or has no data rows"
            }
            
            # Parse header to find email, first name, and last name columns
            $header = $csvContent[0] -split ','
            $emailColumnIndex = -1
            $firstNameColumnIndex = -1
            $lastNameColumnIndex = -1
            
            $possibleEmailColumns = @('Email', 'EmailAddress', 'UserPrincipalName', 'Mail', 'E-mail', 'EMAIL')
            $possibleFirstNameColumns = @('First Name', 'FirstName', 'FIRST_NAME', 'Given Name', 'GivenName')
            $possibleLastNameColumns = @('Last Name', 'LastName', 'LAST_NAME', 'Family Name', 'FamilyName', 'Surname')
            
            # Find column indexes
            for ($i = 0; $i -lt $header.Count; $i++) {
                $columnName = $header[$i].Trim('"')
                
                if ($emailColumnIndex -eq -1 -and $possibleEmailColumns -contains $columnName) {
                    $emailColumnIndex = $i
                    $statusLabel.Text = "Found email column: '$columnName' at position $($i + 1)"
                }
                elseif ($firstNameColumnIndex -eq -1 -and $possibleFirstNameColumns -contains $columnName) {
                    $firstNameColumnIndex = $i
                }
                elseif ($lastNameColumnIndex -eq -1 -and $possibleLastNameColumns -contains $columnName) {
                    $lastNameColumnIndex = $i
                }
            }
            
            if ($emailColumnIndex -eq -1) {
                $availableColumns = $header -join ', '
                throw "Could not find email column in CSV. Available columns: $availableColumns. Expected one of: $($possibleEmailColumns -join ', ')"
            }
            
            $statusLabel.Text = "Processing CSV data with name columns..."
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 300
            
            # Extract user data by column index to avoid property name conflicts
            $userData = @()
            for ($i = 1; $i -lt $csvContent.Count; $i++) {
                $row = $csvContent[$i] -split ','
                if ($row.Count -gt $emailColumnIndex) {
                    $email = $row[$emailColumnIndex].Trim('"').Trim()
                    if (-not [string]::IsNullOrWhiteSpace($email)) {
                        $firstName = if ($firstNameColumnIndex -ne -1 -and $row.Count -gt $firstNameColumnIndex) { 
                            $row[$firstNameColumnIndex].Trim('"').Trim() 
                        } else { $null }
                        
                        $lastName = if ($lastNameColumnIndex -ne -1 -and $row.Count -gt $lastNameColumnIndex) { 
                            $row[$lastNameColumnIndex].Trim('"').Trim() 
                        } else { $null }
                        
                        $userData += [PSCustomObject]@{
                            Email = $email
                            FirstName = $firstName
                            LastName = $lastName
                        }
                    }
                }
            }
        }
        else {
            # Get from Microsoft 365 Group
            $statusLabel.Text = "Getting group members..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Try to find the group by display name or email
            $group = $null
            try {
                $group = Get-MgGroup -Filter "mail eq '$($groupEmailTextBox.Text)'" -ErrorAction Stop | Select-Object -First 1
            }
            catch {
                # Try by display name if email search fails
                try {
                    $group = Get-MgGroup -Filter "displayName eq '$($groupEmailTextBox.Text)'" -ErrorAction Stop | Select-Object -First 1
                }
                catch {
                    throw "Could not find Microsoft 365 Group: $($groupEmailTextBox.Text)"
                }
            }
            
            if (-not $group) {
                throw "Microsoft 365 Group not found: $($groupEmailTextBox.Text)"
            }
            
            $userEmails = Get-M365GroupMembers -GroupId $group.Id
            
            if ($userEmails.Count -eq 0) {
                throw "No users found in the specified group."
            }
            
            # Convert group emails to userData format for consistency
            $userData = @()
            foreach ($email in $userEmails) {
                $userData += [PSCustomObject]@{
                    Email = $email
                    FirstName = $null
                    LastName = $null
                }
            }
        }
        
        # Process users
        $selectedDays = [int]$daysNumericUpDown.Value
        $results = Process-Users -UserData $userData -OutputPath $outputPathTextBox.Text -Days $selectedDays -ProgressBar $progressBar -StatusLabel $statusLabel
        
        # Disconnect from Microsoft Graph
        Disconnect-MgGraph
        
        # Show results - use safe counting to avoid Count property errors
        $successResults = @($results | Where-Object { $_.Status -eq "Success" })
        $errorResults = @($results | Where-Object { $_.Status -eq "Error" })
        $signInColumnName = "SignedInLast${selectedDays}Days"
        $withinDaysResults = @($results | Where-Object { $_.$signInColumnName -eq "Yes" })
        $outsideDaysResults = @($results | Where-Object { $_.$signInColumnName -eq "No" })
        
        $successCount = $successResults.Count
        $errorCount = $errorResults.Count
        $withinDays = $withinDaysResults.Count
        $outsideDays = $outsideDaysResults.Count
        
        # Ensure results is treated as array for safe counting
        $totalResults = @($results)
        
        $resultMessage = @"
Processing completed successfully!

Summary:
• Total users processed: $($totalResults.Count)
• Successful lookups: $successCount
• Failed lookups: $errorCount
• Users signed in within $selectedDays days: $withinDays
• Users NOT signed in within $selectedDays days: $outsideDays

Results saved to: $($outputPathTextBox.Text)
"@
        
        [System.Windows.Forms.MessageBox]::Show($resultMessage, "Processing Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        $statusLabel.Text = "Error occurred"
        $statusLabel.ForeColor = [System.Drawing.Color]::Red
        
        $errorMessage = $_.Exception.Message
        
        # Provide user-friendly error messages for common authentication and processing issues
        $userFriendlyMessage = switch -Regex ($errorMessage) {
            "AADSTS50058|user_consent_required" { 
                "Administrator consent is required for this application.`n`nPlease contact your administrator to grant the necessary permissions." 
            }
            "AADSTS50079|strong_authentication_required" { 
                "Multi-factor authentication is required but not configured properly.`n`nPlease ensure MFA is set up for your account and try again." 
            }
            "AADSTS53003|blocked_by_conditional_access" { 
                "Access is blocked by Conditional Access policies.`n`nPlease contact your administrator or try from a compliant device." 
            }
            "AADSTS70008|expired_password" { 
                "Your password has expired.`n`nPlease reset your password and try again." 
            }
            "AADSTS50076|strong_authentication_required" { 
                "Multi-factor authentication challenge is required.`n`nPlease complete the MFA prompt in your browser." 
            }
            "timeout|timed out" { 
                "Authentication or processing timed out.`n`nThis may happen with slow internet connections or MFA delays. Please try again." 
            }
            "cancelled|canceled" { 
                "Operation was cancelled.`n`nTo use this application, authentication is required." 
            }
            "browser|interactive" { 
                "Browser authentication failed.`n`nEnsure your default browser is working properly and try again." 
            }
            "CSV file not found|Could not find email column" { 
                "CSV file issue: $errorMessage`n`nPlease verify:`n• The file path is correct`n• The file contains an email column`n• You have permission to read the file" 
            }
            "No users found in the specified group" { 
                "Group issue: $errorMessage`n`nPlease verify:`n• The group name/email is correct`n• You have permission to read group membership`n• The group contains members" 
            }
            default { 
                "Processing failed: $errorMessage`n`nCommon solutions:`n" +
                "• Ensure you have the required permissions`n" +
                "• Check your internet connection`n" +
                "• Verify file paths and permissions`n" +
                "• Try running as administrator`n" +
                "• Contact your IT administrator if using managed devices"
            }
        }
        
        [System.Windows.Forms.MessageBox]::Show($userFriendlyMessage, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        # Re-enable controls
        $runButton.Enabled = $true
        $progressBar.Visible = $false
        $statusLabel.Visible = $false
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    }
})

$exitButton.Add_Click({
    $form.Close()
})

# Show the form
[System.Windows.Forms.Application]::EnableVisualStyles()
$form.ShowDialog()