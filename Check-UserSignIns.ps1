# Azure AD Sign-in Checker Script
# This script reads email addresses from a CSV file, checks their last sign-in time in Azure AD,
# and exports the results to a new CSV file with 90-day analysis.

param(
    [Parameter(Mandatory = $true, HelpMessage = "Path to the CSV file containing email addresses")]
    [ValidateNotNullOrEmpty()]
    [string]$InputCsvPath,
    
    [Parameter(Mandatory = $false, HelpMessage = "Path for the output CSV file")]
    [string]$OutputCsvPath = ".\SignInResults.csv"
)

# Set error action preference and strict mode for better error handling
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# Global error handler
trap {
    Write-Host "`n❌ Critical Error Occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "`nStack Trace:" -ForegroundColor Yellow
    Write-Host $_.ScriptStackTrace -ForegroundColor Yellow
    
    # Cleanup
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "`n🔒 Disconnected from Microsoft Graph" -ForegroundColor Yellow
    } catch { }
    
    Write-Host "`n❌ Script terminated due to critical error" -ForegroundColor Red
    exit 1
}

# Input parameter validation
Write-Host "🔍 Validating script parameters..." -ForegroundColor Cyan

if ([string]::IsNullOrWhiteSpace($InputCsvPath)) {
    throw "InputCsvPath parameter cannot be null or empty"
}

if ([string]::IsNullOrWhiteSpace($OutputCsvPath)) {
    $OutputCsvPath = ".\SignInResults.csv"
    Write-Host "Using default output path: $OutputCsvPath" -ForegroundColor Yellow
}

# Convert to absolute paths
try {
    $InputCsvPath = Resolve-Path $InputCsvPath -ErrorAction Stop
    Write-Host "✅ Input path resolved: $InputCsvPath" -ForegroundColor Green
} catch {
    throw "Cannot resolve input CSV path: $InputCsvPath. Error: $($_.Exception.Message)"
}

# Ensure output path is absolute
if (-not [System.IO.Path]::IsPathFullyQualified($OutputCsvPath)) {
    $OutputCsvPath = Join-Path (Get-Location) $OutputCsvPath
}
Write-Host "✅ Output path set: $OutputCsvPath" -ForegroundColor Green

# Function to install and import required modules with enhanced error handling
function Install-RequiredModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Reports'
    )
    
    # Check PowerShell execution policy
    $executionPolicy = Get-ExecutionPolicy
    if ($executionPolicy -eq 'Restricted') {
        Write-Warning "PowerShell execution policy is set to 'Restricted'. You may need to change it to run this script."
        Write-Host "Run: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
    }
    
    # Check if running in constrained language mode
    if ($ExecutionContext.SessionState.LanguageMode -eq 'ConstrainedLanguage') {
        Write-Warning "PowerShell is running in Constrained Language Mode. Some features may not work properly."
    }
    
    foreach ($module in $requiredModules) {
        Write-Host "🔍 Checking module: $module" -ForegroundColor Yellow
        
        # Check if module is available
        $moduleAvailable = Get-Module -ListAvailable -Name $module -ErrorAction SilentlyContinue
        
        if (-not $moduleAvailable) {
            Write-Host "📦 Installing module: $module" -ForegroundColor Yellow
            
            # Try multiple installation methods
            $installSuccess = $false
            $attempts = 0
            $maxAttempts = 3
            
            while (-not $installSuccess -and $attempts -lt $maxAttempts) {
                $attempts++
                Write-Host "  Attempt $attempts of $maxAttempts..." -ForegroundColor Gray
                
                try {
                    # Check internet connectivity first
                    $testConnection = Test-NetConnection -ComputerName "www.powershellgallery.com" -Port 443 -InformationLevel Quiet -ErrorAction SilentlyContinue
                    if (-not $testConnection) {
                        throw "No internet connection to PowerShell Gallery"
                    }
                    
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                    $installSuccess = $true
                    Write-Host "✅ Successfully installed: $module" -ForegroundColor Green
                }
                catch {
                    Write-Warning "  Installation attempt $attempts failed: $($_.Exception.Message)"
                    if ($attempts -lt $maxAttempts) {
                        Write-Host "  Retrying in 5 seconds..." -ForegroundColor Yellow
                        Start-Sleep -Seconds 5
                    }
                }
            }
            
            if (-not $installSuccess) {
                Write-Error "❌ Failed to install module $module after $maxAttempts attempts"
                Write-Host "Manual installation options:" -ForegroundColor Yellow
                Write-Host "1. Run PowerShell as Administrator and try again" -ForegroundColor Yellow
                Write-Host "2. Install manually: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
                Write-Host "3. Download and install Microsoft Graph PowerShell SDK from: https://github.com/microsoftgraph/msgraph-sdk-powershell" -ForegroundColor Yellow
                exit 1
            }
        }
        else {
            Write-Host "✅ Module already installed: $module" -ForegroundColor Green
        }
        
        # Import the module with retry logic
        $importSuccess = $false
        $attempts = 0
        $maxAttempts = 3
        
        while (-not $importSuccess -and $attempts -lt $maxAttempts) {
            $attempts++
            try {
                Import-Module $module -Force -ErrorAction Stop
                $importSuccess = $true
                Write-Host "✅ Successfully imported: $module" -ForegroundColor Green
            }
            catch {
                Write-Warning "  Import attempt $attempts failed: $($_.Exception.Message)"
                if ($attempts -lt $maxAttempts) {
                    Start-Sleep -Seconds 2
                }
            }
        }
        
        if (-not $importSuccess) {
            Write-Error "❌ Failed to import module $module after $maxAttempts attempts"
            exit 1
        }
    }
    
    # Verify all modules are properly loaded
    foreach ($module in $requiredModules) {
        $loadedModule = Get-Module -Name $module -ErrorAction SilentlyContinue
        if (-not $loadedModule) {
            Write-Error "❌ Module $module is not properly loaded"
            exit 1
        }
    }
}

# Install and import required modules
Write-Host "🔧 Setting up required Microsoft Graph modules..." -ForegroundColor Cyan
Install-RequiredModules
Write-Host "✅ All required Microsoft Graph modules are ready" -ForegroundColor Green

# Function to check if a user signed in within the last 90 days
function Test-SignInWithin90Days {
    param([datetime]$LastSignInDate)
    
    if ($null -eq $LastSignInDate -or $LastSignInDate -eq [datetime]::MinValue) {
        return $false
    }
    
    $cutoffDate = (Get-Date).AddDays(-90)
    return $LastSignInDate -gt $cutoffDate
}

# Function to get user sign-in information with enhanced error handling and retry logic
function Get-UserSignInInfo {
    param([string]$UserEmail)
    
    if ([string]::IsNullOrWhiteSpace($UserEmail)) {
        return [PSCustomObject]@{
            Username = "Invalid"
            EmailAddress = "Empty/Null"
            LastSignInDateTime = "Error"
            SignedInLast90Days = "No"
            Status = "Error"
            Error = "Email address is empty or null"
        }
    }
    
    $maxRetries = 3
    $retryDelay = 2
    
    for ($attempt = 1; $attempt -le $maxRetries; $attempt++) {
        try {
            # Validate email format  
            if ($UserEmail -notmatch '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]+$') {
                throw "Invalid email format: $UserEmail"
            }
            
            # Get user information including sign-in activity
            $user = Get-MgUser -UserId $UserEmail -Property "displayName,userPrincipalName,signInActivity" -ErrorAction Stop
            
            if (-not $user) {
                throw "User not found or no data returned"
            }
            
            # Extract last sign-in information
            $lastSignInDateTime = $null
            
            if ($user.SignInActivity) {
                # Try to get the most recent sign-in date from available properties
                if ($user.SignInActivity.LastSignInDateTime) {
                    try {
                        $lastSignInDateTime = [datetime]$user.SignInActivity.LastSignInDateTime
                    }
                    catch {
                        Write-Warning "Failed to parse LastSignInDateTime for $UserEmail"
                    }
                }
                elseif ($user.SignInActivity.LastNonInteractiveSignInDateTime) {
                    try {
                        $lastSignInDateTime = [datetime]$user.SignInActivity.LastNonInteractiveSignInDateTime
                    }
                    catch {
                        Write-Warning "Failed to parse LastNonInteractiveSignInDateTime for $UserEmail"
                    }
                }
            }
            
            # Determine if signed in within last 90 days
            $within90Days = if ($lastSignInDateTime) {
                Test-SignInWithin90Days -LastSignInDate $lastSignInDateTime
            } else {
                $false
            }
            
            return [PSCustomObject]@{
                Username = if ($user.DisplayName) { $user.DisplayName } else { "No Display Name" }
                EmailAddress = if ($user.UserPrincipalName) { $user.UserPrincipalName } else { $UserEmail }
                LastSignInDateTime = if ($lastSignInDateTime) { $lastSignInDateTime.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }
                SignedInLast90Days = if ($within90Days) { "Yes" } else { "No" }
                Status = "Success"
                Error = $null
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            
            # Check for specific error types that might need retry
            $shouldRetry = $false
            if ($errorMessage -match "throttled|rate limit|503|502|timeout" -and $attempt -lt $maxRetries) {
                $shouldRetry = $true
                Write-Warning "Transient error for $UserEmail (attempt $attempt): $errorMessage. Retrying in $retryDelay seconds..."
                Start-Sleep -Seconds $retryDelay
                $retryDelay *= 2  # Exponential backoff
            }
            elseif ($errorMessage -match "not found|does not exist|404") {
                # User not found - don't retry
                Write-Warning "User not found: $UserEmail"
                break
            }
            elseif ($errorMessage -match "forbidden|unauthorized|403|401") {
                # Permission error - don't retry
                Write-Warning "Permission denied for $UserEmail"
                break
            }
            else {
                Write-Warning "Error getting sign-in info for $UserEmail (attempt $attempt): $errorMessage"
                if ($attempt -lt $maxRetries) {
                    $shouldRetry = $true
                    Start-Sleep -Seconds $retryDelay
                }
            }
            
            if (-not $shouldRetry) {
                return [PSCustomObject]@{
                    Username = "Unknown"
                    EmailAddress = $UserEmail
                    LastSignInDateTime = "Error"
                    SignedInLast90Days = "No"
                    Status = "Error"
                    Error = $errorMessage
                }
            }
        }
    }
    
    # If we get here, all retries failed
    return [PSCustomObject]@{
        Username = "Unknown"
        EmailAddress = $UserEmail
        LastSignInDateTime = "Error"
        SignedInLast90Days = "No"
        Status = "Error"
        Error = "Failed after $maxRetries attempts"
    }
}

# Main script execution
Write-Host "🚀 Starting Azure AD Sign-in Checker Script" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# Enhanced input validation and connection handling
Write-Host "📋 Validating input parameters..." -ForegroundColor Yellow

# Validate input file exists and is accessible
if (-not (Test-Path $InputCsvPath)) {
    Write-Error "❌ Input CSV file not found: $InputCsvPath"
    Write-Host "Please verify the file path and ensure the file exists." -ForegroundColor Yellow
    exit 1
}

# Check file size and accessibility
try {
    $fileInfo = Get-Item $InputCsvPath -ErrorAction Stop
    if ($fileInfo.Length -eq 0) {
        Write-Error "❌ Input CSV file is empty: $InputCsvPath"
        exit 1
    }
    Write-Host "✅ Input file validated (Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB)" -ForegroundColor Green
}
catch {
    Write-Error "❌ Cannot access input file: $($_.Exception.Message)"
    exit 1
}

# Validate output path is writable
try {
    $outputDir = Split-Path $OutputCsvPath -Parent
    if ([string]::IsNullOrEmpty($outputDir)) {
        $outputDir = Get-Location
    }
    
    if (-not (Test-Path $outputDir)) {
        New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
        Write-Host "✅ Created output directory: $outputDir" -ForegroundColor Green
    }
    
    # Test write access
    $testFile = Join-Path $outputDir "write_test_$(Get-Random).tmp"
    "test" | Out-File -FilePath $testFile -ErrorAction Stop
    Remove-Item $testFile -ErrorAction SilentlyContinue
    Write-Host "✅ Output location is writable" -ForegroundColor Green
}
catch {
    Write-Error "❌ Cannot write to output location: $($_.Exception.Message)"
    exit 1
}

# Enhanced Microsoft Graph connection with MFA support
Write-Host "🔐 Connecting to Microsoft Graph..." -ForegroundColor Yellow

$connectionAttempts = 0
$maxConnectionAttempts = 3
$connected = $false

while (-not $connected -and $connectionAttempts -lt $maxConnectionAttempts) {
    $connectionAttempts++
    
    try {
        # Check if already connected
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($context -and $context.Account) {
            Write-Host "✅ Already connected to Microsoft Graph" -ForegroundColor Green
            Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
            Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Gray
            $connected = $true
            break
        }
        
        # Prepare for authentication
        Write-Host "📱 Preparing authentication (MFA may be required)..." -ForegroundColor Yellow
        Write-Host "  • A browser window will open for authentication" -ForegroundColor Gray
        Write-Host "  • Complete any multi-factor authentication prompts" -ForegroundColor Gray
        Write-Host "  • The process may take up to 2 minutes" -ForegroundColor Gray
        
        # Use interactive authentication with MFA support
        $connectParams = @{
            Scopes = @("User.Read.All", "AuditLog.Read.All", "Directory.Read.All")
            ErrorAction = "Stop"
            NoWelcome = $true
        }
        
        # Add interactive flag for better MFA handling
        try {
            Connect-MgGraph @connectParams
        }
        catch {
            # If interactive fails, try device code flow as fallback
            if ($_.Exception.Message -match "browser|interactive|AADSTS50058|AADSTS50079") {
                Write-Host "🔄 Trying alternative authentication method (device code)..." -ForegroundColor Yellow
                Write-Host "  • You will receive a device code to enter in your browser" -ForegroundColor Gray
                
                try {
                    Connect-MgGraph @connectParams -UseDeviceAuthentication
                }
                catch {
                    throw
                }
            }
            else {
                throw
            }
        }
        
        # Verify connection with timeout
        $verificationTimeout = 30
        $verificationStart = Get-Date
        $context = $null
        
        while (-not $context -and ((Get-Date) - $verificationStart).TotalSeconds -lt $verificationTimeout) {
            Start-Sleep -Seconds 2
            $context = Get-MgContext -ErrorAction SilentlyContinue
        }
        
        if ($context -and $context.Account) {
            Write-Host "✅ Successfully connected to Microsoft Graph" -ForegroundColor Green
            Write-Host "  Account: $($context.Account)" -ForegroundColor Gray
            Write-Host "  Tenant: $($context.TenantId)" -ForegroundColor Gray
            Write-Host "  Authentication Type: $($context.AuthType)" -ForegroundColor Gray
            
            # Verify required permissions
            Write-Host "🔍 Verifying permissions..." -ForegroundColor Yellow
            try {
                # Test a simple call to verify permissions
                Get-MgUser -Top 1 -Property "displayName" -ErrorAction Stop | Out-Null
                Write-Host "✅ Permissions verified successfully" -ForegroundColor Green
            }
            catch {
                Write-Warning "Permission verification failed. You may encounter issues during processing."
                Write-Host "  Ensure your account has the following permissions:" -ForegroundColor Yellow
                Write-Host "  • User.Read.All (to read user information)" -ForegroundColor Yellow
                Write-Host "  • AuditLog.Read.All (to read sign-in logs)" -ForegroundColor Yellow
                Write-Host "  • Directory.Read.All (to read directory information)" -ForegroundColor Yellow
            }
            
            $connected = $true
        }
        else {
            throw "Authentication completed but no valid context was established"
        }
    }
    catch {
        $errorMessage = $_.Exception.Message
        Write-Warning "Connection attempt $connectionAttempts failed: $errorMessage"
        
        # Provide specific guidance for common MFA/authentication issues
        if ($errorMessage -match "AADSTS50058|AADSTS50079|AADSTS50076") {
            Write-Host "🔐 Multi-Factor Authentication is required:" -ForegroundColor Yellow
            Write-Host "  • Complete the MFA challenge in your browser/authenticator app" -ForegroundColor Yellow
            Write-Host "  • Ensure your MFA method is properly configured" -ForegroundColor Yellow
        }
        elseif ($errorMessage -match "AADSTS50053|AADSTS50057") {
            Write-Host "🔒 Account access issue detected:" -ForegroundColor Yellow
            Write-Host "  • Your account may be locked or have conditional access restrictions" -ForegroundColor Yellow
            Write-Host "  • Contact your administrator if the issue persists" -ForegroundColor Yellow
        }
        elseif ($errorMessage -match "AADSTS65001|AADSTS65004") {
            Write-Host "📱 Consent required:" -ForegroundColor Yellow
            Write-Host "  • Admin consent may be required for the requested permissions" -ForegroundColor Yellow
            Write-Host "  • Contact your Azure AD administrator" -ForegroundColor Yellow
        }
        elseif ($errorMessage -match "timeout|AADSTS90002") {
            Write-Host "⏱️ Authentication timeout:" -ForegroundColor Yellow
            Write-Host "  • The authentication process took too long" -ForegroundColor Yellow
            Write-Host "  • Try again and complete the authentication more quickly" -ForegroundColor Yellow
        }
        
        if ($connectionAttempts -lt $maxConnectionAttempts) {
            Write-Host "Retrying in 10 seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10
        }
        else {
            Write-Error "❌ Failed to connect to Microsoft Graph after $maxConnectionAttempts attempts"
            Write-Host "`nTroubleshooting steps:" -ForegroundColor Yellow
            Write-Host "1. Ensure you have the necessary permissions: User.Read.All, AuditLog.Read.All, Directory.Read.All" -ForegroundColor Yellow
            Write-Host "2. Check your internet connection and ensure you can access portal.azure.com" -ForegroundColor Yellow
            Write-Host "3. Verify your MFA methods are configured and working" -ForegroundColor Yellow
            Write-Host "4. Try running: Disconnect-MgGraph, then run this script again" -ForegroundColor Yellow
            Write-Host "5. Ensure you have Global Reader or equivalent permissions in Azure AD" -ForegroundColor Yellow
            Write-Host "6. If using conditional access, ensure the device is compliant" -ForegroundColor Yellow
            exit 1
        }
    }
}

# Read and validate input CSV file
Write-Host "📄 Reading input CSV file: $InputCsvPath" -ForegroundColor Yellow
try {
    $inputData = Import-Csv $InputCsvPath -ErrorAction Stop
    
    if (-not $inputData -or $inputData.Count -eq 0) {
        throw "CSV file contains no data or only headers"
    }
    
    Write-Host "✅ Successfully read $($inputData.Count) records from input file" -ForegroundColor Green
}
catch {
    Write-Error "❌ Failed to read input CSV file: $($_.Exception.Message)"
    Write-Host "Please ensure the CSV file is properly formatted with headers." -ForegroundColor Yellow
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    exit 1
}

# Enhanced CSV structure validation
Write-Host "🔍 Validating CSV structure..." -ForegroundColor Yellow

if (-not $inputData[0] -or -not $inputData[0].PSObject.Properties) {
    Write-Error "❌ CSV file appears to be malformed or empty"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    exit 1
}

$emailColumn = $null
$possibleEmailColumns = @('Email', 'EmailAddress', 'UserPrincipalName', 'Mail', 'E-mail', 'UPN', 'PrimaryEmail')
$availableColumns = $inputData[0].PSObject.Properties.Name

Write-Host "Available columns in CSV: $($availableColumns -join ', ')" -ForegroundColor Gray

foreach ($column in $possibleEmailColumns) {
    if ($availableColumns -contains $column) {
        $emailColumn = $column
        break
    }
}

if (-not $emailColumn) {
    Write-Error "❌ Could not find email column in CSV."
    Write-Host "Expected one of: $($possibleEmailColumns -join ', ')" -ForegroundColor Yellow
    Write-Host "Available columns: $($availableColumns -join ', ')" -ForegroundColor Yellow
    Write-Host "Please ensure your CSV has a column with one of the expected email column names." -ForegroundColor Yellow
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    exit 1
}

# Validate that the email column has data
$validEmails = $inputData | Where-Object { -not [string]::IsNullOrWhiteSpace($_.$emailColumn) }
if (-not $validEmails -or $validEmails.Count -eq 0) {
    Write-Error "❌ No valid email addresses found in column '$emailColumn'"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    exit 1
}

$emptyEmailCount = $inputData.Count - $validEmails.Count
if ($emptyEmailCount -gt 0) {
    Write-Warning "Found $emptyEmailCount empty email addresses that will be skipped"
}

Write-Host "✅ Using '$emailColumn' column for email addresses ($($validEmails.Count) valid entries)" -ForegroundColor Green

# Initialize results array
$results = @()
$totalUsers = $inputData.Count
$currentUser = 0

Write-Host "🔍 Processing user sign-in information..." -ForegroundColor Yellow

# Process each user
foreach ($row in $inputData) {
    $currentUser++
    $email = $row.$emailColumn
    
    if ([string]::IsNullOrWhiteSpace($email)) {
        Write-Warning "Skipping empty email address at row $currentUser"
        continue
    }
    
    Write-Progress -Activity "Checking sign-in information" -Status "Processing $email - $currentUser of $totalUsers" -PercentComplete (($currentUser / $totalUsers) * 100)
    
    Write-Host "  📧 Checking: $email" -ForegroundColor Gray
    
    # Get sign-in information for the user
    $userInfo = Get-UserSignInInfo -UserEmail $email
    $results += $userInfo
    
    # Brief pause to avoid throttling
    Start-Sleep -Milliseconds 200
}

Write-Progress -Activity "Checking sign-in information" -Completed

# Generate summary statistics
$totalProcessed = $results.Count
$successfulLookups = ($results | Where-Object { $_.Status -eq "Success" }).Count
$errorLookups = ($results | Where-Object { $_.Status -eq "Error" }).Count
$usersWithin90Days = ($results | Where-Object { $_.SignedInLast90Days -eq "Yes" }).Count
$usersOutside90Days = ($results | Where-Object { $_.SignedInLast90Days -eq "No" }).Count

Write-Host "`n📊 Summary Statistics:" -ForegroundColor Cyan
Write-Host "  Total users processed: $totalProcessed" -ForegroundColor White
Write-Host "  Successful lookups: $successfulLookups" -ForegroundColor Green
Write-Host "  Failed lookups: $errorLookups" -ForegroundColor Red
Write-Host "  Users signed in within 90 days: $usersWithin90Days" -ForegroundColor Green
Write-Host "  Users NOT signed in within 90 days: $usersOutside90Days" -ForegroundColor Yellow

# Export results to CSV with enhanced error handling
Write-Host "`n💾 Exporting results to: $OutputCsvPath" -ForegroundColor Yellow

if (-not $results -or $results.Count -eq 0) {
    Write-Warning "No results to export"
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
    exit 1
}

try {
    # Backup existing file if it exists
    if (Test-Path $OutputCsvPath) {
        $backupPath = $OutputCsvPath -replace '\.csv$', "_backup_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        Copy-Item -Path $OutputCsvPath -Destination $backupPath -ErrorAction SilentlyContinue
        Write-Host "  Backed up existing file to: $backupPath" -ForegroundColor Gray
    }
    
    # Export with error handling
    $exportData = $results | Select-Object Username, EmailAddress, LastSignInDateTime, SignedInLast90Days
    $exportData | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    
    # Verify the export
    $exportedFile = Get-Item $OutputCsvPath -ErrorAction Stop
    if ($exportedFile.Length -eq 0) {
        throw "Exported file is empty"
    }
    
    Write-Host "✅ Results successfully exported to $OutputCsvPath" -ForegroundColor Green
    Write-Host "  File size: $([math]::Round($exportedFile.Length/1KB, 2)) KB" -ForegroundColor Gray
    Write-Host "  Records exported: $($exportData.Count)" -ForegroundColor Gray
}
catch {
    Write-Error "❌ Failed to export results: $($_.Exception.Message)"
    Write-Host "Attempting to save to alternative location..." -ForegroundColor Yellow
    
    try {
        $alternativePath = Join-Path $env:TEMP "SignInResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $results | Select-Object Username, EmailAddress, LastSignInDateTime, SignedInLast90Days | Export-Csv -Path $alternativePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Host "✅ Results saved to alternative location: $alternativePath" -ForegroundColor Green
    }
    catch {
        Write-Error "❌ Failed to save to alternative location: $($_.Exception.Message)"
    }
}

# Disconnect from Microsoft Graph
Write-Host "`n🔒 Disconnecting from Microsoft Graph..." -ForegroundColor Yellow
Disconnect-MgGraph
Write-Host "✅ Disconnected successfully" -ForegroundColor Green

Write-Host "`n🎉 Script completed successfully!" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan