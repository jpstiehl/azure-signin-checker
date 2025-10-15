# Quick diagnostic script to test Microsoft Graph permissions

# Clean up any existing modules first
Write-Host "Cleaning existing Graph modules..." -ForegroundColor Yellow
$existingGraphModules = Get-Module -Name "Microsoft.Graph.*" -ErrorAction SilentlyContinue
if ($existingGraphModules) {
    $existingGraphModules | Remove-Module -Force -ErrorAction SilentlyContinue
    Write-Host "Removed $($existingGraphModules.Count) existing Graph modules" -ForegroundColor Green
}

# Import modules with version handling
try {
    Write-Host "Importing Microsoft.Graph.Authentication..." -ForegroundColor Yellow
    Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop
    
    Write-Host "Importing Microsoft.Graph.Users..." -ForegroundColor Yellow  
    Import-Module Microsoft.Graph.Users -Force -ErrorAction Stop
    
    Write-Host "✅ Modules imported successfully" -ForegroundColor Green
}
catch {
    Write-Host "❌ Module import failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Connect with the same parameters as the GUI script
$connectParams = @{
    Scopes = @(
        "User.Read.All",
        "AuditLog.Read.All", 
        "Directory.Read.All",
        "Group.Read.All"
    )
    NoWelcome = $true
    ContextScope = "Process"
}

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Connect-MgGraph @connectParams

# Get context information
$context = Get-MgContext
Write-Host "Connected successfully!" -ForegroundColor Green
Write-Host "Account: $($context.Account)" -ForegroundColor Green
Write-Host "Scopes: $($context.Scopes -join ', ')" -ForegroundColor Green

# Test basic user read
Write-Host "`nTesting basic user read..." -ForegroundColor Yellow
$testUser = $null
try {
    $testUser = Get-MgUser -Top 1 -Property "displayName,userPrincipalName,id" -ErrorAction Stop
    if ($testUser -and $testUser.Id) {
        Write-Host "✅ Basic user read successful: $($testUser.DisplayName)" -ForegroundColor Green
        Write-Host "User ID: $($testUser.Id)" -ForegroundColor Cyan
    }
    else {
        Write-Host "❌ Retrieved user but missing required properties" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ Basic user read failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Test sign-in activity read
Write-Host "`nTesting sign-in activity read..." -ForegroundColor Yellow
try {
    if ($testUser -and $testUser.Id) {
        Write-Host "Testing with user: $($testUser.DisplayName) (ID: $($testUser.Id))" -ForegroundColor Cyan
        $userWithSignIn = Get-MgUser -UserId $testUser.Id -Property "displayName,signInActivity" -ErrorAction Stop
        if ($userWithSignIn.SignInActivity) {
            Write-Host "✅ Sign-in activity read successful" -ForegroundColor Green
            if ($userWithSignIn.SignInActivity.LastSignInDateTime) {
                Write-Host "Last interactive sign-in: $($userWithSignIn.SignInActivity.LastSignInDateTime)" -ForegroundColor Green
            }
            if ($userWithSignIn.SignInActivity.LastNonInteractiveSignInDateTime) {
                Write-Host "Last non-interactive sign-in: $($userWithSignIn.SignInActivity.LastNonInteractiveSignInDateTime)" -ForegroundColor Green
            }
        }
        else {
            Write-Host "⚠️ Sign-in activity property is null (user may not have sign-in data)" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "❌ No valid test user available - skipping sign-in activity test" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ Sign-in activity read failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $($_.Exception.ToString())" -ForegroundColor Red
}

# Test a specific user from the CSV
Write-Host "`nTesting specific user lookup..." -ForegroundColor Yellow
$testEmail = "jbohnens@xula.edu"  # Use the authenticated account as test
try {
    $specificUser = Get-MgUser -Filter "userPrincipalName eq '$testEmail'" -Property "displayName,userPrincipalName,signInActivity" -ErrorAction Stop
    if ($specificUser) {
        Write-Host "✅ Found user: $($specificUser.DisplayName)" -ForegroundColor Green
        if ($specificUser.SignInActivity) {
            Write-Host "Last sign-in: $($specificUser.SignInActivity.LastSignInDateTime)" -ForegroundColor Green
        }
        else {
            Write-Host "⚠️ No sign-in activity data for this user" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "❌ User not found: $testEmail" -ForegroundColor Red
    }
}
catch {
    Write-Host "❌ Specific user lookup failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $($_.Exception.ToString())" -ForegroundColor Red
}

Write-Host "`nDiagnostic complete." -ForegroundColor Cyan