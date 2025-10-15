# Azure AD Sign-in Checker

A PowerShell script that reads email addresses from a CSV file, checks their last sign-in time in Azure AD/Entra ID, and exports the results with 90-day analysis.

## Features

- ✅ Reads email addresses from CSV file
- ✅ Connects to Microsoft Graph API for Azure AD data
- ✅ Checks last sign-in date/time for each user
- ✅ Determines if users signed in within the last 90 days
- ✅ Exports results to CSV with comprehensive information
- ✅ Provides detailed progress reporting and error handling
- ✅ Supports multiple email column formats

## Prerequisites

### Required PowerShell Modules

You must install the Microsoft Graph PowerShell modules before running this script:

```powershell
# Install Microsoft Graph modules
Install-Module Microsoft.Graph -Scope CurrentUser
```

### Required Permissions

The script requires the following Microsoft Graph permissions:
- `User.Read.All` - To read user profiles
- `AuditLog.Read.All` - To read sign-in logs
- `Directory.Read.All` - To read directory information

You must have appropriate Azure AD permissions to grant these scopes when prompted during authentication.

## Usage

### Basic Usage

```powershell
.\Check-UserSignIns.ps1 -InputCsvPath ".\input-emails.csv"
```

### Advanced Usage

```powershell
.\Check-UserSignIns.ps1 -InputCsvPath ".\users.csv" -OutputCsvPath ".\results.csv"
```

### Parameters

- **InputCsvPath** (Required): Path to the CSV file containing email addresses
- **OutputCsvPath** (Optional): Path for the output CSV file (default: `.\SignInResults.csv`)

## Input CSV Format

The script automatically detects the email column from common column names:
- `Email`
- `EmailAddress`
- `UserPrincipalName`
- `Mail`
- `E-mail`

### Sample Input CSV

```csv
Email
john.doe@contoso.com
jane.smith@contoso.com
bob.wilson@contoso.com
```

## Output Format

The script generates a CSV file with the following columns:

| Column | Description |
|--------|-------------|
| **Username** | Display name of the user |
| **EmailAddress** | User's email address (UPN) |
| **LastSignInDateTime** | Last sign-in date and time (YYYY-MM-DD HH:MM:SS format) |
| **SignedInLast90Days** | "Yes" or "No" indicating if user signed in within 90 days |

### Sample Output

```csv
Username,EmailAddress,LastSignInDateTime,SignedInLast90Days
John Doe,john.doe@contoso.com,2024-10-10 14:30:25,Yes
Jane Smith,jane.smith@contoso.com,2024-07-15 09:15:42,No
Bob Wilson,bob.wilson@contoso.com,Never,No
```

## Script Features

### Authentication
- Interactive Microsoft Graph authentication
- Automatic scope request for required permissions
- Secure token handling

### Error Handling
- Validates input file existence
- Checks for required PowerShell modules
- Handles API errors gracefully
- Provides detailed error messages

### Progress Tracking
- Real-time progress bar
- Status updates for each user
- Summary statistics at completion

### Performance
- Built-in throttling to avoid API limits
- Efficient batch processing
- Minimal memory footprint

## Troubleshooting

### Enhanced Error Handling

Both scripts now include comprehensive error handling and recovery mechanisms:

- **Automatic retry logic** for transient failures
- **Network connectivity validation** before attempting connections
- **Input validation** with detailed error messages
- **Graceful degradation** with alternative save locations
- **Progress tracking** with real-time status updates

### Common Issues

**Module Installation Problems**
```powershell
# If automatic installation fails, try manual installation:
Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber

# For restricted environments:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Connection Issues**
- **Internet connectivity**: Scripts now test connection to PowerShell Gallery
- **Azure AD permissions**: Requires Global Reader or equivalent permissions
- **Multi-factor authentication**: May prompt for additional authentication
- **Conditional access policies**: May block programmatic access

**CSV Format Issues**
- **Supported email columns**: Email, EmailAddress, UserPrincipalName, Mail, E-mail, UPN, PrimaryEmail
- **File validation**: Scripts check for empty files and malformed data
- **Encoding support**: UTF-8 encoding for international characters
- **Large file handling**: Progress tracking and memory optimization for large datasets

**GUI-Specific Issues**
- **Windows Forms dependency**: Requires full Windows environment (not Windows Core)
- **Display scaling**: Works with high-DPI displays
- **PowerShell version**: Requires PowerShell 5.0 or later

### Detailed Error Messages

| Error Type | Cause | Solution |
|------------|-------|----------|
| "Module installation failed after X attempts" | Network issues or permissions | Run PowerShell as administrator or check internet connection |
| "User not found or no data returned" | Invalid email or user doesn't exist | Verify email addresses in your input data |
| "Access denied to the group" | Insufficient permissions | Ensure you have Group.Read.All permissions |
| "Throttling detected - retrying" | API rate limits | Script automatically retries with exponential backoff |
| "CSV file is empty or malformed" | Invalid CSV structure | Check CSV has headers and data rows |
| "Cannot write to output location" | File permissions or disk space | Check folder permissions and available disk space |

### Recovery Features

**Automatic Backup**
- Existing output files are automatically backed up before overwriting
- Backup files include timestamp for easy identification

**Alternative Save Locations**
- If primary output location fails, scripts attempt to save to temp directory
- User is notified of alternative save location

**Connection Recovery**
- Multiple connection attempts with increasing delays
- Detailed connection status and troubleshooting guidance

**Data Validation**
- Email format validation before processing
- Empty record detection and skipping
- Progress tracking with ability to identify problematic records

## Examples

### Example 1: Basic Sign-in Check
```powershell
# Create input file
@"
Email
user1@company.com
user2@company.com
"@ | Out-File -FilePath "users.csv" -Encoding UTF8

# Run the script
.\Check-UserSignIns.ps1 -InputCsvPath "users.csv"
```

### Example 2: Custom Output Location
```powershell
.\Check-UserSignIns.ps1 -InputCsvPath "C:\Data\employees.csv" -OutputCsvPath "C:\Reports\signin-report.csv"
```

### Example 3: Batch Processing Large Files
```powershell
# For large user lists, the script automatically handles throttling
.\Check-UserSignIns.ps1 -InputCsvPath "all-users.csv" -OutputCsvPath "comprehensive-report.csv"
```

## Security Considerations

- The script uses secure Microsoft Graph authentication
- No passwords or secrets are stored in the script
- Authentication tokens are handled by the Microsoft Graph SDK
- Always run from a secure, trusted environment

## Support

For issues or questions:
1. Check the troubleshooting section above
2. Verify all prerequisites are met
3. Review the error messages for specific guidance
4. Ensure you have the latest version of Microsoft Graph modules

## License

This script is provided as-is for educational and administrative purposes.

---

**Last Updated:** October 2024
**PowerShell Version:** 5.1+ or PowerShell Core 7+
**Microsoft Graph SDK Version:** Latest