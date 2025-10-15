@{
    Root = 'd:\OneDrive - Xavier University of Louisiana\Documents\Scripts\Entra\Sign-in Last 90 Days\Check-UserSignIns-GUI.ps1'
    OutputPath = 'd:\OneDrive - Xavier University of Louisiana\Documents\Scripts\Entra\Sign-in Last 90 Days\out'
    Package = @{
        Enabled = $true
        Obfuscate = $false
        HideConsoleWindow = $false
        DotNetVersion = 'v4.6.2'
        FileVersion = '1.0.0'
        FileDescription = ''
        ProductName = ''
        ProductVersion = ''
        Copyright = ''
        RequireElevation = $false
        ApplicationIconPath = ''
        PackageType = 'Console'
    }
    Bundle = @{
        Enabled = $true
        Modules = $true
        # IgnoredModules = @()
    }
}
        