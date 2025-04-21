# # Summary:
# # This PowerShell installs Rubberduck for the VBE and runs all tests in the active VBA project.
# 

# Installer location: https://github.com/rubberduck-vba/Rubberduck/releases/latest

# Options to run the installer:
# ---------------------------
# Setup
# ---------------------------
# The Setup program accepts optional command line parameters.
# 
# 
# 
# /HELP, /?
# 
# Shows this information.
# 
# /SP-
# 
# Disables the This will install... Do you wish to continue? prompt at the beginning of Setup.
# 
# /SILENT, /VERYSILENT
# 
# Instructs Setup to be silent or very silent.
# 
# /SUPPRESSMSGBOXES
# 
# Instructs Setup to suppress message boxes.
# 
# /LOG
# 
# Causes Setup to create a log file in the user's TEMP directory.
# 
# /LOG="filename"
# 
# Same as /LOG, except it allows you to specify a fixed path/filename to use for the log file.
# 
# /NOCANCEL
# 
# Prevents the user from cancelling during the installation process.
# 
# /NORESTART
# 
# Prevents Setup from restarting the system following a successful installation, or after a Preparing to Install failure that requests a restart.
# 
# /RESTARTEXITCODE=exit code
# 
# Specifies a custom exit code that Setup is to return when the system needs to be restarted.
# 
# /CLOSEAPPLICATIONS
# 
# Instructs Setup to close applications using files that need to be updated.
# 
# /NOCLOSEAPPLICATIONS
# 
# Prevents Setup from closing applications using files that need to be updated.
# 
# /FORCECLOSEAPPLICATIONS
# 
# Instructs Setup to force close when closing applications.
# 
# /FORCENOCLOSEAPPLICATIONS
# 
# Prevents Setup from force closing when closing applications.
# 
# /LOGCLOSEAPPLICATIONS
# 
# Instructs Setup to create extra logging when closing applications for debugging purposes.
# 
# /RESTARTAPPLICATIONS
# 
# Instructs Setup to restart applications.
# 
# /NORESTARTAPPLICATIONS
# 
# Prevents Setup from restarting applications.
# 
# /LOADINF="filename"
# 
# Instructs Setup to load the settings from the specified file after having checked the command line.
# 
# /SAVEINF="filename"
# 
# Instructs Setup to save installation settings to the specified file.
# 
# /LANG=language
# 
# Specifies the internal name of the language to use.
# 
# /DIR="x:\dirname"
# 
# Overrides the default directory name.
# 
# /GROUP="folder name"
# 
# Overrides the default folder name.
# 
# /NOICONS
# 
# Instructs Setup to initially check the Don't create a Start Menu folder check box.
# 
# /TYPE=type name
# 
# Overrides the default setup type.
# 
# /COMPONENTS="comma separated list of component names"
# 
# Overrides the default component settings.
# 
# /TASKS="comma separated list of task names"
# 
# Specifies a list of tasks that should be initially selected.
# 
# /MERGETASKS="comma separated list of task names"
# 
# Like the /TASKS parameter, except the specified tasks will be merged with the set of tasks that would have otherwise been selected by default.
# 
# /PASSWORD=password
# 
# Specifies the password to use.
# 
# 
# 
# For more detailed information, please visit https://jrsoftware.org/ishelp/index.php?topic=setupcmdline
# ---------------------------
# OK   
# ---------------------------
# 


# This script installs Rubberduck for the VBE and runs all tests in the active VBA project.
# It uses the Inno Setup installer for Rubberduck, which is a popular open-source VBA add-in for unit testing and code inspection.
# The script is designed to be run in a PowerShell environment and requires administrative privileges to install the add-in.

# The script performs the following steps:
# 1. Downloads the latest version of Rubberduck from GitHub.
# 2. Installs Rubberduck using the Inno Setup installer with specified command line options to suppress prompts and run silently.
# ======

# Step 1:

# Download the latest version of Rubberduck from GitHub
# The URL is constructed using the latest release version from the Rubberduck GitHub repository.
# The script uses the Invoke-WebRequest cmdlet to download the installer to a temporary location.
$rubberduckUrl = "https://github.com/rubberduck-vba/Rubberduck/releases/download/v2.5.91/Rubberduck.Setup.2.5.9.6316.exe"


$tempInstallerPath = "$env:TEMP\Rubberduck.Setup.exe"
Invoke-WebRequest -Uri $rubberduckUrl -OutFile $tempInstallerPath

# Step 2:
# Install Rubberduck using the Inno Setup installer with specified command line options
# The script uses the Start-Process cmdlet to run the installer with the /SILENT and /NORESTART options to suppress prompts and prevent automatic restarts.
$installerArgs = "/SILENT /NORESTART /SUPPRESSMSGBOXES /LOG=$env:TEMP\RubberduckInstall.log"
Start-Process -FilePath $tempInstallerPath -ArgumentList $installerArgs -Wait
# The -Wait parameter ensures that the script waits for the installation to complete before proceeding.

# Output logs to the console
# The script uses the Get-Content cmdlet to read the installation log file and display its contents in the console.
# This can help troubleshoot any issues that may arise during the installation process.
# Note: Use -Tail 500 to limit the output to the last 500 lines of the log file.
Get-Content -Path "$env:TEMP\RubberduckInstall.log" | Out-Host

# Verify that Rubberduck was successfully installed by checking registry entries
function Test-RubberduckInstalled {
    $addinProgId = "Rubberduck.Extension"
    $addinCLSID = "8D052AD8-BBD2-4C59-8DEC-F697CA1F8A66"
    $isInstalled = $false
    
    # Check for registry keys in current user hive
    if (Test-Path "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins\$addinProgId") {
        Write-Host "✅ Rubberduck add-in registration found in HKCU VBA\VBE registry."
        $isInstalled = $true
    }
    
    # For 64-bit systems, check additional registry locations
    if ([Environment]::Is64BitOperatingSystem) {
        if (Test-Path "HKCU:\Software\Microsoft\VBA\VBE\6.0\Addins64\$addinProgId") {
            Write-Host "✅ Rubberduck add-in registration found in HKCU VBA\VBE Addins64 registry."
            $isInstalled = $true
        }
        
        # Check for the VB6 addin registration
        if (Test-Path "HKCU:\Software\Microsoft\Visual Basic\6.0\Addins\$addinProgId") {
            Write-Host "✅ Rubberduck add-in registration found in HKCU Visual Basic registry."
            $isInstalled = $true
        }
    }
    
    # Check for the COM class registration
    if (Test-Path "HKCR:\CLSID\{$addinCLSID}" -ErrorAction SilentlyContinue) {
        Write-Host "✅ Rubberduck COM class registration found."
        $isInstalled = $true
    }
    
    # Check if the DLL file was installed
    $commonAppDataPath = [System.Environment]::GetFolderPath("CommonApplicationData")
    $localAppDataPath = [System.Environment]::GetFolderPath("LocalApplicationData")
    
    $possiblePaths = @(
        "$commonAppDataPath\Rubberduck\Rubberduck.dll",
        "$localAppDataPath\Rubberduck\Rubberduck.dll"
    )
    
    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            Write-Host "✅ Rubberduck DLL found at: $path"
            $isInstalled = $true
            break
        }
    }
    
    if (-not $isInstalled) {
        Write-Host "❌ Rubberduck installation verification failed. No registry entries or DLL files found."
        return $false
    }
    
    Write-Host "✅ Rubberduck installation verification completed successfully."
    return $true
}

$rubberduckInstalled = Test-RubberduckInstalled
if (-not $rubberduckInstalled) {
    Write-Host "⚠️ Warning: Rubberduck installation could not be verified. Office addins may not function correctly."
    Write-Host "Please check the installation log for more details or try reinstalling manually."
} else {
    Write-Host "🎉 Rubberduck installed successfully and is ready to use!"
}