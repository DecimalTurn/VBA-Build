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
