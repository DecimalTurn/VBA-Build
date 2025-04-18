# Summary:
# This PowerShell script automates the process of importing VBA modules into an Office document.
# It retrieves the current working directory, constructs the path to the Office file,
# and imports all .bas files from a specified folder into the document.
# It then saves and closes the document, and cleans up the COM objects.

$folderName = $args[0]
$officeAppName = $args[1]

if (-not $folderName) {
    Write-Host "Error: No folder name specified. Usage: Build-VBA.ps1 <FolderName>"
    exit 1
}

if (-not $officeAppName) {
    Write-Host "Error: No Office application specified. Usage: Build-VBA.ps1 <FolderName> <officeAppName>"
    exit 1
}

# Import utility functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Path.ps1"

$currentDir = (Get-Location).Path + "/"
$srcDir = GetAbsPath -path $folderName -basePath $currentDir

$fileName = GetDirName $srcDir
$fileNameNoExt = $fileName.Substring(0, $fileName.LastIndexOf('.'))
$fileExtension = $fileName.Substring($fileName.LastIndexOf('.') + 1)

$outputDir = (DirUp $srcDir) + "out/"
$outputFilePath = $outputDir + $fileName

# Make sure the output file already exists
if (-not (Test-Path $outputFilePath)) {
    Write-Host "Error: Output file not found: $outputFilePath"
    exit 1
}

# Allows to double-ckeck if the VBOM is enabled
# HKCU:\Software\Microsoft\Office\16.0\Common\TrustCenter
# Just check the registry entry
function Test-VBOMAccess {
    param (
        [string]$officeAppName
    )

    # Check if the VBOM is enabled
    $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\TrustCenter"
    if (-not (Test-Path $regPath)) {
        Write-Host "Warning: Registry path not found: $regPath"
        Write-Host "Please enable Access to the VBA project object model in Excel Trust Center settings."
        return $false
    }
    
    # Check if the AccessVBOM property is set to 1
    $accessVBOM = Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue
    if ($null -eq $accessVBOM) {
        Write-Host "Warning: AccessVBOM property not found. Please enable Access to the VBA project object model in Excel Trust Center settings."
        return $false
    } elseif ($accessVBOM.AccessVBOM -ne 1) {
        Write-Host "Warning: AccessVBOM is not enabled. Please enable Access to the VBA project object model in Excel Trust Center settings."
        return $false
    } elseif ($accessVBOM.AccessVBOM -eq 1) {
        Write-Host "AccessVBOM is enabled. Proceeding with import..."
        return $true
    }
    
    # This line should not be reached, but adding as a fallback
    return $false
}

# Check if VBOM access is enabled before attempting imports
if (-not (Test-VBOMAccess -officeAppName $officeAppName)) {
    Write-Host "Error: VBOM access is not enabled. Please enable it in the Trust Center settings."
    exit 1
}

# Create the application instance
$officeApp = New-Object -ComObject "$officeAppName.Application"

# Make app visible (uncomment if needed)
# $officeApp.Visible = $true

# Check if the application instance was created successfully
if ($null -eq $officeApp) {
    Write-Host "Error: Failed to create COM object for $officeApp"
    exit 1
}

# Open the document
if ($officeAppName -eq "Excel") {
    $doc = $officeApp.Workbooks.Open($outputFilePath)
} elseif ($officeApp -eq "Word") {
    $doc = $officeApp.Documents.Open($outputFilePath)
} elseif ($officeApp -eq "PowerPoint") {
    $doc = $officeApp.Presentations.Open($outputFilePath)
} elseif ($officeApp -eq "Access") {
    $doc = $officeApp.OpenCurrentDatabase($outputFilePath)
} else {
    Write-Host "Error: Unsupported Office application: $officeAppName"
    exit 1
}

# Check if the document was opened successfully
if ($null -eq $doc) {
    Write-Host "Error: Failed to open the document: $outputFilePath"
    exit 1
}

# Define the module folder path

$moduleFolder = Join-Path $currentDir "$folderName/Modules"
Write-Host "Module folder path: $moduleFolder"

#Check if the module folder does not exist create an empty one
if (-not (Test-Path $moduleFolder)) {
    Write-Host "Module folder not found: $moduleFolder"
    New-Item -ItemType Directory -Path $moduleFolder -Force | Out-Null
    Write-Host "Created module folder: $moduleFolder"
}

# First check if there are any .bas files
$basFiles = Get-ChildItem -Path $moduleFolder -Filter *.bas
Write-Host "Found $($basFiles.Count) .bas files to import"

# Loop through each file in the module folder
$basFiles | ForEach-Object {
    Write-Host "Importing $($_.Name)..."
    try {
        $vbProject = $doc.VBProject
        # Check if the VBProject is accessible
        if ($null -eq $vbProject) {
            Write-Host "VBProject is not accessible. Attempting to re-open the application..."
            $officeApp.Quit()
            Start-Sleep -Seconds 2
            $officeApp = New-Object -ComObject $officeApp.Application
            $doc = $officeApp.Workbooks.Open($outputFilePath)
            
            $vbProject = $doc.VBProject

            if ($null -eq $vbProject) {
                Write-Host "VBProject is still not accessible after re-opening the application."
                # Throw an error to trigger the catch block
                exit 1
            } else {
                Write-Host "VBProject is now accessible after re-opening the application."
            }
        }

        $vbProject.VBComponents.Import($_.FullName)
        
        Write-Host "Successfully imported $($_.Name)"
    } catch {
        Write-Host "Failed to import $($_.Name): $($_.Exception.Message)"
    }
}

# Take a screenshot of the Office application
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Screenshot.ps1"
Take-Screenshot -OutputPath "${outputDir}Screenshot_${fileNameNoExt}.png"

# Save the document
$doc.Save()
# Close the document
$doc.Close()
# Quit the application
$officeApp.Quit()

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null
Remove-Variable -Name doc, officeApp
Write-Host "VBA imported completed successfully."
