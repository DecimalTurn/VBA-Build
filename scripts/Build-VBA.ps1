# Summary:
# This PowerShell script automates the process of importing VBA modules into an Office document.
# It retrieves the current working directory, constructs the path to the Office file,
# and imports all .bas files from a specified folder into the document.
# It then saves and closes the document, and cleans up the COM objects.

# Load utiliies
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Screenshot.ps1"

# Args
$folderName = $args[0]
$officeAppName = $args[1]

if (-not $folderName) {
    Write-Host "🔴 Error: No folder name specified. Usage: Build-VBA.ps1 <FolderName>"
    exit 1
}

if (-not $officeAppName) {
    Write-Host "🔴 Error: No Office application specified. Usage: Build-VBA.ps1 <FolderName> <officeAppName>"
    exit 1
}

# Import utility functions
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
. "$scriptPath/utils/Path.ps1"

$currentDir = (Get-Location).Path + "/"
Write-Host "Current directory: $currentDir"

$vbaSrcDir = GetAbsPath -path $folderName -basePath $currentDir
Write-Host "Using source directory: $vbaSrcDir"

$fileName = GetDirName $vbaSrcDir
$fileNameNoExt = $fileName.Substring(0, $fileName.LastIndexOf('.'))
$fileExtension = $fileName.Substring($fileName.LastIndexOf('.') + 1)

$outputDir = (DirUp $vbaSrcDir) + "out/"

Write-Host "Output directory: $outputDir"

$outputFilePath = $outputDir + $fileName

# Make sure the output file already exists
if (-not (Test-Path $outputFilePath)) {
    Write-Host "🔴 Error: Output file not found: $outputFilePath"
    exit 1
}

$screenshotDir = (DirUp $outputDir) + "screenshots/"
if (-not (Test-Path $screenshotDir)) {
    New-Item -ItemType Directory -Path $screenshotDir -Force | Out-Null
    Write-Host "Created screenshot directory: $screenshotDir"
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
    Write-Host "🔴 Error: VBOM access is not enabled. Please enable it in the Trust Center settings."
    exit 1
}

# Create the application instance
$officeApp = New-Object -ComObject "$officeAppName.Application"

# Make app visible (uncomment if needed)
$officeApp.Visible = $true

# Check if the application instance was created successfully
if ($null -eq $officeApp) {
    Write-Host "🔴 Error: Failed to create COM object for $officeApp"
    exit 1
}

# Open the document
if ($officeAppName -eq "Excel") {
    $doc = $officeApp.Workbooks.Open($outputFilePath)
} elseif ($officeAppName -eq "Word") {
    $doc = $officeApp.Documents.Open($outputFilePath)
} elseif ($officeAppName -eq "PowerPoint") {
    $doc = $officeApp.Presentations.Open($outputFilePath)
} elseif ($officeAppName -eq "Access") {
    try {
        # Check if the file exists and is accessible
        if (-not (Test-Path $outputFilePath)) {
            Write-Host "🔴 Error: Access database file not found: $outputFilePath"
            exit 1
        }
        
        # Close any existing database
        try {
            if ($null -ne $officeApp.CurrentDb) {
                $officeApp.CloseCurrentDatabase()
                Start-Sleep -Seconds 1
            }
        } catch {
            # Ignore errors when checking CurrentDb - it may throw if no database is open
        }
        
        # First attempt - normal open
        Write-Host "Attempting to open Access database in shared mode..."
        $doc = $officeApp.OpenCurrentDatabase($outputFilePath)
        
        # Check if VBA project is accessible
        $vbaAccessible = $false
        try {
            if ($null -ne $doc -and $null -ne $doc.VBProject) {
                $vbaAccessible = $true
                Write-Host "VBA Project accessible in shared mode"
            }
        } catch {
            Write-Host "Cannot access VBA Project in shared mode: $($_.Exception.Message)"
        }
        
        # If VBA project isn't accessible, try reopening with different flags
        if (-not $vbaAccessible) {
            Write-Host "Attempting to reopen database with exclusive access..."
            try {
                $officeApp.CloseCurrentDatabase()
                Start-Sleep -Seconds 1
                
                # Try a different approach for exclusive access
                # Some versions of Access use different methods
                
                # For newer Access versions
                $officeApp.OpenCurrentDatabase($outputFilePath, $true) # $true = exclusive mode
                $doc = $officeApp.CurrentDb
            } catch {
                Write-Host "Failed with exclusive flag: $($_.Exception.Message)"
                
                # For older Access versions - one last attempt
                try {
                    $officeApp.Quit()
                    Start-Sleep -Seconds 2
                    $officeApp = New-Object -ComObject "Access.Application"
                    $officeApp.Visible = $true
                    
                    # Open in exclusive mode
                    # Use a slightly different technique as a last resort
                    $officeApp.OpenCurrentDatabase($outputFilePath, $true)
                    $doc = $officeApp.CurrentDb
                } catch {
                    Write-Host "🔴 Error: Failed to open the database after multiple attempts: $($_.Exception.Message)"
                    exit 1
                }
            }
        }
    }
    catch {
        Write-Host "🟡 Warning: Error opening Access database: $($_.Exception.Message)"
        Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
        
        # One more attempt after error
        try {
            $officeApp.Quit()
            Start-Sleep -Seconds 2
            $officeApp = New-Object -ComObject "Access.Application"
            $officeApp.Visible = $true
            $officeApp.DefaultOpenExclusive = $true
            $doc = $officeApp.OpenCurrentDatabase($outputFilePath)
        }
        catch {
            Write-Host "🔴 Error: Failed to open the database after multiple attempts: $($_.Exception.Message)"
            exit 1
        }
    }
} else {
    Write-Host "Error: Unsupported Office application: $officeAppName"
    exit 1
}

# Check if the document was opened successfully
if ($null -eq $doc) {
    Write-Host "🔴 Error: Failed to open the document: $outputFilePath"
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    exit 1
}

# Define the module folder path
$moduleFolder = GetAbsPath -path "$folderName/Modules" -basePath $currentDir
Write-Host "Module folder path: $moduleFolder"

# Define the class modules folder path
$classModulesFolder = GetAbsPath -path "$folderName/Class Modules" -basePath $currentDir
Write-Host "Class Modules folder path: $classModulesFolder"

# Define the forms folder path
$formsFolder = GetAbsPath -path "$folderName/Forms" -basePath $currentDir
Write-Host "Forms folder path: $formsFolder"

#Check if the module folder does not exist create an empty one
if (-not (Test-Path $moduleFolder)) {
    Write-Host "Module folder not found: $moduleFolder"
    New-Item -ItemType Directory -Path $moduleFolder -Force | Out-Null
    Write-Host "Created module folder: $moduleFolder"
}

#Check if the class modules folder does not exist create an empty one
if (-not (Test-Path $classModulesFolder)) {
    Write-Host "Class Modules folder not found: $classModulesFolder"
    New-Item -ItemType Directory -Path $classModulesFolder -Force | Out-Null
    Write-Host "Created class modules folder: $classModulesFolder"
}

#Check if the forms folder does not exist create an empty one
if (-not (Test-Path $formsFolder)) {
    Write-Host "Forms folder not found: $formsFolder"
    New-Item -ItemType Directory -Path $formsFolder -Force | Out-Null
    Write-Host "Created forms folder: $formsFolder"
}

# Import class modules first (.cls files)
$clsFiles = Get-ChildItem -Path $classModulesFolder -Filter *.cls -ErrorAction SilentlyContinue
Write-Host "Found $($clsFiles.Count) .cls files to import"

# Loop through each class module file
$clsFiles | ForEach-Object {
    Write-Host "Importing class module $($_.Name)..."
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
        
        Write-Host "Successfully imported class module $($_.Name)"
    } catch {
        Write-Host "Failed to import class module $($_.Name): $($_.Exception.Message)"
    }
}

# Import form modules (.frm files)
$frmFiles = Get-ChildItem -Path $formsFolder -Filter *.frm -ErrorAction SilentlyContinue
Write-Host "Found $($frmFiles.Count) .frm files to import"

# Loop through each form file
$frmFiles | ForEach-Object {
    Write-Host "Importing form $($_.Name)..."
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
        
        Write-Host "Successfully imported form $($_.Name)"
    } catch {
        Write-Host "Failed to import form $($_.Name): $($_.Exception.Message)"
    }
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

Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"

# Save the document
Write-Host "Saving document..."
try {
    if ($officeAppName -eq "PowerPoint") {
        # For PowerPoint, use SaveAs with the same file name to force save
        $doc.SaveAs($outputFilePath)
        Write-Host "PowerPoint presentation saved using SaveAs method"
    } else {
        $doc.Save()
        Write-Host "Document saved successfully"
    }
} catch {
    Write-Host "Warning: Could not save document: $($_.Exception.Message)"
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    
    # Alternative approach for PowerPoint if SaveAs fails
    if ($officeAppName -eq "PowerPoint") {
        try {
            # Try saving with a temporary file name and then renaming
            $tempPath = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.pptm'
            Write-Host "Attempting to save to temporary location: $tempPath"
            $doc.SaveAs($tempPath)
            Write-Host "Temporary file saved successfully at: $tempPath"
            
            # Close the document and application
            $doc.Close()
            $officeApp.Quit()
            
            # Release COM objects
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null
            
            # Wait a moment for resources to be released
            Start-Sleep -Seconds 2
            
            # Copy the temp file to the intended destination
                Copy-Item -Path $tempPath -Destination $outputFilePath -Force
                Remove-Item -Path $tempPath -Force
            
            Write-Host "Document saved using alternative method"
            
            # Skip the rest of the cleanup as we've already done it
            Write-Host "VBA import completed successfully."
            exit 0
        } catch {
            Write-Host "Error: Alternative save method also failed: $($_.Exception.Message)"
            Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
        }
    }
}

# Call the WriteToFile macro to check if the module was imported correctly
try {
    
    $vbaModule = $doc.VBProject.VBComponents.Item(1)
    if ($null -eq $vbaModule) {
        Write-Host "Error: No VBA module found in the document."
        Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
        exit 1
    }
    Write-Host "VBA module found: $($vbaModule.Name)"

    $macroName = "WriteToFile"
    Write-Host "Macro to execute: $macroName"
    Write-Host "Application state before macro execution: Type=$($officeApp.GetType().FullName)"
    $officeApp.Run($macroName)
    Write-Host "Macro finished"

    # Check if the file was created successfully with the correct content
    $outputFile = "$outputDir/$fileNameNoExt.txt"
    if (Test-Path $outputFile) {
        $fileContent = Get-Content -Path $outputFile
        if ($fileContent -eq "Hello, World!") {
            Write-Host "🟢 Macro executed successfully and file content is correct."
        } else {
            Write-Host "🟡 Warning: Macro executed, but file content is incorrect.: $fileContent"
        }

        # Delete the output file after checking
        Remove-Item -Path $outputFile -Force
        Write-Host "Output test file deleted successfully."

    } else {
        Write-Host "🟡 Warning: Macro executed, but output file was not created."
    }

} catch {
    Take-Screenshot -OutputPath "${screenshotDir}Screenshot_${fileNameNoExt}_{{timestamp}}.png"
    Write-Host "🟡 Warning: Could not execute macro ${macroName}: $($_.Exception.Message)"
    
}

# Close the document
try {
    $doc.Close()
    Write-Host "Document closed successfully"
} catch {
    Write-Host "Warning: Could not close document: $($_.Exception.Message)"
}

# Quit the application
try {
    $officeApp.Quit()
    Write-Host "Application closed successfully"
} catch {
    Write-Host "Warning: Could not quit application: $($_.Exception.Message)"
}

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null
Remove-Variable -Name doc, officeApp
Write-Host "VBA import completed successfully."
