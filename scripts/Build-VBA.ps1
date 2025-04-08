# Summary:
# This PowerShell script automates the process of importing VBA modules into an Excel workbook.
# It retrieves the current working directory, constructs the path to the Excel file,
# and imports all .bas files from a specified folder into the workbook.
# It then saves and closes the workbook, and cleans up the COM objects.

# Get the current working directory
$currentDir = (Get-Location).Path
Write-Host "Current working directory: $currentDir"

# Define the Excel file path
$excelFile = Join-Path $currentDir "skeleton.xlsm"

# Create a new Excel application
$excelApp = New-Object -ComObject Excel.Application

# Open the workbook
$wb = $excelApp.Workbooks.Open($excelFile)

# Check if the workbook is opened successfully
if ($null -eq $wb) {
    Write-Host "Error: Failed to open the workbook: $excelFile"
    exit 1
}
# Make Excel visible (uncomment if needed)
# $excelApp.Visible = $true

# Define the module folder path
$moduleFolder = Join-Path $currentDir "src\skeleton.xlsm\Modules"

Write-Host "Module folder path: $moduleFolder"

# Check if the module folder exists
if (Test-Path $moduleFolder) {
    # First check if there are any .bas files
    $basFiles = Get-ChildItem -Path $moduleFolder -Filter *.bas
    if ($basFiles.Count -gt 0) {
        Write-Host "Found $($basFiles.Count) .bas files to import"
        # Loop through each file in the module folder
        $basFiles | ForEach-Object {
            Write-Host "Importing $($_.Name)..."
            # Import the module into the workbook

                # Before making the change, double-ckeck if the VBOM is enabled
                # HKCU:\Software\Microsoft\Office\16.0\Common\TrustCenter
                # Just check the registry entry
                $regPath = "HKCU:\Software\Microsoft\Office\16.0\Common\TrustCenter"
                if (-not (Test-Path $regPath)) {
                    Write-Host "Warning: Registry path not found: $regPath"
                    Write-Host "Please enable Access to the VBA project object model in Excel Trust Center settings."
                    exit 1
                }
                # Check if the AccessVBOM property is set to 1
                $accessVBOM = Get-ItemProperty -Path $regPath -Name AccessVBOM -ErrorAction SilentlyContinue
                if ($null -eq $accessVBOM) {
                    Write-Host "Warning: AccessVBOM property not found. Please enable Access to the VBA project object model in Excel Trust Center settings."
                    exit 1
                } elseif ($accessVBOM.AccessVBOM -ne 1) {
                    Write-Host "Warning: AccessVBOM is not enabled. Please enable Access to the VBA project object model in Excel Trust Center settings."
                    exit 1
                } elseif ($accessVBOM.AccessVBOM -eq 1) {
                    Write-Host "AccessVBOM is enabled. Proceeding with import..."
                }

            try {
                
                $vbProject = $wb.VBProject
                # Check if the VBProject is accessible
                if ($null -eq $vbProject) {
                     
                        # We Close the Excel application and re-open it
                        Write-Host "VBProject is not accessible. Attempting to re-open Excel application..."
                        $excelApp.Quit()
                        Start-Sleep -Seconds 2
                        $excelApp = New-Object -ComObject Excel.Application
                        $wb = $excelApp.Workbooks.Open($excelFile)
                        
                        $vbProject = $wb.VBProject

                        if ($null -eq $vbProject) {
                            Write-Host "VBProject is still not accessible after re-opening Excel."
                            # Throw an error to trigger the catch block
                            exit 1
                        } else {
                            Write-Host "VBProject is now accessible after re-opening Excel."
                        }
                }

                $vbProject.VBComponents.Import($_.FullName)
                
                Write-Host "Successfully imported $($_.Name)"
            } catch {
                Write-Host "Failed to import $($_.Name): $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "Warning: No .bas files found in $moduleFolder"
    }
} else {
    Write-Host "Module folder not found: $moduleFolder"
    exit 1
}

# Save the workbook
$wb.Save()

# Close the workbook
$wb.Close()

# Quit Excel
$excelApp.Quit()

# Clean up
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
Remove-Variable -Name wb, excelApp

Write-Host "Script completed successfully."