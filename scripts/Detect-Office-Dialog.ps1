# The goal of this function is to detect if the Office dialog is open and return a boolean value indicating whether it is open or not.
# For example, if the user is trying to open a file in Excel and the dialog is open, the function should return true.
# It should also detect a VBA MsgBox dialog and return true.

# The function should be able to detect the dialog for all Office applications (Excel, Word, PowerPoint, Access).
# The function will take as an argument the office application as a COM object that is passed to it.

function Detect-OfficeDialog {
    param (
        [Parameter(Mandatory=$true)]
        [object]$officeApp
    )

   
    # Check if the dialog is open
    $dialogOpen = $false

    # Check for common dialog types
    $dialogTypes = @(
        "FileDialog",
        "MsgBox",
        "InputBox",
        "UserForm"
    )

    foreach ($dialogType in $dialogTypes) {
        if ($officeApp.$dialogType.Visible) {
            $dialogOpen = $true
            break
        }
    }



    return $dialogOpen
}

# Example usage
$appName = "Excel" # Change this to the desired Office application (Excel, Word, PowerPoint, Access)
$officeApp = New-Object -ComObject "$appName.Application"
$officeApp.Visible = $true # Make sure the application is visible
$doc = $officeApp.Workbooks.Open("C:\Users\leduc\Workbench\temp_2025\MsgBox-Test.xlsm") # Open a file to trigger the dialog
$officeApp.Run("AsyncFileDialog") # Run a macro that opens a file dialog
# Wait for a few seconds to allow the dialog to open
Start-Sleep -Seconds 5
$dialogOpen = Detect-OfficeDialog -officeApp $officeApp

Write-Host "Is there a dialog open in ${appName}? $dialogOpen"

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($officeApp) | Out-Null