# This script checks if Rubberduck is installed and ready in Excel, then runs all tests in the active VBA project.

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true

# Open the workbook (replace with the path to your workbook)
$workbookPath = "C:\Path\To\YourWorkbook.xlsm"
$workbook = $excel.Workbooks.Open($workbookPath)

# Access the Rubberduck COM interface
$rubberduck = $excel.Application.COMAddIns.Item("Rubberduck").Object

# Ensure Rubberduck is ready
if ($rubberduck.IsReady()) {
    Write-Host "Rubberduck is ready. Running all tests..."

    # Run all tests in the active VBA project
    $rubberduck.RunAllTests()

    # Wait for tests to complete (optional: add a timeout)
    Start-Sleep -Seconds 10

    # Retrieve test results
    $results = $rubberduck.GetTestResults()
    foreach ($result in $results) {
        Write-Host "Test: $($result.Name) - Outcome: $($result.Outcome)"
    }
} else {
    Write-Host "Rubberduck is not ready. Please ensure it is properly configured."
}

# Close Excel
$workbook.Close($false)
$excel.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null