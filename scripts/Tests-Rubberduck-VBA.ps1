# Create a function for Rubberduck testing
function Test-WithRubberduck {
    param (
        [Parameter(Mandatory=$true)]
        $officeApp
    )
    
    $rubberduckAddin = $null
    $rubberduck = $null
    try {
        $rubberduckAddin = $officeApp.COMAddIns.Item("Rubberduck.Extension")
        if ($null -eq $rubberduckAddin) {
            Write-Host "🔴 Error: Rubberduck add-in not found. Please install it first."
            return $false
        }
        Write-Host "Rubberduck add-in found."

        $rubberduck = $rubberduckAddin.Object
        if ($null -eq $rubberduck) {
            Write-Host "🔴 Error: Rubberduck object not found. Please ensure it is properly installed."
            return $false
        }
        Write-Host "Rubberduck object found."

        # Run all tests in the active VBA project
        $logPath = "$env:temp\RubberduckTestLog.txt"
        $rubberduck.RunAllTests($logPath)
        Write-Host "All tests executed successfully."

        # Wait for tests to complete (optional: add a timeout)
        Start-Sleep -Seconds 10

        # Retrieve test results from the log file and display each line in the console
        # For each line if it starts with "Succeeded", add "✅" to the line, otherwise add "❌"
        if (Test-Path $logPath) {
            $results = Get-Content -Path $logPath
            Write-Host "Test results:"
            foreach ($line in $results) {
                if ($line -match "Succeeded") {
                    Write-Host "✅ $line"
                } else {
                    Write-Host "❌ $line"
                }
            }
        } else {
            Write-Host "🔴 Error: Log file not found. Please check the installation."
            return $false
        }
        
        # Make sure to release the COM object
        if ($null -ne $rubberduck) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rubberduck) | Out-Null
            Write-Host "Released Rubberduck COM object"
        }
        
        return $true
    }
    catch {
        Write-Host "🔴 Error: Could not access Rubberduck add-in: $($_.Exception.Message)"
        return $false
    }
}