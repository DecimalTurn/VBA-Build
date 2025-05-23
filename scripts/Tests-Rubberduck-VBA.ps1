# Create a function for Rubberduck testing
function Test-WithRubberduck {
    param (
        [Parameter(Mandatory=$true)]
        $officeApp
    )
    
    # Display the VBE to activate the Rubberduck COM add-in
    $officeApp.CommandBars.ExecuteMso("VisualBasic")

    # Wait for a moment to ensure the VBE is fully loaded
    Start-Sleep -Seconds 5

    $rubberduckAddin = $null
    $rubberduck = $null
    try {
        $rubberduckAddin = $officeApp.VBE.AddIns("Rubberduck.Extension")
        if ($null -eq $rubberduckAddin) {
            Write-Host "🔴 Error: Rubberduck add-in not found."
            return $false
        }
        Write-Host "Rubberduck add-in found."

        $rubberduck = $rubberduckAddin.Object
        if ($null -eq $rubberduck) {
            Write-Host "🔴 Error: Rubberduck object not found."
            return $false
        }
        Write-Host "Rubberduck object found."

        # Check if Rubberduck is actually ready
        try {
            $isConnected = $rubberduck.IsConnected
            Write-Host "Rubberduck connection status: $isConnected"
            if (-not $isConnected) {
                Write-Host "Waiting for Rubberduck to connect..."
                Start-Sleep -Seconds 5  # Give it more time to connect
            }
        }
        catch {
            Write-Host "⚠️ Warning: Unable to check Rubberduck connection status: $($_.Exception.Message)"
            # Let's try to refresh the object
            Start-Sleep -Seconds 3
            $rubberduck = $rubberduckAddin.Object
        }
        

        # Run all tests in the active VBA project
        $logPath = "${env:temp}\RubberduckTestLog.txt"
        $rubberduck.RunAllTestsAndGetResults($logPath)
        
        # Wait for tests to complete with a timeout of 3 minutes
        $timeout = New-TimeSpan -Minutes 3
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
        $testCompleted = $false
        
        Write-Host "Waiting for tests to complete (timeout: 3 minutes)..."
        while ($stopwatch.Elapsed -lt $timeout -and -not $testCompleted) {
            if (Test-Path $logPath) {
                $content = Get-Content -Path $logPath -ErrorAction SilentlyContinue
                if ($null -ne $content -and $content.Count -gt 0) {
                    $testCompleted = $true
                    Write-Host "Tests completed in $([math]::Round($stopwatch.Elapsed.TotalSeconds, 2)) seconds."
                    break
                }
            }
            Start-Sleep -Seconds 2
        }
        
        $stopwatch.Stop()
        
        if (-not $testCompleted) {
            Write-Host "🔴 Error: Test execution timed out after 3 minutes."
            return $false
        }

        # Retrieve test results from the log file and display each line in the console
        # For each line if it starts with "Succeeded", add "✅" to the line, otherwise add "❌"
        if (Test-Path $logPath) {
            $results = Get-Content -Path $logPath
            Write-Host "Test results:"
            foreach ($line in $results) {
                if ($line -match "Succeeded") {
                    Write-Host "✅ $line"
                } elseif ($line -match "🟡 No tests were run") {
                    Write-Host "$line"
                } else {
                    Write-Host "❌ $line"
                }
            }
        } else {
            Write-Host "🔴 Error: Log file not found."
            return $false
        }

        # Delete the log file after processing
        Remove-Item -Path $logPath -ErrorAction SilentlyContinue
        Write-Host "Log file deleted."
        
        # Make sure to release the COM object
        if ($null -ne $rubberduck) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($rubberduck) | Out-Null
            Write-Host "Released Rubberduck COM object"
        }
        
        return $true
    }
    catch {
        Write-Host "🔴 Error while using Rubberduck add-in on line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
        return $false
    }
}