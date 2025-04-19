
function Invoke-ScriptWithTimeout {
    param (
        [string]$ScriptPath,
        [array]$Arguments,
        [int]$TimeoutSeconds = 300  # 5 minutes default timeout
    )
    
    $job = Start-Job -ScriptBlock {
        param($scriptPath, $args)
        & $scriptPath $args
    } -ArgumentList $ScriptPath, $Arguments
    
    $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
    
    if ($completed -eq $null) {
        Write-Host "Script execution timed out after $TimeoutSeconds seconds" -ForegroundColor Red
        Stop-Job -Job $job
        Remove-Job -Job $job -Force
        return $false
    } else {
        $result = Receive-Job -Job $job
        Remove-Job -Job $job
        Write-Host $result
        return $true
    }
}