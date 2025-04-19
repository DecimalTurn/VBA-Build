function Write-TimedMessage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [switch]$StartNewStep
    )
    
    $stepTime = $stepTimer.Elapsed.ToString("hh\:mm\:ss\.fff")
    $totalTime = $mainTimer.Elapsed.ToString("hh\:mm\:ss\.fff")
    
    Write-Host "[$totalTime total | $stepTime step] $Message"
    
    if ($StartNewStep) {
        $script:stepTimer.Restart()
    }
}