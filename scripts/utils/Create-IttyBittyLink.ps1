function Create-IttyBittyLink {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ScreenshotDirectory,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath
    )
    
    # Check if the directory exists
    if (-not (Test-Path -Path $ScreenshotDirectory)) {
        Write-Error "Screenshot directory does not exist: $ScreenshotDirectory"
        return $null
    }
    
    # Get all PNG files in the directory
    $screenshots = Get-ChildItem -Path $ScreenshotDirectory -Filter "*.png"
    
    if ($screenshots.Count -eq 0) {
        Write-Warning "No screenshots found in $ScreenshotDirectory"
        return $null
    }
    
    # Create HTML content with embedded data URLs for screenshots
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>VBA Build Screenshots</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .screenshot { margin-bottom: 30px; }
        img { max-width: 100%; border: 1px solid #ddd; }
    </style>
</head>
<body>
    <h1>VBA Build Screenshots</h1>
    <p>Build Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
"@
    
    foreach ($screenshot in $screenshots) {
        Write-Verbose "Processing $($screenshot.Name)"
        
        # Read image file as bytes and convert to base64
        $bytes = [System.IO.File]::ReadAllBytes($screenshot.FullName)
        $base64Image = [Convert]::ToBase64String($bytes)
        $dataUrl = "data:image/png;base64,$base64Image"
        
        # Add image to HTML content
        $htmlContent += @"
    <div class="screenshot">
        <h2>$($screenshot.Name)</h2>
        <img src="$dataUrl" alt="$($screenshot.Name)" />
    </div>
"@
    }
    
    # Close HTML document
    $htmlContent += @"
</body>
</html>
"@
    
    Write-Verbose "HTML content created, size: $($htmlContent.Length) chars"
    
    # Create temporary directory for processing
    $tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    
    try {
        # Path for HTML content
        $htmlPath = [System.IO.Path]::Combine($tempDir, "content.html")
        
        # Save HTML content to file
        [System.IO.File]::WriteAllText($htmlPath, $htmlContent)
        
        # Check if xz is available (needed for LZMA compression)
        $xzPath = $null
        
        # Try to find xz in common paths or download it if needed
        $xzPaths = @(
            "xz",                          # If in PATH
            "C:\Program Files\Git\usr\bin\xz.exe", # Git for Windows
            "C:\msys64\usr\bin\xz.exe",    # MSYS2
            "C:\xz\xz.exe"                 # Custom location
        )
        
        foreach ($path in $xzPaths) {
            try {
                if (Get-Command $path -ErrorAction SilentlyContinue) {
                    $xzPath = $path
                    break
                }
            } catch {
                # Continue to next path
            }
        }
        
        # If xz not found, use alternative method
        if (-not $xzPath) {
            Write-Warning "xz utility not found. Using PowerShell for compression (not optimal)."
            
            # Use PowerShell's Compress-Archive as fallback (not LZMA, but will work)
            $compressedPath = [System.IO.Path]::Combine($tempDir, "content.zip")
            Compress-Archive -Path $htmlPath -DestinationPath $compressedPath
            $compressedBytes = [System.IO.File]::ReadAllBytes($compressedPath)
            $base64Compressed = [Convert]::ToBase64String($compressedBytes)
        }
        else {
            # Use xz for LZMA compression
            $compressedPath = [System.IO.Path]::Combine($tempDir, "content.lzma")
            
            # Execute xz command to compress HTML content
            $xzCommand = "$xzPath --format=lzma -9 -c `"$htmlPath`" > `"$compressedPath`""
            $result = Invoke-Expression "cmd /c $xzCommand 2>&1"
            
            if (-not (Test-Path $compressedPath)) {
                Write-Error "Failed to compress content: $result"
                return $null
            }
            
            # Read compressed file and convert to base64
            $compressedBytes = [System.IO.File]::ReadAllBytes($compressedPath)
            $base64Compressed = [Convert]::ToBase64String($compressedBytes)
        }
        
        # Generate itty.bitty link
        $ittyBittyLink = "https://itty.bitty.site/#/$base64Compressed"
        
        # Save the link to a file if an output path was provided
        if ($OutputPath) {
            Set-Content -Path $OutputPath -Value $ittyBittyLink
            Write-Verbose "Itty.bitty link saved to: $OutputPath"
        }
        
        return $ittyBittyLink
    }
    finally {
        # Clean up temporary directory
        if (Test-Path $tempDir) {
            Remove-Item -Path $tempDir -Recurse -Force
        }
    }
}

# This makes the function available when the script is dot-sourced
# Example: . .\Create-IttyBittyLink.ps1