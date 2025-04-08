# Normalize the path by replacing backslashes with forward slashes
function NormalizeFilePath {
    param (
        [string]$path
    )

    $normalizeFilePath = $path -replace '\\', '/'

    return $normalizeFilePath
}

# Normalize the path by replacing backslashes with forward slashes
# and ensuring it ends with a slash
function NormalizeDirPath {
    param (
        [string]$path
    )

    $normalizeFilePath = NormalizeFilePath($path)

    # Add a trailing slash if it doesn't exist
    if ($normalizeFilePath -notmatch '/$') {
        $normalizeFilePath += '/'
    }

    return $normalizeFilePath
}

function IsRelativePath {
    param (
        [string]$path
    )
    $path = NormalizeFilePath($path)
    return -not ($path -match '^[a-zA-Z]:/|^file://|^//')
}

function GetFileNameFromPath {
    param (
        [string]$path
    )
    $path = NormalizeFilePath($path)
    # If path has no "/", return the whole path since it might be a relative path from the current directory
    if ($path -notmatch '/') {
        return $path
    }
    return $path.Substring($path.LastIndexOf('/') + 1)
}

# Extract the file name without extension from the path
function GetFileNameWithoutExtension {
    param (
        [string]$path
    )
    $path = NormalizeFilePath($path)
    $temp = GetFileNameFromPath($path) 
    return $temp.Substring(0, $temp.LastIndexOf('.'))
}

function GetAbsPath {
    param (
        [string]$path,
        [string]$basePath
    )

    $path = NormalizeFilePath($path)
    $basePath = NormalizeDirPath($basePath)

    if (-not (IsRelativePath $path)) {
        return $path
    }    

    # If path starts with a ., remove it
    if ($path.StartsWith('.')) {
        $path = $path.Substring(1)
    }

    # If path starts with a /, remove it
    if ($path.StartsWith('/')) {
        $path = $path.Substring(1)
    }

    return $basePath + $path

    # Alternative:
    # return (Resolve-Path $path).Path
}

function DirUp{
    param (
        [string]$path
    )

    $path = NormalizeDirPath($path)
    $path = $path.TrimEnd('/')

    # Remove the last segment of the path
    $lastSlashIndex = $path.LastIndexOf('/')
    if ($lastSlashIndex -eq -1) {
        return ''
    }

    return $path.Substring(0, $lastSlashIndex) + '/'
}

function GetDirName {
    param (
        [string]$path
    )

    $path = NormalizeFilePath($path)

    if (IsRelativePath $path) {
        # If it stats with "/", add a "." to the beginning
        if ($path.StartsWith('/')) {
            $path = '.' + $path
        }
        $path = (Resolve-Path $path).Path
    }

    # Check if path is a file. If so, get the directory containing it
    if (Test-Path -Path $path -PathType Leaf) {
        # If it's a file, get the directory containing it
        $path = Split-Path -Parent $path
    }

    $path = NormalizeDirPath($path)
    $path = $path.TrimEnd('/')
    
    $DirName = $path.Substring($path.LastIndexOf('/') + 1)
    return $DirName
}