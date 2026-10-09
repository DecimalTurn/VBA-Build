<#
.SYNOPSIS
    Returns headers for GitHub API requests.

.DESCRIPTION
    Builds the HTTP headers used for GitHub API calls. When a token is available
    via the GH_TOKEN or GITHUB_TOKEN environment variable, an Authorization
    header is added only for HTTPS requests to api.github.com. This avoids hitting the
    unauthenticated API rate limit that GitHub-hosted runners frequently share.
#>
function Get-GitHubHeaders {
    param(
        [Parameter(Mandatory = $true)]
        [uri]$Uri
    )

    $headers = @{
        "User-Agent"           = "msaccess-vcs-build"
        "Accept"               = "application/vnd.github+json"
        "X-GitHub-Api-Version" = "2022-11-28"
    }

    $token = $env:GH_TOKEN
    if ([string]::IsNullOrEmpty($token)) {
        $token = $env:GITHUB_TOKEN
    }

    if ($Uri.IsAbsoluteUri -and $Uri.Scheme -eq "https" -and
        $Uri.Host -eq "api.github.com" -and
        -not [string]::IsNullOrEmpty($token)) {
        $headers["Authorization"] = "Bearer $token"
    }

    return $headers
}
