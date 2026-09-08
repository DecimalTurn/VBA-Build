<#
.SYNOPSIS
    Refreshes the committed vendors/msaccess-vcs-build tree from the
    submodules/msaccess-vcs-build submodule (the DecimalTurn/msaccess-vcs-build
    rate-limiting fork at its pinned commit).

.DESCRIPTION
    Copies the whole upstream tree (Build.ps1, action.yml, README.md, LICENSE,
    scripts/, examples/, ...) except .git into vendors/msaccess-vcs-build, then
    writes SOURCE.txt (source URL + commit SHA, NO timestamp so the output is
    byte-idempotent). Run by maintainers/devs and by the CI sync-guard job.
    At runtime Main.ps1 only uses the committed files under vendors/ - the
    submodule is never needed there.

.EXAMPLE
    pwsh -File scripts/Bundle-Vendors.ps1
#>
[CmdletBinding()]
param(
    [string]$SubmodulePath = "submodules/msaccess-vcs-build",
    [string]$VendorPath     = "vendors/msaccess-vcs-build"
)
$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot   # scripts/ -> repo root

# --- 1) Ensure the submodule is loaded --------------------------------
Push-Location $RepoRoot
if (-not (Test-Path "$SubmodulePath/.git") -or
    -not (Test-Path "$SubmodulePath/Build.ps1")) {
    Write-Host "Initializing submodule $SubmodulePath ..."
    git submodule update --init --recursive
    if ($LASTEXITCODE -ne 0) { throw "git submodule update failed." }
}
if (-not (Test-Path "$SubmodulePath/Build.ps1")) {
    throw "Submodule not populated at $SubmodulePath. Run:  git submodule update --init --recursive"
}

# --- 2) Record source commit for provenance ---------------------------
$srcCommit = git -C $SubmodulePath rev-parse HEAD
$srcUrl    = git -C $SubmodulePath remote get-url origin

# --- 3) Copy the whole upstream tree (decision: vendor everything) -----
# Copy every file/folder in the submodule (Build.ps1, action.yml, README.md,
# LICENSE, scripts/, examples/, ...) except .git, so the vendored scripts always
# match the pinned commit exactly and optional features (compile, app-config,
# AccUnit) never hit a missing-file error later. Clear the destination first so a
# submodule commit that removed a file also removes its stale vendored copy.
if (Test-Path $VendorPath) {
    Remove-Item -Path $VendorPath -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $VendorPath | Out-Null
Get-ChildItem -Path $SubmodulePath -Force | Where-Object { $_.Name -ne '.git' } |
    Copy-Item -Destination $VendorPath -Recurse -Force

# --- 4) Apply the re-appliable edits from the plan (currently none) ----
# Reserved for deterministic string replaces. The bundler always re-copies from
# the pristine submodule, so any future edits re-apply cleanly and stay idempotent.

# --- 5) Write provenance file (NO "Bundled:" timestamp) -----------------
# A timestamp would make every regeneration produce a different SOURCE.txt,
# breaking the idempotency check and the CI sync-guard (`git diff --exit-code`).
# Write with explicit LF so the bytes are identical on Windows and Linux
# regardless of the platform's default newline.
$sourceContent = "Source:  $srcUrl`n" +
                 "Commit:  $srcCommit`n" +
                 "Note: vendors/msaccess-vcs-build/action.yml is reference-only and is NOT executed.`n" +
                 "Run 'pwsh -File scripts/Bundle-Vendors.ps1' to regenerate.`n"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$sourceFilePath = Join-Path $RepoRoot (Join-Path $VendorPath 'SOURCE.txt')
[System.IO.File]::WriteAllText($sourceFilePath, $sourceContent, $utf8NoBom)

Pop-Location
Write-Host "Vendored msaccess-vcs-build into $VendorPath (commit $srcCommit)"
