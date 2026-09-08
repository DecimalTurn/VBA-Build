# Plan: Fold Access builds into `Main.ps1` (drop the `msaccess-vcs-build-all` subaction)

## 1. Goal & motivation

Today Access is built by the root action calling a static-unroll local composite (`msaccess-vcs-build-all`) that invokes `DecimalTurn/msaccess-vcs-build` up to 5 times (`Call action 0..4`) via `uses:`. This exists only because composite actions can't loop over `uses:`.

Target: `Main.ps1` detects Access folders and builds each one **directly** by invoking vendored copies of the `msaccess-vcs-build` PowerShell scripts. Result:
- No subaction, no `Call action 0..4` cap, no nested composite.
- Detection → digest-verify → build all folders happens in one PowerShell flow.
- The dependency is pinned by the submodule commit (dev) + committed vendored files (runtime).

## 2. Target layout

```
scripts/
  Bundle-Vendors.ps1          # (new) refresh vendors/ from the submodule
  Build-VBA.ps1 ...          # existing, untouched
vendors/
  msaccess-vcs-build/
    LICENSE                  # copied from submodule (see §7)
    SOURCE.txt               # generated: repo URL + commit SHA
    Build.ps1                # copied (light edits, §6)
    scripts/
      Build-Accdb.ps1
      Compile-Accdb.ps1
      Export-Accdb.ps1
      Install-AccUnit.ps1
      Install-msaccess-vcs.ps1
      Prepare-Application.ps1
      Remove-TrustedLocation.ps1
      Run-AccUnit-Tests.ps1
      Set-TrustedLocation.ps1
      Get-GitHubHeaders.ps1  # from the rate-limiting fork
submodules/
  msaccess-vcs-build/        # (dev-only source, kept; NOT needed at runtime)
```

**Design decision:** `vendors/` is **committed**. CI never needs the submodule checked out — `Main.ps1` only touches committed files under `vendors/`. The submodule exists purely as the *source* that `Bundle-Vendors.ps1` re-syncs from.

## 3. `scripts/Bundle-Vendors.ps1` (new)

Run by maintainers/devs (and by the error message when files are missing). Pinned source: the submodule at `msaccess-vcs-build`.

```powershell
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

# --- 3) Copy the whole upstream tree (decision: vendor everything) -------
# Copy every file/folder in the submodule (Build.ps1, action.yml, README.md,
# LICENSE, scripts/, examples/, ...) except .git, so the vendored scripts
# always match the pinned commit exactly and optional features (compile,
# app-config, AccUnit) never hit a missing-file error later.
New-Item -ItemType Directory -Force -Path $VendorPath | Out-Null
Get-ChildItem -Path $SubmodulePath -Force | Where-Object { $_.Name -ne '.git' } |
    Copy-Item -Destination $VendorPath -Recurse -Force

# --- 4) Make minor edits (see §6) -------------------------------------
# e.g. sed-style patch on the copied files, implemented as small Replace calls

# --- 5) Write provenance file -----------------------------------------
@"
Source:  $srcUrl
Commit:  $srcCommit
Bundled: $(Get-Date -Format o)
Run 'pwsh -File scripts/Bundle-Vendors.ps1' to regenerate.
"@ | Set-Content "$VendorPath/SOURCE.txt"

Write-Host "Vendored msaccess-vcs-build into $VendorPath (commit $srcCommit)"
```

## 4. `Main.ps1` integration

`Main.ps1` already computes `$accessFolders`. Add:

```powershell
$VendorBuildDir = Join-Path $RepoRoot 'vendors/msaccess-vcs-build'

function Assert-VendoredVcsBuild {
    $required = @('Build.ps1','scripts/Build-Accdb.ps1','scripts/Install-msaccess-vcs.ps1','scripts/Get-GitHubHeaders.ps1')
    $missing = $required | Where-Object { -not (Test-Path (Join-Path $VendorBuildDir $_)) }
    if ($missing) {
        Write-Error @"
Vendored msaccess-vcs-build files are missing from '$VendorBuildDir':
  $($missing -join "`n  ")

Generate them by running (from the repo root, with the submodule initialized):
  pwsh -File scripts/Bundle-Vendors.ps1
If the submodule isn't loaded yet:
  git submodule update --init --recursive
  pwsh -File scripts/Bundle-Vendors.ps1
"@
        exit 1
    }
}

function Invoke-AccessBuilds {
    param([string[]]$AccessFolders)

    Assert-VendoredVcsBuild

    # --- digest-verify the add-in release ONCE (authenticated via GH_TOKEN) ---
    # ... reuse the existing curl + expected-sha256 compare from the subaction ...

    # --- install the add-in + trusted location ONCE; reuse for all folders (see §5) ---
    $install = & "$VendorBuildDir/scripts/Install-msaccess-vcs.ps1" `
        -vcsUrl $VcsUrl -TargetDir $RepoRoot -SetTrustedLocation $true
    if (-not $install.Success) {
        throw "Failed to install msaccess-vcs add-in."
    }
    $addInPath = $install.AddInPath   # e.g. <RepoRoot>\MSAccessVCS\Version Control.accda

    # --- one FRESH Access instance per build, reusing the installed add-in ---
    foreach ($folder in $AccessFolders) {
        Write-Host "Building Access database: $folder"
        $accdbPath = ( & pwsh -NoProfile -File "$VendorBuildDir/scripts/Build-Accdb.ps1" `
            -SourceDir    $folder `
            -TargetDir    "$SourceDir/out" `
            -VcsAddInPath $addInPath ) | Select-Object -Last 1
        if (-not $accdbPath -or -not (Test-Path $accdbPath)) {
            throw "Access build failed for $folder (no output file)."
        }
    }
}
```

> This mirrors today's behavior (one fresh Access instance per folder) while
> installing the add-in only once, minimizing download/install time. If
> in-process reuse is ever desired for benchmarking, it can be toggled later
> behind a flag.

## 5. Chosen approach: install once, fresh Access instance per build

`Build.ps1` has no `-VcsAddInPath` parameter — it *installs* the add-in when `vcsUrl` is set and only then passes the installed path to `Build-Accdb.ps1`. With an empty `vcsUrl` it falls back to the APPDATA default path, so a pre-installed add-in elsewhere would be ignored. Because we install once but still want a clean Access process per build, we bypass `Build.ps1` and call `scripts/Build-Accdb.ps1` directly:

- **Once per run:** `Install-msaccess-vcs.ps1` installs the add-in and sets the trusted location; the returned `AddInPath` is reused for every folder.
- **Per folder (fresh):** spawn a new child `pwsh` that opens a brand-new Access instance and builds that one folder via `Build-Accdb.ps1 -SourceDir … -VcsAddInPath <AddInPath>`.
- **Net effect:** N add-in downloads/installs → 1; N fresh Access instances → still N (same isolation as today).

Rationale: minimizes download/install time while keeping each build isolated and matching today's per-`uses:` behavior. If in-process reuse is ever wanted (benchmarking only), add a `-ReuseProcess` flag later.

## 6. "Make minor edits" — what they actually are

Keep edits minimal and *re-appliable* in `Bundle-Vendors.ps1` (never hand-edit `vendors/`):

1. **Provenance header** on `Build.ps1` (comment): source repo + commit (also stored in `SOURCE.txt`).
2. **Default `vcsUrl`** on `Build.ps1` / `Install-msaccess-vcs.ps1`: change the `josef-poetzl/...latest` default to the `joyfullservice v5.0.1` API URL so a bare invocation isn't surprising (values are always passed explicitly by `Main.ps1` anyway).
3. **Auth** (`Get-GitHubHeaders.ps1`): already present in the fork — no edit needed; verify it's the file bundled.
4. **Trusted-location uniqueness**: `Set-TrustedLocation.ps1` already uses timestamped names; confirm repeated invocations in one process don't collide.
5. **`action.yml`**: keep in `vendors/` only for reference/documentation; it is **not** executed. (Document this in `SOURCE.txt`.)

`Bundle-Vendors.ps1` applies these as string replaces on the copied files, so a submodule update that overwrites them re-applies cleanly.

## 7. License & provenance

- Copy the submodule's `LICENSE` into `vendors/msaccess-vcs-build/LICENSE` **unchanged**.
- Generate `vendors/msaccess-vcs-build/SOURCE.txt` with: source URL, bundled commit SHA, bundle date, regeneration command.
- If any script files carry their own license headers, keep them intact (no stripping).
- Optional: add a top-level note in `README.md` that `vendors/msaccess-vcs-build` is vendored from `DecimalTurn/msaccess-vcs-build` at a pinned commit.

## 8. Root `action.yml` cleanup

After `Main.ps1` handles Access:

```yaml
- name: "Run VBA Build"        # now also builds Access internally
  ...
- name: "Expose GitHub token to environment"   # KEEP: Main.ps1 + digest curl read GH_TOKEN
  ...
```
- Delete the `Build Access Database (if detected)` step that `uses: ./subactions/msaccess-vcs-build-all`.
- Delete `msaccess-vcs-build-all` (and its `verify-vcs-digest` logic — fold into `Main.ps1`/helper).
- Keep `access-folders` / `has-access-database` outputs (Main.ps1 already writes them).
- Keep the submodule (it feeds `Bundle-Vendors.ps1`).

## 9. Missing-file experience (the user-facing error)

When `vendors/` isn't present (fresh clone where the maintainer forgot to run the bundler):

```
Error: Vendored msaccess-vcs-build files are missing from 'vendors/msaccess-vcs-build':
  Build.ps1
  scripts/Build-Accdb.ps1
  scripts/Install-msaccess-vcs.ps1

Generate them by running (from the repo root, with the submodule initialized):
  git submodule update --init --recursive
  pwsh -File scripts/Bundle-Vendors.ps1
```

CI only hits this if the commit didn't include `vendors/` — so add a CI guard job that runs `scripts/Bundle-Vendors.ps1` and fails on `git diff --exit-code` if it produces changes (keeps `vendors/` in sync).

## 10. Testing plan

1. `pwsh -File scripts/Bundle-Vendors.ps1` → `git status` shows only expected vendor changes + `SOURCE.txt`.
2. Dry-run edit check: re-run bundler → no further diff (idempotent).
3. Local compile of vendored scripts: `[System.Management.Automation.Language.Parser]::ParseFile(...)` on each copied `.ps1`.
4. CI: `Test Build VBA` on `vcs-v5` must still detect `Testing.accdb.src`, verify digest, and produce `tests/out/Testing.accdb`.
5. Negative test: temporarily rename `vendors/msaccess-vcs-build` → expect the clear error message above.
6. >5 Access folders regression: add >5 `.src` fixtures temporarily to confirm the cap is gone.

## 11. Rollback

The subaction and `submodules` stay in git history; reverting the fold-in is a `git revert`. `vendors/` commit is additive and doesn't affect non-Access builds.

---

**Decisions:**
1. Build execution: **install once, fresh Access instance per folder** (see §5).
2. Vendoring scope: **copy the whole upstream tree** into `vendors/msaccess-vcs-build` (see §3).
3. Keep the submodule in this repo — it is the source `Bundle-Vendors.ps1` copies from (runtime CI only uses committed `vendors/`).