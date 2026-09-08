# Plan: Fold Access builds into `Main.ps1` (drop the `msaccess-vcs-build-all` subaction)

## 1. Goal & motivation

Today Access is built by the root action calling a static-unroll local composite (`msaccess-vcs-build-all`) that invokes `DecimalTurn/msaccess-vcs-build` up to 5 times (`Call action 0..4`) via `uses:`. This exists only because composite actions can't loop over `uses:`.

Target: `Main.ps1` detects Access folders and builds each one **directly** by invoking vendored copies of the `msaccess-vcs-build` PowerShell scripts. Result:
- No subaction, no `Call action 0..4` cap, no nested composite.
- Detection → digest-verify → install → per-folder build all happens in one PowerShell flow.
- The dependency is pinned by the submodule commit (dev) + committed vendored files (runtime).

> **Scope note:** `.accde` compile and `app-config` (`Prepare-Application.ps1`) support —
> including the `access-vcs-compile` / `access-vcs-config` inputs and the name-based
> compile rule — were added to this repo *after* the first draft of this plan. They are
> in scope: the folded-in flow must reproduce them (see §5), and the CI fixture
> `tests/AccessExecuteOnlyDatabase.accde` exercises the compile path (see §10).

## 2. Target layout

```
scripts/
  Bundle-Vendors.ps1          # (new) refresh vendors/ from the submodule
  Build-VBA.ps1 ...          # existing, untouched
vendors/
  msaccess-vcs-build/
    LICENSE                  # copied from submodule (see §7)
    SOURCE.txt               # generated: repo URL + commit SHA (NO timestamp, see §3)
    Build.ps1                # copied; EXECUTED per Access folder (approach (b), §5)
    action.yml               # reference only; NOT executed (documented in SOURCE.txt)
    scripts/
      Build-Accdb.ps1
      Compile-Accdb.ps1      # used for .accde builds
      Export-Accdb.ps1
      Install-AccUnit.ps1
      Install-msaccess-vcs.ps1   # used for the one-time install (§5)
      Prepare-Application.ps1    # used when access-vcs-config is set
      Remove-TrustedLocation.ps1
      Run-AccUnit-Tests.ps1
      Set-TrustedLocation.ps1
      Get-GitHubHeaders.ps1  # from the rate-limiting fork; auth for digest + install
submodules/
  msaccess-vcs-build/        # (dev-only source, kept; NOT needed at runtime)
```

**Design decision:** `vendors/` is **committed**. At runtime `Main.ps1` only touches
committed files under `vendors/` — the submodule is never needed. The submodule exists
purely as the *source* that `Bundle-Vendors.ps1` re-syncs from (only the bundler and the
CI sync-guard job in §9 need it initialized).

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

# --- 4) Apply the re-appliable edits from §6 (currently none required) ---
# Reserved for deterministic string replaces. The bundler always re-copies from
# the pristine submodule, so any future edits re-apply cleanly and stay idempotent.

# --- 5) Write provenance file (NO "Bundled:" timestamp) ------------------
# A timestamp would make every regeneration produce a different SOURCE.txt,
# breaking the idempotency test (§10 #2) and the CI sync-guard (`git diff --exit-code`).
@"
Source:  $srcUrl
Commit:  $srcCommit
Note: vendors/msaccess-vcs-build/action.yml is reference-only and is NOT executed.
Run 'pwsh -File scripts/Bundle-Vendors.ps1' to regenerate.
"@ | Set-Content "$VendorPath/SOURCE.txt"

Write-Host "Vendored msaccess-vcs-build into $VendorPath (commit $srcCommit)"
```

## 4. `Main.ps1` integration

Two working roots, kept distinct (see review #5): the **action checkout** where the
vendored scripts live is `$PSScriptRoot` (never `Get-Location`); the **consumer
workspace** holding `SourceDir` is the current directory. There is no shared `$RepoRoot`.

Extend the `param` block and plumb the values from `action.yml` (§8):

```powershell
param(
    [string]$SourceDir = "src",
    [string]$TestFramework = "none",
    [string]$OfficeAppDetection = "automatic",
    # --- Access (new) ---
    [string]$AccessVcsUrl     = "https://api.github.com/repos/joyfullservice/msaccess-vcs-addin/releases/tags/v5.0.1",
    [string]$AccessVcsSha     = "",  # expected sha256 of the add-in asset; empty = skip verify
    [string]$AccessVcsConfig  = "",  # app-config file for Prepare-Application; empty = skip
    [string]$AccessVcsCompile = ""   # force compile for ALL Access folders; empty = name-based
)
```

Add (near the other state, or just before use):

```powershell
$VendorBuildDir = Join-Path $PSScriptRoot 'vendors/msaccess-vcs-build'
```

Missing-file guard (unchanged intent):

```powershell
function Assert-VendoredVcsBuild {
    $required = @('Build.ps1','scripts/Build-Accdb.ps1','scripts/Install-msaccess-vcs.ps1','scripts/Get-GitHubHeaders.ps1')
    $missing = $required | Where-Object { -not (Test-Path (Join-Path $VendorBuildDir $_)) }
    if ($missing) {
        Write-Error @"
Vendored msaccess-vcs-build files are missing from '$VendorBuildDir':
  $($missing -join "`n  ")

Generate them by running (from the repo root, with the submodule initialized):
  git submodule update --init --recursive
  pwsh -File scripts/Bundle-Vendors.ps1
"@
        exit 1
    }
}
```

Digest verification — PowerShell port of the old bash `verify-vcs-digest`, reimplemented
because `Install-msaccess-vcs.ps1` never verifies the zip it downloads. Asset selection
**mirrors the installer** (`Version*.zip` first asset), so we verify exactly the file that
will be installed, not `.assets[0]` (review #7.3):

```powershell
function Assert-VcsAddinDigest {
    param([string]$ReleaseUrl, [string]$ExpectedSha)
    if ([string]::IsNullOrEmpty($ExpectedSha)) {
        Write-Host "No access-vcs-sha provided - skipping add-in digest verification."
        return
    }
    . "$VendorBuildDir/scripts/Get-GitHubHeaders.ps1"
    Write-Host "Verifying add-in digest against: $ReleaseUrl"
    $release = Invoke-RestMethod -Uri $ReleaseUrl -Headers (Get-GitHubHeaders)
    $asset = $release.assets | Where-Object { $_.name -like 'Version*.zip' } |
        Select-Object -First 1
    if (-not $asset -or -not $asset.digest) {
        throw "No sha256 digest found for the Version*.zip asset. Update the pin."
    }
    $actual   = ($asset.digest -replace '^sha256:','').ToLower()
    $expected = $ExpectedSha.Trim().ToLower()
    Write-Host "Expected SHA256: $expected`nActual SHA256:   $actual"
    if ($actual -ne $expected) {
        throw "SHA256 digest mismatch for '$($asset.name)'. Expected $expected, got $actual."
    }
    Write-Host "SHA256 digest verification passed."
}
```

> The expected value must be a 64-hex SHA-256 (not a git SHA-1, which is 40 hex) —
> a known CI pitfall recorded in the repo notes.

Main orchestration — approach (b) (see §5):

```powershell
function Invoke-AccessBuilds {
    param([string[]]$AccessFolders)

    Assert-VendoredVcsBuild

    # --- digest-verify the add-in release ONCE (auth via GH_TOKEN) --------
    Assert-VcsAddinDigest -ReleaseUrl $AccessVcsUrl -ExpectedSha $AccessVcsSha

    # --- install the add-in ONCE into the APPDATA default location ---------
    # Install-msaccess-vcs.ps1 expands to $TargetDir\MSAccessVCS. TargetDir =
    # $env:APPDATA puts it exactly where Build-Accdb.ps1 looks when vcsUrl is
    # empty, so per-folder Build.ps1 calls with -vcsUrl "" reuse it without
    # re-downloading. Run from a temp dir so msaccess-vcs.zip doesn't pollute
    # the workspace.
    New-Item -ItemType Directory -Force -Path (Join-Path $env:TEMP "vba-build-$PID") | Out-Null
    Push-Location (Join-Path $env:TEMP "vba-build-$PID")
    try {
        $install = & "$VendorBuildDir/scripts/Install-msaccess-vcs.ps1" `
            -vcsUrl $AccessVcsUrl -TargetDir $env:APPDATA -SetTrustedLocation $true
    } finally { Pop-Location }
    if (-not $install.AddInPath) {
        throw "Failed to install the msaccess-vcs add-in (see output above)."
    }
    Write-Host "msaccess-vcs add-in installed: $($install.AddInPath)"

    # --- per folder: fresh child pwsh running the vendored Build.ps1 -------
    # Phase 1 uses a separate pwsh per folder for isolation while we validate
    # the new flow. Phase 2 (follow-up) drops the child pwsh and calls Build.ps1
    # in-process for speed.
    foreach ($folder in $AccessFolders) {
        # Replicate the old bash parse-step: per-folder compile flag.
        $compile = if (-not [string]::IsNullOrEmpty($AccessVcsCompile)) { $AccessVcsCompile }
                   elseif ($folder -match '\.accde(\.src)?$')             { 'true' }
                   else                                                   { 'false' }
        Write-Host "Building Access database: $folder (compile=$compile)"
        & pwsh -NoProfile -File "$VendorBuildDir/Build.ps1" `
            -SourceDir      $folder `
            -TargetDir      "$SourceDir/out" `
            -Compile        $compile `
            -AppConfigFile  $AccessVcsConfig `
            -vcsUrl         ""
        if ($LASTEXITCODE -ne 0) {
            throw "Access build failed for $folder (exit code $LASTEXITCODE)."
        }
    }
}
```

After the existing per-folder loop (Access folders are currently collected there and
`continue`d) and **before** the `GITHUB_OUTPUT` writes:

```powershell
if ($hasAccessDatabase) {
    Invoke-AccessBuilds -AccessFolders $accessFolders
}
```

## 5. Chosen approach (b): install add-in once to APPDATA; run vendored `Build.ps1` per folder

Why not call `Build-Accdb.ps1` directly (the earlier plan draft): it bypasses `Build.ps1`
and therefore silently drops compile (`.accde`) and `app-config`
(`Prepare-Application.ps1`), which this repo now needs.

Mechanics of (b):

- `Build.ps1` has **no** `-VcsAddInPath` parameter — it *installs* the add-in when
  `vcsUrl` is set, and with an **empty** `vcsUrl` it passes an empty add-in path to
  `Build-Accdb.ps1`, which then falls back to the **APPDATA default**:
  `%APPDATA%\MSAccessVCS\Version Control.accd[ae]`.
- **Once per run:** `Install-msaccess-vcs.ps1 -TargetDir $env:APPDATA` expands the add-in
  to exactly that default location (and sets its trusted location). One download, one install.
- **Per folder (Phase 1):** spawn a fresh child `pwsh` that runs the vendored `Build.ps1`
  with `-vcsUrl ""` (plus `-Compile <flag>` and `-AppConfigFile <config>` when applicable).
  Because `vcsUrl` is empty, `Build.ps1` skips the install, finds the pre-installed add-in
  at the APPDATA default, and otherwise keeps its full upstream behavior: per-folder
  timestamped trusted location (set + removed), build, optional compile → `.accde`, optional
  `Prepare-Application` (Pre/PostCompile staging around the compile). Each call still gets a
  fresh Access instance (`Build-Accdb.ps1` starts, quits, and force-kills its own
  `MSACCESS`).
- **Per-folder compile flag:** name-based (`*.accde[.src]`) unless `access-vcs-compile` is
  set, exactly reproducing the old bash `parse` step.
- **Net effect:** 1 add-in download/install; N fresh Access instances; full feature parity
  (incl. `.accde` + `app-config`); no `Call action 0..4` cap.

**Phase 2 (follow-up, NOT part of this change):** once the fold-in is validated, drop the
child `pwsh` and invoke the vendored `Build.ps1` in-process (`&`, checking
`$LASTEXITCODE`) — `Build-Accdb.ps1` already isolates Access per call, so the child process
adds no Access isolation, only process fault isolation.

## 6. "Make minor edits" — what they actually are (updated for (b))

Under (b) `Build.ps1` is always invoked with an explicit `-vcsUrl ""`, and
`Install-msaccess-vcs.ps1` with explicit `-vcsUrl` / `-TargetDir`, so their **defaults are
never hit**. Auth already ships in the fork. Therefore **no string edits are required
today**; bundler step 4 is reserved for future deterministic replaces. Items to verify (no
edit):

1. **Auth** (`Get-GitHubHeaders.ps1`): present in the fork and used by
   `Install-msaccess-vcs.ps1` (and now by the digest check). Verify it is the file bundled.
2. **`action.yml`**: vendored for reference only; not executed (documented in `SOURCE.txt`).
3. **Trusted-location uniqueness**: `Build.ps1` uses timestamped location names per call;
   repeated per-folder invocations don't collide. Leftover registry entries after a run are
   acceptable (CI is ephemeral; dev machines accumulate a harmless entry).

## 7. License & provenance

- Copy the submodule's `LICENSE` into `vendors/msaccess-vcs-build/LICENSE` **unchanged**.
- Generate `vendors/msaccess-vcs-build/SOURCE.txt` with: source URL, bundled commit SHA,
  regeneration command. **No bundle date** (idempotency, §3/§10).
- If any script files carry their own license headers, keep them intact (no stripping).
- Optional: add a top-level note in `README.md` that `vendors/msaccess-vcs-build` is
  vendored from `DecimalTurn/msaccess-vcs-build` at a pinned commit.

## 8. Root `action.yml` cleanup

After `Main.ps1` handles Access:

- **"Run VBA Build" step** now also passes the Access inputs:

```yaml
- name: "Run VBA Build"
  id: "run_vba_build"
  shell: pwsh
  run: |
    ${{ github.action_path }}\Main.ps1 `
      -SourceDir        "${{ inputs['source-dir'] }}" `
      -TestFramework    "${{ inputs['test-framework'] }}" `
      -OfficeAppDetection "${{ inputs['office-app'] }}" `
      -AccessVcsUrl     "${{ inputs['access-vcs-url'] }}" `
      -AccessVcsSha     "${{ inputs['access-vcs-sha'] }}" `
      -AccessVcsConfig  "${{ inputs['access-vcs-config'] }}" `
      -AccessVcsCompile "${{ inputs['access-vcs-compile'] }}"
```
- Delete the `Build Access Database (if detected)` step that `uses:
  ./subactions/msaccess-vcs-build-all`.
- Delete `subactions/msaccess-vcs-build-all/` (its `parse` + `verify-vcs-digest` logic now
  lives in `Main.ps1`).
- Keep `access-folders` / `has-access-database` outputs (Main.ps1 already writes them).
- Keep the "Expose GitHub token to environment" step — `Main.ps1`'s digest check and the
  install read `GH_TOKEN`.
- Keep the submodule (it feeds `Bundle-Vendors.ps1`).
- Adjacent cleanup while here: the `first-access-folder` output references a non-existent
  `steps.extract_access_folder` step (pre-existing dead output) — remove or fix it.

## 9. Missing-file experience & CI sync-guard

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

CI only hits this if the commit didn't include `vendors/`. Add a CI guard job that re-runs
the bundler and fails if it produces changes (keeps `vendors/` in sync). Note
`actions/checkout` does **not** fetch submodules by default, so the guard must request them:

```yaml
guard-vendors:
  runs-on: ubuntu-latest
  steps:
    - uses: actions/checkout@v6
      with:
        submodules: recursive
    - name: "Check vendors/ is in sync with the submodule"
      shell: pwsh
      run: |
        pwsh -File scripts/Bundle-Vendors.ps1
        git diff --exit-code -- vendors/
        if ($LASTEXITCODE -ne 0) {
          Write-Error "vendors/ is out of sync. Run scripts/Bundle-Vendors.ps1 and commit the result."
          exit 1
        }
        Write-Host "vendors/ is in sync."
```

## 10. Testing plan

1. `pwsh -File scripts/Bundle-Vendors.ps1` → `git status` shows only the expected vendor
   tree + `SOURCE.txt`.
2. **Idempotency:** re-run the bundler → no further diff (holds because `SOURCE.txt` has no
   timestamp).
3. Local parse check of every vendored `.ps1`:
   `[System.Management.Automation.Language.Parser]::ParseFile(...)`.
4. CI `Test Build VBA` (source-dir `./tests`) must still detect both Access exports, verify
   the digest against the default pin (`9504af3d…`), and produce:
   - `tests/AccessDatabase.accdb` (non-compile, name-based) → `tests/out/Testing.accdb`;
   - `tests/AccessExecuteOnlyDatabase.accde` (compile, name-based) → compiled `.accde` in
     `tests/out/`.
   (This covers both the plain build and the new-in-scope compile path.)
5. Compile / app-config smoke: run with `access-vcs-compile: true` and
   `access-vcs-config: tests/Application-Config.json` to exercise the `Prepare-Application`
   Pre/PostCompile paths.
6. Negative test: temporarily rename `vendors/msaccess-vcs-build` → expect the clear error
   message above.
7. >5 Access folders regression: temporarily copy the real fixtures (they must stay
   UTF-8-with-BOM + CRLF, per the repo notes — hand-made fixtures fail the add-in build)
   into >5 folders and confirm all build (cap gone).

## 11. Rollback

The subaction and `submodules` stay in git history; reverting the fold-in is a `git revert`.
`vendors/` commit is additive and doesn't affect non-Access builds.

---

**Decisions:**
1. Build execution **(b)**: install the add-in **once** into `%APPDATA%` (the default
   location `Build-Accdb.ps1` falls back to), then run the vendored `Build.ps1` per folder
   with `-vcsUrl ""` in a **fresh child pwsh** (Phase 1). Phase 2 = in-process optimization
   (follow-up) (see §5).
2. Feature parity is in scope: compile (`.accde`) and `app-config`
   (`Prepare-Application`) are preserved by calling full `Build.ps1`; the per-folder
   compile-flag logic moves from bash into `Main.ps1`.
3. Vendoring scope: **copy the whole upstream tree** into `vendors/msaccess-vcs-build`
   (see §3).
4. Digest verification is reimplemented in PowerShell, selects the same `Version*.zip`
   asset the installer downloads, and runs **before** the install; skipped when
   `access-vcs-sha` is empty (see §4, review #7.3).
5. `$VendorBuildDir = $PSScriptRoot/vendors/msaccess-vcs-build` (action checkout); install
   target = `$env:APPDATA`. No shared `$RepoRoot` (review #5).
6. `SOURCE.txt` has **no timestamp** so the bundler is byte-idempotent and the CI
   sync-guard works (see §3, §9).
7. Keep the submodule in this repo — it is the source `Bundle-Vendors.ps1` copies from
   (runtime CI only uses committed `vendors/`).