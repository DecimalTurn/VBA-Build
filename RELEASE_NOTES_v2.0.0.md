# VBA-Build v2.0.0 (Breaking Change)

## Summary

`VBA-Build` now focuses only on **building and testing VBA-enabled files** from source.

Starting with **v2.0.0**, environment initialization is no longer handled inside this action.
You must run [`DecimalTurn/setup-vba`](https://github.com/DecimalTurn/setup-vba) before `VBA-Build`.

## What changed

In previous versions, `VBA-Build` handled setup tasks like:

- Installing Microsoft Office on the runner
- Initializing Office applications
- Configuring VBA security (VBOM access / macro settings)

In **v2.0.0**, these responsibilities were removed from `VBA-Build` and moved to `setup-vba`.

This is an intentional separation of concerns:

- `setup-vba` = prepare runner/runtime
- `VBA-Build` = build documents from source

## Migration guide (existing workflows)

### Before (v1.x style)

```yaml
jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v6

      - name: Build VBA
        uses: DecimalTurn/VBA-Build@v1
        with:
          source-dir: ./src
          test-framework: rubberduck
```

### After (v2.0.0)

```yaml
jobs:
  build:
    runs-on: windows-2025
    steps:
      - uses: actions/checkout@v6

      - name: Setup VBA runtime
        uses: DecimalTurn/setup-vba@75c6ce5e714186234ef9090c1c77537a60bd7339 # v0.1.1
        with:
          office-apps: "Excel,Word,PowerPoint,Access"
          install-office: "true"

      - name: Build VBA
        uses: DecimalTurn/VBA-Build@v2.0.0
        with:
          source-dir: ./src
          test-framework: rubberduck
```

## Recommended migration steps

1. Add a `setup-vba` step **before** `VBA-Build`.
2. Keep your existing `VBA-Build` inputs (`source-dir`, `test-framework`, `office-app`) as needed.
3. (Optional) Pin your `setup-vba` version to a commit SHA (recommended for supply-chain hardening).
4. Run the workflow once and verify Office setup + build output artifacts.

## Notes

- If Office is already available on your runner image, set `install-office: "false"`.
- If you only target specific apps, narrow `office-apps` (example: `"Excel"`).
- This release is breaking by design to enforce modular setup/build responsibilities.
