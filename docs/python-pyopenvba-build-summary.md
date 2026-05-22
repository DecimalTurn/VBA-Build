# Python + pyOpenVBA Build Summary

This document summarizes the Python-based workbook build work done in this repository for the `tests/ExcelWorkbook.xlsm` fixture.

## Scope

The goal was to keep pyOpenVBA code and docs locally available, install pyOpenVBA in the workspace environment, and add scriptable build paths for `.xlsm` outputs from fixture source files.

## What Was Added

- Git submodule:
  - `external/pyOpenVBA`
- Python package install in the workspace virtual environment:
  - `pip install pyOpenVBA`
- Fixture class module file:
  - `tests/ExcelWorkbook.xlsm/Class Modules/Class1.cls`

## Build Scripts

### 1) Full XML + modules build (Approach currently not working => XML mixup)

- Script: `scripts/build_excel_workbook.py`
- Behavior:
  - Creates a workbook from `pyOpenVBA` template (`ExcelFile.create_new`)
  - Merges `tests/ExcelWorkbook.xlsm/XMLsource` into the package
  - Pushes standard modules from `tests/ExcelWorkbook.xlsm/Modules`
- Default output:
  - `build/ExcelWorkbook.xlsm`

### 2) Modules-only build

- Script: `scripts/build_excel_module_only.py`
- Behavior:
  - Creates a clean workbook from `pyOpenVBA` template
  - Pushes only standard modules (`.bas`) from `Modules`
  - Intentionally ignores workbook/sheet code-behind files in `Microsoft Excel Objects`
- Default output:
  - `build/ExcelWorkbook_modules_only.xlsm`

### 3) Modules + class modules build

- Script: `scripts/build_excel_modules_and_classes.py`
- Behavior:
  - Creates a clean workbook from `pyOpenVBA` template
  - Seeds class modules in the workbook using `vba_project().add_module(..., kind=VBAModuleKind.other)`
  - Stages and pushes both:
    - `Modules/*.bas`
    - `Class Modules/*.cls`
- Default output:
  - `build/ExcelWorkbook_modules_and_classes.xlsm`

## Important pyOpenVBA Constraint

`pyOpenVBA.push()` updates existing modules by matching file stem names.

That means:
- standard modules work when a matching module already exists,
- class modules must also exist in the workbook before push,
- to add class source from disk, the workbook must be seeded with matching class module entries first.

## Current Outputs

Depending on which script is run, build artifacts are created under:

- `build/ExcelWorkbook.xlsm`
- `build/ExcelWorkbook_modules_only.xlsm`
- `build/ExcelWorkbook_modules_and_classes.xlsm`
