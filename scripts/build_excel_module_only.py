#!/usr/bin/env python3
"""Build a clean Excel workbook that imports only standard VBA modules.

This script intentionally ignores the workbook and worksheet code-behind
files in the fixture. It starts from a known-good macro-enabled template
and pushes only the contents of the Modules folder.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from pyopenvba import ExcelFile, push


def build_workbook(modules_dir: Path, output_file: Path) -> Path:
    if not modules_dir.is_dir():
        raise NotADirectoryError(f"Missing Modules folder: {modules_dir}")

    output_file.parent.mkdir(parents=True, exist_ok=True)

    with ExcelFile.create_new(output_file) as workbook:
        workbook.save(output_file)

    push(modules_dir, output_file)
    return output_file


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build ExcelWorkbook.xlsm with only standard VBA modules."
    )
    parser.add_argument(
        "--modules-dir",
        type=Path,
        default=Path("tests/ExcelWorkbook.xlsm/Modules"),
        help="Directory containing .bas standard modules.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("build/ExcelWorkbook_modules_only.xlsm"),
        help="Destination .xlsm file.",
    )
    args = parser.parse_args()

    built_file = build_workbook(args.modules_dir, args.output)

    with ExcelFile(built_file) as workbook:
        print("Built:", built_file)
        print("Modules:", ", ".join(workbook.module_names()))
        issues = workbook.validate()
        if issues:
            print("Validation issues:")
            for issue in issues:
                print("-", issue)
        else:
            print("Validation: OK")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())