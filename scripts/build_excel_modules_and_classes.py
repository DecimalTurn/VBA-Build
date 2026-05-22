#!/usr/bin/env python3
"""Build the ExcelWorkbook fixture with standard modules and class modules.

This script seeds a workbook with a matching class module so pyOpenVBA's
push() function can update both the standard module and the class module
sources from disk.
"""

from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

from pyopenvba import ExcelFile, VBAModuleKind, push


def _stage_sources(modules_dir: Path, class_modules_dir: Path, staged_dir: Path) -> None:
    for source_dir in (modules_dir, class_modules_dir):
        if not source_dir.is_dir():
            raise NotADirectoryError(f"Missing source folder: {source_dir}")
        for item in source_dir.iterdir():
            if item.is_file() and item.suffix.lower() in {".bas", ".cls"}:
                shutil.copy2(item, staged_dir / item.name)


def build_workbook(source_root: Path, output_file: Path) -> Path:
    modules_dir = source_root / "Modules"
    class_modules_dir = source_root / "Class Modules"

    output_file.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        staged_sources = temp_path / "staged_sources"
        staged_sources.mkdir()
        _stage_sources(modules_dir, class_modules_dir, staged_sources)

        with ExcelFile.create_new(output_file) as workbook:
            for class_file in sorted(class_modules_dir.glob("*.cls")):
                class_name = class_file.stem
                if class_name not in workbook.module_names():
                    workbook.vba_project().add_module(
                        class_name,
                        "",
                        kind=VBAModuleKind.other,
                    )
            workbook.save(output_file)

        push(staged_sources, output_file)

    return output_file


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build ExcelWorkbook.xlsm with standard modules and class modules."
    )
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path("tests/ExcelWorkbook.xlsm"),
        help="Fixture root containing Modules/ and Class Modules/.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("build/ExcelWorkbook_modules_and_classes.xlsm"),
        help="Destination .xlsm file.",
    )
    args = parser.parse_args()

    built_file = build_workbook(args.source_root, args.output)

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