#!/usr/bin/env python3
"""Build the ExcelWorkbook.xlsm fixture into a macro-enabled workbook.

This script uses pyOpenVBA to create a valid macro-enabled workbook
template, merges in the XML source tree from the fixture, and then
pushes the VBA modules from the fixture's Modules folder.
"""

from __future__ import annotations

import argparse
import tempfile
import zipfile
from pathlib import Path

from pyopenvba import ExcelFile, push


def _merge_xml_source_into_workbook(template_workbook: Path, source_dir: Path, output_file: Path) -> None:
    source_files = {
        item.relative_to(source_dir).as_posix(): item
        for item in source_dir.rglob("*")
        if item.is_file()
    }

    with zipfile.ZipFile(template_workbook, mode="r") as template_zip, zipfile.ZipFile(
        output_file, mode="w", compression=zipfile.ZIP_DEFLATED
    ) as merged_zip:
        written_names: set[str] = set()

        for info in template_zip.infolist():
            source_path = source_files.get(info.filename)
            if source_path is not None:
                merged_zip.writestr(info, source_path.read_bytes())
            else:
                merged_zip.writestr(info, template_zip.read(info.filename))
            written_names.add(info.filename)

        for relative_name, source_path in source_files.items():
            if relative_name in written_names:
                continue
            merged_zip.write(source_path, relative_name)


def build_workbook(source_root: Path, output_file: Path) -> Path:
    xml_source = source_root / "XMLsource"
    modules_source = source_root / "Modules"

    if not xml_source.is_dir():
        raise NotADirectoryError(f"Missing XML source folder: {xml_source}")
    if not modules_source.is_dir():
        raise NotADirectoryError(f"Missing Modules folder: {modules_source}")

    output_file.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_workbook = Path(temp_dir) / output_file.name
        with ExcelFile.create_new(temp_workbook) as workbook:
            workbook.save(temp_workbook)

        _merge_xml_source_into_workbook(temp_workbook, xml_source, output_file)

    push(modules_source, output_file)
    return output_file


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the ExcelWorkbook.xlsm fixture.")
    parser.add_argument(
        "--source-root",
        type=Path,
        default=Path("tests/ExcelWorkbook.xlsm"),
        help="Fixture root containing XMLsource/ and Modules/.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("build/ExcelWorkbook.xlsm"),
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