# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] - 2026-02-18

## [1.4.0] - 2025-09-14
### Added
- Support for building Access files (`.accdb`) via `msaccess-vcs-build`.
- Support for macro-enabled template formats:
	- `.xltm` (Excel template)
	- `.potm` (PowerPoint template)
	- `.dotm` (Word template)

### Security
- Added `sha256` checksum verification for dependencies.

### Changed
- Updated multiple GitHub Actions dependencies.

## [1.3.0] - 2025-06-02
### Added
- Support for Excel and Word objects.
- Support for PowerPoint `.ppam` add-in format.

## [1.2.0] - 2025-05-23
### Added
- Support for Excel Binary Workbook (`.xlsb`) format.

## [1.1.0] - 2025-05-21
### Added
- Experimental support for running VBA unit tests via Rubberduck.

## [1.0.0] - 2025-04-20
### Added
- Support for Word (`.docm`) and PowerPoint (`.pptm`) files.
- Expanded Excel format support.
- Support for forms (`.frm`) and class modules (`.cls`).
- Support for custom source folder/file naming in generated outputs.

## [0.1.0] - 2025-06-02
### Added
- Initial demo release showing Excel file generation from VBA and XML source code.

[1.4.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v1.4.0
[1.3.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v1.3.0
[1.2.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v1.2.0
[1.1.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v1.1.0
[1.0.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v1.0.0
[0.1.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v0.1.0
[2.0.0]: https://github.com/DecimalTurn/VBA-Build/releases/tag/v2.0.0
