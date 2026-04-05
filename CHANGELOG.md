# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [2.0.0] - 2026-02-18

## Breaking changes

-  Use setup action by @DecimalTurn ([#31])

	`VBA-Build` now focuses only on **building and testing VBA-enabled files** from source.
	
	Starting with **v2.0.0**, environment initialization is no longer handled inside this action.
	You must run [`DecimalTurn/setup-vba`](https://github.com/DecimalTurn/setup-vba) before `VBA-Build`.

	In previous versions, `VBA-Build` handled setup tasks like:
	
	- Installing Microsoft Office on the runner
	- Initializing Office applications
	- Configuring VBA security (VBOM access / macro settings)
	
	In **v2.0.0**, these responsibilities were removed from `VBA-Build` and moved to `setup-vba`.
	
	This is an intentional separation of concerns:
	
	- `setup-vba` = prepare runner/runtime
	- `VBA-Build` = build documents from source

	See [Migration Guide](https://github.com/DecimalTurn/VBA-Build/blob/7270b97b7e155824a85d055eb9bbda607f36fd56/RELEASE_NOTES_v2.0.0.md#migration-guide-existing-workflows) for more details on how to upgrade.

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
[#31]: https://github.com/DecimalTurn/VBA-Build/pull/31
