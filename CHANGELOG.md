# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [Released]
## [0.8.1] - 2026-01-10
- Pinned PDF extraction library to >=6.5 <6.6
- Added support for the Open Office file formats .odg (drawings) and .odf (formulas)
- Added support for 7zip archives

## [0.8.0] - 2026-01-04
- Added sharepoint connector to read files from a Sharepoint via MS Graph API
- Added instructions how to configure/add needed permissions for this to work
- Added support for .zip/.tar archives for supported file types

## [0.7.0] - 2026-01-03
- New `sharepoint2text` CLI entry point with JSON (`--json`) and unit-level JSON (`--json-unit`) output, plus optional base64 binary payloads via `--binary`.
- Standardized unit processing across formats (including heading-based units for DOCX) to support chunked extraction and per-unit metadata workflows.
- Extractor refactors for modern/legacy Office and OpenDocument formats improve text/table/image consistency and shared image handling utilities.
- PDF decryption enhancements add AES fallback support, optional `pdf-crypto` acceleration, and clearer encryption errors for protected files.
- ZIP bomb mitigation for ZIP-based formats (OOXML/ODF) introduces `ExtractionZipBombError` safeguards.
- File routing now prefers extensions before MIME types for more deterministic extractor selection.
- Packaging/dev workflow updates: version sourced from `pyproject.toml`, `py.typed` restored, uv/pre-commit adoption, and refreshed docs/tests.

PRs
- Allow calling the tool from the CLI (#17)
- added --json switch to cli tool (#18)
- Feature/extractor refactoring (#19)
- use file-ending as primary routing mechanism and fallback to mime type (#20)
- Feature/cli extension (#21)
- Feature/zip bomb mitigation (#22)
- Feature/docx iteration (#23)
- Feature/refined iterate unit (#24)
- Feature/unit processing (#25)
- Feature/pdf decryption (#26)

## [0.6.0] - 2026-01-01
- 0.6.0 focuses on reliability and richer extraction coverage: core data types and interfaces now capture tables, images, and formula output more consistently across formats, and extractor behavior has been tightened with broader fixtures and tests to validate real-world edge cases.
- Email handling expands beyond bodies to include binary attachment extraction and downstream processing; attachments now carry MIME metadata, support checks, and can be iterated through the same extractor pipeline when supported.
- A dedicated encryption detection layer now prevents unsupported protected files from being partially parsed, with a consistent exception raised across PDF, modern OOXML, ODF, and legacy Office formats.
- Parser internals for document formats (PDF, DOCX, PPTX, XLSX, ODT/ODS/ODP, legacy DOC) received upgrades to improve table/image handling, formula parsing (OMML-to-LaTeX), and metadata consistency.
- XML handling is hardened to avoid unsafe entity expansion in supported formats.
- PPTX extraction now includes table parsing, and legacy Office formats gained image extraction across PPT/XLS/DOC.
- JSON (de)serialization was added for structured extractor outputs, with image metadata included and test coverage for serialization/deserialization.
- Refactoring and typing fixes improved interface consistency, and email body extraction now strips HTML more reliably.
- Documentation refreshes (README + docstrings) capture the updated behavior and usage.
- Tests now cover JSON conversion paths and image metadata serialization to guard regression risk.

PRs
- Feature/separated formula parser (#5)
- Feature/image context (#6)
- Feature/table interface (#7)
- Feature/interface polishing (#8)
- Feature/msg support attachment extraction (#9)
- Feature/encryption detection (#10)
- Feature/refactoring (#11)
- Added a serialization to json (#12)
- Added a deserialization (#13)
- Feature/legacy image extraction (#14)
- Feature/pptx tables (#15)
- Feature/security xml (#16)

## [0.5.0] - 2025-12-29
- Added support for open office file formats
- Reduced dependency footprint
- Re-implemented modern .docx and .pptx extraction

## [0.4.1] - 2025-12-28
- Added support for .html files

## [0.4.0] - 2025-12-28
- Dropped Pandas/Numpy dependencies for reading Excel documents
- legacy .xls is read directly via `xlrd`
- modern .xlsx is now read via `openpyxl`
- Cut dependency foot-print in half

## [0.3.0] - 2025-12-28
- Added support for .rtf files
- Added support for .md files
- throw custom exception when not supported files are encounted instead of RunTimException
- .docx and .pptx have no formula parsing capabilities which aims at re-constructing latex-styled formulas from found formulas

## [0.2.0] - 2025-12-27

### Added
- Added support for email file formats
- All extractors are now generators
- Some email formats may contain multiple email entries


## [0.1.1] - 2025-12-25

### Added

- Initial public release
- Text extraction support for modern Office formats:
  - Word documents (.docx)
  - Excel spreadsheets (.xlsx)
  - PowerPoint presentations (.pptx)
- Text extraction support for legacy Office formats:
  - Word 97-2003 documents (.doc)
  - Excel 97-2003 spreadsheets (.xls)
  - PowerPoint 97-2003 presentations (.ppt)
- PDF document text extraction (.pdf)
- Plain text file support (.txt, .json, .csv, .tsv)
- Router module for automatic file type detection
- Comprehensive metadata extraction for all supported formats
- Image extraction from PDFs and PowerPoint presentations
