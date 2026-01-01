# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]
## [0.6.0]
- 0.6.0 focuses on reliability and richer extraction coverage: core data types and interfaces now capture tables, images, and formula output more consistently across formats, and extractor behavior has been tightened with broader fixtures and tests to validate real-world edge cases.
- Email handling expands beyond bodies to include binary attachment extraction and downstream processing; attachments now carry MIME metadata, support checks, and can be iterated through the same extractor pipeline when supported.
- A dedicated encryption detection layer now prevents unsupported protected files from being partially parsed, with a consistent exception raised across PDF, modern OOXML, ODF, and legacy Office formats.
- Parser internals for document formats (PDF, DOCX, PPTX, XLSX, ODT/ODS/ODP, legacy DOC) received upgrades to improve table/image handling, formula parsing (OMML-to-LaTeX), and metadata consistency.
- Image extraction for MS legacy file formats (PPT/XLS/DOC)
- JSON (de)serialization

PRs
- Feature/separated formula parser (#5)
- Feature/image context (#6)
- Feature/table interface (#7)
- Feature/interface polishing (#8)
- Feature/msg support attachment extraction (#9)
- Feature/encryption detection (#10)

## [Released]

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
