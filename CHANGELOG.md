# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.6.0]
- Added OMML-to-LaTeX conversion utilities and tests for formula extraction.
- Expanded extraction coverage and robustness across PDF, DOCX, PPTX, XLSX, ODT/ODS/ODP, and legacy DOC files, including improved table/image handling.
- Added MSG attachment extraction with MIME-type support awareness and attachment content iteration.
- Introduced shared MIME type mapping module for consistent routing and attachment support checks.
- Added new fixture files and broader extraction tests, plus CI/test workflow updates.

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
