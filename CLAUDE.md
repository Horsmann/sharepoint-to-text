# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Development Commands

```bash
# Install dependencies (using uv - recommended)
uv venv && source .venv/bin/activate && uv pip install -e ".[dev]"

# Run all tests
pytest

# Run a single test file
pytest sharepoint2text/tests/test_extractions.py

# Run a specific test
pytest sharepoint2text/tests/test_extractions.py::test_function_name -v

# Format code
black sharepoint2text

# Run pre-commit hooks manually
pre-commit run --all-files
```

## Architecture

This is a pure Python library for extracting text from document formats (no LibreOffice/Java dependencies). All extractors return **generators** that yield content objects implementing `ExtractionInterface`.

### Router Pattern

`sharepoint2text/router.py` maps file extensions/MIME types to extractors via `_EXTRACTOR_REGISTRY`. Extractors are lazy-loaded to minimize startup time. Entry point is `sharepoint2text.read_file(path)` which auto-detects format.

### Extractor Organization

```
sharepoint2text/extractors/
├── ms_modern/          # .docx, .xlsx, .pptx (OOXML/ZIP-based)
├── ms_legacy/          # .doc, .xls, .ppt, .rtf (OLE2 binary)
├── open_office/        # .odt, .odp, .ods (ODF/ZIP-based)
├── mail/               # .eml, .msg, .mbox
├── util/               # Shared utilities (encryption detection, ZIP handling, OMML->LaTeX)
├── pdf_extractor.py
├── html_extractor.py
├── plain_extractor.py
└── data_types.py       # All content dataclasses and interfaces
```

### Content Interface

All extractors yield objects implementing:
- `get_full_text()` - complete text as single string
- `iterate_units()` - yields logical units (pages/slides/sheets) as `UnitInterface` objects
- `iterate_images()` - yields `ImageInterface` objects
- `iterate_tables()` - yields `TableInterface` objects
- `get_metadata()` - returns `FileMetadataInterface`

### Adding a New Format

1. Create extractor in appropriate subdirectory following pattern: `read_{format}(file: io.BytesIO, path: str | None) -> Generator[ContentType, Any, None]`
2. Add content dataclass to `data_types.py` implementing `ExtractionInterface`
3. Register in `router.py` `_EXTRACTOR_REGISTRY` and `mime_types.py` if MIME-detectable
4. Add tests in `sharepoint2text/tests/test_extractions.py` with fixtures in `tests/resources/`

### Key Dependencies

- `olefile`: OLE2 parsing for legacy Office formats
- `defusedxml`: Secure XML parsing for OOXML/ODF
- `openpyxl`: .xlsx parsing
- `xlrd`: .xls parsing
- `pypdf`: PDF extraction
- `mail-parser`/`msg-parser`: Email formats
