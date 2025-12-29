"""
Modern Microsoft Office Extractor Package
==========================================

This package provides extractors for parsing and extracting text content from
modern Microsoft Office formats (Office 2007 and later). These formats use
the Office Open XML (OOXML) standard, which stores documents as ZIP archives
containing XML files.

Supported Formats
-----------------

.docx (Word 2007+):
    Office Open XML Word document. Uses direct XML parsing of the ZIP archive
    for all content extraction. Extracts paragraphs, tables, headers/footers,
    images, footnotes, comments, and math formulas (converted to LaTeX).

.pptx (PowerPoint 2007+):
    Office Open XML Presentation. Uses `python-pptx` for parsing with
    additional XML handling for comments and embedded formulas. Extracts
    slides, shapes, images, speaker notes, and formulas.

.xlsx (Excel 2007+):
    Office Open XML Spreadsheet. Uses `openpyxl` for parsing cells and
    sheets. Extracts data as both structured dictionaries and formatted
    text tables.

File Format Background
----------------------
Office Open XML (OOXML) was introduced with Microsoft Office 2007 and became
an ISO/IEC standard (ISO/IEC 29500). Key characteristics:

    - ZIP-compressed archive containing XML and binary parts
    - Human-readable XML structure (when unzipped)
    - Relationships between parts defined in .rels files
    - Media (images, etc.) stored in separate folders
    - Core properties in docProps/core.xml

Common Archive Structure:
    document.docx/
    ├── [Content_Types].xml    # MIME types for parts
    ├── _rels/
    │   └── .rels              # Package relationships
    ├── docProps/
    │   ├── core.xml           # Title, author, dates
    │   └── app.xml            # Application properties
    └── word/                   # (or ppt/, xl/)
        ├── document.xml       # Main content
        ├── styles.xml         # Style definitions
        └── media/             # Images and other media

XML Namespaces:
    - http://schemas.openxmlformats.org/wordprocessingml/2006/main (w:)
    - http://schemas.openxmlformats.org/presentationml/2006/main (p:)
    - http://schemas.openxmlformats.org/spreadsheetml/2006/main (s:)
    - http://schemas.openxmlformats.org/officeDocument/2006/math (m:)

Dependencies
------------
docx_extractor: No external dependencies (uses stdlib zipfile and xml.etree)
    Direct XML parsing of OOXML structure provides:
    - Paragraph and run extraction with formatting
    - Table parsing
    - Header/footer extraction
    - Image extraction with metadata
    - Core properties (metadata) access
    - Footnote and endnote extraction
    - Comment extraction
    - Formula conversion to LaTeX

python-pptx: https://github.com/scanny/python-pptx
    pip install python-pptx

    Provides:
    - Slide enumeration
    - Shape parsing (text, images, placeholders)
    - Slide layout handling
    - Core properties access

openpyxl: https://github.com/theorchard/openpyxl
    pip install openpyxl

    Provides:
    - Cell and sheet iteration
    - Formula evaluation (cached results)
    - Named ranges and tables
    - Core properties access

Math Formula Support
--------------------
Office documents store math formulas in OMML (Office Math Markup Language),
an XML format within the http://schemas.openxmlformats.org/officeDocument/2006/math
namespace. The extractors convert OMML to LaTeX for text representation.

Supported math elements:
    - Fractions (\\frac)
    - Super/subscripts (^, _)
    - Square roots (\\sqrt)
    - Summation/integration (\\sum, \\int)
    - Matrices (\\begin{matrix})
    - Greek letters and math symbols

Known Limitations
-----------------
- Password-protected/encrypted files are not supported
- Macros (VBA) are not extracted
- Embedded OLE objects are not fully extracted
- Complex SmartArt may not extract all text
- Charts are not converted to text
- Very large documents may use significant memory

Usage Example
-------------
    >>> from sharepoint2text.extractors.ms_modern import read_docx, read_pptx
    >>> import io
    >>>
    >>> with open("report.docx", "rb") as f:
    ...     for doc in read_docx(io.BytesIO(f.read())):
    ...         print(doc.full_text)
    ...
    >>> with open("slides.pptx", "rb") as f:
    ...     for ppt in read_pptx(io.BytesIO(f.read())):
    ...         for slide in ppt.slides:
    ...             print(f"Slide {slide.slide_number}: {slide.title}")

Comparison with Legacy Formats
------------------------------
Modern formats offer several advantages over legacy formats:
    - No OLE complexity - just ZIP + XML
    - Better Unicode support
    - Smaller file sizes (ZIP compression)
    - Easier to parse with standard XML tools
    - Open standard (ISO/IEC 29500)

See Also
--------
- sharepoint2text.extractors.ms_legacy: For Office 97-2003 formats
- sharepoint2text.extractors.data_types: Data structures for extracted content
- OOXML specification: https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
"""

from sharepoint2text.extractors.ms_modern.docx_extractor import read_docx
from sharepoint2text.extractors.ms_modern.pptx_extractor import read_pptx
from sharepoint2text.extractors.ms_modern.xlsx_extractor import read_xlsx

__all__ = [
    "read_docx",
    "read_pptx",
    "read_xlsx",
]
