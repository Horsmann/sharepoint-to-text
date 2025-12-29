"""
Legacy Microsoft Office Extractor Package
==========================================

This package provides extractors for parsing and extracting text content from
legacy Microsoft Office binary file formats. These formats were used by
Microsoft Office versions 97 through 2003 and are based on the OLE2 (Object
Linking and Embedding) Compound File Binary Format (CFBF).

Supported Formats
-----------------

.doc (Word 97-2003):
    Microsoft Word binary document format. Uses `olefile` for OLE container
    parsing and custom binary parsing for the WordDocument stream. Extracts
    main text, footnotes, headers/footers, and annotations.

.ppt (PowerPoint 97-2003):
    Microsoft PowerPoint binary presentation format. Uses `olefile` for OLE
    container parsing and parses the PowerPoint Document stream according
    to the MS-PPT specification. Extracts slides, notes, and master text.

.xls (Excel 97-2003):
    Microsoft Excel binary workbook format. Uses `xlrd` for cell/sheet parsing
    and `olefile` for metadata extraction. Extracts sheet data as both
    structured dictionaries and formatted text tables.

.rtf (Rich Text Format):
    Cross-platform document format developed by Microsoft. Plain text format
    with control words for formatting. No external dependencies beyond Python
    stdlib. Extracts text, metadata, fonts, colors, headers/footers, and more.

File Format Background
----------------------
All Office 97-2003 binary formats (except RTF) are based on the Compound File
Binary Format (CFBF), also known as OLE2 or structured storage. Key concepts:

    - Files are containers with internal "streams" (like a file system)
    - Streams contain binary data specific to each application
    - Common streams include document content, metadata, and summary info
    - The format uses little-endian byte ordering throughout

Common OLE Streams:
    - WordDocument: Main document content for .doc files
    - PowerPoint Document: Presentation data for .ppt files
    - Workbook: Sheet data for .xls files
    - \\x05SummaryInformation: Standard metadata (title, author, etc.)
    - \\x05DocumentSummaryInformation: Extended metadata

Dependencies
------------
olefile: https://github.com/decalage2/olefile
    pip install olefile
    Core dependency for all OLE-based formats (.doc, .ppt, .xls).
    Provides OLE container parsing and metadata extraction.

xlrd: https://github.com/python-excel/xlrd
    pip install xlrd
    Required for .xls files. Parses Excel binary format cells and sheets.
    Note: xlrd 2.x only supports .xls, not .xlsx (use openpyxl for .xlsx).

Known Limitations
-----------------
- Password-protected/encrypted files are not supported
- Embedded OLE objects are not extracted (e.g., embedded Excel in Word)
- Macros (VBA) are not extracted
- Some legacy files from Office 95 or earlier may not parse correctly
- Complex formatting (tables, images) extraction varies by format
- Files with severe corruption may fail to parse

Character Encoding
------------------
Legacy Office files may use various encodings:
    - Unicode (UTF-16LE) for modern documents
    - Windows code pages (CP1252, CP1250, etc.) for older documents
    - Mixed encoding within single documents

All extractors handle encoding with fallback strategies and use replacement
characters for undecodable bytes.

Usage Example
-------------
    >>> from sharepoint2text.extractors.ms_legacy import read_doc, read_ppt
    >>> import io
    >>>
    >>> with open("document.doc", "rb") as f:
    ...     for doc in read_doc(io.BytesIO(f.read()), path="document.doc"):
    ...         print(doc.main_text)
    ...
    >>> with open("presentation.ppt", "rb") as f:
    ...     for ppt in read_ppt(io.BytesIO(f.read())):
    ...         for slide in ppt.slides:
    ...             print(f"Slide {slide.slide_number}: {slide.title}")

Maintenance Notes
-----------------
- The doc_extractor uses AI-generated binary parsing code; verify against
  MS-DOC specification for edge cases
- The ppt_extractor follows MS-PPT specification for record types
- xlrd is unmaintained but stable for .xls files
- olefile is actively maintained and reliable
- All extractors follow the generator pattern for consistency

See Also
--------
- sharepoint2text.extractors.ms_modern: For Office 2007+ formats (.docx, .pptx, .xlsx)
- sharepoint2text.extractors.data_types: Data structures for extracted content
- MS-DOC, MS-PPT, MS-XLS specifications: https://docs.microsoft.com/en-us/openspecs/
"""

from sharepoint2text.extractors.ms_legacy.doc_extractor import read_doc
from sharepoint2text.extractors.ms_legacy.ppt_extractor import read_ppt
from sharepoint2text.extractors.ms_legacy.rtf_extractor import read_rtf
from sharepoint2text.extractors.ms_legacy.xls_extractor import read_xls

__all__ = [
    "read_doc",
    "read_ppt",
    "read_rtf",
    "read_xls",
]
