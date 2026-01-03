"""
OpenDocument Format (ODF) Extractor Package
============================================

This package provides extractors for parsing and extracting text content from
OpenDocument Format (ODF) files. ODF is an open XML-based document format
standardized by OASIS and ISO (ISO/IEC 26300).

Supported Formats
-----------------

.odt (OpenDocument Text):
    Word processing documents. Similar to Microsoft Word's .docx format.
    Extracts paragraphs, tables, headers/footers, footnotes, images,
    hyperlinks, bookmarks, and annotations.

.odp (OpenDocument Presentation):
    Presentation documents. Similar to Microsoft PowerPoint's .pptx format.
    Extracts slides with titles, body text, speaker notes, tables, and images.

.ods (OpenDocument Spreadsheet):
    Spreadsheet documents. Similar to Microsoft Excel's .xlsx format.
    Extracts sheets with cell data, annotations, and images.

File Format Background
----------------------
ODF files are ZIP archives containing XML files. The structure is defined
by the OASIS OpenDocument specification. Key components:

    content.xml: Main document content (text, tables, drawings)
    meta.xml: Document metadata (title, author, dates)
    styles.xml: Style definitions and master pages
    settings.xml: Application settings
    Pictures/: Embedded images folder
    mimetype: File type identifier (uncompressed, first in archive)

All ODF formats share a common XML vocabulary based on these namespaces:
    - office: urn:oasis:names:tc:opendocument:xmlns:office:1.0
    - text: urn:oasis:names:tc:opendocument:xmlns:text:1.0
    - table: urn:oasis:names:tc:opendocument:xmlns:table:1.0
    - draw: urn:oasis:names:tc:opendocument:xmlns:drawing:1.0
    - style: urn:oasis:names:tc:opendocument:xmlns:style:1.0
    - meta: urn:oasis:names:tc:opendocument:xmlns:meta:1.0
    - dc: http://purl.org/dc/elements/1.1/ (Dublin Core metadata)

Dependencies
------------
Python Standard Library only:
    - zipfile: Archive extraction
    - xml.etree.ElementTree: XML parsing
    - mimetypes: Image type detection

No external dependencies required for basic ODF parsing.

Metadata (Dublin Core)
----------------------
ODF uses Dublin Core metadata elements in meta.xml:
    - dc:title: Document title
    - dc:description: Document description
    - dc:subject: Document subject
    - dc:creator: Last modifier
    - dc:date: Last modification date
    - dc:language: Document language

Additional ODF-specific metadata:
    - meta:initial-creator: Original creator
    - meta:creation-date: Creation date
    - meta:editing-cycles: Number of edit sessions
    - meta:editing-duration: Total editing time (ISO 8601 duration)
    - meta:generator: Application that created the document

Text Element Handling
---------------------
ODF has special elements for preserving whitespace and formatting:
    - text:s (with text:c attribute): Multiple spaces
    - text:tab: Tab character
    - text:line-break: Line break (soft return)

The extractors handle these elements to preserve document formatting.

Known Limitations
-----------------
- Encrypted/password-protected files are not supported
- Embedded OLE objects are not extracted
- Complex drawings (shapes, charts) may not extract all text
- Form controls are not extracted
- Change tracking (revisions) is not separately reported
- Math formulas (ODF formula objects) are not converted to LaTeX

Usage Example
-------------
    >>> from sharepoint2text.parsing.extractors.open_office import read_odt, read_odp
    >>> import io
    >>>
    >>> with open("document.odt", "rb") as f:
    ...     for doc in read_odt(io.BytesIO(f.read())):
    ...         print(doc.full_text)
    ...
    >>> with open("slides.odp", "rb") as f:
    ...     for ppt in read_odp(io.BytesIO(f.read())):
    ...         for slide in ppt.slides:
    ...             print(f"Slide {slide.slide_number}: {slide.title}")

Comparison with Microsoft Formats
---------------------------------
ODF files are conceptually similar to Microsoft OOXML formats:
    - .odt <-> .docx (Word)
    - .odp <-> .pptx (PowerPoint)
    - .ods <-> .xlsx (Excel)

Both are ZIP archives with XML content, but use different schemas and
namespaces. ODF is fully open and standardized by ISO.

Applications that create ODF files include:
    - LibreOffice / OpenOffice
    - Google Docs (export)
    - Apple iWork (export)
    - Microsoft Office (with ODF support enabled)

See Also
--------
- OASIS OpenDocument specification: https://www.oasis-open.org/committees/office/
- ISO/IEC 26300: https://www.iso.org/standard/66363.html
- sharepoint2text.parsing.extractors.ms_modern: For Microsoft OOXML formats
"""

from sharepoint2text.parsing.extractors.open_office.odp_extractor import read_odp
from sharepoint2text.parsing.extractors.open_office.ods_extractor import read_ods
from sharepoint2text.parsing.extractors.open_office.odt_extractor import read_odt

__all__ = [
    "read_odt",
    "read_odp",
    "read_ods",
]
