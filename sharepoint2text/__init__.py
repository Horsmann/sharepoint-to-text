"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
from pathlib import Path

from sharepoint2text.router import get_extractor, is_supported_file

__version__ = "0.1.1"


def read_docx(file_like: io.BytesIO, path: str | None = None):
    """Extract content from a DOCX file."""
    from sharepoint2text.extractors.docx_extractor import read_docx as _read_docx

    return _read_docx(file_like, path)


def read_doc(file_like: io.BytesIO, path: str | None = None):
    """Extract content from a DOC file."""
    from sharepoint2text.extractors.doc_extractor import read_doc as _read_doc

    return _read_doc(file_like, path)


def read_xlsx(file_like: io.BytesIO, path: str | None = None):
    """Extract content from an XLSX file."""
    from sharepoint2text.extractors.xlsx_extractor import read_xlsx as _read_xlsx

    return _read_xlsx(file_like, path)


def read_xls(file_like: io.BytesIO, path: str | None = None):
    """Extract content from an XLS file."""
    from sharepoint2text.extractors.xls_extractor import read_xls as _read_xls

    return _read_xls(file_like, path)


def read_pptx(file_like: io.BytesIO, path: str | None = None):
    """Extract content from a PPTX file."""
    from sharepoint2text.extractors.pptx_extractor import read_pptx as _read_pptx

    return _read_pptx(file_like, path)


def read_ppt(file_like: io.BytesIO, path: str | None = None):
    """Extract content from a PPT file."""
    from sharepoint2text.extractors.ppt_extractor import read_ppt as _read_ppt

    return _read_ppt(file_like, path)


def read_pdf(file_like: io.BytesIO, path: str | None = None):
    """Extract content from a PDF file."""
    from sharepoint2text.extractors.pdf_extractor import read_pdf as _read_pdf

    return _read_pdf(file_like, path)


def read_plain_text(file_like: io.BytesIO, path: str | None = None):
    """Extract content from a plain text file."""
    from sharepoint2text.extractors.plain_extractor import (
        read_plain_text as _read_plain_text,
    )

    return _read_plain_text(file_like, path)


def read_file(path: str | Path):
    """
    Read and extract content from a file.

    Automatically detects the file type based on extension and uses
    the appropriate extractor.

    Args:
        path: Path to the file to read.

    Returns:
        A dataclass containing extracted content and metadata.
        The specific type depends on the file format:
        - .docx -> MicrosoftDocxContent
        - .doc  -> MicrosoftDocContent
        - .xlsx -> MicrosoftXlsxContent
        - .xls  -> MicrosoftXlsContent
        - .pptx -> MicrosoftPptxContent
        - .ppt  -> PPTContent
        - .pdf  -> PdfContent
        - .txt  -> PlainTextContent

    Raises:
        RuntimeError: If the file type is not supported.
        FileNotFoundError: If the file does not exist.

    Example:
        >>> import sharepoint2text
        >>> result = sharepoint2text.read_file("document.docx")
        >>> print(result.get_full_text())
    """
    path = Path(path)
    extractor = get_extractor(str(path))
    with open(path, "rb") as f:
        return extractor(io.BytesIO(f.read()), str(path))


__all__ = [
    # Version
    "__version__",
    # Main functions
    "read_file",
    "is_supported_file",
    "get_extractor",
    # Format-specific extractors
    "read_docx",
    "read_doc",
    "read_xlsx",
    "read_xls",
    "read_pptx",
    "read_ppt",
    "read_pdf",
    "read_plain_text",
]
