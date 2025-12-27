"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
from pathlib import Path
from typing import Any, Generator

from sharepoint2text.extractors.data_types import (
    DocContent,
    DocxContent,
    EmailContent,
    ExtractionInterface,
    PdfContent,
    PlainTextContent,
    PptContent,
    PptxContent,
    XlsContent,
    XlsxContent,
)
from sharepoint2text.router import get_extractor, is_supported_file

__version__ = "0.1.1"


def read_docx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocxContent, Any, None]:
    """Extract content from a DOCX file."""
    from sharepoint2text.extractors.docx_extractor import read_docx as _read_docx

    return _read_docx(file_like, path)


def read_doc(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocContent, Any, None]:
    """Extract content from a DOC file."""
    from sharepoint2text.extractors.doc_extractor import read_doc as _read_doc

    return _read_doc(file_like, path)


def read_xlsx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsxContent, Any, None]:
    """Extract content from an XLSX file."""
    from sharepoint2text.extractors.xlsx_extractor import read_xlsx as _read_xlsx

    return _read_xlsx(file_like, path)


def read_xls(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsContent, Any, None]:
    """Extract content from an XLS file."""
    from sharepoint2text.extractors.xls_extractor import read_xls as _read_xls

    return _read_xls(file_like, path)


def read_pptx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PptxContent, Any, None]:
    """Extract content from a PPTX file."""
    from sharepoint2text.extractors.pptx_extractor import read_pptx as _read_pptx

    return _read_pptx(file_like, path)


def read_ppt(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PptContent, Any, None]:
    """Extract content from a PPT file."""
    from sharepoint2text.extractors.ppt_extractor import read_ppt as _read_ppt

    return _read_ppt(file_like, path)


def read_pdf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PdfContent, Any, None]:
    """Extract content from a PDF file."""
    from sharepoint2text.extractors.pdf_extractor import read_pdf as _read_pdf

    return _read_pdf(file_like, path)


def read_plain_text(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PlainTextContent, Any, None]:
    """Extract content from a plain text file."""
    from sharepoint2text.extractors.plain_extractor import (
        read_plain_text as _read_plain_text,
    )

    return _read_plain_text(file_like, path)


def read_email__msg_format(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Extract content from an email in msg format."""
    from sharepoint2text.extractors.mail.msg_email_extractor import (
        read_msg_format_mail as _read_msg_format_mail,
    )

    return _read_msg_format_mail(file_like, path)


def read_email__eml_format(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Extract content from an email in eml format."""
    from sharepoint2text.extractors.mail.eml_email_extractor import (
        read_eml_format_mail as _read_eml_format_mail,
    )

    return _read_eml_format_mail(file_like, path)


def read_email__mbox_format(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Extract content from an email in mbox format."""
    from sharepoint2text.extractors.mail.mbox_email_extractor import (
        read_mbox_format_mail as _read_mbox_format_mail,
    )

    return _read_mbox_format_mail(file_like, path)


def read_file(
    path: str | Path,
) -> Generator[ExtractionInterface, Any, None]:
    """
    Read and extract content from a file.

    Automatically detects the file type based on extension and uses
    the appropriate extractor.

    Args:
        path: Path to the file to read.

    Yields:
        A dataclass containing extracted content and metadata.
        The specific type depends on the file format:
        - .docx -> DocxContent
        - .doc  -> DocContent
        - .xlsx -> XlsxContent
        - .xls  -> XlsContent
        - .pptx -> PptxContent
        - .ppt  -> PptContent
        - .pdf  -> PdfContent
        - .txt  -> PlainTextContent
        - .msg  -> EmailContent
        - .mbox -> EmailContent
        - .eml  -> EmailContent

    Raises:
        RuntimeError: If the file type is not supported.
        FileNotFoundError: If the file does not exist.

    The individual extractors are callable separately

    Example:
        >>> import sharepoint2text
        >>> for result in sharepoint2text.read_file("document.docx"):
        ...     print(result.get_full_text())
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
    "read_email__msg_format",
    "read_email__eml_format",
    "read_email__mbox_format",
]
