"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
import logging
import re
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any, Generator

from sharepoint2text.extractors.data_types import (
    DocContent,
    DocxContent,
    EmailContent,
    EpubContent,
    ExtractionInterface,
    HtmlContent,
    OdpContent,
    OdsContent,
    OdtContent,
    PdfContent,
    PlainTextContent,
    PptContent,
    PptxContent,
    RtfContent,
    XlsContent,
    XlsxContent,
)
from sharepoint2text.router import get_extractor, is_supported_file

logger = logging.getLogger(__name__)

_PRERELEASE_NORMALIZE_RE = re.compile(r"(?<=\d)\.(a|b|rc)(0|[1-9]\d*)\b", re.IGNORECASE)


def _normalize_version(value: str) -> str:
    def repl(match: re.Match[str]) -> str:
        tag = match.group(1).lower()
        number = str(int(match.group(2)))
        return f"{tag}{number}"

    return _PRERELEASE_NORMALIZE_RE.sub(repl, value)


def _version_from_pyproject() -> str | None:
    here = Path(__file__).resolve()
    for parent in list(here.parents)[:5]:
        pyproject = parent / "pyproject.toml"
        if not pyproject.is_file():
            continue
        text = pyproject.read_text(encoding="utf-8", errors="ignore")
        match = re.search(
            r'(?ms)^\[project\]\s.*?^version\s*=\s*["\']([^"\']+)["\']\s*$',
            text,
        )
        return match.group(1) if match else None
    return None


try:
    raw_version = _version_from_pyproject() or version("sharepoint-to-text")
    __version__ = _normalize_version(raw_version)
except PackageNotFoundError:  # pragma: no cover
    __version__ = "unknown"


#############
# Modern MS
#############
def read_docx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocxContent, Any, None]:
    """Extract content from a DOCX file."""
    from sharepoint2text.extractors.ms_modern.docx_extractor import (
        read_docx as _read_docx,
    )

    logger.debug("Reading MS docx file: %s", path)
    return _read_docx(file_like, path)


def read_xlsx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsxContent, Any, None]:
    """Extract content from an XLSX file."""
    from sharepoint2text.extractors.ms_modern.xlsx_extractor import (
        read_xlsx as _read_xlsx,
    )

    logger.debug("Reading MS xlsx file: %s", path)
    return _read_xlsx(file_like, path)


def read_pptx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PptxContent, Any, None]:
    """Extract content from a PPTX file."""
    from sharepoint2text.extractors.ms_modern.pptx_extractor import (
        read_pptx as _read_pptx,
    )

    logger.debug("Reading MS pptx file: %s", path)
    return _read_pptx(file_like, path)


#############
# Legacy MS
#############


def read_doc(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocContent, Any, None]:
    """Extract content from a DOC file."""
    from sharepoint2text.extractors.ms_legacy.doc_extractor import read_doc as _read_doc

    logger.debug("Reading legacy MS doc file: %s", path)
    return _read_doc(file_like, path)


def read_xls(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsContent, Any, None]:
    """Extract content from an XLS file."""
    from sharepoint2text.extractors.ms_legacy.xls_extractor import read_xls as _read_xls

    logger.debug("Reading legacy MS xls file: %s", path)
    return _read_xls(file_like, path)


def read_ppt(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PptContent, Any, None]:
    """Extract content from a PPT file."""
    from sharepoint2text.extractors.ms_legacy.ppt_extractor import read_ppt as _read_ppt

    logger.debug("Reading legacy MS ppt file: %s", path)
    return _read_ppt(file_like, path)


def read_rtf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[RtfContent, Any, None]:
    """Extract content from a RTF file."""
    from sharepoint2text.extractors.ms_legacy.rtf_extractor import read_rtf as _read_rtf

    logger.debug("Reading legacy MS rtf file: %s", path)
    return _read_rtf(file_like, path)


#############
# Open Office
#############


def read_odt(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdtContent, Any, None]:
    """Extract content from an ODT (OpenDocument Text) file."""
    from sharepoint2text.extractors.open_office.odt_extractor import (
        read_odt as _read_odt,
    )

    logger.debug("Reading open office odt file: %s", path)
    return _read_odt(file_like, path)


def read_odp(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdpContent, Any, None]:
    """Extract content from an ODP (OpenDocument Presentation) file."""
    from sharepoint2text.extractors.open_office.odp_extractor import (
        read_odp as _read_odp,
    )

    logger.debug("Reading open office odp file: %s", path)
    return _read_odp(file_like, path)


def read_ods(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdsContent, Any, None]:
    """Extract content from an ODS (OpenDocument Spreadsheet) file."""
    from sharepoint2text.extractors.open_office.ods_extractor import (
        read_ods as _read_ods,
    )

    logger.debug("Reading open office ods file: %s", path)
    return _read_ods(file_like, path)


#############
# Plain Text
#############


def read_plain_text(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PlainTextContent, Any, None]:
    """Extract content from a plain text file."""
    from sharepoint2text.extractors.plain_extractor import (
        read_plain_text as _read_plain_text,
    )

    logger.debug("Reading plain text file: %s", path)
    return _read_plain_text(file_like, path)


#############
# PDF
#############
def read_pdf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PdfContent, Any, None]:
    """Extract content from a PDF file."""
    from sharepoint2text.extractors.pdf.pdf_extractor import read_pdf as _read_pdf

    logger.debug("Reading PDF file: %s", path)
    return _read_pdf(file_like, path)


#############
# HTML
#############
def read_html(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[HtmlContent, Any, None]:
    """Extract content from an HTML file."""
    from sharepoint2text.extractors.html_extractor import read_html as _read_html

    logger.debug("Reading HTML file: %s", path)
    return _read_html(file_like, path)


#############
# EPUB
#############
def read_epub(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EpubContent, Any, None]:
    """Extract content from an EPUB eBook file."""
    from sharepoint2text.extractors.epub_extractor import read_epub as _read_epub

    logger.debug("Reading EPUB file: %s", path)
    return _read_epub(file_like, path)


#############
# MHTML
#############
def read_mhtml(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[HtmlContent, Any, None]:
    """Extract content from an MHTML (web archive) file."""
    from sharepoint2text.extractors.mhtml_extractor import read_mhtml as _read_mhtml

    logger.debug("Reading MHTML file: %s", path)
    return _read_mhtml(file_like, path)


#############
# Emails
#############
def read_email__msg_format(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Extract content from an email in msg format."""
    from sharepoint2text.extractors.mail.msg_email_extractor import (
        read_msg_format_mail as _read_msg_format_mail,
    )

    logger.debug("Reading mail .msg file: %s", path)
    return _read_msg_format_mail(file_like, path)


def read_email__eml_format(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Extract content from an email in eml format."""
    from sharepoint2text.extractors.mail.eml_email_extractor import (
        read_eml_format_mail as _read_eml_format_mail,
    )

    logger.debug("Reading mail .eml file: %s", path)
    return _read_eml_format_mail(file_like, path)


def read_email__mbox_format(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[EmailContent, Any, None]:
    """Extract content from an email in mbox format."""
    from sharepoint2text.extractors.mail.mbox_email_extractor import (
        read_mbox_format_mail as _read_mbox_format_mail,
    )

    logger.debug("Reading mail .mbox file: %s", path)
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
        - .html -> HtmlContent
        - .htm  -> HtmlContent
        - .odt  -> OdtContent
        - .odp  -> OdpContent
        - .ods  -> OdsContent
        - .msg  -> EmailContent
        - .mbox -> EmailContent
        - .eml  -> EmailContent
        - .epub -> EpubContent
        - .mhtml -> HtmlContent
        - .mht  -> HtmlContent

    Raises:
        sharepoint2text.exceptions.ExtractionFileFormatNotSupportedError:
            If the file type is not supported.
        sharepoint2text.exceptions.ExtractionFileEncryptedError:
            If the file is encrypted or password-protected.
        sharepoint2text.exceptions.LegacyMicrosoftParsingError:
            If parsing a legacy Office file fails.
        sharepoint2text.exceptions.ExtractionFailedError:
            If extraction fails for an unexpected reason (with `__cause__` set).
        FileNotFoundError: If the file does not exist.

    The individual extractors are callable separately

    Example:
        >>> import sharepoint2text
        >>> for result in sharepoint2text.read_file("document.docx"):
        ...     print(result.get_full_text())
    """
    from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError

    path = Path(path)
    logger.info("Starting extraction: %s", path)
    extractor = get_extractor(str(path))
    with open(path, "rb") as f:
        try:
            for result in extractor(io.BytesIO(f.read()), str(path)):
                logger.info("Extraction complete: %s", path)
                yield result
        except ExtractionError:
            raise
        except Exception as exc:
            raise ExtractionFailedError(
                f"Failed to extract file: {path}", cause=exc
            ) from exc


__all__ = [
    # Version
    "__version__",
    # Main functions
    "read_file",
    "is_supported_file",
    "get_extractor",
    # legacy MS office
    "read_doc",
    "read_xls",
    "read_ppt",
    # modern office
    "read_docx",
    "read_xlsx",
    "read_pptx",
    "read_rtf",
    "read_plain_text",
    "read_html",
    # open office
    "read_odt",
    "read_odp",
    "read_ods",
    "read_email__msg_format",
    "read_email__eml_format",
    "read_email__mbox_format",
    # other
    "read_pdf",
    "read_epub",
    "read_mhtml",
]
