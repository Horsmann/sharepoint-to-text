"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
import logging
import re
import sys
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any, Generator

from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
    ExtractionFileFormatNotSupportedError,
    ExtractionFileTooLargeError,
    ExtractionLegacyMicrosoftParsingError,
    ExtractionZipBombError,
)
from sharepoint2text.parsing.extractors.data_types import (
    CsvContent,
    DocContent,
    DocxContent,
    EmailContent,
    EpubContent,
    ExtractionInterface,
    HtmlContent,
    OdfContent,
    OdgContent,
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
from sharepoint2text.parsing.mime_types import MIME_TYPE_MAPPING
from sharepoint2text.parsing.router import get_extractor, is_supported_file

logger = logging.getLogger(__name__)

# Default pypdf decompression limit (bytes). pypdf uses 75 MB internally.
_PYPDF_DEFAULT_DECOMPRESSION_LIMIT = 75_000_000

# Sentinel value to effectively disable pypdf limits.  pypdf checks
# ``length > MAX_…`` so any stream smaller than this passes.  We cannot
# use ``0`` because that would make *every* stream exceed the limit.
# ``sys.maxsize`` is the platform's ``ssize_t`` max, accepted by zlib.
_PYPDF_NO_LIMIT = sys.maxsize


def _configure_pypdf_limits(max_file_size: int) -> None:
    """Adjust pypdf's internal decompression limits to match *max_file_size*.

    pypdf caps zlib/LZW/RunLength/JBIG2 output and declared stream lengths
    at 75 MB.  When the caller raises (or disables) the sharepoint2text size
    limit we must propagate that to pypdf, otherwise large but legitimate
    PDFs still hit ``LimitReachedError``.

    Args:
        max_file_size: The user-supplied limit in bytes.
                       ``0`` means "no limit" – pypdf limits are disabled.
    """
    import pypdf.filters as _filters  # noqa: PLC0415

    if max_file_size == 0:
        # Disable all pypdf decompression limits.
        target = _PYPDF_NO_LIMIT
    elif max_file_size > _PYPDF_DEFAULT_DECOMPRESSION_LIMIT:
        # Use the user's limit (compressed content can expand, so match it).
        target = max_file_size
    else:
        # User limit is at or below pypdf's default – leave pypdf defaults.
        return

    for attr in (
        "ZLIB_MAX_OUTPUT_LENGTH",
        "LZW_MAX_OUTPUT_LENGTH",
        "RUN_LENGTH_MAX_OUTPUT_LENGTH",
        "JBIG2_MAX_OUTPUT_LENGTH",
        "MAX_DECLARED_STREAM_LENGTH",
        "MAX_ARRAY_BASED_STREAM_OUTPUT_LENGTH",
    ):
        if hasattr(_filters, attr):
            setattr(_filters, attr, target)


_PRERELEASE_NORMALIZE_RE = re.compile(r"(?<=\d)\.(a|b|rc)(0|[1-9]\d*)\b", re.IGNORECASE)


def _normalize_version(value: str) -> str:
    def repl(match: re.Match[str]) -> str:
        tag = match.group(1).lower()
        number = str(int(match.group(2)))
        return f"{tag}{number}"

    return _PRERELEASE_NORMALIZE_RE.sub(repl, value)


def _version_from_pyproject() -> str | None:
    here = Path(__file__).resolve()
    for parent in list(here.parents)[:3]:
        pyproject = parent / "pyproject.toml"
        if not pyproject.is_file():
            continue
        text = pyproject.read_text(encoding="utf-8", errors="ignore")
        # Verify this is our project's pyproject.toml
        if 'name = "sharepoint-to-text"' not in text:
            continue
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
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[DocxContent, Any, None]:
    """Extract content from a DOCX file."""
    from sharepoint2text.parsing.extractors.ms_modern.docx_extractor import (
        read_docx as _read_docx,
    )

    logger.debug("Reading MS docx file: %s", path)
    return _read_docx(file_like, path, ignore_images=ignore_images)


def read_xlsx(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[XlsxContent, Any, None]:
    """Extract content from an XLSX file."""
    from sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor import (
        read_xlsx as _read_xlsx,
    )

    logger.debug("Reading MS xlsx file: %s", path)
    return _read_xlsx(file_like, path, ignore_images=ignore_images)


def read_pptx(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[PptxContent, Any, None]:
    """Extract content from a PPTX file."""
    from sharepoint2text.parsing.extractors.ms_modern.pptx_extractor import (
        read_pptx as _read_pptx,
    )

    logger.debug("Reading MS pptx file: %s", path)
    return _read_pptx(file_like, path, ignore_images=ignore_images)


#############
# Legacy MS
#############


def read_doc(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[DocContent, Any, None]:
    """Extract content from a DOC file."""
    from sharepoint2text.parsing.extractors.ms_legacy.doc_extractor import (
        read_doc as _read_doc,
    )

    logger.debug("Reading legacy MS doc file: %s", path)
    return _read_doc(file_like, path, ignore_images=ignore_images)


def read_xls(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[XlsContent, Any, None]:
    """Extract content from an XLS file."""
    from sharepoint2text.parsing.extractors.ms_legacy.xls_extractor import (
        read_xls as _read_xls,
    )

    logger.debug("Reading legacy MS xls file: %s", path)
    return _read_xls(file_like, path, ignore_images=ignore_images)


def read_ppt(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[PptContent, Any, None]:
    """Extract content from a PPT file."""
    from sharepoint2text.parsing.extractors.ms_legacy.ppt_extractor import (
        read_ppt as _read_ppt,
    )

    logger.debug("Reading legacy MS ppt file: %s", path)
    return _read_ppt(file_like, path, ignore_images=ignore_images)


def read_rtf(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[RtfContent, Any, None]:
    """Extract content from a RTF file."""
    from sharepoint2text.parsing.extractors.ms_legacy.rtf_extractor import (
        read_rtf as _read_rtf,
    )

    logger.debug("Reading legacy MS rtf file: %s", path)
    return _read_rtf(file_like, path, ignore_images=ignore_images)


#############
# Open Office
#############


def read_odt(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[OdtContent, Any, None]:
    """Extract content from an ODT (OpenDocument Text) file."""
    from sharepoint2text.parsing.extractors.open_office.odt_extractor import (
        read_odt as _read_odt,
    )

    logger.debug("Reading open office odt file: %s", path)
    return _read_odt(file_like, path, ignore_images=ignore_images)


def read_odp(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[OdpContent, Any, None]:
    """Extract content from an ODP (OpenDocument Presentation) file."""
    from sharepoint2text.parsing.extractors.open_office.odp_extractor import (
        read_odp as _read_odp,
    )

    logger.debug("Reading open office odp file: %s", path)
    return _read_odp(file_like, path, ignore_images=ignore_images)


def read_ods(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[OdsContent, Any, None]:
    """Extract content from an ODS (OpenDocument Spreadsheet) file."""
    from sharepoint2text.parsing.extractors.open_office.ods_extractor import (
        read_ods as _read_ods,
    )

    logger.debug("Reading open office ods file: %s", path)
    return _read_ods(file_like, path, ignore_images=ignore_images)


def read_odg(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[OdgContent, Any, None]:
    """Extract content from an ODG (OpenDocument Drawing) file."""
    from sharepoint2text.parsing.extractors.open_office.odg_extractor import (
        read_odg as _read_odg,
    )

    logger.debug("Reading open office odg file: %s", path)
    return _read_odg(file_like, path, ignore_images=ignore_images)


def read_odf(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[OdfContent, Any, None]:
    """Extract content from an ODF (OpenDocument Formula) file."""
    from sharepoint2text.parsing.extractors.open_office.odf_extractor import (
        read_odf as _read_odf,
    )

    logger.debug("Reading open office odf file: %s", path)
    return _read_odf(file_like, path, ignore_images=ignore_images)


#############
# Plain Text
#############


def read_plain_text(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[PlainTextContent, Any, None]:
    """Extract content from a plain text file."""
    from sharepoint2text.parsing.extractors.plain_extractor import (
        read_plain_text as _read_plain_text,
    )

    logger.debug("Reading plain text file: %s", path)
    return _read_plain_text(file_like, path, ignore_images=ignore_images)


#############
# CSV / TSV
#############
def read_csv(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[CsvContent, Any, None]:
    """Extract structured content from a CSV or TSV file."""
    from sharepoint2text.parsing.extractors.csv_extractor import read_csv as _read_csv

    logger.debug("Reading CSV/TSV file: %s", path)
    return _read_csv(file_like, path, ignore_images=ignore_images)


#############
# PDF
#############
def read_pdf(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[PdfContent, Any, None]:
    """Extract content from a PDF file."""
    from sharepoint2text.parsing.extractors.pdf.pdf_extractor import (
        read_pdf as _read_pdf,
    )

    logger.debug("Reading PDF file: %s", path)
    return _read_pdf(file_like, path, ignore_images=ignore_images)


#############
# HTML
#############
def read_html(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[HtmlContent, Any, None]:
    """Extract content from an HTML file."""
    from sharepoint2text.parsing.extractors.html_extractor import (
        read_html as _read_html,
    )

    logger.debug("Reading HTML file: %s", path)
    return _read_html(file_like, path, ignore_images=ignore_images)


#############
# EPUB
#############
def read_epub(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[EpubContent, Any, None]:
    """Extract content from an EPUB eBook file."""
    from sharepoint2text.parsing.extractors.epub_extractor import (
        read_epub as _read_epub,
    )

    logger.debug("Reading EPUB file: %s", path)
    return _read_epub(file_like, path, ignore_images=ignore_images)


#############
# MHTML
#############
def read_mhtml(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[HtmlContent, Any, None]:
    """Extract content from an MHTML (web archive) file."""
    from sharepoint2text.parsing.extractors.mhtml_extractor import (
        read_mhtml as _read_mhtml,
    )

    logger.debug("Reading MHTML file: %s", path)
    return _read_mhtml(file_like, path, ignore_images=ignore_images)


#############
# Emails
#############
def read_msg_email(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[EmailContent, Any, None]:
    """Extract content from an Outlook MSG email file."""
    from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
        read_msg_format_mail as _read_msg_format_mail,
    )

    logger.debug("Reading mail .msg file: %s", path)
    return _read_msg_format_mail(file_like, path, ignore_images=ignore_images)


def read_eml_email(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[EmailContent, Any, None]:
    """Extract content from an EML email file."""
    from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
        read_eml_format_mail as _read_eml_format_mail,
    )

    logger.debug("Reading mail .eml file: %s", path)
    return _read_eml_format_mail(file_like, path, ignore_images=ignore_images)


def read_mbox_email(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[EmailContent, Any, None]:
    """Extract content from an MBOX email file."""
    from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
        read_mbox_format_mail as _read_mbox_format_mail,
    )

    logger.debug("Reading mail .mbox file: %s", path)
    return _read_mbox_format_mail(file_like, path, ignore_images=ignore_images)


def read_file(
    path: str | Path,
    max_file_size: int = 100 * 1024 * 1024,  # 100MB default
    *,
    ignore_images: bool = False,
    force_plain_text: bool = False,
) -> Generator[ExtractionInterface, Any, None]:
    """
    Read and extract content from a file.

    Automatically detects the file type based on extension and uses
    the appropriate extractor.

    Args:
        path: Path to the file to read.
        max_file_size: Maximum file size in bytes (default: 100MB).
                      Set to 0 to disable size checking.
        ignore_images: If True, skip image extraction. This can significantly
                      improve performance for files with many images.
                      Default is False.
        force_plain_text: If True, route extraction to plain text handling
                      regardless of extension/MIME detection.
                      Useful for unknown or custom plain-text file formats.

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
        - .odg  -> OdgContent
        - .odf  -> OdfContent
        - .msg  -> EmailContent
        - .mbox -> EmailContent
        - .eml  -> EmailContent
        - .epub -> EpubContent
        - .mhtml -> HtmlContent
        - .mht  -> HtmlContent

    Raises:
        sharepoint2text.parsing.exceptions.ExtractionFileFormatNotSupportedError:
            If the file type is not supported.
        sharepoint2text.parsing.exceptions.ExtractionFileEncryptedError:
            If the file is encrypted or password-protected.
        sharepoint2text.parsing.exceptions.ExtractionLegacyMicrosoftParsingError:
            If parsing a legacy Office file fails.
        sharepoint2text.parsing.exceptions.ExtractionFailedError:
            If extraction fails for an unexpected reason (with `__cause__` set).
        sharepoint2text.parsing.exceptions.ExtractionFileTooLargeError:
            If the file exceeds the maximum allowed size.
        FileNotFoundError: If the file does not exist.

    The individual extractors are callable separately

    Example:
        >>> import sharepoint2text
        >>> for result in sharepoint2text.read_file("document.docx"):
        ...     print(result.get_full_text())
        >>> # Skip image extraction for faster processing
        >>> for result in sharepoint2text.read_file("document.docx", ignore_images=True):
        ...     print(result.get_full_text())
    """
    from sharepoint2text.parsing.exceptions import (
        ExtractionError,
        ExtractionFailedError,
        ExtractionFileTooLargeError,
    )

    path = Path(path)

    _configure_pypdf_limits(max_file_size)

    # Check file size before reading
    if max_file_size > 0:
        file_size = path.stat().st_size
        if file_size > max_file_size:
            raise ExtractionFileTooLargeError(
                f"File size {file_size} bytes exceeds maximum allowed size of {max_file_size} bytes",
                max_size=max_file_size,
                actual_size=file_size,
            )

    logger.info("Starting extraction: %s", path)
    extractor = get_extractor(
        str(path),
        ignore_images=ignore_images,
        force_plain_text=force_plain_text,
    )
    with open(path, "rb") as f:
        try:
            # Pass the file stream directly to avoid an unnecessary full-copy
            # buffer before extraction.
            for result in extractor(f, str(path)):
                logger.info("Extraction complete: %s", path)
                yield result
        except ExtractionError:
            raise
        except (OSError, ValueError, TypeError, UnicodeDecodeError) as exc:
            raise ExtractionFailedError(
                f"Failed to extract file: {path}", cause=exc
            ) from exc


def read_bytes(
    data: bytes | io.BytesIO,
    *,
    mime_type: str | None = None,
    extension: str | None = None,
    max_file_size: int = 100 * 1024 * 1024,  # 100MB default
    ignore_images: bool = False,
    force_plain_text: bool = False,
) -> Generator[ExtractionInterface, Any, None]:
    """
    Read and extract content from in-memory bytes.

    This is an alternative to ``read_file`` for content that is already loaded
    in memory. Routing is done via ``extension`` (preferred) or ``mime_type``.

    Args:
        data: Raw file bytes or a ``io.BytesIO`` buffer.
        mime_type: MIME type hint (for example ``"application/pdf"``).
        extension: File extension hint (for example ``"pdf"`` or ``".pdf"``).
        max_file_size: Maximum file size in bytes (default: 100MB).
                      Set to 0 to disable size checking.
        ignore_images: If True, skip image extraction. This can significantly
                      improve performance for files with many images.
                      Default is False.
        force_plain_text: If True, route extraction to plain text handling
                      regardless of extension/MIME detection.
                      Useful for unknown or custom plain-text file formats.

    Yields:
        A dataclass containing extracted content and metadata.

    Raises:
        ValueError: If both ``mime_type`` and ``extension`` are missing/empty,
            unless ``force_plain_text=True``.
        TypeError: If ``data`` is not ``bytes`` or ``io.BytesIO``.
        sharepoint2text.parsing.exceptions.ExtractionFileFormatNotSupportedError:
            If the provided extension/MIME type is unsupported.
        sharepoint2text.parsing.exceptions.ExtractionFileEncryptedError:
            If the file is encrypted or password-protected.
        sharepoint2text.parsing.exceptions.ExtractionLegacyMicrosoftParsingError:
            If parsing a legacy Office file fails.
        sharepoint2text.parsing.exceptions.ExtractionFailedError:
            If extraction fails for an unexpected reason (with ``__cause__`` set).
        sharepoint2text.parsing.exceptions.ExtractionFileTooLargeError:
            If the file exceeds the maximum allowed size.
    """
    if not isinstance(data, (bytes, io.BytesIO)):
        raise TypeError("data must be bytes or io.BytesIO")

    _configure_pypdf_limits(max_file_size)

    normalized_extension = extension.strip().lower() if extension else ""
    normalized_mime_type = mime_type.strip().lower() if mime_type else ""

    if normalized_extension.startswith("."):
        normalized_extension = normalized_extension[1:]

    if isinstance(data, bytes):
        file_size = len(data)
        file_like = io.BytesIO(data)
    else:
        file_size = data.getbuffer().nbytes
        data.seek(0)
        file_like = data

    # Check file size before extraction
    if max_file_size > 0 and file_size > max_file_size:
        raise ExtractionFileTooLargeError(
            f"File size {file_size} bytes exceeds maximum allowed size of {max_file_size} bytes",
            max_size=max_file_size,
            actual_size=file_size,
        )

    extractor = None
    virtual_path = "<in-memory>"
    extension_error: ExtractionFileFormatNotSupportedError | None = None

    if force_plain_text:
        virtual_path = "in_memory.txt"
        extractor = get_extractor(
            virtual_path,
            ignore_images=ignore_images,
            force_plain_text=True,
        )
    else:
        if not normalized_extension and not normalized_mime_type:
            raise ValueError("Either mime_type or extension must be provided")

        if normalized_extension:
            virtual_path = f"in_memory.{normalized_extension}"
            try:
                extractor = get_extractor(virtual_path, ignore_images=ignore_images)
            except ExtractionFileFormatNotSupportedError as exc:
                extension_error = exc
                if not normalized_mime_type:
                    raise

        if extractor is None and normalized_mime_type:
            file_type = MIME_TYPE_MAPPING.get(normalized_mime_type)
            if file_type is None:
                if extension_error is not None:
                    raise extension_error
                raise ExtractionFileFormatNotSupportedError(
                    f"File type not supported for MIME type '{normalized_mime_type}'"
                )
            virtual_path = f"in_memory.{file_type}"
            extractor = get_extractor(virtual_path, ignore_images=ignore_images)

        if extractor is None and extension_error is not None:
            raise extension_error

        if extractor is None:
            raise ExtractionFileFormatNotSupportedError(
                "Could not resolve extractor from provided extension/MIME type"
            )

    logger.info("Starting in-memory extraction: %s", virtual_path)
    try:
        for result in extractor(file_like, virtual_path):
            logger.info("In-memory extraction complete: %s", virtual_path)
            yield result
    except ExtractionError:
        raise
    except (OSError, ValueError, TypeError, UnicodeDecodeError) as exc:
        raise ExtractionFailedError(
            f"Failed to extract in-memory data: {virtual_path}",
            cause=exc,
        ) from exc


__all__ = [
    # Version
    "__version__",
    # Main functions
    "read_file",
    "read_bytes",
    "is_supported_file",
    "get_extractor",
    # exceptions
    "ExtractionError",
    "ExtractionFileFormatNotSupportedError",
    "ExtractionLegacyMicrosoftParsingError",
    "ExtractionFileEncryptedError",
    "ExtractionZipBombError",
    "ExtractionFailedError",
    "ExtractionFileTooLargeError",
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
    "read_csv",
    "read_html",
    # open office
    "read_odt",
    "read_odp",
    "read_ods",
    "read_odg",
    "read_odf",
    "read_msg_email",
    "read_eml_email",
    "read_mbox_email",
    # other
    "read_pdf",
    "read_epub",
    "read_mhtml",
    # content types
    "CsvContent",
    "DocContent",
    "DocxContent",
    "EmailContent",
    "EpubContent",
    "HtmlContent",
    "OdfContent",
    "OdgContent",
    "OdpContent",
    "OdsContent",
    "OdtContent",
    "PdfContent",
    "PlainTextContent",
    "PptContent",
    "PptxContent",
    "RtfContent",
    "XlsContent",
    "XlsxContent",
]
