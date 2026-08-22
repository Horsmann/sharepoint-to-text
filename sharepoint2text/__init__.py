"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
import logging
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
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    ZipBombLimits,
    get_zip_bomb_limits,
    reset_zip_bomb_limits,
    set_zip_bomb_limits,
)
from sharepoint2text.parsing.mime_types import MIME_TYPE_MAPPING
from sharepoint2text.parsing.router import (
    _resolve_file_type,
    get_extractor,
    is_supported_file,
)

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


try:
    __version__ = version("sharepoint-to-text")
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
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
    include_attachments: bool = True,
) -> Generator[EmailContent, Any, None]:
    """Extract content from an Outlook MSG email file."""
    from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
        read_msg_format_mail as _read_msg_format_mail,
    )

    logger.debug("Reading mail .msg file: %s", path)
    return _read_msg_format_mail(
        file_like,
        path,
        ignore_images=ignore_images,
        include_attachments=include_attachments,
    )


def read_eml_email(
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
    include_attachments: bool = True,
) -> Generator[EmailContent, Any, None]:
    """Extract content from an EML email file."""
    from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
        read_eml_format_mail as _read_eml_format_mail,
    )

    logger.debug("Reading mail .eml file: %s", path)
    return _read_eml_format_mail(
        file_like,
        path,
        ignore_images=ignore_images,
        include_attachments=include_attachments,
    )


def read_mbox_email(
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
    include_attachments: bool = True,
) -> Generator[EmailContent, Any, None]:
    """Extract content from an MBOX email file."""
    from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
        read_mbox_format_mail as _read_mbox_format_mail,
    )

    logger.debug("Reading mail .mbox file: %s", path)
    return _read_mbox_format_mail(
        file_like,
        path,
        ignore_images=ignore_images,
        include_attachments=include_attachments,
    )


class InvalidConfigurationError(ValueError):
    """Raised when incompatible configuration options are provided."""


def read_many(
    folder_path: str | Path,
    suffixes: list[str] | None = None,
    *,
    extract_all_supported: bool = False,
    max_file_size: int = 100 * 1024 * 1024,  # 100MB default
    ignore_images: bool = False,
    force_plain_text: bool = False,
    include_attachments: bool = True,
    recursive: bool = True,
) -> Generator[ExtractionInterface, Any, None]:
    """
    Extract content from multiple files in a folder.

    Traverses a folder (optionally recursively) and extracts content from files
    matching the specified suffixes, or all supported files if extract_all_supported
    is True.

    Args:
        folder_path: Path to the folder to traverse.
        suffixes: List of file suffixes to extract (e.g., [".docx", ".pdf"]).
                 Suffixes should include the leading dot.
                 Required if extract_all_supported is False.
        extract_all_supported: If True, extract all files with supported formats,
            ignoring the suffixes parameter. When combined with
            ``force_plain_text=True``, extract every regular file. Default is False.
        max_file_size: Maximum file size in bytes (default: 100MB).
                      Set to 0 to disable size checking.
        ignore_images: If True, skip image extraction. Default is False.
        force_plain_text: If True, treat selected files as plain text, including
            files with unknown extensions in ``extract_all_supported`` mode.
        include_attachments: If False, skip email attachment extraction.
        recursive: If True, traverse subdirectories recursively. Default is True.

    Yields:
        ExtractionInterface objects for each successfully extracted file.

    Raises:
        InvalidConfigurationError: If both suffixes are provided and
            extract_all_supported is True.
        ValueError: If neither suffixes nor extract_all_supported is specified.
        NotADirectoryError: If folder_path is not a directory.
        FileNotFoundError: If folder_path does not exist.

    Example:
        >>> import sharepoint2text
        >>> # Extract only Word and PDF files
        >>> for result in sharepoint2text.read_many("/path/to/folder", [".docx", ".pdf"]):
        ...     print(f"{result.get_metadata().file_path}: {len(result.get_full_text())} chars")
        >>> # Extract all supported file types
        >>> for result in sharepoint2text.read_many("/path/to/folder", extract_all_supported=True):
        ...     print(result.get_full_text())
    """
    import glob as glob_module

    folder = Path(folder_path)

    # Validate folder exists and is a directory
    if not folder.exists():
        raise FileNotFoundError(f"Folder not found: {folder_path}")
    if not folder.is_dir():
        raise NotADirectoryError(f"Path is not a directory: {folder_path}")

    # Validate configuration
    has_suffixes = suffixes is not None and len(suffixes) > 0
    if has_suffixes and extract_all_supported:
        raise InvalidConfigurationError(
            "Cannot specify both 'suffixes' and 'extract_all_supported=True'. "
            "Use either suffixes to filter specific file types, or "
            "extract_all_supported=True to extract all supported formats."
        )

    if not has_suffixes and not extract_all_supported:
        raise ValueError(
            "Must specify either 'suffixes' (list of file extensions to extract) "
            "or 'extract_all_supported=True' to extract all supported formats."
        )

    # Normalize suffixes to ensure they start with a dot
    normalized_suffixes: set[str] = set()
    if has_suffixes and suffixes is not None:
        for suffix in suffixes:
            normalized = suffix.strip().lower()
            if not normalized.startswith("."):
                normalized = f".{normalized}"
            normalized_suffixes.add(normalized)

    # Build glob pattern
    pattern = "**/*" if recursive else "*"
    glob_path = str(folder / pattern)

    # Track statistics for logging
    files_found = 0
    files_extracted = 0
    files_skipped = 0

    logger.info(
        "Starting batch extraction from folder: %s (recursive=%s, extract_all_supported=%s)",
        folder_path,
        recursive,
        extract_all_supported,
    )

    # Iterate through all files matching the glob pattern
    for file_path_str in glob_module.glob(glob_path, recursive=recursive):
        file_path = Path(file_path_str)

        # Skip directories
        if file_path.is_dir():
            continue

        files_found += 1

        # Check if file should be processed
        if extract_all_supported:
            # Forced plain-text extraction also accepts unknown file extensions.
            if not force_plain_text and not is_supported_file(str(file_path)):
                files_skipped += 1
                logger.debug("Skipping unsupported file: %s", file_path)
                continue
        else:
            # Check against provided suffixes
            file_suffix = file_path.suffix.lower()
            if file_suffix not in normalized_suffixes:
                files_skipped += 1
                continue

        # Extract the file
        try:
            for result in read_file(
                file_path,
                max_file_size=max_file_size,
                ignore_images=ignore_images,
                force_plain_text=force_plain_text,
                include_attachments=include_attachments,
            ):
                files_extracted += 1
                yield result
        except ExtractionError as e:
            logger.warning("Failed to extract %s: %s", file_path, e)
            files_skipped += 1
            continue
        except (OSError, IOError) as e:
            logger.warning("IO error reading %s: %s", file_path, e)
            files_skipped += 1
            continue

    logger.info(
        "Batch extraction complete: %d files found, %d extracted, %d skipped",
        files_found,
        files_extracted,
        files_skipped,
    )


def read_file(
    path: str | Path,
    max_file_size: int = 100 * 1024 * 1024,  # 100MB default
    *,
    ignore_images: bool = False,
    force_plain_text: bool = False,
    include_attachments: bool = True,
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
        include_attachments: If False, skip extracting/storing email attachment
                      payloads for email file formats.

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
        include_attachments=include_attachments,
    )
    with open(path, "rb") as f:
        try:
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
    include_attachments: bool = True,
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
        include_attachments: If False, skip extracting/storing email attachment
            payloads for email file formats.

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
    resolved_file_type: str | None = None
    virtual_path = "<in-memory>"
    extension_error: ExtractionFileFormatNotSupportedError | None = None

    if force_plain_text:
        resolved_file_type = "txt"
        virtual_path = "in_memory.txt"
        extractor = get_extractor(
            virtual_path,
            ignore_images=ignore_images,
            force_plain_text=True,
            include_attachments=include_attachments,
        )
    else:
        if not normalized_extension and not normalized_mime_type:
            raise ValueError("Either mime_type or extension must be provided")

        if normalized_extension:
            resolved_file_type = _resolve_file_type(f"in_memory.{normalized_extension}")
            virtual_path = f"in_memory.{normalized_extension}"
            try:
                extractor = get_extractor(
                    virtual_path,
                    ignore_images=ignore_images,
                    include_attachments=include_attachments,
                )
            except ExtractionFileFormatNotSupportedError as exc:
                extension_error = exc
                if not normalized_mime_type:
                    raise

        if extractor is None and normalized_mime_type:
            resolved_file_type = MIME_TYPE_MAPPING.get(normalized_mime_type)
            if resolved_file_type is None:
                if extension_error is not None:
                    raise extension_error
                raise ExtractionFileFormatNotSupportedError(
                    f"File type not supported for MIME type '{normalized_mime_type}'"
                )
            virtual_path = f"in_memory.{resolved_file_type}"
            extractor = get_extractor(
                virtual_path,
                ignore_images=ignore_images,
                include_attachments=include_attachments,
            )

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
    "read_many",
    "is_supported_file",
    "get_extractor",
    # ZIP-bomb limit configuration
    "ZipBombLimits",
    "get_zip_bomb_limits",
    "set_zip_bomb_limits",
    "reset_zip_bomb_limits",
    # exceptions
    "ExtractionError",
    "ExtractionFileFormatNotSupportedError",
    "ExtractionLegacyMicrosoftParsingError",
    "ExtractionFileEncryptedError",
    "ExtractionZipBombError",
    "ExtractionFailedError",
    "ExtractionFileTooLargeError",
    "InvalidConfigurationError",
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
