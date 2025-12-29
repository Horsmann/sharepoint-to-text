import io
import logging
import mimetypes
import os
from typing import Any, Callable, Generator

from sharepoint2text.exceptions import ExtractionFileFormatNotSupportedError
from sharepoint2text.extractors.data_types import ExtractionInterface

logger = logging.getLogger(__name__)

# File extensions that require fallback detection (not reliably detected via MIME type)
FALLBACK_EXTENSIONS = frozenset({".msg", ".eml", ".mbox", ".md"})

# Mapping from file type identifiers to their extractor module paths and function names
# Format: file_type -> (module_path, function_name)
_EXTRACTOR_REGISTRY: dict[str, tuple[str, str]] = {
    # Modern MS Office
    "xlsx": ("sharepoint2text.extractors.ms_modern.xlsx_extractor", "read_xlsx"),
    "docx": ("sharepoint2text.extractors.ms_modern.docx_extractor", "read_docx"),
    "pptx": ("sharepoint2text.extractors.ms_modern.pptx_extractor", "read_pptx"),
    # Legacy MS Office
    "xls": ("sharepoint2text.extractors.ms_legacy.xls_extractor", "read_xls"),
    "doc": ("sharepoint2text.extractors.ms_legacy.doc_extractor", "read_doc"),
    "ppt": ("sharepoint2text.extractors.ms_legacy.ppt_extractor", "read_ppt"),
    "rtf": ("sharepoint2text.extractors.ms_legacy.rtf_extractor", "read_rtf"),
    # OpenDocument formats
    "odt": ("sharepoint2text.extractors.open_office.odt_extractor", "read_odt"),
    "odp": ("sharepoint2text.extractors.open_office.odp_extractor", "read_odp"),
    "ods": ("sharepoint2text.extractors.open_office.ods_extractor", "read_ods"),
    # Email formats
    "msg": (
        "sharepoint2text.extractors.mail.msg_email_extractor",
        "read_msg_format_mail",
    ),
    "mbox": (
        "sharepoint2text.extractors.mail.mbox_email_extractor",
        "read_mbox_format_mail",
    ),
    "eml": (
        "sharepoint2text.extractors.mail.eml_email_extractor",
        "read_eml_format_mail",
    ),
    # Plain text variants (all use the same extractor)
    "csv": ("sharepoint2text.extractors.plain_extractor", "read_plain_text"),
    "json": ("sharepoint2text.extractors.plain_extractor", "read_plain_text"),
    "txt": ("sharepoint2text.extractors.plain_extractor", "read_plain_text"),
    "tsv": ("sharepoint2text.extractors.plain_extractor", "read_plain_text"),
    "md": ("sharepoint2text.extractors.plain_extractor", "read_plain_text"),
    # Other formats
    "pdf": ("sharepoint2text.extractors.pdf_extractor", "read_pdf"),
    "html": ("sharepoint2text.extractors.html_extractor", "read_html"),
}

MIME_TYPE_MAPPING = {
    # Legacy MS Office
    "application/vnd.ms-powerpoint": "ppt",
    "application/vnd.ms-excel": "xls",
    "application/msword": "doc",
    "application/rtf": "rtf",
    "text/rtf": "rtf",
    # Modern MS Office
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    # OpenDocument formats
    "application/vnd.oasis.opendocument.text": "odt",
    "application/vnd.oasis.opendocument.presentation": "odp",
    "application/vnd.oasis.opendocument.spreadsheet": "ods",
    # Plain text variants
    "text/csv": "csv",
    "application/csv": "csv",
    "application/json": "json",
    "text/json": "json",
    "text/plain": "txt",
    "text/markdown": "md",
    "text/tab-separated-values": "tsv",
    "application/tab-separated-values": "tsv",
    # Email formats
    "application/vnd.ms-outlook": "msg",
    "message/rfc822": "eml",
    "application/mbox": "mbox",
    # Other formats
    "text/html": "html",
    "application/xhtml+xml": "html",
    "application/pdf": "pdf",
}


def _get_extractor(
    file_type: str,
) -> Callable[[io.BytesIO, str | None], Generator[ExtractionInterface, Any, None]]:
    """
    Return the extractor function for a file type using lazy import.

    Uses a registry-based lookup pattern to map file types to their
    corresponding extractor modules and functions. Imports are performed
    lazily to minimize startup time and memory usage.

    Args:
        file_type: File type identifier (e.g., "docx", "pdf", "xlsx").

    Returns:
        Callable extractor function that accepts (BytesIO, path) arguments.

    Raises:
        ExtractionFileFormatNotSupportedError: If no extractor exists for the file type.
    """
    if file_type not in _EXTRACTOR_REGISTRY:
        raise ExtractionFileFormatNotSupportedError(
            f"No extractor for file type: {file_type}"
        )

    module_path, function_name = _EXTRACTOR_REGISTRY[file_type]

    # Lazy import of the extractor module
    import importlib

    module = importlib.import_module(module_path)
    return getattr(module, function_name)


def is_supported_file(path: str) -> bool:
    """
    Check if a file path corresponds to a supported file format.

    Detection uses MIME type guessing based on extension, with fallback
    to explicit extension checking for formats not reliably detected.

    Args:
        path: File path or filename to check.

    Returns:
        True if the file format is supported, False otherwise.
    """
    path_lower = path.lower()
    mime_type, _ = mimetypes.guess_type(path_lower)

    if mime_type in MIME_TYPE_MAPPING:
        return True

    # Fallback: check extensions not reliably detected via MIME type
    extension = os.path.splitext(path_lower)[1]
    return extension in FALLBACK_EXTENSIONS


def get_extractor(
    path: str,
) -> Callable[[io.BytesIO, str | None], Generator[ExtractionInterface, Any, None]]:
    """
    Analyze a file path and return the appropriate extractor function.

    The file does not need to exist; the path or filename alone is sufficient
    to determine the correct extractor based on extension and MIME type.

    Args:
        path: File path or filename to analyze.

    Returns:
        Extractor function that accepts (BytesIO, path) arguments.

    Raises:
        ExtractionFileFormatNotSupportedError: If no extractor exists for the file type.
    """
    path_lower = path.lower()
    mime_type, _ = mimetypes.guess_type(path_lower)
    logger.debug("Guessed MIME type: [%s]", mime_type)

    # Primary detection: MIME type lookup
    if mime_type is not None and mime_type in MIME_TYPE_MAPPING:
        file_type = MIME_TYPE_MAPPING[mime_type]
        logger.debug(
            "Detected file type: %s (MIME: %s) for file: %s", file_type, mime_type, path
        )
        return _get_extractor(file_type)

    # Fallback detection: explicit extension checking
    extension = os.path.splitext(path_lower)[1]
    if extension in FALLBACK_EXTENSIONS:
        file_type = extension[1:]  # Remove leading dot
        logger.debug("Detected file type: %s for file: %s", file_type, path)
        return _get_extractor(file_type)

    logger.debug("File [%s] with MIME type [%s] is not supported", path, mime_type)
    raise ExtractionFileFormatNotSupportedError(f"File type not supported: {mime_type}")
