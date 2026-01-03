import io
import logging
import mimetypes
import os
from typing import Any, Callable, Generator

from sharepoint2text.exceptions import ExtractionFileFormatNotSupportedError
from sharepoint2text.extractors.data_types import ExtractionInterface
from sharepoint2text.mime_types import MIME_TYPE_MAPPING

logger = logging.getLogger(__name__)

# Mapping from file type identifiers to their extractor module paths and function names
# Format: file_type -> (module_path, function_name)
_EXTRACTOR_REGISTRY: dict[str, tuple[str, str]] = {
    # Modern MS Office
    "xlsx": ("sharepoint2text.extractors.ms_modern.xlsx_extractor", "read_xlsx"),
    "docx": ("sharepoint2text.extractors.ms_modern.docx_extractor", "read_docx"),
    "pptx": ("sharepoint2text.extractors.ms_modern.pptx_extractor", "read_pptx"),
    # Macro-enabled variants (same OOXML structure)
    "xlsm": ("sharepoint2text.extractors.ms_modern.xlsx_extractor", "read_xlsx"),
    "docm": ("sharepoint2text.extractors.ms_modern.docx_extractor", "read_docx"),
    "pptm": ("sharepoint2text.extractors.ms_modern.pptx_extractor", "read_pptx"),
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
    "pdf": ("sharepoint2text.extractors.pdf.pdf_extractor", "read_pdf"),
    "html": ("sharepoint2text.extractors.html_extractor", "read_html"),
    "epub": ("sharepoint2text.extractors.epub_extractor", "read_epub"),
    "mhtml": ("sharepoint2text.extractors.mhtml_extractor", "read_mhtml"),
}

_EXTENSION_ALIASES: dict[str, str] = {
    "htm": "html",
    "mht": "mhtml",
}

_SUPPORTED_EXTENSIONS: frozenset[str] = frozenset(
    {f".{ext}" for ext in _EXTRACTOR_REGISTRY.keys()}
    | {f".{ext}" for ext in _EXTENSION_ALIASES.keys()}
)


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


def _file_type_from_extension(path_lower: str) -> str | None:
    extension = os.path.splitext(path_lower)[1]
    if not extension:
        return None
    ext = extension[1:]
    if not ext:
        return None
    ext = _EXTENSION_ALIASES.get(ext, ext)
    return ext if ext in _EXTRACTOR_REGISTRY else None


def is_supported_file(path: str) -> bool:
    """
    Check if a file path corresponds to a supported file format.

    Detection is extension-first (OS-independent), then falls back to MIME.

    Args:
        path: File path or filename to check.

    Returns:
        True if the file format is supported, False otherwise.
    """
    path_lower = path.lower()

    extension = os.path.splitext(path_lower)[1]
    if extension in _SUPPORTED_EXTENSIONS:
        return True

    mime_type, _ = mimetypes.guess_type(path_lower)
    return bool(mime_type and mime_type in MIME_TYPE_MAPPING)


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

    # Primary detection: file extension (platform-independent)
    file_type = _file_type_from_extension(path_lower)
    if file_type:
        logger.debug("Detected file type: %s (extension) for file: %s", file_type, path)
        logger.info("Using extractor for file type: %s", file_type)
        return _get_extractor(file_type)

    # Secondary detection: MIME type lookup (may vary by OS configuration)
    if mime_type is not None and mime_type in MIME_TYPE_MAPPING:
        file_type = MIME_TYPE_MAPPING[mime_type]
        logger.debug(
            "Detected file type: %s (MIME: %s) for file: %s", file_type, mime_type, path
        )
        logger.info("Using extractor for file type: %s", file_type)
        return _get_extractor(file_type)

    logger.warning("Unsupported file type: %s (MIME: %s)", path, mime_type)
    raise ExtractionFileFormatNotSupportedError(f"File type not supported: {mime_type}")
