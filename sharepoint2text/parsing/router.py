import functools
import logging
import mimetypes
import os
from typing import Any, BinaryIO, Callable, Generator

from sharepoint2text.parsing.exceptions import ExtractionFileFormatNotSupportedError
from sharepoint2text.parsing.extractors.data_types import ExtractionInterface
from sharepoint2text.parsing.mime_types import MIME_TYPE_MAPPING

logger = logging.getLogger(__name__)

# Mapping from file type identifiers to their extractor module paths and function names
# Format: file_type -> (module_path, function_name)
_EXTRACTOR_REGISTRY: dict[str, tuple[str, str]] = {
    # Modern MS Office
    "xlsx": (
        "sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor",
        "read_xlsx",
    ),
    "xlsb": (
        "sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor",
        "read_xlsx",
    ),
    "docx": (
        "sharepoint2text.parsing.extractors.ms_modern.docx_extractor",
        "read_docx",
    ),
    "pptx": (
        "sharepoint2text.parsing.extractors.ms_modern.pptx_extractor",
        "read_pptx",
    ),
    # Macro-enabled variants (same OOXML structure)
    "xlsm": (
        "sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor",
        "read_xlsx",
    ),
    "docm": (
        "sharepoint2text.parsing.extractors.ms_modern.docx_extractor",
        "read_docx",
    ),
    "pptm": (
        "sharepoint2text.parsing.extractors.ms_modern.pptx_extractor",
        "read_pptx",
    ),
    # Legacy MS Office
    "xls": ("sharepoint2text.parsing.extractors.ms_legacy.xls_extractor", "read_xls"),
    "doc": ("sharepoint2text.parsing.extractors.ms_legacy.doc_extractor", "read_doc"),
    "ppt": ("sharepoint2text.parsing.extractors.ms_legacy.ppt_extractor", "read_ppt"),
    "rtf": ("sharepoint2text.parsing.extractors.ms_legacy.rtf_extractor", "read_rtf"),
    # OpenDocument formats
    "odt": ("sharepoint2text.parsing.extractors.open_office.odt_extractor", "read_odt"),
    "odp": ("sharepoint2text.parsing.extractors.open_office.odp_extractor", "read_odp"),
    "ods": ("sharepoint2text.parsing.extractors.open_office.ods_extractor", "read_ods"),
    "odg": ("sharepoint2text.parsing.extractors.open_office.odg_extractor", "read_odg"),
    "odf": ("sharepoint2text.parsing.extractors.open_office.odf_extractor", "read_odf"),
    # Email formats
    "msg": (
        "sharepoint2text.parsing.extractors.mail.msg_email_extractor",
        "read_msg_format_mail",
    ),
    "mbox": (
        "sharepoint2text.parsing.extractors.mail.mbox_email_extractor",
        "read_mbox_format_mail",
    ),
    "eml": (
        "sharepoint2text.parsing.extractors.mail.eml_email_extractor",
        "read_eml_format_mail",
    ),
    # Plain text variants (all use the same extractor)
    "csv": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "json": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "txt": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "tsv": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "md": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    # Configuration and data formats
    "yaml": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "yml": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "xml": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "log": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "ini": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "cfg": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "conf": ("sharepoint2text.parsing.extractors.plain_extractor", "read_plain_text"),
    "properties": (
        "sharepoint2text.parsing.extractors.plain_extractor",
        "read_plain_text",
    ),
    # Other formats
    "pdf": ("sharepoint2text.parsing.extractors.pdf.pdf_extractor", "read_pdf"),
    "html": ("sharepoint2text.parsing.extractors.html_extractor", "read_html"),
    "epub": ("sharepoint2text.parsing.extractors.epub_extractor", "read_epub"),
    "mhtml": ("sharepoint2text.parsing.extractors.mhtml_extractor", "read_mhtml"),
    # Archive formats
    "zip": ("sharepoint2text.parsing.extractors.archive_extractor", "read_archive"),
    "tar": ("sharepoint2text.parsing.extractors.archive_extractor", "read_archive"),
    "tgz": ("sharepoint2text.parsing.extractors.archive_extractor", "read_archive"),
    "tbz2": ("sharepoint2text.parsing.extractors.archive_extractor", "read_archive"),
    "txz": ("sharepoint2text.parsing.extractors.archive_extractor", "read_archive"),
    "7z": ("sharepoint2text.parsing.extractors.archive_extractor", "read_archive"),
}

_EXTENSION_ALIASES: dict[str, str] = {
    "htm": "html",
    "mht": "mhtml",
    # Office templates / slide shows (map to existing extractors)
    "dot": "doc",
    "dotx": "docx",
    "dotm": "docm",
    "xlt": "xls",
    "xltx": "xlsx",
    "xltm": "xlsm",
    "pot": "ppt",
    "potx": "pptx",
    "potm": "pptm",
    "pps": "ppt",
    "ppsx": "pptx",
    "ppsm": "pptm",
    # OpenDocument templates (map to existing extractors)
    "ott": "odt",
    "ots": "ods",
    "otp": "odp",
    "gz": "tgz",  # .gz alone treated as gzip-compressed tar
    "bz2": "tbz2",  # .bz2 alone treated as bzip2-compressed tar
    "xz": "txz",  # .xz alone treated as xz-compressed tar
}

# Compound extensions that need special handling (checked before single extension)
_COMPOUND_EXTENSIONS: dict[str, str] = {
    ".tar.gz": "tgz",
    ".tar.bz2": "tbz2",
    ".tar.xz": "txz",
}

_SUPPORTED_EXTENSIONS: frozenset[str] = frozenset(
    {f".{ext}" for ext in _EXTRACTOR_REGISTRY.keys()}
    | {f".{ext}" for ext in _EXTENSION_ALIASES.keys()}
    | set(_COMPOUND_EXTENSIONS.keys())
)


def _get_extractor(
    file_type: str,
    ignore_images: bool = False,
) -> Callable[[BinaryIO, str | None], Generator[ExtractionInterface, Any, None]]:
    """
    Return the extractor function for a file type using lazy import.

    Uses a registry-based lookup pattern to map file types to their
    corresponding extractor modules and functions. Imports are performed
    lazily to minimize startup time and memory usage.

    Args:
        file_type: File type identifier (e.g., "docx", "pdf", "xlsx").
        ignore_images: If True, skip image extraction for supported formats.

    Returns:
        Callable extractor function that accepts (binary stream, path) arguments.

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
    extractor = getattr(module, function_name)

    # If ignore_images is True, wrap the extractor with the flag
    if ignore_images:
        return functools.partial(extractor, ignore_images=True)
    return extractor


def _file_type_from_extension(path_lower: str) -> str | None:
    """Resolve a normalized path to an internal file-type key via extension.

    Checks compound extensions first (for example ``.tar.gz``), then single
    extensions with alias mapping (for example ``.htm`` -> ``html``).

    Args:
        path_lower: Lower-cased path or filename.

    Returns:
        Internal extractor key (for example ``"docx"``, ``"tgz"``), or ``None``
        when the extension is missing or unsupported.
    """
    # Check compound extensions first (e.g., .tar.gz)
    for compound_ext, file_type in _COMPOUND_EXTENSIONS.items():
        if path_lower.endswith(compound_ext):
            return file_type

    extension = os.path.splitext(path_lower)[1]
    if not extension:
        return None
    ext = extension[1:]
    if not ext:
        return None
    ext = _EXTENSION_ALIASES.get(ext, ext)
    return ext if ext in _EXTRACTOR_REGISTRY else None


def is_supported_file(path: str | os.PathLike[str]) -> bool:
    """
    Check if a path/filename appears to be supported by the extractor registry.

    Detection is extension-first (OS-independent), then falls back to MIME.
    This function does not open or inspect file contents.

    Args:
        path: File path or filename to check.

    Returns:
        ``True`` if routing would likely succeed, else ``False``.
    """
    path_lower = os.fspath(path).lower()

    # Check compound extensions first (e.g., .tar.gz)
    for compound_ext in _COMPOUND_EXTENSIONS:
        if path_lower.endswith(compound_ext):
            return True

    extension = os.path.splitext(path_lower)[1]
    if extension in _SUPPORTED_EXTENSIONS:
        return True

    mime_type, _ = mimetypes.guess_type(path_lower)
    return bool(mime_type and mime_type in MIME_TYPE_MAPPING)


def get_extractor(
    path: str | os.PathLike[str],
    ignore_images: bool = False,
    force_plain_text: bool = False,
) -> Callable[[BinaryIO, str | None], Generator[ExtractionInterface, Any, None]]:
    """
    Analyze a path/filename and return the appropriate extractor callable.

    The file does not need to exist; routing is based on path text only.
    Detection order is:
    1) extension / alias lookup, 2) MIME mapping fallback.

    Args:
        path: File path or filename to analyze.
        ignore_images: If True, skip image extraction for supported formats.
        force_plain_text: If True, always route to the plain text extractor,
            even when extension/MIME detection does not recognize the file.

    Returns:
        Extractor function with signature ``(binary stream, path) -> Generator`` that
        yields one or more ``ExtractionInterface`` results.

    Raises:
        ExtractionFileFormatNotSupportedError: If no extractor exists for the file type.
    """
    path_str = os.fspath(path)
    path_lower = path_str.lower()
    mime_type, _ = mimetypes.guess_type(path_lower)
    logger.debug("Guessed MIME type: [%s]", mime_type)

    if force_plain_text:
        logger.info("Force plain text extraction for file: %s", path_str)
        return _get_extractor("txt", ignore_images=ignore_images)

    # Primary detection: file extension (platform-independent)
    file_type = _file_type_from_extension(path_lower)
    if file_type:
        logger.debug(
            "Detected file type: %s (extension) for file: %s", file_type, path_str
        )
        logger.info("Using extractor for file type: %s", file_type)
        return _get_extractor(file_type, ignore_images=ignore_images)

    # Secondary detection: MIME type lookup (may vary by OS configuration)
    if mime_type is not None and mime_type in MIME_TYPE_MAPPING:
        file_type = MIME_TYPE_MAPPING[mime_type]
        logger.debug(
            "Detected file type: %s (MIME: %s) for file: %s",
            file_type,
            mime_type,
            path_str,
        )
        logger.info("Using extractor for file type: %s", file_type)
        return _get_extractor(file_type, ignore_images=ignore_images)

    extension = ""
    for compound_ext in _COMPOUND_EXTENSIONS:
        if path_lower.endswith(compound_ext):
            extension = compound_ext
            break
    if not extension:
        extension = os.path.splitext(path_lower)[1]

    mime_display = mime_type if mime_type is not None else "<unknown>"
    extension_display = extension if extension else "<none>"
    logger.warning("Unsupported file type: %s (MIME: %s)", path_str, mime_type)
    raise ExtractionFileFormatNotSupportedError(
        "File type not supported for path "
        f"'{path_str}' (extension: {extension_display}, MIME: {mime_display})"
    )
