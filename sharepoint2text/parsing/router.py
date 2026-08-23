import functools
import logging
import mimetypes
import os
from typing import Any, BinaryIO, Callable, Generator, cast

from sharepoint2text.parsing.exceptions import ExtractionFileFormatNotSupportedError
from sharepoint2text.parsing.mime_types import MIME_TYPE_MAPPING
from sharepoint2text.parsing.models import ExtractedDocument

logger = logging.getLogger(__name__)

_ATTACHMENT_AWARE_FILE_TYPES = frozenset({"msg", "eml", "mbox"})
_ARCHIVE_FILE_TYPES = frozenset({"zip", "tar", "tgz", "tbz2", "txz", "7z"})
_ATTACHMENT_OPTION_FILE_TYPES = _ATTACHMENT_AWARE_FILE_TYPES | _ARCHIVE_FILE_TYPES

# Mapping from file type identifiers to allowlisted extractor keys.
# Format: file_type -> extractor_key
_EXTRACTOR_REGISTRY: dict[str, str] = {
    # Modern MS Office
    "xlsx": "xlsx",
    "xlsb": "xlsx",
    "docx": "docx",
    "pptx": "pptx",
    # Macro-enabled variants (same OOXML structure)
    "xlsm": "xlsx",
    "docm": "docx",
    "pptm": "pptx",
    # Legacy MS Office
    "xls": "xls",
    "doc": "doc",
    "ppt": "ppt",
    "rtf": "rtf",
    # OpenDocument formats
    "odt": "odt",
    "odp": "odp",
    "ods": "ods",
    "odg": "odg",
    "odf": "odf",
    # Email formats
    "msg": "msg",
    "mbox": "mbox",
    "eml": "eml",
    # Structured delimited formats
    "csv": "csv",
    "tsv": "csv",
    # Plain text variants (all use the same extractor)
    "json": "plain_text",
    "txt": "plain_text",
    "md": "plain_text",
    # Configuration and data formats
    "yaml": "plain_text",
    "yml": "plain_text",
    "xml": "plain_text",
    "log": "plain_text",
    "ini": "plain_text",
    "cfg": "plain_text",
    "conf": "plain_text",
    "properties": "plain_text",
    # Other formats
    "pdf": "pdf",
    "html": "html",
    "epub": "epub",
    "mhtml": "mhtml",
    # Archive formats
    "zip": "archive",
    "tar": "archive",
    "tgz": "archive",
    "tbz2": "archive",
    "txz": "archive",
    "7z": "archive",
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

ExtractorFunction = Callable[
    [BinaryIO, str | None], Generator[ExtractedDocument, Any, None]
]


@functools.lru_cache(maxsize=None)
def _load_registered_extractor(extractor_key: str) -> Any:
    """Return an allowlisted extractor function for a registered key.

    Args:
        extractor_key: Internal extractor key from ``_EXTRACTOR_REGISTRY``.

    Returns:
        Extractor callable associated with the approved key.

    Raises:
        ExtractionFileFormatNotSupportedError: If the key is not allowlisted.
    """
    match extractor_key:
        case "xlsx":
            from sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor import (
                read_xlsx,
            )

            return read_xlsx
        case "docx":
            from sharepoint2text.parsing.extractors.ms_modern.docx_extractor import (
                read_docx,
            )

            return read_docx
        case "pptx":
            from sharepoint2text.parsing.extractors.ms_modern.pptx_extractor import (
                read_pptx,
            )

            return read_pptx
        case "xls":
            from sharepoint2text.parsing.extractors.ms_legacy.xls_extractor import (
                read_xls,
            )

            return read_xls
        case "doc":
            from sharepoint2text.parsing.extractors.ms_legacy.doc_extractor import (
                read_doc,
            )

            return read_doc
        case "ppt":
            from sharepoint2text.parsing.extractors.ms_legacy.ppt_extractor import (
                read_ppt,
            )

            return read_ppt
        case "rtf":
            from sharepoint2text.parsing.extractors.ms_legacy.rtf_extractor import (
                read_rtf,
            )

            return read_rtf
        case "odt":
            from sharepoint2text.parsing.extractors.open_office.odt_extractor import (
                read_odt,
            )

            return read_odt
        case "odp":
            from sharepoint2text.parsing.extractors.open_office.odp_extractor import (
                read_odp,
            )

            return read_odp
        case "ods":
            from sharepoint2text.parsing.extractors.open_office.ods_extractor import (
                read_ods,
            )

            return read_ods
        case "odg":
            from sharepoint2text.parsing.extractors.open_office.odg_extractor import (
                read_odg,
            )

            return read_odg
        case "odf":
            from sharepoint2text.parsing.extractors.open_office.odf_extractor import (
                read_odf,
            )

            return read_odf
        case "msg":
            from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
                read_msg_format_mail,
            )

            return read_msg_format_mail
        case "mbox":
            from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
                read_mbox_format_mail,
            )

            return read_mbox_format_mail
        case "eml":
            from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
                read_eml_format_mail,
            )

            return read_eml_format_mail
        case "csv":
            from sharepoint2text.parsing.extractors.csv_extractor import read_csv

            return read_csv
        case "plain_text":
            from sharepoint2text.parsing.extractors.plain_extractor import (
                read_plain_text,
            )

            return read_plain_text
        case "pdf":
            from sharepoint2text.parsing.extractors.pdf.pdf_extractor import read_pdf

            return read_pdf
        case "html":
            from sharepoint2text.parsing.extractors.html_extractor import read_html

            return read_html
        case "epub":
            from sharepoint2text.parsing.extractors.epub_extractor import read_epub

            return read_epub
        case "mhtml":
            from sharepoint2text.parsing.extractors.mhtml_extractor import read_mhtml

            return read_mhtml
        case "archive":
            from sharepoint2text.parsing.extractors.archive_extractor import (
                read_archive,
            )

            return read_archive
        case _:
            raise ExtractionFileFormatNotSupportedError(
                f"No allowlisted extractor for key: {extractor_key}"
            )


def _get_extractor_for_type(
    file_type: str,
    ignore_images: bool = False,
    include_attachments: bool = True,
) -> ExtractorFunction:
    """
    Return the extractor function for a file type using lazy import.

    Uses a registry-based lookup pattern to map file types to a fixed
    allowlist of extractor implementations. Imports are performed lazily
    through literal import statements to minimize startup time while
    preventing untrusted module loading.

    Args:
        file_type: File type identifier (e.g., "docx", "pdf", "xlsx").
        ignore_images: If True, skip image extraction for supported formats.
        include_attachments: If False, omit email attachments, including from
            emails contained in archives.

    Returns:
        Callable extractor function that accepts (binary stream, path) arguments.

    Raises:
        ExtractionFileFormatNotSupportedError: If no extractor exists for the file type.
    """
    if file_type not in _EXTRACTOR_REGISTRY:
        raise ExtractionFileFormatNotSupportedError(
            f"No extractor for file type: {file_type}"
        )

    extractor_key = _EXTRACTOR_REGISTRY[file_type]
    extractor = cast(ExtractorFunction, _load_registered_extractor(extractor_key))

    if ignore_images or (
        file_type in _ATTACHMENT_OPTION_FILE_TYPES and not include_attachments
    ):
        partial_kwargs: dict[str, Any] = {}
        if ignore_images:
            partial_kwargs["ignore_images"] = True
        if file_type in _ATTACHMENT_OPTION_FILE_TYPES and not include_attachments:
            partial_kwargs["include_attachments"] = False
        return cast(ExtractorFunction, functools.partial(extractor, **partial_kwargs))
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


def _is_archive_file_type(file_type: str | None) -> bool:
    """Return whether a resolved file type is handled by the archive extractor.

    Args:
        file_type: Internal extractor key returned by the router.

    Returns:
        ``True`` when the file type maps to archive extraction, else ``False``.
    """
    return file_type in _ARCHIVE_FILE_TYPES


def _resolve_file_type(
    path: str | os.PathLike[str], *, force_plain_text: bool = False
) -> str | None:
    """Resolve a path or filename to an internal file-type key.

    Detection order mirrors ``_get_extractor()`` so callers can make routing
    decisions without duplicating extension and MIME lookup logic.

    Args:
        path: File path or filename to analyze.
        force_plain_text: If True, return the plain-text file type regardless of
            path suffix or MIME detection.

    Returns:
        Internal file-type key, or ``None`` when the path cannot be resolved.
    """
    if force_plain_text:
        return "txt"

    path_str = os.fspath(path)
    path_lower = path_str.lower()

    file_type = _file_type_from_extension(path_lower)
    if file_type:
        return file_type

    mime_type, _ = mimetypes.guess_type(path_lower)
    if mime_type is not None and mime_type in MIME_TYPE_MAPPING:
        return MIME_TYPE_MAPPING[mime_type]

    return None


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


def _get_extractor(
    path: str | os.PathLike[str],
    ignore_images: bool = False,
    force_plain_text: bool = False,
    include_attachments: bool = True,
) -> ExtractorFunction:
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
        include_attachments: If False, skip extracting/storing supported email
            attachment payloads.

    Returns:
        Extractor function with signature ``(binary stream, path) -> Generator`` that
        yields one or more canonical ``ExtractedDocument`` results.

    Raises:
        ExtractionFileFormatNotSupportedError: If no extractor exists for the file type.
    """
    path_str = os.fspath(path)
    path_lower = path_str.lower()
    mime_type, _ = mimetypes.guess_type(path_lower)

    file_type = _resolve_file_type(path_str, force_plain_text=force_plain_text)
    if file_type is not None:
        if force_plain_text:
            logger.debug("Forcing plain text extraction for file: %s", path_str)
        if _file_type_from_extension(path_lower):
            logger.debug(
                "Detected file type: %s (extension) for file: %s", file_type, path_str
            )
        elif mime_type is not None:
            logger.debug(
                "Detected file type: %s (MIME: %s) for file: %s",
                file_type,
                mime_type,
                path_str,
            )
        extractor = _get_extractor_for_type(
            file_type,
            ignore_images=ignore_images,
            include_attachments=include_attachments,
        )
        return extractor

    extension = ""
    for compound_ext in _COMPOUND_EXTENSIONS:
        if path_lower.endswith(compound_ext):
            extension = compound_ext
            break
    if not extension:
        extension = os.path.splitext(path_lower)[1]

    mime_display = mime_type if mime_type is not None else "<unknown>"
    extension_display = extension if extension else "<none>"
    raise ExtractionFileFormatNotSupportedError(
        "File type not supported for path "
        f"'{path_str}' (extension: {extension_display}, MIME: {mime_display})"
    )
