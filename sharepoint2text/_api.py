"""
Implement the package's extraction entry points.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
import logging
import sys
import threading
from contextlib import contextmanager
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any, Generator, Iterator, TypeVar

from sharepoint2text.parsing._normalization import (
    _normalize_record,
)
from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileFormatNotSupportedError,
    ExtractionFileTooLargeError,
)
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    ZipBombLimits,
    _validate_zip_bomb_limits,
    _zip_bomb_limits_scope,
)
from sharepoint2text.parsing.mime_types import MIME_TYPE_MAPPING
from sharepoint2text.parsing.models import ExtractedDocument
from sharepoint2text.parsing.router import (
    _get_extractor,
    _resolve_file_type,
    is_supported_file,
)

logger = logging.getLogger(__name__)

_T = TypeVar("_T")

# Default pypdf decompression limit (bytes). pypdf uses 75 MB internally.
_PYPDF_DEFAULT_DECOMPRESSION_LIMIT = 75_000_000

# Sentinel value to effectively disable pypdf limits.  pypdf checks
# ``length > MAX_…`` so any stream smaller than this passes.  We cannot
# use ``0`` because that would make *every* stream exceed the limit.
# ``sys.maxsize`` is the platform's ``ssize_t`` max, accepted by zlib.
_PYPDF_NO_LIMIT = sys.maxsize

_PYPDF_LIMIT_ATTRIBUTES = (
    "ZLIB_MAX_OUTPUT_LENGTH",
    "LZW_MAX_OUTPUT_LENGTH",
    "RUN_LENGTH_MAX_OUTPUT_LENGTH",
    "JBIG2_MAX_OUTPUT_LENGTH",
    "MAX_DECLARED_STREAM_LENGTH",
    "MAX_ARRAY_BASED_STREAM_OUTPUT_LENGTH",
)

# pypdf exposes its limits as process-wide module globals. Serialize PDF
# extraction while an override is active so concurrent calls cannot observe
# another call's relaxed limits. An RLock keeps nested use in one thread safe.
_PYPDF_LIMIT_LOCK = threading.RLock()


def _pypdf_limit_target(max_file_size: int) -> int | None:
    """Return the pypdf override required for one extraction call.

    Args:
        max_file_size: User-supplied input size limit in bytes. Zero disables
            the limit.

    Returns:
        The process-wide pypdf limit to apply, or ``None`` to retain pypdf's
        existing defaults.
    """
    if max_file_size == 0:
        return _PYPDF_NO_LIMIT
    if max_file_size > _PYPDF_DEFAULT_DECOMPRESSION_LIMIT:
        return max_file_size
    return None


@contextmanager
def _pypdf_limits_scope(max_file_size: int) -> Iterator[None]:
    """Temporarily adjust and then restore pypdf decompression limits.

    Args:
        max_file_size: User-supplied input size limit in bytes. Zero disables
            pypdf's decompression limits for the duration of this scope.

    Yields:
        Control while the per-call pypdf limits are active.
    """
    import pypdf.filters as _filters  # noqa: PLC0415

    target = _pypdf_limit_target(max_file_size)
    with _PYPDF_LIMIT_LOCK:
        original_limits = {
            attribute: getattr(_filters, attribute)
            for attribute in _PYPDF_LIMIT_ATTRIBUTES
            if hasattr(_filters, attribute)
        }
        try:
            if target is not None:
                for attribute in original_limits:
                    setattr(_filters, attribute, target)
            yield
        finally:
            for attribute, original_limit in original_limits.items():
                setattr(_filters, attribute, original_limit)


try:
    __version__ = version("sharepoint-to-text")
except PackageNotFoundError:  # pragma: no cover
    __version__ = "unknown"


class InvalidConfigurationError(ValueError):
    """Raised when incompatible configuration options are provided."""


def _iterate_with_zip_bomb_limits(
    iterator: Iterator[_T],
    limits: ZipBombLimits | None,
) -> Generator[_T, None, None]:
    """Advance an iterator under isolated ZIP-bomb limits.

    Each iterator step leaves the scope before yielding control to the caller.
    This ensures a suspended generator cannot expose its relaxed limits to
    another extraction running in the same thread or task.

    Args:
        iterator: Underlying extractor iterator to advance.
        limits: Per-call ZIP-bomb limits, or ``None`` for library defaults.

    Yields:
        Values produced by the underlying iterator.

    Raises:
        TypeError: If ``limits`` is neither ``None`` nor ``ZipBombLimits``.
    """
    try:
        while True:
            with _zip_bomb_limits_scope(limits):
                try:
                    value = next(iterator)
                except StopIteration:
                    return
            yield value
    finally:
        close = getattr(iterator, "close", None)
        if callable(close):
            with _zip_bomb_limits_scope(limits):
                close()


def _iterate_with_pypdf_limits(
    iterator: Iterator[_T], max_file_size: int
) -> Generator[_T, None, None]:
    """Advance a PDF iterator under isolated pypdf limits.

    Args:
        iterator: Underlying PDF extractor iterator to advance.
        max_file_size: Per-call limit used to configure pypdf.

    Yields:
        Values produced by the PDF extractor after restoring global limits.
    """
    try:
        while True:
            with _pypdf_limits_scope(max_file_size):
                try:
                    value = next(iterator)
                except StopIteration:
                    return
            yield value
    finally:
        close = getattr(iterator, "close", None)
        if callable(close):
            with _pypdf_limits_scope(max_file_size):
                close()


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
    zip_bomb_limits: ZipBombLimits | None = None,
) -> Generator[ExtractedDocument, Any, None]:
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
        zip_bomb_limits: ZIP-bomb limits for each selected file. When ``None``,
            enforce the library defaults independently for every file.

    Yields:
        Normalized documents for each successfully extracted file.

    Raises:
        InvalidConfigurationError: If both suffixes are provided and
            extract_all_supported is True.
        ValueError: If neither suffixes nor extract_all_supported is specified.
        NotADirectoryError: If folder_path is not a directory.
        FileNotFoundError: If folder_path does not exist.
        TypeError: If ``zip_bomb_limits`` is not ``None`` or ``ZipBombLimits``.

    Example:
        >>> import sharepoint2text
        >>> # Extract only Word and PDF files
        >>> for result in sharepoint2text.read_many("/path/to/folder", [".docx", ".pdf"]):
        ...     print(f"{result.source.path}: {len(result.full_text)} chars")
        >>> # Extract all supported file types
        >>> for result in sharepoint2text.read_many("/path/to/folder", extract_all_supported=True):
        ...     print(result.full_text)
    """
    _validate_zip_bomb_limits(zip_bomb_limits)

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
    documents_extracted = 0
    files_skipped = 0

    logger.info(
        "Starting batch extraction from folder: %s (recursive=%s, extract_all_supported=%s)",
        folder_path,
        recursive,
        extract_all_supported,
    )

    # Iterate through all files matching the glob pattern
    for file_path_str in glob_module.iglob(glob_path, recursive=recursive):
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
            path_lower = str(file_path).lower()
            if not any(path_lower.endswith(suffix) for suffix in normalized_suffixes):
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
                zip_bomb_limits=zip_bomb_limits,
            ):
                documents_extracted += 1
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
        "Batch extraction complete: files_found=%d, documents_extracted=%d, "
        "files_skipped=%d",
        files_found,
        documents_extracted,
        files_skipped,
    )


def read_file(
    path: str | Path,
    max_file_size: int = 100 * 1024 * 1024,  # 100MB default
    *,
    ignore_images: bool = False,
    force_plain_text: bool = False,
    include_attachments: bool = True,
    zip_bomb_limits: ZipBombLimits | None = None,
) -> Generator[ExtractedDocument, Any, None]:
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
        zip_bomb_limits: ZIP-bomb limits for this extraction call. When
            ``None``, enforce the library defaults.

    Yields:
        Normalized documents for every supported source format.

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
        TypeError: If ``zip_bomb_limits`` is not ``None`` or ``ZipBombLimits``.

    Example:
        >>> import sharepoint2text
        >>> for result in sharepoint2text.read_file("document.docx"):
        ...     print(result.full_text)
        >>> # Skip image extraction for faster processing
        >>> for result in sharepoint2text.read_file("document.docx", ignore_images=True):
        ...     print(result.full_text)
    """
    from sharepoint2text.parsing.exceptions import (
        ExtractionError,
        ExtractionFailedError,
        ExtractionFileTooLargeError,
    )

    _validate_zip_bomb_limits(zip_bomb_limits)
    path = Path(path)

    # Check file size before reading
    if max_file_size > 0:
        file_size = path.stat().st_size
        if file_size > max_file_size:
            raise ExtractionFileTooLargeError(
                f"File size {file_size} bytes exceeds maximum allowed size of {max_file_size} bytes",
                max_size=max_file_size,
                actual_size=file_size,
            )

    logger.debug("Extracting file: %s", path)
    extractor = _get_extractor(
        str(path),
        ignore_images=ignore_images,
        force_plain_text=force_plain_text,
        include_attachments=include_attachments,
    )
    with open(path, "rb") as f:
        try:
            records = extractor(f, str(path))
            if _resolve_file_type(path, force_plain_text=force_plain_text) == "pdf":
                records = _iterate_with_pypdf_limits(records, max_file_size)
            documents_extracted = 0
            for result in _iterate_with_zip_bomb_limits(records, zip_bomb_limits):
                documents_extracted += 1
                yield _normalize_record(result)
            logger.debug(
                "Extracted file: %s (%d document%s)",
                path,
                documents_extracted,
                "" if documents_extracted == 1 else "s",
            )
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
    zip_bomb_limits: ZipBombLimits | None = None,
) -> Generator[ExtractedDocument, Any, None]:
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
        zip_bomb_limits: ZIP-bomb limits for this extraction call. When
            ``None``, enforce the library defaults.

    Yields:
        A normalized extraction document.

    Raises:
        ValueError: If both ``mime_type`` and ``extension`` are missing/empty,
            unless ``force_plain_text=True``.
        TypeError: If ``data`` has the wrong type or ``zip_bomb_limits`` is not
            ``None`` or ``ZipBombLimits``.
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
    _validate_zip_bomb_limits(zip_bomb_limits)
    if not isinstance(data, (bytes, io.BytesIO)):
        raise TypeError("data must be bytes or io.BytesIO")

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
        extractor = _get_extractor(
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
                extractor = _get_extractor(
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
            extractor = _get_extractor(
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

    logger.debug("Extracting in-memory file: %s", virtual_path)
    try:
        records = extractor(file_like, virtual_path)
        if resolved_file_type == "pdf":
            records = _iterate_with_pypdf_limits(records, max_file_size)
        documents_extracted = 0
        for result in _iterate_with_zip_bomb_limits(records, zip_bomb_limits):
            documents_extracted += 1
            yield _normalize_record(result)
        logger.debug(
            "Extracted in-memory file: %s (%d document%s)",
            virtual_path,
            documents_extracted,
            "" if documents_extracted == 1 else "s",
        )
    except ExtractionError:
        raise
    except (OSError, ValueError, TypeError, UnicodeDecodeError) as exc:
        raise ExtractionFailedError(
            f"Failed to extract in-memory data: {virtual_path}",
            cause=exc,
        ) from exc
