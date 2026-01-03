"""
Archive Content Extractor
==========================

Extracts content from archive files (ZIP and TAR formats) by recursively
processing supported files within the archive.

File Format Background
----------------------
Archive formats bundle multiple files into a single container:

ZIP (.zip):
    - Most common archive format, widely supported
    - Supports compression (deflate, bzip2, lzma)
    - Random access to individual files
    - Optional encryption (not supported by this extractor)

TAR (.tar, .tar.gz, .tgz, .tar.bz2, .tbz2, .tar.xz, .txz):
    - Unix tape archive format
    - Sequential access (optimized for streaming)
    - Often combined with compression (gzip, bzip2, xz)
    - No native encryption

Common archive sources include:
    - Document bundles and backups
    - Email attachment collections
    - Data exports and migrations
    - Software distribution packages

Dependencies
------------
Python Standard Library only:
    - zipfile: ZIP archive handling
    - tarfile: TAR archive handling
    - io: BytesIO for in-memory file handling

No external dependencies required.

Implementation Details
----------------------
The extractor:
    1. Detects archive type from file signature (magic bytes)
    2. Opens the archive using the appropriate library
    3. Iterates through archive members
    4. For each supported file type, extracts to memory and processes
    5. Yields extraction results with archive-aware metadata

Nested archives are NOT recursively extracted to prevent zip bombs
and excessive memory usage.

Known Limitations
-----------------
- Encrypted ZIP files are not supported
- Nested archives are not recursively processed
- Very large files within archives may cause memory issues
- Symbolic links in TAR archives are skipped
- Archive comments and extended attributes are not preserved

Encoding Handling
-----------------
- ZIP: Filenames use CP437 or UTF-8 (flag-dependent)
- TAR: Filenames use UTF-8 or system encoding
- Both handle encoding errors gracefully with replacement

Usage
-----
    >>> import io
    >>> from sharepoint2text.parsing.extractors.archive_extractor import read_archive
    >>>
    >>> with open("documents.zip", "rb") as f:
    ...     for content in read_archive(io.BytesIO(f.read()), path="documents.zip"):
    ...         print(f"File: {content.get_metadata().filename}")
    ...         print(f"Text: {content.get_full_text()[:100]}...")

See Also
--------
- zipfile: https://docs.python.org/3/library/zipfile.html
- tarfile: https://docs.python.org/3/library/tarfile.html

Maintenance Notes
-----------------
- Uses lazy import of router to avoid circular dependencies
- Archive type detection uses magic bytes, not extension
- Memory-efficient: files extracted one at a time
- Skips unsupported files silently (logged at debug level)
"""

import io
import logging
import os
import tarfile
import zipfile
from typing import Any, Generator

from sharepoint2text.parsing.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.parsing.extractors.data_types import ExtractionInterface

logger = logging.getLogger(__name__)

# Magic bytes for archive detection
ZIP_MAGIC = b"PK\x03\x04"
ZIP_EMPTY_MAGIC = b"PK\x05\x06"
GZIP_MAGIC = b"\x1f\x8b"
BZIP2_MAGIC = b"BZ"
XZ_MAGIC = b"\xfd7zXZ\x00"
TAR_MAGIC_OFFSET = 257
TAR_MAGIC = b"ustar"


def _detect_archive_type(file_like: io.BytesIO) -> str | None:
    """
    Detect archive type from file magic bytes.

    Args:
        file_like: BytesIO containing archive data.

    Returns:
        Archive type string ("zip", "tar", "tar.gz", "tar.bz2", "tar.xz")
        or None if not a recognized archive.
    """
    file_like.seek(0)
    header = file_like.read(512)
    file_like.seek(0)

    if not header:
        return None

    # Check for ZIP
    if header[:4] == ZIP_MAGIC or header[:4] == ZIP_EMPTY_MAGIC:
        return "zip"

    # Check for compressed TAR variants
    if header[:2] == GZIP_MAGIC:
        return "tar.gz"
    if header[:2] == BZIP2_MAGIC:
        return "tar.bz2"
    if header[:6] == XZ_MAGIC:
        return "tar.xz"

    # Check for uncompressed TAR (magic at offset 257)
    if len(header) >= TAR_MAGIC_OFFSET + 5:
        if header[TAR_MAGIC_OFFSET : TAR_MAGIC_OFFSET + 5] == TAR_MAGIC:
            return "tar"

    return None


def _is_supported_file(filename: str) -> bool:
    """Check if a filename corresponds to a supported extraction format."""
    # Lazy import to avoid circular dependency
    from sharepoint2text.parsing.router import is_supported_file

    return is_supported_file(filename)


def _get_file_extractor(filename: str):
    """Get the extractor function for a filename."""
    # Lazy import to avoid circular dependency
    from sharepoint2text.parsing.router import get_extractor

    return get_extractor(filename)


def _extract_from_zip(
    file_like: io.BytesIO, archive_path: str | None
) -> Generator[ExtractionInterface, Any, None]:
    """
    Extract supported files from a ZIP archive.

    Args:
        file_like: BytesIO containing the ZIP archive.
        archive_path: Optional path to the archive file for metadata.

    Yields:
        ExtractionInterface objects for each supported file in the archive.
    """
    try:
        with zipfile.ZipFile(file_like, "r") as zf:
            for info in zf.infolist():
                # Skip directories
                if info.is_dir():
                    continue

                filename = info.filename
                basename = os.path.basename(filename)

                # Skip hidden files, macOS resource forks, and unsupported types
                if basename.startswith(".") or filename.startswith("__MACOSX/"):
                    logger.debug("Skipping hidden/system file: %s", filename)
                    continue

                if not _is_supported_file(basename):
                    logger.debug("Skipping unsupported file: %s", filename)
                    continue

                # Skip nested archives to prevent zip bombs
                if basename.lower().endswith(
                    (
                        ".zip",
                        ".tar",
                        ".tar.gz",
                        ".tgz",
                        ".tar.bz2",
                        ".tbz2",
                        ".tar.xz",
                        ".txz",
                    )
                ):
                    logger.debug("Skipping nested archive: %s", filename)
                    continue

                try:
                    logger.debug("Extracting from ZIP: %s", filename)
                    file_data = zf.read(info)
                    file_bytes = io.BytesIO(file_data)

                    # Build path that includes archive context
                    if archive_path:
                        full_path = f"{archive_path}!/{filename}"
                    else:
                        full_path = filename

                    extractor = _get_file_extractor(basename)
                    for content in extractor(file_bytes, path=full_path):
                        yield content

                except Exception as e:
                    logger.warning("Failed to extract %s from archive: %s", filename, e)
                    continue

    except zipfile.BadZipFile as e:
        raise ExtractionFailedError(f"Invalid ZIP archive: {e}", cause=e) from e


def _extract_from_tar(
    file_like: io.BytesIO, archive_path: str | None, mode: str = "r:*"
) -> Generator[ExtractionInterface, Any, None]:
    """
    Extract supported files from a TAR archive.

    Args:
        file_like: BytesIO containing the TAR archive.
        archive_path: Optional path to the archive file for metadata.
        mode: TAR open mode (r:* for auto-detect compression).

    Yields:
        ExtractionInterface objects for each supported file in the archive.
    """
    try:
        with tarfile.open(fileobj=file_like, mode=mode) as tf:
            for member in tf.getmembers():
                # Skip directories and non-regular files
                if not member.isreg():
                    continue

                filename = member.name
                basename = os.path.basename(filename)

                # Skip hidden files, macOS resource forks, and unsupported types
                if basename.startswith(".") or filename.startswith("__MACOSX/"):
                    logger.debug("Skipping hidden/system file: %s", filename)
                    continue

                if not _is_supported_file(basename):
                    logger.debug("Skipping unsupported file: %s", filename)
                    continue

                # Skip nested archives to prevent zip bombs
                if basename.lower().endswith(
                    (
                        ".zip",
                        ".tar",
                        ".tar.gz",
                        ".tgz",
                        ".tar.bz2",
                        ".tbz2",
                        ".tar.xz",
                        ".txz",
                    )
                ):
                    logger.debug("Skipping nested archive: %s", filename)
                    continue

                try:
                    logger.debug("Extracting from TAR: %s", filename)
                    extracted = tf.extractfile(member)
                    if extracted is None:
                        continue

                    file_data = extracted.read()
                    file_bytes = io.BytesIO(file_data)

                    # Build path that includes archive context
                    if archive_path:
                        full_path = f"{archive_path}!/{filename}"
                    else:
                        full_path = filename

                    extractor = _get_file_extractor(basename)
                    for content in extractor(file_bytes, path=full_path):
                        yield content

                except Exception as e:
                    logger.warning("Failed to extract %s from archive: %s", filename, e)
                    continue

    except tarfile.TarError as e:
        raise ExtractionFailedError(f"Invalid TAR archive: {e}", cause=e) from e


def read_archive(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[ExtractionInterface, Any, None]:
    """
    Extract content from supported files within a ZIP or TAR archive.

    Primary entry point for archive extraction. Automatically detects the
    archive format and iterates through contents, yielding extraction results
    for each supported file type.

    This function uses a generator pattern to yield multiple extraction
    results, one for each supported file in the archive.

    Args:
        file_like: BytesIO object containing the complete archive data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source archive. If provided,
            the path is included in extracted file metadata using the
            format "archive_path!/internal_path".

    Yields:
        ExtractionInterface: Extraction results for each supported file
            in the archive. The specific type depends on the file format
            (e.g., DocxContent for .docx files, PdfContent for .pdf).

    Raises:
        ExtractionFailedError: If the archive cannot be read or is corrupt.

    Note:
        - Nested archives are skipped to prevent zip bombs
        - Unsupported file types are silently skipped (logged at debug level)
        - Hidden files (starting with '.') are skipped
        - Encrypted archives are not supported

    Example:
        >>> import io
        >>> with open("documents.zip", "rb") as f:
        ...     archive_data = io.BytesIO(f.read())
        ...     for content in read_archive(archive_data, path="documents.zip"):
        ...         meta = content.get_metadata()
        ...         print(f"File: {meta.filename}")
        ...         print(f"Path: {meta.file_path}")
        ...         text = content.get_full_text()
        ...         print(f"Text length: {len(text)}")
    """
    try:
        file_like.seek(0)
        archive_type = _detect_archive_type(file_like)

        if archive_type is None:
            raise ExtractionFailedError("Unable to detect archive type")

        logger.debug("Detected archive type: %s", archive_type)

        if archive_type == "zip":
            yield from _extract_from_zip(file_like, path)
        elif archive_type in ("tar", "tar.gz", "tar.bz2", "tar.xz"):
            yield from _extract_from_tar(file_like, path)
        else:
            raise ExtractionFailedError(f"Unsupported archive type: {archive_type}")

    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError(
            "Failed to extract archive file", cause=exc
        ) from exc


# Convenience aliases for specific formats
read_zip = read_archive
read_tar = read_archive
