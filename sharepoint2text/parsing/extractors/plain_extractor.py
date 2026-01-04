"""
Plain Text Content Extractor
=============================

Extracts content from plain text files with automatic encoding detection
and normalization.

File Format Background
----------------------
Plain text files contain unformatted text without embedded structure
or binary content. This extractor handles common text file extensions:
    - .txt: Generic text files
    - .log: Log files
    - .csv: Comma-separated values (treated as plain text)
    - .md, .markdown: Markdown files (content only, no rendering)
    - .json, .xml, .yaml: Data files (treated as plain text)
    - Source code files (.py, .js, .java, etc.)

Encoding Handling
-----------------
The extractor uses charset_normalizer for automatic encoding detection:
    - Detects encoding from file content (UTF-8, Latin-1, Windows-1252, etc.)
    - Falls back to UTF-8 with error replacement if detection fails
    - Handles legacy files with non-UTF-8 encodings common in enterprise environments
    - The detected encoding is available in metadata.detected_encoding

Dependencies
------------
    - charset_normalizer: Encoding detection library

Extracted Content
-----------------
The extractor produces:
    - content: Full text content as a single string (decoded using detected encoding)
    - metadata: FileMetadataInterface with file information including detected_encoding

The content is returned as-is without modification, preserving:
    - Line endings (\\n, \\r\\n, \\r)
    - Whitespace and indentation
    - Empty lines

Known Limitations
-----------------
- Very short files may have unreliable encoding detection
- Binary files may produce garbled output
- Very large files are loaded entirely into memory
- No line ending normalization

Usage
-----
    >>> import io
    >>> from sharepoint2text.parsing.extractors.plain_extractor import read_plain_text
    >>>
    >>> with open("notes.txt", "rb") as f:
    ...     for doc in read_plain_text(io.BytesIO(f.read()), path="notes.txt"):
    ...         print(f"Encoding: {doc.metadata.detected_encoding}")
    ...         print(f"Characters: {len(doc.content)}")
    ...         print(doc.content[:200])

See Also
--------
- charset_normalizer: https://github.com/Ousret/charset_normalizer
- Python codecs module: https://docs.python.org/3/library/codecs.html
- Unicode HOWTO: https://docs.python.org/3/howto/unicode.html

Maintenance Notes
-----------------
- Uses charset_normalizer for encoding detection
- Falls back to UTF-8 with errors="replace" if detection fails
- FileMetadataInterface provides basic file info population
- Generator pattern for API consistency with other extractors
- Content returned unmodified (no stripping or normalization)
"""

import io
import logging
from typing import Any, Generator

from charset_normalizer import from_bytes

from sharepoint2text.parsing.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.parsing.extractors.data_types import (
    FileMetadataInterface,
    PlainTextContent,
)

logger = logging.getLogger(__name__)


def _detect_and_decode(content: bytes) -> tuple[str, str]:
    """
    Detect encoding and decode bytes to string.

    Uses charset_normalizer to detect the most likely encoding,
    then decodes the content. Falls back to UTF-8 with replacement
    if detection fails or returns no results.

    Args:
        content: Raw bytes to decode.

    Returns:
        Tuple of (decoded_text, detected_encoding).
        If detection fails, encoding will be "utf-8" (fallback).
    """
    if not content:
        return "", "utf-8"

    # Use charset_normalizer to detect encoding
    results = from_bytes(content)
    best_match = results.best()

    if best_match is not None:
        encoding = best_match.encoding
        logger.debug(
            "Detected encoding: %s (confidence: %.2f)",
            encoding,
            best_match.encoding_aliases,
        )
        try:
            # Use the detected encoding
            text = str(best_match)
            return text, encoding
        except Exception as e:
            logger.warning(
                "Failed to decode with detected encoding %s: %s", encoding, e
            )

    # Fallback to UTF-8 with replacement characters
    logger.debug("Encoding detection failed, falling back to UTF-8")
    return content.decode("utf-8", errors="replace"), "utf-8"


def read_plain_text(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PlainTextContent, Any, None]:
    """
    Extract content from a plain text file with automatic encoding detection.

    Primary entry point for plain text extraction. Reads the entire file
    content, detects the encoding using charset_normalizer, and decodes
    it appropriately.

    This function uses a generator pattern for API consistency with other
    extractors, even though text files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete text file data.
            The stream position is reset to the beginning before reading.
            Can contain either bytes or str (bytes is typical).
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned PlainTextContent.metadata.

    Yields:
        PlainTextContent: Single PlainTextContent object containing:
            - content: Full text content as a string
            - metadata: FileMetadataInterface with file information and
              detected_encoding field

    Note:
        Encoding is automatically detected from file content. For very
        short files (< 32 bytes), detection may be unreliable and UTF-8
        is used as the default.

    Example:
        >>> import io
        >>> with open("readme.txt", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_plain_text(data, path="readme.txt"):
        ...         print(f"Encoding: {doc.metadata.detected_encoding}")
        ...         lines = doc.content.splitlines()
        ...         print(f"Lines: {len(lines)}")
        ...         print(f"First line: {lines[0] if lines else '(empty)'}")
    """
    try:
        logger.debug("Reading plain text file")
        file_like.seek(0)

        content = file_like.read()

        if isinstance(content, bytes):
            text, detected_encoding = _detect_and_decode(content)
        else:
            text = content
            detected_encoding = "utf-8"  # Already a string, assume UTF-8

        metadata = FileMetadataInterface()
        metadata.populate_from_path(path)
        metadata.detected_encoding = detected_encoding

        yield PlainTextContent(content=text, metadata=metadata)
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError(
            "Failed to extract plain text file", cause=exc
        ) from exc
