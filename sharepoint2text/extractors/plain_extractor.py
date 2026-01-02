"""
Plain Text Content Extractor
=============================

Extracts content from plain text files with encoding detection and
normalization.

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
The extractor currently assumes UTF-8 encoding with error replacement:
    - UTF-8 is the most common encoding for modern text files
    - Invalid byte sequences are replaced with the Unicode replacement
      character (U+FFFD) rather than raising an exception
    - This provides best-effort extraction for mixed-encoding files

Future enhancement could add encoding detection using chardet or
charset_normalizer libraries.

Dependencies
------------
Python Standard Library only:
    - io: BytesIO handling
    - No external dependencies required

Extracted Content
-----------------
The extractor produces:
    - content: Full text content as a single string
    - metadata: FileMetadataInterface with file information

The content is returned as-is without modification, preserving:
    - Line endings (\\n, \\r\\n, \\r)
    - Whitespace and indentation
    - Empty lines

Known Limitations
-----------------
- Only UTF-8 encoding is supported (with error replacement)
- No automatic encoding detection (e.g., chardet)
- Binary files may produce garbled output or replacement characters
- Very large files are loaded entirely into memory
- No line ending normalization

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.plain_extractor import read_plain_text
    >>>
    >>> with open("notes.txt", "rb") as f:
    ...     for doc in read_plain_text(io.BytesIO(f.read()), path="notes.txt"):
    ...         print(f"Characters: {len(doc.content)}")
    ...         print(doc.content[:200])

See Also
--------
- Python codecs module: https://docs.python.org/3/library/codecs.html
- Unicode HOWTO: https://docs.python.org/3/howto/unicode.html

Maintenance Notes
-----------------
- Uses errors="ignore" for decoding (could change to "replace" for visibility)
- FileMetadataInterface provides basic file info population
- Generator pattern for API consistency with other extractors
- Content returned unmodified (no stripping or normalization)
"""

import io
import logging
from typing import Any, Generator

from sharepoint2text.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.extractors.data_types import (
    FileMetadataInterface,
    PlainTextContent,
)

logger = logging.getLogger(__name__)


def read_plain_text(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PlainTextContent, Any, None]:
    """
    Extract content from a plain text file.

    Primary entry point for plain text extraction. Reads the entire file
    content and decodes it as UTF-8 with invalid byte replacement.

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
            - metadata: FileMetadataInterface with file information

    Note:
        Non-UTF-8 bytes are silently ignored. For files with mixed
        or unknown encodings, some characters may be lost. Consider
        using a dedicated encoding detection library for critical
        applications.

    Example:
        >>> import io
        >>> with open("readme.txt", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_plain_text(data, path="readme.txt"):
        ...         lines = doc.content.splitlines()
        ...         print(f"Lines: {len(lines)}")
        ...         print(f"First line: {lines[0] if lines else '(empty)'}")
    """
    try:
        logger.debug("Reading plain text file")
        file_like.seek(0)

        content = file_like.read()

        if isinstance(content, bytes):
            text = content.decode("utf-8", errors="ignore")
        else:
            text = content

        metadata = FileMetadataInterface()
        metadata.populate_from_path(path)

        yield PlainTextContent(content=text, metadata=metadata)
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError(
            "Failed to extract plain text file", cause=exc
        ) from exc
