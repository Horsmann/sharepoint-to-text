"""
CSV/TSV Content Extractor
=========================

Extracts structured tabular content from CSV and TSV files with automatic
encoding and dialect detection.

Unlike the plain text extractor, this module parses the delimited structure
and returns a ``CsvContent`` object whose ``iterate_tables()`` yields a
``TableData`` instance suitable for downstream processing (DataFrames,
Markdown rendering, etc.).

Encoding Detection
------------------
Uses charset_normalizer (same as plain_extractor) for automatic encoding
detection with UTF-8 fallback.

Delimiter Detection
-------------------
``csv.Sniffer`` is tried first for automatic dialect detection.  When
sniffing fails (very short files, unusual quoting) the extractor falls
back to tab-delimited for ``.tsv`` files and comma-delimited otherwise.

Dependencies
------------
    - charset_normalizer: Encoding detection library
    - csv (stdlib): CSV parsing

Extracted Content
-----------------
    - content: Full text (decoded, unmodified)
    - table: ``TableData`` with each row as a list of cell strings
    - metadata: ``FileMetadataInterface`` with detected_encoding
"""

import csv
import io
import logging
from typing import Any, Generator

from charset_normalizer import from_bytes

from sharepoint2text.parsing.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.parsing.extractors._legacy_types import (
    CsvContent,
    FileMetadataInterface,
    TableData,
)

logger = logging.getLogger(__name__)

# Maximum bytes used by csv.Sniffer for dialect detection
_SNIFF_SAMPLE_SIZE = 8192


def _detect_and_decode(content: bytes) -> tuple[str, str]:
    """Detect encoding and decode bytes to string."""
    if not content:
        return "", "utf-8"

    results = from_bytes(content)
    best_match = results.best()

    if best_match is not None:
        encoding = best_match.encoding
        try:
            return str(best_match), encoding
        except (UnicodeDecodeError, LookupError, ValueError):
            pass

    return content.decode("utf-8", errors="replace"), "utf-8"


def _is_tsv(path: str | None) -> bool:
    """Return True when the path hints at tab-separated values."""
    if not path:
        return False
    return path.lower().endswith(".tsv")


def read_csv(
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
) -> Generator[CsvContent, Any, None]:
    """Extract structured content from a CSV or TSV file.

    Args:
        file_like: BytesIO object containing the file data.
        path: Optional filesystem path (used for metadata and TSV detection).
        ignore_images: Unused, accepted for interface consistency.

    Yields:
        A single ``CsvContent`` with both the raw text and a ``TableData``.
    """
    source_path = path or "<in-memory>"
    logger.info("Entering CSV/TSV extraction: %s", source_path)

    try:
        file_like.seek(0)
        raw = file_like.read()

        if isinstance(raw, bytes):
            text, detected_encoding = _detect_and_decode(raw)
        else:
            text = raw
            detected_encoding = "utf-8"

        # Parse the delimited data
        rows: list[list[str]] = []
        if text.strip():
            dialect: Any = None
            try:
                sample = text[:_SNIFF_SAMPLE_SIZE]
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                pass

            if dialect is None:
                delimiter = "\t" if _is_tsv(path) else ","
                reader = csv.reader(io.StringIO(text), delimiter=delimiter)
            else:
                reader = csv.reader(io.StringIO(text), dialect=dialect)

            for row in reader:
                rows.append(row)

        metadata = FileMetadataInterface()
        metadata.populate_from_path(path)
        metadata.detected_encoding = detected_encoding

        yield CsvContent(
            content=text,
            table=TableData(data=rows),
            metadata=metadata,
        )
    except ExtractionError:
        raise
    except (OSError, UnicodeDecodeError, ValueError, TypeError, csv.Error) as exc:
        raise ExtractionFailedError(
            "Failed to extract CSV/TSV file", cause=exc
        ) from exc
