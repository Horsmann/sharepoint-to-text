"""
CSV/TSV Content Extractor
=========================

Extracts structured tabular content from CSV and TSV files with automatic
encoding and dialect detection.

Unlike the plain text extractor, this module parses the delimited structure
and returns a canonical document with a table suitable for downstream
processing (DataFrames, Markdown rendering, etc.).

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
    - table: canonical table with each row as a list of cell strings
    - source: canonical source metadata with detected encoding
"""

import csv
import io
import logging
from pathlib import Path
from typing import Any, Generator, cast

from charset_normalizer import from_bytes

from sharepoint2text.parsing.exceptions import ExtractionError, ExtractionFailedError
from sharepoint2text.parsing.extractors._model import source_metadata
from sharepoint2text.parsing.models import (
    CellValue,
    ContentUnit,
    ExtractedDocument,
    Table,
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
) -> Generator[ExtractedDocument, Any, None]:
    """Extract structured content from a CSV or TSV file.

    Args:
        file_like: BytesIO object containing the file data.
        path: Optional filesystem path (used for metadata and TSV detection).
        ignore_images: Unused, accepted for interface consistency.

    Yields:
        A single canonical document with both the raw text and table data.
    """
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

        logger.debug(
            "Extracted delimited text: rows=%d, encoding=%s",
            len(rows),
            detected_encoding,
        )

        source_format = Path(path).suffix.lower().lstrip(".") if path else "csv"
        tables = [Table(rows=cast(list[list[CellValue]], rows))] if rows else []
        yield ExtractedDocument(
            format=source_format or "csv",
            source=source_metadata(path, encoding=detected_encoding),
            units=[
                ContentUnit(
                    number=1,
                    kind="document",
                    text=text,
                    tables=tables,
                )
            ],
        )
    except ExtractionError:
        raise
    except (OSError, UnicodeDecodeError, ValueError, TypeError, csv.Error) as exc:
        raise ExtractionFailedError(
            "Failed to extract CSV/TSV file", cause=exc
        ) from exc
