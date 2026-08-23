"""Build canonical extraction models from parser results."""

from __future__ import annotations

import mimetypes
from pathlib import Path

from sharepoint2text.parsing.models import SourceMetadata


def source_metadata(
    path: str | None,
    *,
    encoding: str | None = None,
    media_type: str | None = None,
) -> SourceMetadata:
    """Build canonical source metadata from an optional source path.

    Args:
        path: Source path supplied to the extractor, or ``None``.
        encoding: Detected source encoding, when available.
        media_type: Explicit source media type, when known.

    Returns:
        Canonical source identity populated without requiring the path to exist.

    Example:
        >>> source_metadata("docs/report.pdf").filename
        'report.pdf'
    """
    if path is None:
        return SourceMetadata(encoding=encoding, media_type=media_type)
    source_path = Path(path)
    return SourceMetadata(
        filename=source_path.name,
        extension=source_path.suffix or None,
        path=str(source_path.resolve()) if source_path.exists() else str(source_path),
        folder=(
            str(source_path.parent.resolve())
            if source_path.parent.exists()
            else str(source_path.parent)
        ),
        media_type=media_type or mimetypes.guess_type(source_path.name)[0],
        encoding=encoding,
    )
