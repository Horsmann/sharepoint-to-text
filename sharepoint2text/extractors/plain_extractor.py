"""
Plain text content extractor.
"""

import io
import logging
from typing import Any, Generator

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

    Args:
        file_like: A BytesIO object containing the text file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        PlainTextContent dataclass with the extracted content.
    """
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
