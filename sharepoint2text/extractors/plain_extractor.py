"""
Plain text content extractor.
"""

import io
import logging
import typing
from dataclasses import dataclass

from sharepoint2text.extractors.abstract_extractor import ExtractionInterface

logger = logging.getLogger(__name__)


@dataclass
class PlainTextContent(ExtractionInterface):
    content: str = ""

    def iterator(self) -> typing.Iterator[str]:
        yield self.content

    def get_full_text(self) -> str:
        return self.content


def read_plain_text(file_like: io.BytesIO) -> PlainTextContent:
    """
    Extract content from a plain text file.

    Args:
        file_like: A BytesIO object containing the text file data.

    Returns:
        PlainTextContent dataclass with the extracted content.
    """
    logger.debug("Reading plain text file")
    file_like.seek(0)

    content = file_like.read()

    if isinstance(content, bytes):
        text = content.decode("utf-8", errors="ignore")
    else:
        text = content

    return PlainTextContent(content=text)
