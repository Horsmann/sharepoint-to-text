import io
from typing import Any, Generator

from parsing.extractors.data_types import ApplePagesContent

# =============================================================================
# Main entry point
# =============================================================================


def read_apple_pages(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[ApplePagesContent, Any, None]:
    pass
