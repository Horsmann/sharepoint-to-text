"""Parsing and extraction utilities for sharepoint2text."""

from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
    ExtractionFileFormatNotSupportedError,
    ExtractionFileTooLargeError,
    ExtractionLegacyMicrosoftParsingError,
    ExtractionZipBombError,
)
from sharepoint2text.parsing.router import get_extractor, is_supported_file

__all__ = [
    "ExtractionError",
    "ExtractionFailedError",
    "ExtractionFileEncryptedError",
    "ExtractionFileFormatNotSupportedError",
    "ExtractionFileTooLargeError",
    "ExtractionLegacyMicrosoftParsingError",
    "ExtractionZipBombError",
    "get_extractor",
    "is_supported_file",
]
