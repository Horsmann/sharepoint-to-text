"""Extract content from SharePoint file formats into one normalized model."""

from sharepoint2text._api import (
    InvalidConfigurationError,
    __version__,
    is_supported_file,
    read_bytes,
    read_file,
    read_many,
)
from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
    ExtractionFileFormatNotSupportedError,
    ExtractionFileTooLargeError,
    ExtractionLegacyMicrosoftParsingError,
    ExtractionZipBombError,
)
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    ZipBombLimits,
)
from sharepoint2text.parsing.models import (
    Annotation,
    Attachment,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    SourceMetadata,
    Table,
    document_from_dict,
    document_from_json,
    document_to_dict,
    document_to_json,
    render_markdown,
)

__all__ = [
    "Annotation",
    "Attachment",
    "ContentUnit",
    "DocumentMetadata",
    "ExtractedDocument",
    "ExtractionError",
    "ExtractionFailedError",
    "ExtractionFileEncryptedError",
    "ExtractionFileFormatNotSupportedError",
    "ExtractionFileTooLargeError",
    "ExtractionLegacyMicrosoftParsingError",
    "ExtractionZipBombError",
    "ImageAsset",
    "InvalidConfigurationError",
    "SourceMetadata",
    "Table",
    "ZipBombLimits",
    "__version__",
    "document_from_dict",
    "document_from_json",
    "document_to_dict",
    "document_to_json",
    "is_supported_file",
    "read_bytes",
    "read_file",
    "read_many",
    "render_markdown",
]
