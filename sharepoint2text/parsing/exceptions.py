"""
Custom exceptions for the sharepoint2text extraction library.

These exceptions provide clear error semantics for common failure modes
during document extraction.
"""

from typing import Optional


class ExtractionError(Exception):
    """Base exception for sharepoint2text errors."""


class ExtractionFileFormatNotSupportedError(ExtractionError):
    """
    Raised when the file format is not supported for extraction.

    This exception indicates that the library does not have an extractor
    capable of handling the given file type.

    Attributes:
        message: Human-readable description of the error.
        __cause__: Optional underlying exception that triggered this error.
    """

    def __init__(self, message: str, *, cause: Optional[Exception] = None):
        super().__init__(message)
        if cause is not None:
            self.__cause__ = cause


class LegacyMicrosoftParsingError(ExtractionError):
    """
    Raised when parsing a legacy Microsoft Office file fails.

    This exception indicates issues with parsing binary formats such as
    .doc, .xls, or .ppt files. Common causes include corrupted files,
    encrypted content, or unsupported format variations.

    Attributes:
        message: Human-readable description of the error.
        __cause__: Optional underlying exception that triggered this error.
    """

    def __init__(
        self,
        message: str = "Error when processing legacy Microsoft Office file",
        *,
        cause: Optional[Exception] = None,
    ):
        super().__init__(message)
        if cause is not None:
            self.__cause__ = cause


class ExtractionFileEncryptedError(ExtractionError):
    """
    Raised when the file appears to be encrypted or password-protected.

    This exception is used when an extractor detects encryption and
    cannot proceed without a password or decryption step.

    Attributes:
        message: Human-readable description of the error.
        __cause__: Optional underlying exception that triggered this error.
    """

    def __init__(
        self,
        message: str = "File is encrypted or password-protected",
        *,
        cause: Optional[Exception] = None,
    ):
        super().__init__(message)
        if cause is not None:
            self.__cause__ = cause


class ExtractionZipBombError(ExtractionError):
    """
    Raised when a ZIP container is deemed unsafe (probable ZIP bomb).

    Applies to formats that are ZIP containers (OOXML/ODF: .docx/.pptx/.odt/.ods/.odp)
    and other ZIP-based processing.
    """

    def __init__(
        self,
        message: str = "ZIP container rejected (possible ZIP bomb)",
        *,
        cause: Optional[Exception] = None,
    ):
        super().__init__(message)
        if cause is not None:
            self.__cause__ = cause


class ExtractionFailedError(ExtractionError):
    """
    Raised when extraction fails for an unexpected reason.

    Intended for use at API boundaries (e.g. `read_file`) to provide a
    stable exception surface while preserving the original exception in
    `__cause__`.
    """

    def __init__(
        self,
        message: str = "Extraction failed",
        *,
        cause: Optional[Exception] = None,
    ):
        super().__init__(message)
        if cause is not None:
            self.__cause__ = cause
