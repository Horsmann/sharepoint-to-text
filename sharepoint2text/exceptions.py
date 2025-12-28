class ExtractionFileFormatNotSupportedError(Exception):
    """Raised when the file format for extraction is not supported."""

    def __init__(self, file_path: str, message: str = None, *, cause: Exception = None):
        self.file_path = file_path
        if message is None:
            message = f"Extraction file format not supported: {file_path}"
        # Use exception chaining if cause is provided
        super().__init__(message)
        self.__cause__ = cause  # Optional chaining for debugging


class LegacyMicrosoftParsingError(Exception):
    """Raised when parsing a legacy doc file."""

    def __init__(self, message: str = None, *, cause: Exception = None):
        if message is None:
            message = "Error when processing legacy doc file"
        # Use exception chaining if cause is provided
        super().__init__(message)
        # Optional chaining for debugging
        self.__cause__ = cause
