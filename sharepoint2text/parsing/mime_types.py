MIME_TYPE_MAPPING = {
    # Legacy MS Office
    "application/vnd.ms-powerpoint": "ppt",
    "application/vnd.ms-excel": "xls",
    "application/msword": "doc",
    "application/rtf": "rtf",
    "text/rtf": "rtf",
    # Modern MS Office
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/vnd.ms-excel.sheet.binary.macroEnabled.12": "xlsb",
    # Macro-enabled variants
    "application/vnd.ms-powerpoint.presentation.macroEnabled.12": "pptm",
    "application/vnd.ms-word.document.macroEnabled.12": "docm",
    "application/vnd.ms-excel.sheet.macroEnabled.12": "xlsm",
    # OpenDocument formats
    "application/vnd.oasis.opendocument.text": "odt",
    "application/vnd.oasis.opendocument.presentation": "odp",
    "application/vnd.oasis.opendocument.spreadsheet": "ods",
    "application/vnd.oasis.opendocument.graphics": "odg",
    "application/vnd.oasis.opendocument.formula": "odf",
    # Plain text variants
    "text/csv": "csv",
    "application/csv": "csv",
    "application/json": "json",
    "text/json": "json",
    "text/plain": "txt",
    "text/markdown": "md",
    "text/tab-separated-values": "tsv",
    "application/tab-separated-values": "tsv",
    # Email formats
    "application/vnd.ms-outlook": "msg",
    "message/rfc822": "eml",
    "application/mbox": "mbox",
    # Other formats
    "text/html": "html",
    "application/xhtml+xml": "html",
    "application/pdf": "pdf",
    "application/epub+zip": "epub",
    # MHTML detection is extension-based (.mhtml, .mht) since MIME overlaps with EML
    # Archive formats
    "application/zip": "zip",
    "application/x-zip-compressed": "zip",
    "application/x-tar": "tar",
}


def is_supported_mime_type(mime_type: str | None) -> bool:
    """Return whether a MIME type maps to a registered extractor.

    Args:
        mime_type: MIME type to check. ``None`` and empty strings are unsupported.

    Returns:
        True when the normalized media type is supported.
    """
    if not mime_type:
        return False
    return mime_type in MIME_TYPE_MAPPING
