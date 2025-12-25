"""
sharepoint-to-text: Text extraction library for SharePoint file formats.

A Python library for extracting plain text content from files typically found
in SharePoint repositories. Supports both modern Office Open XML formats and
legacy binary formats, plus PDF documents.
"""

import io
from pathlib import Path

from sharepoint2text.extractors.doc_extractor import MicrosoftDocContent, read_doc
from sharepoint2text.extractors.docx_extractor import MicrosoftDocxContent, read_docx
from sharepoint2text.extractors.pdf_extractor import PdfContent, read_pdf
from sharepoint2text.extractors.plain_extractor import PlainTextContent, read_plain_text
from sharepoint2text.extractors.ppt_extractor import PPTContent, read_ppt
from sharepoint2text.extractors.pptx_extractor import MicrosoftPptxContent, read_pptx
from sharepoint2text.extractors.xls_extractor import MicrosoftXlsContent, read_xls
from sharepoint2text.extractors.xlsx_extractor import MicrosoftXlsxContent, read_xlsx
from sharepoint2text.router import get_extractor, is_supported_file

__version__ = "0.1.1.dev31"


def read_file(path: str | Path):
    """
    Read and extract content from a file.

    Automatically detects the file type based on extension and uses
    the appropriate extractor.

    Args:
        path: Path to the file to read.

    Returns:
        A dataclass containing extracted content and metadata.
        The specific type depends on the file format:
        - .docx -> MicrosoftDocxContent
        - .doc  -> MicrosoftDocContent
        - .xlsx -> MicrosoftXlsxContent
        - .xls  -> MicrosoftXlsContent
        - .pptx -> MicrosoftPptxContent
        - .ppt  -> PPTContent
        - .pdf  -> PdfContent
        - .txt  -> PlainTextContent

    Raises:
        RuntimeError: If the file type is not supported.
        FileNotFoundError: If the file does not exist.

    Example:
        >>> import sharepoint2text
        >>> result = sharepoint2text.read_file("document.docx")
        >>> print(result.get_full_text())
    """
    path = Path(path)
    extractor = get_extractor(str(path))
    with open(path, "rb") as f:
        return extractor(io.BytesIO(f.read()))


__all__ = [
    # Version
    "__version__",
    # Main functions
    "read_file",
    "is_supported_file",
    "get_extractor",
    # Format-specific extractors
    "read_docx",
    "read_doc",
    "read_xlsx",
    "read_xls",
    "read_pptx",
    "read_ppt",
    "read_pdf",
    "read_plain_text",
    # Content dataclasses (return types)
    "MicrosoftDocxContent",
    "MicrosoftDocContent",
    "MicrosoftXlsxContent",
    "MicrosoftXlsContent",
    "MicrosoftPptxContent",
    "PPTContent",
    "PdfContent",
    "PlainTextContent",
]
