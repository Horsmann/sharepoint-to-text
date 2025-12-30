"""
PDF Content Extractor
=====================

Extracts text content, metadata, and embedded images from Portable Document
Format (PDF) files using the pypdf library.

File Format Background
----------------------
PDF (Portable Document Format) is a file format developed by Adobe for
document exchange. Key characteristics:
    - Fixed-layout format preserving visual appearance
    - Can contain text, images, vector graphics, annotations
    - Text may be stored as character codes with font mappings
    - Images stored as XObject resources with various compressions
    - Page-based structure with independent page content streams

PDF Internal Structure
----------------------
Relevant components for text extraction:
    - Page objects: Define content streams and resources
    - Content streams: Drawing operators including text operators
    - Font resources: Character encoding mappings
    - XObject resources: Images and reusable graphics
    - Catalog/Info: Document metadata

Image Compression Types
-----------------------
PDF supports multiple image compression filters:
    - /DCTDecode: JPEG compression (lossy)
    - /JPXDecode: JPEG 2000 compression
    - /FlateDecode: PNG-style deflate compression
    - /CCITTFaxDecode: TIFF Group 3/4 fax compression
    - /JBIG2Decode: JBIG2 compression for bi-level images
    - /LZWDecode: LZW compression (legacy)

Dependencies
------------
pypdf (https://pypdf.readthedocs.io/):
    - Pure Python PDF library (no external dependencies)
    - Successor to PyPDF2
    - Provides text extraction via content stream parsing
    - Handles encrypted PDFs (with password)
    - Image extraction from XObject resources

Extracted Content
-----------------
Per-page content includes:
    - text: Extracted text content (may have layout artifacts)
    - images: List of PdfImage objects with:
        - Binary data in original format
        - Dimensions (width, height)
        - Color space information
        - Compression filter type

Metadata extraction includes:
    - total_pages: Number of pages in document
    - File metadata from path (if provided)

Text Extraction Caveats
-----------------------
PDF text extraction is inherently imperfect:
    - Text order depends on content stream order, not visual layout
    - Columns may interleave incorrectly
    - Hyphenation at line breaks may not be detected
    - Ligatures may extract as single characters
    - Some fonts use custom encodings (CID fonts, symbolic)
    - Rotated text may extract in unexpected order

Known Limitations
-----------------
- Scanned PDFs (image-only) return empty text (no OCR)
- Form field values (AcroForms/XFA) are not extracted
- Annotations and comments are not extracted
- Digital signatures are not reported
- Embedded files/attachments are not extracted
- Very complex layouts may have garbled text order
- Password-protected PDFs require the password

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.pdf_extractor import read_pdf
    >>>
    >>> with open("document.pdf", "rb") as f:
    ...     for doc in read_pdf(io.BytesIO(f.read()), path="document.pdf"):
    ...         print(f"Pages: {doc.metadata.total_pages}")
    ...         for page_num, page in enumerate(doc.pages, start=1):
    ...             print(f"Page {page_num}: {len(page.text)} chars, {len(page.images)} images")

See Also
--------
- pypdf documentation: https://pypdf.readthedocs.io/
- PDF Reference: https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf

Maintenance Notes
-----------------
- pypdf handles most PDF quirks internally
- Image extraction accesses raw XObject data
- Failed image extractions are logged and skipped (not raised)
- Color space reported as string for debugging
- Format detection based on compression filter type
"""

import io
import logging
from typing import Any, Generator, List

from pypdf import PdfReader
from pypdf.generic import ContentStream

from sharepoint2text.extractors.data_types import (
    PdfContent,
    PdfImage,
    PdfMetadata,
    PdfPage,
)

logger = logging.getLogger(__name__)


def read_pdf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PdfContent, Any, None]:
    """
    Extract all relevant content from a PDF file.

    Primary entry point for PDF extraction. Uses pypdf to parse the document
    structure, extract text from each page's content stream, and extract
    embedded images from XObject resources.

    This function uses a generator pattern for API consistency with other
    extractors, even though PDF files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete PDF file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned PdfContent.metadata.

    Yields:
        PdfContent: Single PdfContent object containing:
            - pages: List of PdfPage objects in document order
            - metadata: PdfMetadata with total_pages and file info

    Note:
        Scanned PDFs containing only images will yield pages with empty
        text strings. OCR is not performed. For scanned documents, the
        images are still extracted and could be processed separately.

    Example:
        >>> import io
        >>> with open("report.pdf", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_pdf(data, path="report.pdf"):
        ...         print(f"Total pages: {doc.metadata.total_pages}")
        ...         for page_num, page in enumerate(doc.pages, start=1):
        ...             print(f"Page {page_num}:")
        ...             print(f"  Text: {page.text[:100]}...")
        ...             print(f"  Images: {len(page.images)}")
    """
    file_like.seek(0)
    reader = PdfReader(file_like)
    logger.debug("Parsing PDF with %d pages", len(reader.pages))

    pages = []
    total_images = 0
    for page_num, page in enumerate(reader.pages, start=1):
        images = _extract_image_bytes(page, page_num)
        total_images += len(images)
        pages.append(
            PdfPage(
                text=page.extract_text() or "",
                images=images,
            )
        )

    metadata = PdfMetadata(total_pages=len(reader.pages))
    metadata.populate_from_path(path)

    logger.info(
        "Extracted PDF: %d pages, %d images",
        len(reader.pages),
        total_images,
    )

    yield PdfContent(
        pages=pages,
        metadata=metadata,
    )


def _extract_image_bytes(page, page_num: int) -> List[PdfImage]:
    """
    Extract all images from a PDF page's XObject resources.

    Iterates through the page's /Resources/XObject dictionary, identifies
    image objects by their /Subtype attribute, and extracts each image's
    binary data and properties.

    Args:
        page: A pypdf PageObject to extract images from.

    Returns:
        List of PdfImage objects for successfully extracted images.
        Failed extractions are logged and skipped.
    """
    found_images = []
    if "/XObject" not in page.get("/Resources", {}):
        return found_images

    x_objects = page["/Resources"]["/XObject"].get_object()
    image_occurrences, mcid_order, mcid_text = _extract_page_mcid_data(page)

    if image_occurrences:
        image_index = 1
        for occurrence in image_occurrences:
            obj_name = occurrence["name"]
            obj = x_objects.get(obj_name)
            if obj is None or obj.get("/Subtype") != "/Image":
                continue
            caption = _lookup_caption(
                occurrence.get("mcid"),
                mcid_order,
                mcid_text,
            )
            try:
                image_data = _extract_image(
                    obj,
                    obj_name,
                    image_index,
                    page_num,
                    caption,
                )
                found_images.append(image_data)
                image_index += 1
            except Exception as e:
                logger.warning(
                    f"Silently ignoring - Failed to extract image [{obj_name}] [{image_index}]: %s",
                    e,
                )
        return found_images

    attempt_index = 1
    image_index = 1
    for obj_name in x_objects:
        obj = x_objects[obj_name]

        if obj.get("/Subtype") == "/Image":
            try:
                image_data = _extract_image(
                    obj,
                    obj_name,
                    image_index,
                    page_num,
                    "",
                )
                found_images.append(image_data)
                image_index += 1
            except Exception as e:
                logger.warning(
                    f"Silently ignoring - Failed to extract image [{obj_name}] [{attempt_index}]: %s",
                    e,
                )
            attempt_index += 1
    return found_images


def _extract_image(
    image_obj,
    name,
    index: int,
    page_num: int,
    caption: str,
) -> PdfImage:
    """
    Extract image data and properties from a PDF image XObject.

    Reads the image object's attributes to determine dimensions, color
    space, and compression filter. Maps the filter type to a standard
    image format identifier and extracts the raw binary data.

    Args:
        image_obj: A pypdf image object from the XObject dictionary.
        name: The XObject name (e.g., "/Im0") for identification.
        index: 1-based index for ordering extracted images on the page.

    Returns:
        PdfImage with binary data and image properties.
    """

    width = image_obj.get("/Width", 0)
    height = image_obj.get("/Height", 0)
    color_space = str(image_obj.get("/ColorSpace", "unknown"))
    bits = image_obj.get("/BitsPerComponent", 8)

    # Determine image format based on filter
    filter_type = image_obj.get("/Filter", "")
    if isinstance(filter_type, list):
        filter_type = filter_type[-1] if filter_type else ""
    filter_type = str(filter_type)

    # Map filter to format
    format_map = {
        "/DCTDecode": "jpeg",
        "/JPXDecode": "jp2",
        "/FlateDecode": "png",
        "/CCITTFaxDecode": "tiff",
        "/JBIG2Decode": "jbig2",
        "/LZWDecode": "png",
    }
    img_format = format_map.get(filter_type, "raw")

    content_type_map = {
        "/DCTDecode": "image/jpeg",
        "/JPXDecode": "image/jp2",
        "/FlateDecode": "image/png",
        "/CCITTFaxDecode": "image/tiff",
        "/JBIG2Decode": "image/jbig2",
        "/LZWDecode": "image/png",
    }
    content_type = content_type_map.get(filter_type, "image/unknown")

    # Get raw image data
    try:
        data = image_obj.get_data()
    except Exception as e:
        logger.warning("Failed to extract image data: %s", e)
        data = image_obj._data if hasattr(image_obj, "_data") else b""

    resolved_caption = caption or _extract_image_alt_text(image_obj)

    return PdfImage(
        index=index,
        name=str(name),
        caption=resolved_caption,
        width=int(width),
        height=int(height),
        color_space=color_space,
        bits_per_component=int(bits),
        filter=filter_type,
        data=data,
        format=img_format,
        content_type=content_type,
        unit_index=page_num,
    )


def _extract_page_mcid_data(page) -> tuple[list[dict], list[int], dict[int, str]]:
    """Collect MCID text and image occurrence order from the page content stream."""
    contents = page.get_contents()
    if contents is None:
        return [], [], {}

    try:
        stream = ContentStream(contents, page.pdf)
    except Exception as e:
        logger.debug("Failed to parse content stream: %s", e)
        return [], [], {}

    mcid_stack: list[int | None] = []
    actual_text_stack: list[str | None] = []
    mcid_order: list[int] = []
    mcid_text: dict[int, str] = {}
    image_occurrences: list[dict] = []

    for operands, operator in stream.operations:
        op = (
            operator.decode("utf-8", errors="ignore")
            if isinstance(operator, bytes)
            else operator
        )
        if op in ("BDC", "BMC"):
            current_mcid = mcid_stack[-1] if mcid_stack else None
            actual_text = None
            if op == "BDC" and len(operands) >= 2:
                props = operands[1]
                if isinstance(props, dict):
                    if "/MCID" in props:
                        current_mcid = props.get("/MCID")
                    actual_text = props.get("/ActualText")
            mcid_stack.append(current_mcid)
            actual_text_stack.append(actual_text)
            if current_mcid is not None and current_mcid not in mcid_order:
                mcid_order.append(current_mcid)
            continue

        if op == "EMC":
            if mcid_stack:
                mcid_stack.pop()
            if actual_text_stack:
                actual_text_stack.pop()
            continue

        if op == "Do":
            if not operands:
                continue
            current_mcid = mcid_stack[-1] if mcid_stack else None
            image_occurrences.append(
                {
                    "name": operands[0],
                    "mcid": current_mcid,
                }
            )
            continue

        if op in ("Tj", "TJ", "'", '"'):
            current_mcid = mcid_stack[-1] if mcid_stack else None
            if current_mcid is None:
                continue
            actual_text = actual_text_stack[-1] if actual_text_stack else None
            if actual_text:
                text = str(actual_text)
                actual_text_stack[-1] = None
            else:
                text = _extract_text_from_operands(op, operands)
            if text:
                mcid_text[current_mcid] = mcid_text.get(current_mcid, "") + text
                if current_mcid not in mcid_order:
                    mcid_order.append(current_mcid)

    return image_occurrences, mcid_order, mcid_text


def _extract_text_from_operands(operator: str, operands: list) -> str:
    if not operands:
        return ""
    if operator == "TJ":
        parts = []
        for item in operands[0]:
            if isinstance(item, (str, bytes)):
                parts.append(_normalize_text(item))
        return "".join(parts)
    if isinstance(operands[0], (str, bytes)):
        return _normalize_text(operands[0])
    return ""


def _normalize_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", errors="ignore")
    return str(value)


def _lookup_caption(
    mcid: int | None,
    mcid_order: list[int],
    mcid_text: dict[int, str],
) -> str:
    if mcid is None:
        return ""
    text = mcid_text.get(mcid, "").strip()
    if text:
        return text
    if not mcid_order:
        return ""
    try:
        start_index = mcid_order.index(mcid)
    except ValueError:
        return ""
    for next_mcid in mcid_order[start_index + 1 :]:
        next_text = mcid_text.get(next_mcid, "").strip()
        if next_text:
            return next_text
    return ""


def _extract_image_alt_text(image_obj) -> str:
    """Extract alt text or title for a PDF image XObject if present."""
    caption_keys = ("/Alt", "/Title", "/Caption", "/TU")
    for key in caption_keys:
        value = image_obj.get(key)
        if isinstance(value, str):
            if value.strip():
                return value
        elif value is not None:
            text = str(value).strip()
            if text:
                return text
    return ""
