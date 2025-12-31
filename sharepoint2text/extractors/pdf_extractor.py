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
import re
import unicodedata
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
    total_tables = 0
    for page_num, page in enumerate(reader.pages, start=1):
        images = _extract_image_bytes(page, page_num)
        total_images += len(images)
        page_text = page.extract_text() or ""
        tables = _extract_tables_from_lines(page_text.splitlines())
        total_tables += len(tables)
        pages.append(
            PdfPage(
                text=page_text,
                images=images,
                tables=tables,
            )
        )

    metadata = PdfMetadata(total_pages=len(reader.pages))
    metadata.populate_from_path(path)

    logger.info(
        "Extracted PDF: %d pages, %d images, %d tables",
        len(reader.pages),
        total_images,
        total_tables,
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


def _extract_tables_from_lines(lines: List[str]) -> List[List[List[str]]]:
    """Extract basic tables from page lines using numeric tail parsing."""
    lines = [line.strip() for line in lines]
    tables: List[List[List[str]]] = []
    current_rows: List[List[str]] = []
    column_count = 0
    gap_count = 0
    max_gap = 2
    min_value_columns = 2
    date_header_pattern = re.compile(r"\d{2}/\d{2}/\d{4}")
    month_pattern = re.compile(
        r"\b(January|February|March|April|May|June|July|August|September|October|November|December)\b",
        re.IGNORECASE,
    )
    known_words = _collect_known_words(lines)

    def is_numeric_token(token: str) -> bool:
        cleaned = token.strip()
        if not cleaned:
            return False
        if cleaned[0] in ("(", "-", "–") and cleaned[-1] == ")":
            cleaned = cleaned[1:-1]
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        return all(ch.isdigit() or ch in {",", "."} for ch in cleaned) and any(
            ch.isdigit() for ch in cleaned
        )

    def normalize_label(label: str) -> str:
        normalized = unicodedata.normalize("NFKC", label)
        normalized = re.sub(r"[\u2010-\u2013\u2212]", "-", normalized)
        normalized = re.sub(r"(?<=\\w)\\s*-\\s*(?=\\w)", "-", normalized)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        normalized = _split_compound_words(normalized, known_words)
        normalized = _normalize_phrasing(normalized)
        return normalized

    def is_footnote_leader(token: str) -> bool:
        return len(token) == 1 and token.isdigit()

    def normalize_values(values: list[str], expected_count: int) -> list[str]:
        if not values or expected_count <= 0:
            return values
        if len(values) == expected_count + 1 and is_footnote_leader(values[0]):
            return values[1:]
        merged = values[:]
        while len(merged) > expected_count:
            merged_any = False
            for idx in range(len(merged) - 1):
                if merged[idx].isdigit() and merged[idx + 1].isdigit():
                    merged[idx] = merged[idx] + merged[idx + 1]
                    del merged[idx + 1]
                    merged_any = True
                    break
            if not merged_any:
                merged[0] = merged[0] + merged[1]
                del merged[1]
        return merged

    def split_numeric_blob(blob: str, expected_count: int) -> list[str]:
        if expected_count != 2:
            return [blob]
        match = re.match(r"^(\d[\d,]*\.\d)(\d[\d,]*\.\d+)$", blob)
        if match:
            return [match.group(1), match.group(2)]
        return [blob]

    def is_date_header(line: str) -> bool:
        return len(date_header_pattern.findall(line)) >= 2

    def extract_date_header(line: str) -> tuple[str, list[str], str] | None:
        dates = date_header_pattern.findall(line)
        if len(dates) < 2:
            return None
        first_start = line.find(dates[0])
        label = line[:first_start].strip()
        label = normalize_label(label) if label else ""
        second_end = line.find(dates[1]) + len(dates[1])
        unit_text = line[second_end:].strip()
        return label, dates[:2], unit_text

    def extract_row(line: str) -> tuple[str, list[str]]:
        tokens = line.split()
        values: list[str] = []
        idx = len(tokens) - 1
        while idx >= 0 and is_numeric_token(tokens[idx]):
            values.append(tokens[idx])
            idx -= 1
        values.reverse()
        label_tokens = tokens[: idx + 1]
        label = " ".join(label_tokens).strip()
        if label:
            label = _strip_label_footnote(label)
            label = normalize_label(label)
        return label, values

    def _strip_label_footnote(label: str) -> str:
        cleaned = []
        for token in label.split():
            if token and token[-1].isdigit() and token[:-1].isalpha():
                cleaned.append(token[:-1])
            else:
                cleaned.append(token)
        return " ".join(cleaned).strip()

    def flush_current() -> None:
        if len(current_rows) >= 2:
            tables.append(current_rows.copy())

    pending_header_label = ""

    for idx, line in enumerate(lines):
        if not line:
            if current_rows:
                gap_count += 1
                if gap_count > max_gap:
                    flush_current()
                    current_rows = []
                    column_count = 0
                    gap_count = 0
            continue

        next_line = _next_non_empty_line(lines, idx)
        if (
            current_rows
            and not re.search(r"\d", line)
            and _looks_like_section_break(line)
        ):
            next_header = extract_date_header(next_line)
            if next_header:
                flush_current()
                current_rows = []
                column_count = 0
                gap_count = 0
                pending_header_label = normalize_label(line)
                continue

        if not current_rows and not re.search(r"\d", line):
            pending_header_label = normalize_label(line)

        date_header = extract_date_header(line)
        if date_header:
            label, dates, unit_text = date_header
            if current_rows:
                flush_current()
                current_rows = []
                column_count = 0
                gap_count = 0
            if not label and pending_header_label:
                label = pending_header_label
            pending_header_label = ""
            column_count = len(dates) + 1
            current_rows.append([label] + dates)
            if unit_text:
                current_rows.append(
                    [normalize_label(unit_text)] + [""] * (column_count - 1)
                )
            continue

        label, values = extract_row(line)
        has_values = bool(values)
        if has_values:
            pending_header_label = ""
        if has_values and month_pattern.search(line):
            if values and values[-1].isdigit() and len(values[-1]) == 4:
                values = []
                has_values = False
        if not has_values and current_rows and current_rows[-1][0] == "":
            if not line[:1].isdigit() and not _looks_like_unit_line(line):
                flush_current()
                current_rows = []
                column_count = 0
                gap_count = 0
                continue
        if not has_values and current_rows:
            if re.match(r"^\d+\.", line):
                flush_current()
                current_rows = []
                column_count = 0
                gap_count = 0
                continue
            if (
                len(line) >= 60
                and "." in line
                and not line[:1].isdigit()
                and not re.search(r"[\d.,]+$", line)
            ):
                flush_current()
                current_rows = []
                column_count = 0
                gap_count = 0
                continue
        values_from_trailing_blob = False
        if not values:
            trailing_match = re.search(r"([\d.,]+)$", line)
            if trailing_match and "/" not in trailing_match.group(1):
                blob = trailing_match.group(1)
                label = line[: trailing_match.start()].strip()
                if label:
                    label = normalize_label(label)
                values = split_numeric_blob(
                    blob, column_count - 1 if column_count else 2
                )
                values_from_trailing_blob = True
            if not values and not current_rows:
                continue

        if values and len(values) < min_value_columns and line[:1].isdigit():
            values = []
            label = normalize_label(line)

        if values and column_count == 0:
            if len(values) < min_value_columns:
                continue
            column_count = len(values) + 1

        if values or current_rows:
            if column_count == 0 and values:
                if len(values) < min_value_columns:
                    continue
                column_count = len(values) + 1
            if column_count == 0:
                continue
            expected_values = column_count - 1
            values = normalize_values(values, expected_values)
            if values and current_rows and values_from_trailing_blob:
                last_row = current_rows[-1]
                if last_row[0] and all(not cell for cell in last_row[1:]):
                    if label and label[:1].islower():
                        label = f"{last_row[0]} {label}"
                        current_rows.pop()
            if values:
                row = [label] + values
                if len(row) < column_count:
                    row.extend([""] * (column_count - len(row)))
                elif len(row) > column_count:
                    row = row[:column_count]
            else:
                row = [label] + [""] * (column_count - 1)
            current_rows.append(row)
            gap_count = 0

    flush_current()
    if tables:
        return tables
    return _extract_tables_from_text_simple(lines)


def _extract_tables_from_text_simple(lines: List[str]) -> List[List[List[str]]]:
    """Fallback extractor for non-numeric tables with consistent columns."""
    tables: List[List[List[str]]] = []
    current_rows: List[List[str]] = []
    current_cols = 0

    def flush_current() -> None:
        if len(current_rows) >= 2:
            tables.append(current_rows.copy())

    for line in lines:
        if not line:
            flush_current()
            current_rows = []
            current_cols = 0
            continue

        tokens = line.split()
        if len(tokens) < 2:
            flush_current()
            current_rows = []
            current_cols = 0
            continue

        if current_cols == 0:
            current_cols = len(tokens)
            current_rows = [tokens]
            continue

        if len(tokens) == current_cols:
            current_rows.append(tokens)
            continue

        flush_current()
        current_rows = [tokens]
        current_cols = len(tokens)

    flush_current()
    return tables


def _collect_known_words(lines: list[str]) -> dict[str, int]:
    words: dict[str, int] = {}
    for line in lines:
        for token in re.findall(r"[A-Za-z]+", line):
            key = token.lower()
            words[key] = words.get(key, 0) + 1
    return words


def _split_compound_words(text: str, known_words: dict[str, int]) -> str:
    tokens = text.split()
    if not tokens:
        return text
    line_words = {token.lower() for token in re.findall(r"[A-Za-z]+", text)}
    candidates = set(known_words) | line_words
    new_tokens: list[str] = []
    for idx, token in enumerate(tokens):
        if "-" in token:
            new_tokens.append(token)
            continue
        alpha = re.sub(r"[^A-Za-z]", "", token)
        if not alpha or alpha != alpha.lower() or len(alpha) < 6 or idx == 0:
            new_tokens.append(token)
            continue
        if known_words.get(alpha, 0) > 1:
            new_tokens.append(token)
            continue
        split = _find_compound_split(alpha, candidates, allow_short=True)
        if not split:
            new_tokens.append(token)
            continue
        prefix, suffix = split
        new_tokens.append(token[: len(prefix)])
        new_tokens.append(token[len(prefix) :])
    return " ".join(new_tokens)


def _find_compound_split(
    token: str, candidates: set[str], allow_short: bool = False
) -> tuple[str, str] | None:
    for idx in range(len(token) - 3, 1, -1):
        prefix = token[:idx]
        suffix = token[idx:]
        if len(prefix) < 3 and not allow_short:
            continue
        if len(prefix) < 2:
            continue
        if prefix in candidates and suffix.isalpha() and len(suffix) >= 3:
            return prefix, suffix
    return None


def _normalize_phrasing(text: str) -> str:
    text = re.sub(r"\b([A-Za-z]{4,})s(?=\s+for\b)", r"\1", text)
    text = re.sub(
        r"\b(last|previous|prior)\s+([A-Za-z]{3,})ed\b",
        r"\1 \2",
        text,
        flags=re.IGNORECASE,
    )
    return text


def _looks_like_unit_line(line: str) -> bool:
    if re.search(r"\d", line):
        return False
    if re.search(r"[.;:]", line):
        return False
    if re.search(r"[^A-Za-z\\s&]", line):
        return True
    return False


def _looks_like_section_break(line: str) -> bool:
    if re.search(r"[.;:]", line):
        return False
    if len(line) > 50:
        return False
    return 0 < len(line.split()) <= 6


def _next_non_empty_line(lines: list[str], start_index: int) -> str:
    for idx in range(start_index + 1, len(lines)):
        if lines[idx]:
            return lines[idx]
    return ""


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
