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
import statistics
import unicodedata
from typing import Any, Generator, Iterable, Optional, Protocol

from pypdf import PdfReader
from pypdf.generic import ContentStream

from sharepoint2text.exceptions import ExtractionFileEncryptedError
from sharepoint2text.extractors.data_types import (
    PdfContent,
    PdfImage,
    PdfMetadata,
    PdfPage,
)

logger = logging.getLogger(__name__)

# =============================================================================
# Type Aliases
# =============================================================================
TableRows = list[list[str]]
TextSegment = tuple[float, float, str, float]  # (y, x, text, font_size)

# =============================================================================
# PDF Filter to Image Format Mapping
# =============================================================================
# Maps PDF compression filter names to image format identifiers
FILTER_TO_FORMAT: dict[str, str] = {
    "/DCTDecode": "jpeg",  # JPEG compression
    "/JPXDecode": "jp2",  # JPEG 2000 compression
    "/FlateDecode": "png",  # PNG-style deflate compression
    "/CCITTFaxDecode": "tiff",  # TIFF Group 3/4 fax compression
    "/JBIG2Decode": "jbig2",  # JBIG2 bi-level compression
    "/LZWDecode": "png",  # LZW compression (legacy)
}

# Maps PDF compression filter names to MIME content types
FILTER_TO_CONTENT_TYPE: dict[str, str] = {
    "/DCTDecode": "image/jpeg",
    "/JPXDecode": "image/jp2",
    "/FlateDecode": "image/png",
    "/CCITTFaxDecode": "image/tiff",
    "/JBIG2Decode": "image/jbig2",
    "/LZWDecode": "image/png",
}


class PageLike(Protocol):
    def extract_text(self, *args: Any, **kwargs: Any) -> str: ...

    def get(self, key: str, default: Any = None) -> Any: ...

    def get_contents(self) -> Any: ...

    @property
    def pdf(self) -> Any: ...


def read_pdf(
    file_like: io.BytesIO, path: Optional[str] = None
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
    if reader.is_encrypted:
        raise ExtractionFileEncryptedError("PDF is encrypted or password-protected")
    logger.debug("Parsing PDF with %d pages", len(reader.pages))

    pages = []
    total_images = 0
    total_tables = 0
    for page_num, page in enumerate(reader.pages, start=1):
        images = _extract_image_bytes(page, page_num)
        total_images += len(images)
        page_text = page.extract_text() or ""
        raw_lines = page_text.splitlines()
        spatial_lines = _extract_lines_with_spacing(page)
        raw_tables = _TableExtractor.extract(raw_lines)
        spatial_tables = _TableExtractor.extract(spatial_lines)
        tables = _TableExtractor.choose_tables(raw_tables, spatial_tables)
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


def _extract_image_bytes(page: PageLike, page_num: int) -> list[PdfImage]:
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
    found_images: list[PdfImage] = []
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


def _extract_lines_with_spacing(page: PageLike) -> list[str]:
    """
    Extract text lines from a PDF page with spatial awareness.

    Uses the text visitor API to capture text positioning, then reconstructs
    lines based on vertical position and adds spacing between text segments
    based on horizontal gaps.

    This is more accurate than simple text extraction for:
        - Preserving column layouts
        - Detecting paragraph breaks
        - Separating table data from prose

    Returns:
        List of text lines with appropriate spacing.
    """
    # Collect text segments with position information
    segments: list[TextSegment] = []

    def visitor(
        text: str,
        _cm: Any,
        tm: Iterable[float],
        _font_dict: Any,
        font_size: Any,
    ) -> None:
        """Callback to capture text with its transformation matrix."""
        if not text:
            return
        try:
            tm_list = list(tm)
            x = float(tm_list[4])  # Horizontal position
            y = float(tm_list[5])  # Vertical position
        except Exception:
            return
        size = float(font_size) if font_size else 0.0
        segments.append((y, x, text, size))

    # Try spatial extraction, fall back to simple extraction on failure
    try:
        page.extract_text(visitor_text=visitor)
    except Exception:
        raw_text = page.extract_text() or ""
        return raw_text.splitlines()

    if not segments:
        raw_text = page.extract_text() or ""
        return raw_text.splitlines()

    # Calculate dynamic tolerances based on median font size
    font_sizes = [size for _, _, _, size in segments if size > 0]
    median_size = float(statistics.median(font_sizes)) if font_sizes else 10.0
    line_tolerance = max(1.0, median_size * 0.5)  # Y-distance for same line
    word_gap_threshold = max(1.0, median_size * 0.4)  # X-distance for word break

    # Sort segments by Y (descending) then X (ascending) for reading order
    segments.sort(key=lambda item: (-item[0], item[1]))

    # Check if all segments are on approximately the same line (degenerate case)
    y_values = [item[0] for item in segments]
    if max(y_values) - min(y_values) < line_tolerance:
        raw_text = page.extract_text() or ""
        return raw_text.splitlines()

    # Group segments into lines based on vertical proximity
    lines: list[dict[str, Any]] = []
    for y, x, text, size in segments:
        if not lines or abs(y - lines[-1]["y"]) > line_tolerance:
            lines.append({"y": y, "segments": [(x, text)]})
            continue
        lines[-1]["segments"].append((x, text))

    # Reconstruct line text with appropriate word spacing
    line_texts: list[str] = []
    line_positions: list[float] = []
    for line in lines:
        line_positions.append(line["y"])
        parts = []
        last_x = None
        for x, text in sorted(line["segments"], key=lambda item: item[0]):
            # Add space if there's a significant horizontal gap
            if last_x is not None and x - last_x > word_gap_threshold:
                # Don't add space if there's a hyphen at the break point
                if not (
                    (parts and parts[-1].endswith(("-", "–", "—")))
                    or text.startswith(("-", "–", "—"))
                ):
                    parts.append(" ")
            parts.append(text)
            last_x = x
        line_text = "".join(parts)
        line_texts.append(re.sub(r"\s+", " ", line_text).strip())

    if len(line_positions) < 2:
        return line_texts

    # Calculate paragraph break threshold based on typical line spacing
    gaps = [
        line_positions[idx] - line_positions[idx + 1]
        for idx in range(len(line_positions) - 1)
    ]
    median_gap = float(statistics.median(gaps)) if gaps else 0.0
    paragraph_gap_threshold = max(median_gap * 1.8, median_size * 1.2)

    # Track where to insert extra blank lines for paragraph breaks
    extra_after: dict[int, int] = {}
    last_non_empty: Optional[int] = None
    numeric_cache: dict[int, int] = {}

    for idx, text in enumerate(line_texts):
        if not text:
            continue
        numeric_cache[idx] = _TableExtractor.count_numeric_tokens(text)
        if last_non_empty is not None:
            gap_between = line_positions[last_non_empty] - line_positions[idx]

            # Check if this gap suggests a paragraph break
            if gap_between > paragraph_gap_threshold:
                prev_count = numeric_cache.get(last_non_empty, 0)
                next_count = numeric_cache.get(idx, 0)
                prev_numeric = prev_count > 0
                next_numeric = next_count > 0

                # Add break when transitioning from numeric to non-numeric content
                add_break = prev_numeric and not next_numeric

                # Also add break for large gaps with changing numeric density
                if (
                    not add_break
                    and prev_count >= 2
                    and next_count == 1
                    and gap_between > paragraph_gap_threshold * 1.6
                ):
                    add_break = True

                if add_break:
                    existing_empty = idx - last_non_empty - 1
                    needed = max(0, 3 - existing_empty)
                    if needed:
                        extra_after[last_non_empty] = max(
                            extra_after.get(last_non_empty, 0), needed
                        )
        last_non_empty = idx

    # Insert paragraph breaks as extra blank lines
    spaced_lines: list[str] = []
    for idx, text in enumerate(line_texts):
        spaced_lines.append(text)
        if idx in extra_after:
            spaced_lines.extend([""] * extra_after[idx])
    return spaced_lines


class _TableExtractor:
    """
    Extracts tabular data from PDF text lines.

    This class identifies and parses tables from extracted PDF text by analyzing
    patterns like date headers, numeric values, and column alignment. It uses
    heuristics to detect table boundaries and separate data from prose.

    The extraction process:
        1. Scan lines for date headers (MM/DD/YYYY patterns)
        2. Identify rows with trailing numeric values
        3. Track column counts for consistency
        4. Handle gaps between table rows
        5. Split compound words and normalize labels
    """

    # -------------------------------------------------------------------------
    # Regex Patterns for Table Detection
    # -------------------------------------------------------------------------
    # Date patterns (e.g., "12/31/2024")
    DATE_HEADER_PATTERN = re.compile(r"\d{2}/\d{2}/\d{4}")
    MONTH_PATTERN = re.compile(
        r"\b(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\b",
        re.IGNORECASE,
    )

    # Numeric detection patterns
    NUMERIC_RE = re.compile(r"\d")
    TRAILING_NUMBER_RE = re.compile(r"([\d.,]+)$")
    TRAILING_NUMBER_BLOCK_RE = re.compile(r"[\d.,]+$")
    LINE_NUMBER_PREFIX_RE = re.compile(r"^\d+\.")

    # Text structure patterns
    SECTION_BREAK_RE = re.compile(r"[.;:]")
    NON_UNIT_CHARS_RE = re.compile(r"[^A-Za-z\\s&]")

    # -------------------------------------------------------------------------
    # Extraction Configuration
    # -------------------------------------------------------------------------
    MAX_GAP = 2  # Maximum blank lines allowed within a table
    MIN_VALUE_COLUMNS = 2  # Minimum numeric columns required for a table row

    # -------------------------------------------------------------------------
    # Initialization
    # -------------------------------------------------------------------------
    def __init__(self, lines: list[str]) -> None:
        self.lines = [line.strip() for line in lines]
        self.known_words = self._collect_known_words(self.lines)

    # -------------------------------------------------------------------------
    # Public API
    # -------------------------------------------------------------------------
    @classmethod
    def extract(cls, lines: list[str]) -> list[TableRows]:
        """Extract all tables from the given text lines."""
        return cls(lines)._extract()

    @classmethod
    def choose_tables(
        cls, raw_tables: list[TableRows], spatial_tables: list[TableRows]
    ) -> list[TableRows]:
        """Choose the better table extraction result based on scoring."""
        if cls._score_tables(spatial_tables) > cls._score_tables(raw_tables):
            return spatial_tables
        return raw_tables

    # -------------------------------------------------------------------------
    # Core Extraction Logic
    # -------------------------------------------------------------------------
    def _extract(self) -> list[TableRows]:
        tables: list[TableRows] = []
        current_rows: TableRows = []
        column_count = 0
        gap_count = 0
        pending_header_label = ""

        for idx, line in enumerate(self.lines):
            if not line:
                if current_rows:
                    gap_count += 1
                    if gap_count > self.MAX_GAP:
                        self._flush_current(tables, current_rows)
                        current_rows = []
                        column_count = 0
                        gap_count = 0
                continue

            has_digits = bool(self.NUMERIC_RE.search(line))
            next_line = self._next_non_empty_line(idx)
            if current_rows and not has_digits and self._looks_like_section_break(line):
                next_header = self._extract_date_header(next_line)
                if next_header:
                    self._flush_current(tables, current_rows)
                    current_rows = []
                    column_count = 0
                    gap_count = 0
                    pending_header_label = self._normalize_label(line)
                    continue

            if not current_rows and not has_digits:
                pending_header_label = self._normalize_label(line)

            date_header = self._extract_date_header(line)
            if date_header:
                label, dates, unit_text = date_header
                if current_rows:
                    self._flush_current(tables, current_rows)
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
                        [self._normalize_label(unit_text)]
                        + ["" for _ in range(column_count - 1)]
                    )
                continue

            label, values = self._extract_row(line)
            has_values = bool(values)
            if has_values:
                pending_header_label = ""
            if has_values and self.MONTH_PATTERN.search(line):
                if values and values[-1].isdigit() and len(values[-1]) == 4:
                    values = []
                    has_values = False
            if not has_values and current_rows and current_rows[-1][0] == "":
                if not line[:1].isdigit() and not self._looks_like_unit_line(line):
                    self._flush_current(tables, current_rows)
                    current_rows = []
                    column_count = 0
                    gap_count = 0
                    continue
            if not has_values and current_rows:
                if self.LINE_NUMBER_PREFIX_RE.match(line):
                    self._flush_current(tables, current_rows)
                    current_rows = []
                    column_count = 0
                    gap_count = 0
                    continue
                if (
                    len(line) >= 60
                    and "." in line
                    and not line[:1].isdigit()
                    and not self.TRAILING_NUMBER_BLOCK_RE.search(line)
                ):
                    self._flush_current(tables, current_rows)
                    current_rows = []
                    column_count = 0
                    gap_count = 0
                    continue
            values_from_trailing_blob = False
            if not values:
                trailing_match = self.TRAILING_NUMBER_RE.search(line)
                if trailing_match and "/" not in trailing_match.group(1):
                    blob = trailing_match.group(1)
                    label = line[: trailing_match.start()].strip()
                    if label:
                        label = self._normalize_label(label)
                    values = self._split_numeric_blob(
                        blob, column_count - 1 if column_count else 2
                    )
                    values_from_trailing_blob = True
                if not values and not current_rows:
                    continue

            if values and len(values) < self.MIN_VALUE_COLUMNS and line[:1].isdigit():
                values = []
                label = self._normalize_label(line)

            if values and column_count == 0:
                if len(values) < self.MIN_VALUE_COLUMNS:
                    continue
                column_count = len(values) + 1

            if values or current_rows:
                if column_count == 0 and values:
                    if len(values) < self.MIN_VALUE_COLUMNS:
                        continue
                    column_count = len(values) + 1
                if column_count == 0:
                    continue
                expected_values = column_count - 1
                values = self._normalize_values(values, expected_values)
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

        self._flush_current(tables, current_rows)
        if tables:
            return tables
        return self._extract_tables_from_text_simple()

    # -------------------------------------------------------------------------
    # Label and Value Processing
    # -------------------------------------------------------------------------
    def _normalize_label(self, label: str) -> str:
        normalized = unicodedata.normalize("NFKC", label)
        normalized = re.sub(r"[\u2010-\u2013\u2212]", "-", normalized)
        normalized = re.sub(r"(?<=\\w)\\s*-\\s*(?=\\w)", "-", normalized)
        normalized = re.sub(r"\s+", " ", normalized).strip()
        return self._split_compound_words(normalized, self.known_words)

    def _extract_date_header(self, line: str) -> Optional[tuple[str, list[str], str]]:
        matches = list(self.DATE_HEADER_PATTERN.finditer(line))
        if len(matches) < 2:
            return None
        first = matches[0]
        second = matches[1]
        label = line[: first.start()].strip()
        label = self._normalize_label(label) if label else ""
        unit_text = line[second.end() :].strip()
        return label, [first.group(), second.group()], unit_text

    def _extract_row(self, line: str) -> tuple[str, list[str]]:
        tokens = line.split()
        values: list[str] = []
        idx = len(tokens) - 1
        while idx >= 0 and self.is_numeric_token(tokens[idx]):
            values.append(tokens[idx])
            idx -= 1
        values.reverse()
        label_tokens = tokens[: idx + 1]
        label = " ".join(label_tokens).strip()
        if label:
            label = self._strip_label_footnote(label)
            label = self._normalize_label(label)
        return label, values

    @staticmethod
    def _strip_label_footnote(label: str) -> str:
        cleaned = []
        for token in label.split():
            if token and token[-1].isdigit() and token[:-1].isalpha():
                cleaned.append(token[:-1])
            else:
                cleaned.append(token)
        return " ".join(cleaned).strip()

    @staticmethod
    def _is_footnote_leader(token: str) -> bool:
        return len(token) == 1 and token.isdigit()

    @classmethod
    def _normalize_values(cls, values: list[str], expected_count: int) -> list[str]:
        if not values or expected_count <= 0:
            return values
        if len(values) == expected_count + 1 and cls._is_footnote_leader(values[0]):
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

    # -------------------------------------------------------------------------
    # Numeric Detection and Processing
    # -------------------------------------------------------------------------
    @staticmethod
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

    @classmethod
    def count_numeric_tokens(cls, text: str) -> int:
        if not text:
            return 0
        return sum(1 for token in text.split() if cls.is_numeric_token(token))

    @staticmethod
    def _split_numeric_blob(blob: str, expected_count: int) -> list[str]:
        if expected_count != 2:
            return [blob]
        match = re.match(r"^(\d[\d,]*\.\d)(\d[\d,]*\.\d+)$", blob)
        if match:
            return [match.group(1), match.group(2)]
        return [blob]

    # -------------------------------------------------------------------------
    # Table Management
    # -------------------------------------------------------------------------
    @staticmethod
    def _flush_current(tables: list[TableRows], current_rows: TableRows) -> None:
        """Save completed table rows if they meet minimum requirements."""
        if len(current_rows) >= 2:
            tables.append(current_rows.copy())

    def _extract_tables_from_text_simple(self) -> list[TableRows]:
        tables: list[TableRows] = []
        current_rows: TableRows = []
        current_cols = 0

        def flush_current() -> None:
            if len(current_rows) >= 2:
                tables.append(current_rows.copy())

        for line in self.lines:
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

    # -------------------------------------------------------------------------
    # Word and Text Analysis
    # -------------------------------------------------------------------------
    @staticmethod
    def _collect_known_words(lines: list[str]) -> dict[str, int]:
        words: dict[str, int] = {}
        for line in lines:
            for token in re.findall(r"[A-Za-z]+", line):
                key = token.lower()
                words[key] = words.get(key, 0) + 1
        return words

    @staticmethod
    def _score_tables(tables: list[TableRows]) -> int:
        score = 0
        table_penalty = 2
        for table in tables:
            if not table:
                continue
            score -= table_penalty
            for row in table:
                non_empty = sum(1 for cell in row if str(cell).strip())
                if non_empty >= 2:
                    score += 2
                elif non_empty == 1:
                    score -= 1
        return score

    @classmethod
    def _split_compound_words(cls, text: str, known_words: dict[str, int]) -> str:
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
            split = cls._find_compound_split(alpha, candidates, allow_short=True)
            if not split:
                new_tokens.append(token)
                continue
            prefix, suffix = split
            new_tokens.append(token[: len(prefix)])
            new_tokens.append(token[len(prefix) :])
        return " ".join(new_tokens)

    @staticmethod
    def _find_compound_split(
        token: str, candidates: set[str], allow_short: bool = False
    ) -> Optional[tuple[str, str]]:
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

    # -------------------------------------------------------------------------
    # Line Classification
    # -------------------------------------------------------------------------
    @classmethod
    def _looks_like_unit_line(cls, line: str) -> bool:
        if cls.NUMERIC_RE.search(line):
            return False
        if cls.SECTION_BREAK_RE.search(line):
            return False
        if cls.NON_UNIT_CHARS_RE.search(line):
            return True
        return False

    @classmethod
    def _looks_like_section_break(cls, line: str) -> bool:
        if cls.SECTION_BREAK_RE.search(line):
            return False
        if len(line) > 50:
            return False
        return 0 < len(line.split()) <= 6

    def _next_non_empty_line(self, start_index: int) -> str:
        for idx in range(start_index + 1, len(self.lines)):
            candidate = self.lines[idx]
            if candidate:
                return candidate
        return ""


def _extract_image(
    image_obj: Any,
    name: Any,
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

    # Determine image format based on compression filter
    filter_type = image_obj.get("/Filter", "")
    if isinstance(filter_type, list):
        filter_type = filter_type[-1] if filter_type else ""
    filter_type = str(filter_type)

    img_format = FILTER_TO_FORMAT.get(filter_type, "raw")
    content_type = FILTER_TO_CONTENT_TYPE.get(filter_type, "image/unknown")

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


# =============================================================================
# MCID (Marked Content ID) Extraction
# =============================================================================
# MCID is used in Tagged PDFs to associate content with structure elements.
# This enables features like accessibility (alt text) and logical ordering.


def _extract_page_mcid_data(
    page: PageLike,
) -> tuple[list[dict[str, Any]], list[int], dict[int, str]]:
    """
    Extract MCID (Marked Content Identifier) data from a PDF page.

    Parses the page content stream to find:
        - Image occurrences with their associated MCIDs
        - Text content associated with each MCID
        - Order in which MCIDs appear (for caption association)

    Args:
        page: A pypdf PageObject to extract MCID data from.

    Returns:
        Tuple of:
            - image_occurrences: List of dicts with 'name' and 'mcid' keys
            - mcid_order: List of MCIDs in document order
            - mcid_text: Dict mapping MCID to accumulated text content
    """
    contents = page.get_contents()
    if contents is None:
        return [], [], {}

    try:
        stream = ContentStream(contents, page.pdf)
    except Exception as e:
        logger.debug("Failed to parse content stream: %s", e)
        return [], [], {}

    # State tracking for nested marked content
    mcid_stack: list[int | None] = []  # Current MCID context
    actual_text_stack: list[str | None] = []  # ActualText overrides
    mcid_order: list[int] = []  # Order of MCID occurrences
    mcid_text: dict[int, str] = {}  # Text content per MCID
    image_occurrences: list[dict[str, Any]] = []

    for operands, operator in stream.operations:
        op = (
            operator.decode("utf-8", errors="ignore")
            if isinstance(operator, bytes)
            else operator
        )

        # BDC/BMC: Begin Marked Content (with/without properties)
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

        # EMC: End Marked Content
        if op == "EMC":
            if mcid_stack:
                mcid_stack.pop()
            if actual_text_stack:
                actual_text_stack.pop()
            continue

        # Do: Invoke XObject (images)
        if op == "Do":
            if not operands:
                continue
            current_mcid = mcid_stack[-1] if mcid_stack else None
            image_occurrences.append({"name": operands[0], "mcid": current_mcid})
            continue

        # Text operators: Tj, TJ, ', "
        if op in ("Tj", "TJ", "'", '"'):
            current_mcid = mcid_stack[-1] if mcid_stack else None
            if current_mcid is None:
                continue
            # Use ActualText if available (accessibility text)
            actual_text = actual_text_stack[-1] if actual_text_stack else None
            if actual_text:
                text = str(actual_text)
                actual_text_stack[-1] = None  # Only use once
            else:
                text = _extract_text_from_operands(op, operands)
            if text:
                mcid_text[current_mcid] = mcid_text.get(current_mcid, "") + text
                if current_mcid not in mcid_order:
                    mcid_order.append(current_mcid)

    return image_occurrences, mcid_order, mcid_text


def _extract_text_from_operands(operator: str, operands: list[Any]) -> str:
    """Extract text string from PDF text operator operands."""
    if not operands:
        return ""
    if operator == "TJ":
        # TJ operator: array of strings and positioning values
        parts = []
        for item in operands[0]:
            if isinstance(item, (str, bytes)):
                parts.append(_normalize_text(item))
        return "".join(parts)
    if isinstance(operands[0], (str, bytes)):
        return _normalize_text(operands[0])
    return ""


def _normalize_text(value: Any) -> str:
    """Convert PDF text value to Python string."""
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
    """
    Find caption text for an image based on its MCID.

    Looks for text in the same MCID or the next MCID in document order.
    This captures captions that immediately follow images.
    """
    if mcid is None:
        return ""
    # Check for text in the same MCID
    text = mcid_text.get(mcid, "").strip()
    if text:
        return text
    if not mcid_order:
        return ""
    # Look for text in subsequent MCIDs (caption after image)
    try:
        start_index = mcid_order.index(mcid)
    except ValueError:
        return ""
    for next_mcid in mcid_order[start_index + 1 :]:
        next_text = mcid_text.get(next_mcid, "").strip()
        if next_text:
            return next_text
    return ""


def _extract_image_alt_text(image_obj: Any) -> str:
    """
    Extract alt text or title from a PDF image XObject.

    Checks standard accessibility attributes in order of preference:
        - /Alt: Alternative text (most common)
        - /Title: Image title
        - /Caption: Caption text
        - /TU: Tool tip (user-facing description)
    """
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
