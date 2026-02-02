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
    >>> from sharepoint2text.parsing.extractors.pdf.pdf_extractor import read_pdf
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

import contextlib
import io
import logging
import re
import statistics
import struct
from typing import Any, Callable, Generator, Iterable, Optional, Protocol

from pypdf import PdfReader
from pypdf.errors import DependencyError
from pypdf.generic import ContentStream

from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    PdfContent,
    PdfImage,
    PdfMetadata,
    PdfPage,
)
from sharepoint2text.parsing.extractors.pdf._pypdf_aes_fallback import (
    patch_pypdf_fallback_aes,
)

logger = logging.getLogger(__name__)

# PERFORMANCE OPTIMIZATION: Font analysis cache to avoid repeated TTF parsing
_FONT_CACHE: dict[bytes, Optional[tuple[int, dict[int, tuple[int, int]]]]] = {}

# PERFORMANCE OPTIMIZATION: Pre-compiled regex patterns for better performance
_DATE_HEADER_PATTERN = re.compile(r"\d{2}/\d{2}/\d{4}")
_MONTH_PATTERN = re.compile(
    r"\b(January|February|March|April|May|June|July|August|"
    r"September|October|November|December)\b",
    re.IGNORECASE,
)
_NUMERIC_RE = re.compile(r"\d")
_TRAILING_NUMBER_RE = re.compile(r"([\d.,]+)$")
_TRAILING_NUMBER_BLOCK_RE = re.compile(r"[\d.,]+$")
_LINE_NUMBER_PREFIX_RE = re.compile(r"^\d+\.")
_SECTION_BREAK_RE = re.compile(r"[.;:]")
_NON_UNIT_CHARS_RE = re.compile(r"[^A-Za-z\\s&]")
_DASH_RANGE_RE = re.compile(r"[\u2010-\u2013\u2212]")
_WHITESPACE_RE = re.compile(r"\s+")

# =============================================================================
# Type Aliases
# =============================================================================
TextSegment = tuple[float, float, str, float]  # (y, x, text, font_size)

# Reference digit bounding boxes (width, height) for a common sans font.
# Values are in font units (units-per-em = 2048).
_REFERENCE_DIGIT_UNITS_PER_EM = 2048
_REFERENCE_DIGIT_FEATURES: dict[int, tuple[int, int]] = {
    0: (956, 1497),
    1: (540, 1472),
    2: (971, 1472),
    3: (960, 1498),
    4: (1014, 1466),
    5: (972, 1471),
    6: (968, 1497),
    7: (949, 1447),
    8: (966, 1497),
    9: (964, 1497),
}

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

_AES_FALLBACK_IMAGE_SKIP_THRESHOLD_BYTES = 10 * 1024 * 1024


class PageLike(Protocol):
    def extract_text(self, *args: Any, **kwargs: Any) -> str: ...

    def get(self, key: str, default: Any = None) -> Any: ...

    def get_contents(self) -> Any: ...

    @property
    def pdf(self) -> Any: ...


def _open_pdf_reader(file_like: io.BytesIO) -> PdfReader:
    file_like.seek(0)
    try:
        return PdfReader(file_like)
    except DependencyError as exc:
        if "AES algorithm" not in str(exc):
            raise
        if not patch_pypdf_fallback_aes():
            raise
        file_like.seek(0)
        return PdfReader(file_like)


def _should_skip_images(reader: PdfReader, file_like: io.BytesIO) -> bool:
    if not reader.is_encrypted:
        return False
    try:
        import pypdf._crypt_providers as providers
    except Exception:
        return False
    if providers.crypt_provider[0] != "local_crypt_fallback":
        return False
    try:
        data_size = file_like.getbuffer().nbytes
    except Exception:
        return False
    return data_size >= _AES_FALLBACK_IMAGE_SKIP_THRESHOLD_BYTES


def _ttf_read_table_directory(
    font_data: bytes,
) -> dict[str, tuple[int, int]]:
    if len(font_data) < 12:
        return {}
    num_tables = struct.unpack(">H", font_data[4:6])[0]
    tables: dict[str, tuple[int, int]] = {}
    for idx in range(num_tables):
        entry = font_data[12 + idx * 16 : 12 + (idx + 1) * 16]
        if len(entry) < 16:
            break
        tag, _check, offset, length = struct.unpack(">4sIII", entry)
        tables[tag.decode("ascii", errors="ignore")] = (offset, length)
    return tables


def _ttf_read_head(
    font_data: bytes, tables: dict[str, tuple[int, int]]
) -> Optional[tuple[int, int]]:
    if "head" not in tables:
        return None
    offset, length = tables["head"]
    if offset + 52 > len(font_data):
        return None
    head = font_data[offset : offset + length]
    units_per_em = struct.unpack(">H", head[18:20])[0]
    index_to_loc_format = struct.unpack(">h", head[50:52])[0]
    return units_per_em, index_to_loc_format


def _ttf_read_maxp(
    font_data: bytes, tables: dict[str, tuple[int, int]]
) -> Optional[int]:
    if "maxp" not in tables:
        return None
    offset, length = tables["maxp"]
    if offset + 6 > len(font_data):
        return None
    maxp = font_data[offset : offset + length]
    return struct.unpack(">H", maxp[4:6])[0]


def _ttf_read_loca(
    font_data: bytes,
    tables: dict[str, tuple[int, int]],
    index_to_loc_format: int,
    num_glyphs: int,
) -> Optional[list[int]]:
    if "loca" not in tables:
        return None
    offset, length = tables["loca"]
    if offset >= len(font_data):
        return None
    loca = font_data[offset : offset + length]
    offsets: list[int] = []
    if index_to_loc_format == 0:
        for idx in range(num_glyphs + 1):
            start = idx * 2
            end = start + 2
            if end > len(loca):
                return None
            val = struct.unpack(">H", loca[start:end])[0]
            offsets.append(val * 2)
    else:
        for idx in range(num_glyphs + 1):
            start = idx * 4
            end = start + 4
            if end > len(loca):
                return None
            val = struct.unpack(">I", loca[start:end])[0]
            offsets.append(val)
    return offsets


def _ttf_read_glyph_dimensions(
    font_data: bytes, tables: dict[str, tuple[int, int]], glyph_offset: int
) -> Optional[tuple[int, int]]:
    if "glyf" not in tables:
        return None
    offset, length = tables["glyf"]
    start = offset + glyph_offset
    if start + 10 > offset + length or start + 10 > len(font_data):
        return None
    header = font_data[start : start + 10]
    if len(header) < 10:
        return None
    _contours, x_min, y_min, x_max, y_max = struct.unpack(">hhhhh", header)
    return x_max - x_min, y_max - y_min


def _ttf_get_glyph_features(
    font_data: bytes, glyph_ids: list[int]
) -> Optional[tuple[int, dict[int, tuple[int, int]]]]:
    # PERFORMANCE OPTIMIZATION: Use cached font analysis
    if font_data in _FONT_CACHE:
        return _FONT_CACHE[font_data]

    tables = _ttf_read_table_directory(font_data)
    head = _ttf_read_head(font_data, tables)
    num_glyphs = _ttf_read_maxp(font_data, tables)
    if head is None or num_glyphs is None:
        _FONT_CACHE[font_data] = None
        return None
    units_per_em, index_to_loc_format = head
    offsets = _ttf_read_loca(font_data, tables, index_to_loc_format, num_glyphs)
    if offsets is None:
        _FONT_CACHE[font_data] = None
        return None
    features: dict[int, tuple[int, int]] = {}
    for gid in glyph_ids:
        if gid < 0 or gid >= len(offsets) - 1:
            continue
        dims = _ttf_read_glyph_dimensions(font_data, tables, offsets[gid])
        if dims is None:
            continue
        features[gid] = dims

    result = (units_per_em, features)
    _FONT_CACHE[font_data] = result
    return result


def _assign_digit_glyphs(
    features: dict[int, tuple[int, int]], units_per_em: int
) -> dict[int, int]:
    if not features:
        return {}
    scale = units_per_em / _REFERENCE_DIGIT_UNITS_PER_EM
    ref_scaled = {
        digit: (width * scale, height * scale)
        for digit, (width, height) in _REFERENCE_DIGIT_FEATURES.items()
    }
    candidates: list[tuple[float, int, int]] = []
    for gid, (width, height) in features.items():
        for digit, (ref_width, ref_height) in ref_scaled.items():
            cost = abs(width - ref_width) + abs(height - ref_height)
            candidates.append((cost, gid, digit))
    candidates.sort()
    assigned: dict[int, int] = {}
    used_digits: set[int] = set()
    for _cost, gid, digit in candidates:
        if gid in assigned or digit in used_digits:
            continue
        assigned[gid] = digit
        used_digits.add(digit)
        if len(assigned) == len(features):
            break
    return assigned


def _patch_font_digit_map(
    font_map: dict[Any, Any], font_dict: Optional[dict[str, Any]]
) -> None:
    if not font_dict:
        return
    null_char = chr(0)
    null_keys = [
        ord(key)
        for key, value in font_map.items()
        if isinstance(key, str)
        and isinstance(value, str)
        and value == null_char
        and len(key) == 1
    ]
    if not null_keys:
        return
    font_data: Optional[bytes] = None
    try:
        desc = font_dict["/DescendantFonts"][0].get_object()
        font_desc = desc["/FontDescriptor"].get_object()
        font_data = font_desc["/FontFile2"].get_object().get_data()
    except Exception:
        font_data = None
    if not font_data:
        return
    glyph_info = _ttf_get_glyph_features(font_data, null_keys)
    if glyph_info is None:
        return
    units_per_em, features = glyph_info
    digit_map = _assign_digit_glyphs(features, units_per_em)
    if not digit_map:
        return
    for gid, digit in digit_map.items():
        try:
            font_map[chr(gid)] = str(digit)
        except ValueError:
            continue


def _get_pypdf_char_map_patcher() -> (
    tuple[list[tuple[Any, str]], "Callable[[Any], Callable[..., Any]]"]
):
    """
    Get the appropriate pypdf module(s) and patcher for character map building.

    pypdf 6.6.0+ removed build_char_map from _page and uses get_encoding from _cmap.
    This function detects the API version and returns the appropriate patching targets.

    Note: In pypdf 6.6.0+, get_encoding is imported into _font module, so we need
    to patch both _cmap and _font to ensure the patch takes effect.

    Returns:
        Tuple of ([(module, function_name), ...], wrapper_factory) for patching.
    """
    # Try pypdf 6.6.0+ API first (get_encoding in _cmap and _font)
    try:
        import pypdf._cmap as pypdf_cmap
        import pypdf._font as pypdf_font

        if hasattr(pypdf_cmap, "get_encoding") and hasattr(pypdf_font, "get_encoding"):

            def make_wrapper(
                original: Callable[..., Any],
            ) -> Callable[..., Any]:
                def patched(ft: Any) -> Any:
                    encoding, char_map = original(ft)
                    # In the new API, we get the font dict directly
                    _patch_font_digit_map(
                        char_map, ft if isinstance(ft, dict) else None
                    )
                    return encoding, char_map

                return patched

            # Need to patch both modules since _font imports get_encoding
            return [
                (pypdf_cmap, "get_encoding"),
                (pypdf_font, "get_encoding"),
            ], make_wrapper
    except (ImportError, AttributeError):
        pass

    # Fall back to pypdf < 6.6.0 API (build_char_map in _page)
    import pypdf._page as pypdf_page

    if not hasattr(pypdf_page, "build_char_map"):
        raise AttributeError(
            "pypdf version not supported: neither get_encoding nor build_char_map found"
        )

    def make_wrapper(original: Callable[..., Any]) -> Callable[..., Any]:
        def patched(font_name: str, space_width: float, obj: Any) -> Any:
            font_subtype, font_halfspace, font_encoding, font_map, font = original(
                font_name, space_width, obj
            )
            _patch_font_digit_map(font_map, font)
            return font_subtype, font_halfspace, font_encoding, font_map, font

        return patched

    return [(pypdf_page, "build_char_map")], make_wrapper


@contextlib.contextmanager
def _patched_build_char_map() -> Iterable[None]:
    """
    Context manager that patches pypdf's character map building function.

    This patches the font character map building to fix digit mappings in
    fonts that use null characters for digits. Supports both pypdf < 6.6.0
    (build_char_map) and pypdf >= 6.6.0 (get_encoding).
    """
    try:
        patch_targets, make_wrapper = _get_pypdf_char_map_patcher()
    except AttributeError as e:
        logger.warning("Cannot patch pypdf char map: %s", e)
        yield
        return

    # Store originals and apply patches
    originals: list[tuple[Any, str, Any]] = []
    for module, func_name in patch_targets:
        original = getattr(module, func_name)
        originals.append((module, func_name, original))
        patched = make_wrapper(original)
        setattr(module, func_name, patched)

    try:
        yield
    finally:
        # Restore all originals
        for module, func_name, original in originals:
            setattr(module, func_name, original)


def read_pdf(
    file_like: io.BytesIO, path: Optional[str] = None, *, ignore_images: bool = False
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
        ignore_images: If True, skip image extraction.

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
    try:
        reader = _open_pdf_reader(file_like)
        if reader.is_encrypted:
            try:
                decrypt_result = reader.decrypt("")
            except Exception:
                decrypt_result = 0
            if decrypt_result == 0:
                raise ExtractionFileEncryptedError(
                    "PDF is encrypted or password-protected"
                )
        logger.debug("Parsing PDF with %d pages", len(reader.pages))

        skip_images = ignore_images or _should_skip_images(reader, file_like)
        if skip_images and not ignore_images:
            logger.info(
                "Skipping image extraction for large AES-encrypted PDF using fallback crypto"
            )

        pages = []
        total_images = 0
        for page_num, page in enumerate(reader.pages, start=1):
            images = [] if skip_images else _extract_image_bytes(page, page_num)
            total_images += len(images)
            page_text, _spatial_lines = _extract_text_with_spacing(page)
            pages.append(
                PdfPage(
                    text=page_text,
                    images=images,
                    tables=[],
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
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract PDF file", cause=exc) from exc


def _extract_image_bytes(page: PageLike, page_num: int) -> list[PdfImage]:
    """
    Extract all images from a PDF page's XObject resources.

    Iterates through the page's /Resources/XObject dictionary, identifies
    image objects by their /Subtype attribute, and extracts each image's
    binary data and properties.

    Args:
        page: A pypdf PageObject to extract images from.
        page_num: 1-based page number for image metadata.

    Returns:
        List of PdfImage objects for successfully extracted images.
        Failed extractions are logged and skipped.
    """
    resources = page.get("/Resources", {})
    if "/XObject" not in resources:
        return []

    x_objects = resources["/XObject"].get_object()
    image_occurrences, mcid_order, mcid_text = _extract_page_mcid_data(page)

    # Build list of (obj_name, obj, caption) tuples to extract
    candidates: list[tuple[Any, Any, str]] = []

    if image_occurrences:
        # Use MCID data for document order and captions
        for occurrence in image_occurrences:
            obj_name = occurrence["name"]
            obj = x_objects.get(obj_name)
            if obj is None or obj.get("/Subtype") != "/Image":
                continue
            caption = _lookup_caption(occurrence.get("mcid"), mcid_order, mcid_text)
            candidates.append((obj_name, obj, caption))
    else:
        # Fall back to XObject dictionary order
        for obj_name in x_objects:
            obj = x_objects[obj_name]
            if obj.get("/Subtype") == "/Image":
                candidates.append((obj_name, obj, ""))

    # Extract images from candidates
    found_images: list[PdfImage] = []
    for image_index, (obj_name, obj, caption) in enumerate(candidates, start=1):
        try:
            image_data = _extract_image(obj, obj_name, image_index, page_num, caption)
            found_images.append(image_data)
        except Exception as e:
            logger.warning(
                "Failed to extract image [%s] [%d]: %s", obj_name, image_index, e
            )

    return found_images


def _extract_text_with_spacing(page: PageLike) -> tuple[str, list[str]]:
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
        Tuple of:
            - raw page text (string)
            - List of text lines with appropriate spacing.
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
    with _patched_build_char_map():
        try:
            page_text = page.extract_text(visitor_text=visitor) or ""
        except Exception:
            page_text = page.extract_text() or ""
            return page_text, page_text.splitlines()

    if not segments:
        return page_text, page_text.splitlines()

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
        return page_text, page_text.splitlines()

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
        return page_text, line_texts

    return page_text, line_texts


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
        image_index=index,
        name=str(name),
        caption=resolved_caption,
        width=int(width),
        height=int(height),
        color_space=color_space,
        bits_per_component=int(bits),
        filter=filter_type,
        data=io.BytesIO(data) if data else None,
        format=img_format,
        content_type=content_type,
        unit_number=page_num,
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
