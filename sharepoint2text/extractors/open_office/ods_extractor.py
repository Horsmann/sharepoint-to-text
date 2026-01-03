"""
ODS Spreadsheet Extractor
=========================

Extracts text content, metadata, and structure from OpenDocument Spreadsheet
(.ods) files created by LibreOffice Calc, OpenOffice, and other ODF-compatible
applications.

File Format Background
----------------------
ODS files are ZIP archives containing XML files following the OASIS OpenDocument
specification (ISO/IEC 26300). Key components:

    content.xml: Spreadsheet data (sheets, rows, cells)
    meta.xml: Metadata (title, author, dates)
    styles.xml: Cell and table styles
    Pictures/: Embedded images

Spreadsheet Structure in content.xml:
    - office:document-content: Root element
    - office:body: Container for content
    - office:spreadsheet: Spreadsheet body
    - table:table: Individual sheets
    - table:table-row: Row containers
    - table:table-cell: Cell containers with values

Cell Value Types
----------------
ODS cells have explicit type information via office:value-type:
    - float: Numeric values (office:value attribute)
    - currency: Currency values (office:value + office:currency)
    - percentage: Percentage values (office:value)
    - date: Date values (office:date-value in ISO format)
    - time: Time values (office:time-value in ISO duration)
    - boolean: Boolean values (office:boolean-value)
    - string: Text values (content in text:p element)

The extractor prioritizes typed values when available, falling back
to text content for display values.

Row and Cell Repetition
-----------------------
ODS uses repetition attributes for efficient storage:
    - table:number-rows-repeated: Repeat row N times
    - table:number-columns-repeated: Repeat cell N times

The extractor handles these to avoid extracting huge empty areas
while preserving actual data.

Column Naming
-------------
Columns are named using Excel-style letters (A, B, ..., Z, AA, AB, ...).
The _get_column_name function converts 0-based indices to letter names.

Dependencies
------------
Python Standard Library only:
    - zipfile: ZIP archive handling
    - xml.etree.ElementTree: XML parsing
    - mimetypes: Image content type detection

Extracted Content
-----------------
Per-sheet content includes:
    - name: Sheet name
    - data: List of dicts with column-letter keys (A, B, C, ...)
    - text: Tab-separated text representation
    - annotations: Cell comments with creator and date
    - images: Embedded images in the sheet

Data Structure:
    Each row is a dictionary with column names as keys:
    {"A": "value1", "B": "value2", "C": "value3"}

Known Limitations
-----------------
- Formulas are not extracted (only calculated values)
- Charts are not extracted
- Named ranges are not reported
- Merged cells may have unexpected behavior
- Very large spreadsheets may use significant memory
- Password-protected files are not supported
- Pivot tables show only cached data

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.open_office.ods_extractor import read_ods
    >>>
    >>> with open("data.ods", "rb") as f:
    ...     for workbook in read_ods(io.BytesIO(f.read()), path="data.ods"):
    ...         print(f"Title: {workbook.metadata.title}")
    ...         for sheet in workbook.sheets:
    ...             print(f"Sheet: {sheet.name}")
    ...             print(f"Rows: {len(sheet.data)}")

See Also
--------
- odt_extractor: For OpenDocument Text files
- odp_extractor: For OpenDocument Presentation files
- xlsx_extractor: For Microsoft Excel files

Maintenance Notes
-----------------
- Cell repetition is handled but large repeats are limited to 1
- Empty rows and cells with repetition are not expanded
- Column naming uses standard Excel convention
- Images can be inline in cells or anchored to sheet
"""

import io
import logging
import mimetypes
from functools import lru_cache
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    OdsContent,
    OdsSheet,
    OpenDocumentAnnotation,
    OpenDocumentImage,
    OpenDocumentMetadata,
)
from sharepoint2text.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.extractors.util.zip_context import ZipContext

logger = logging.getLogger(__name__)

# ODF namespaces
NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
}

_TEXT_SPACE_TAG = f"{{{NS['text']}}}s"
_TEXT_TAB_TAG = f"{{{NS['text']}}}tab"
_TEXT_LINE_BREAK_TAG = f"{{{NS['text']}}}line-break"
_OFFICE_ANNOTATION_TAG = f"{{{NS['office']}}}annotation"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"
_ATTR_TABLE_NAME = f"{{{NS['table']}}}name"
_ATTR_TABLE_REPEAT_ROWS = f"{{{NS['table']}}}number-rows-repeated"
_ATTR_TABLE_REPEAT_COLS = f"{{{NS['table']}}}number-columns-repeated"

_ATTR_OFFICE_VALUE_TYPE = f"{{{NS['office']}}}value-type"
_ATTR_OFFICE_VALUE = f"{{{NS['office']}}}value"
_ATTR_OFFICE_DATE_VALUE = f"{{{NS['office']}}}date-value"
_ATTR_OFFICE_TIME_VALUE = f"{{{NS['office']}}}time-value"
_ATTR_OFFICE_BOOLEAN_VALUE = f"{{{NS['office']}}}boolean-value"

_ATTR_DRAW_NAME = f"{{{NS['draw']}}}name"
_ATTR_SVG_WIDTH = f"{{{NS['svg']}}}width"
_ATTR_SVG_HEIGHT = f"{{{NS['svg']}}}height"
_ATTR_XLINK_HREF = f"{{{NS['xlink']}}}href"


@lru_cache(maxsize=512)
def _guess_content_type(path: str) -> str:
    return mimetypes.guess_type(path)[0] or "application/octet-stream"


class _OdsContext(ZipContext):
    """Cached context for ODS extraction."""

    def __init__(self, file_like: io.BytesIO):
        super().__init__(file_like)
        self._content_root: ET.Element | None = (
            self.read_xml_root("content.xml") if self.exists("content.xml") else None
        )
        self._meta_root: ET.Element | None = (
            self.read_xml_root("meta.xml") if self.exists("meta.xml") else None
        )

    @property
    def content_root(self) -> ET.Element | None:
        return self._content_root

    @property
    def meta_root(self) -> ET.Element | None:
        return self._meta_root


def _get_text_recursive(element: ET.Element) -> str:
    """Recursively extract all text from an element and its children."""
    parts: list[str] = []

    text = element.text
    if text:
        parts.append(text)

    for child in element:
        tag = child.tag

        if tag == _TEXT_SPACE_TAG:
            count = int(child.get(_ATTR_TEXT_C, "1"))
            parts.append(" " * count)
        elif tag == _TEXT_TAB_TAG:
            parts.append("\t")
        elif tag == _TEXT_LINE_BREAK_TAG:
            parts.append("\n")
        elif tag == _OFFICE_ANNOTATION_TAG:
            # Skip annotations in main text extraction.
            pass
        else:
            parts.append(_get_text_recursive(child))

        tail = child.tail
        if tail:
            parts.append(tail)

    return "".join(parts)


def _extract_metadata(meta_root: ET.Element | None) -> OpenDocumentMetadata:
    """Extract metadata from meta.xml."""
    logger.debug("Extracting ODS metadata")
    metadata = OpenDocumentMetadata()

    if meta_root is None:
        return metadata

    meta_elem = meta_root.find(".//office:meta", NS)
    if meta_elem is None:
        return metadata

    # Extract Dublin Core elements
    title = meta_elem.find("dc:title", NS)
    if title is not None and title.text:
        metadata.title = title.text

    description = meta_elem.find("dc:description", NS)
    if description is not None and description.text:
        metadata.description = description.text

    subject = meta_elem.find("dc:subject", NS)
    if subject is not None and subject.text:
        metadata.subject = subject.text

    creator = meta_elem.find("dc:creator", NS)
    if creator is not None and creator.text:
        metadata.creator = creator.text

    date = meta_elem.find("dc:date", NS)
    if date is not None and date.text:
        metadata.date = date.text

    language = meta_elem.find("dc:language", NS)
    if language is not None and language.text:
        metadata.language = language.text

    # Extract meta elements
    keywords = meta_elem.find("meta:keyword", NS)
    if keywords is not None and keywords.text:
        metadata.keywords = keywords.text

    initial_creator = meta_elem.find("meta:initial-creator", NS)
    if initial_creator is not None and initial_creator.text:
        metadata.initial_creator = initial_creator.text

    creation_date = meta_elem.find("meta:creation-date", NS)
    if creation_date is not None and creation_date.text:
        metadata.creation_date = creation_date.text

    editing_cycles = meta_elem.find("meta:editing-cycles", NS)
    if editing_cycles is not None and editing_cycles.text:
        try:
            metadata.editing_cycles = int(editing_cycles.text)
        except ValueError:
            pass

    editing_duration = meta_elem.find("meta:editing-duration", NS)
    if editing_duration is not None and editing_duration.text:
        metadata.editing_duration = editing_duration.text

    generator = meta_elem.find("meta:generator", NS)
    if generator is not None and generator.text:
        metadata.generator = generator.text

    return metadata


def _extract_cell_value(cell: ET.Element) -> tuple[Any, str]:
    """Extract the value from a table cell.

    Returns:
        A tuple of (typed_value, display_text) where:
        - typed_value: The value in appropriate Python type (int, float, str, None)
        - display_text: The string representation for text output
    """
    # Check for value type attribute
    value_type = cell.get(_ATTR_OFFICE_VALUE_TYPE, "")

    # For numeric values, get the value attribute and convert to proper type
    if value_type in ("float", "currency", "percentage"):
        value = cell.get(_ATTR_OFFICE_VALUE, "")
        if value:
            try:
                float_val = float(value)
                # Return int if it's a whole number
                if float_val == int(float_val):
                    return int(float_val), value
                return float_val, value
            except ValueError:
                return value, value

    # For date/time values
    if value_type == "date":
        value = cell.get(_ATTR_OFFICE_DATE_VALUE, "")
        if value:
            return value, value

    if value_type == "time":
        value = cell.get(_ATTR_OFFICE_TIME_VALUE, "")
        if value:
            return value, value

    # For boolean values
    if value_type == "boolean":
        value = cell.get(_ATTR_OFFICE_BOOLEAN_VALUE, "")
        if value:
            return value.lower() == "true", value

    # For string values or fallback, get text from paragraphs
    text_parts: list[str] = []
    for p in cell.iterfind(".//text:p", NS):
        text_parts.append(_get_text_recursive(p))

    text = "\n".join(text_parts)
    if text:
        return text, text
    return None, ""


def _extract_annotations(cell: ET.Element) -> list[OpenDocumentAnnotation]:
    """Extract annotations/comments from a cell."""
    annotations = []

    for annotation in cell.findall(".//office:annotation", NS):
        creator_elem = annotation.find("dc:creator", NS)
        creator = (
            creator_elem.text if creator_elem is not None and creator_elem.text else ""
        )

        date_elem = annotation.find("dc:date", NS)
        date = date_elem.text if date_elem is not None and date_elem.text else ""

        text_parts = []
        for p in annotation.findall(".//text:p", NS):
            text_parts.append(_get_text_recursive(p))
        text = "\n".join(text_parts)

        annotations.append(
            OpenDocumentAnnotation(creator=creator, date=date, text=text)
        )

    return annotations


def _extract_images(
    ctx: _OdsContext,
    table: ET.Element,
    image_counter: int,
) -> tuple[list[OpenDocumentImage], int]:
    """Extract images from a table/sheet.

    Extracts images with their metadata:
    - caption: Always empty (ODS sheets don't have captions like ODT documents)
    - description: Combined from svg:title and svg:desc elements (with newline separator)
    - image_index: Sequential index of the image across all sheets
    - unit_index: None (ODS has sheets, not pages/slides)

    Args:
        ctx: Cached ZIP context for this spreadsheet.
        table: The table:table XML element for this sheet.
        image_counter: The current global image counter across all sheets.

    Returns:
        A tuple of (list of OpenDocumentImage, updated image_counter).
    """
    images: list[OpenDocumentImage] = []

    for frame in table.findall(".//draw:frame", NS):
        name = frame.get(_ATTR_DRAW_NAME, "")
        width = frame.get(_ATTR_SVG_WIDTH)
        height = frame.get(_ATTR_SVG_HEIGHT)

        # Extract title and description from frame
        # ODF uses svg:title and svg:desc elements for accessibility
        # In ODS, we combine title and desc into description (no caption support)
        title_elem = frame.find("svg:title", NS)
        title = title_elem.text if title_elem is not None and title_elem.text else ""

        desc_elem = frame.find("svg:desc", NS)
        desc = desc_elem.text if desc_elem is not None and desc_elem.text else ""

        # Combine title and description with newline separator
        if title and desc:
            description = f"{title}\n{desc}"
        else:
            description = title or desc

        # ODS sheets don't have captions like ODT documents
        caption = ""

        image_elem = frame.find("draw:image", NS)
        if image_elem is None:
            continue

        href = image_elem.get(_ATTR_XLINK_HREF, "")
        if not href:
            continue

        image_counter += 1

        if href.startswith("http"):
            # External image reference
            images.append(
                OpenDocumentImage(
                    href=href,
                    name=name,
                    width=width,
                    height=height,
                    image_index=image_counter,
                    caption=caption,
                    description=description,
                    unit_index=None,
                )
            )
        else:
            # Internal image reference
            try:
                if ctx.exists(href):
                    img_data = ctx.read_bytes(href)
                    images.append(
                        OpenDocumentImage(
                            href=href,
                            name=name or href.split("/")[-1],
                            content_type=_guess_content_type(href),
                            data=io.BytesIO(img_data),
                            size_bytes=len(img_data),
                            width=width,
                            height=height,
                            image_index=image_counter,
                            caption=caption,
                            description=description,
                            unit_index=None,
                        )
                    )
            except Exception as e:
                logger.debug("Failed to extract image %s: %s", href, e)
                images.append(
                    OpenDocumentImage(
                        href=href,
                        name=name or href,
                        error=str(e),
                        width=width,
                        height=height,
                        image_index=image_counter,
                        caption=caption,
                        description=description,
                        unit_index=None,
                    )
                )

    return images, image_counter


def _extract_sheet(
    ctx: _OdsContext,
    table: ET.Element,
    sheet_number: int,
    image_counter: int,
) -> tuple[OdsSheet, int]:
    """Extract content from a single sheet (table:table element).

    Args:
        z: The open zipfile containing the spreadsheet.
        table: The table:table XML element for this sheet.
        sheet_number: The 1-based sheet number.
        image_counter: The current global image counter across all sheets.

    Returns:
        A tuple of (OdsSheet, updated_image_counter).
    """
    sheet = OdsSheet()

    # Get sheet name
    sheet.name = table.get(_ATTR_TABLE_NAME, "")

    # First pass: collect all rows and find max column count
    raw_rows: list[list[tuple[Any, str]]] = []  # list of (typed_value, display_text)
    all_annotations = []
    max_cols = 0

    for row in table.findall("table:table-row", NS):
        row_values: list[tuple[Any, str]] = []

        # Check for repeated rows
        row_repeat = int(row.get(_ATTR_TABLE_REPEAT_ROWS, "1"))

        for cell in row.findall("table:table-cell", NS):
            # Check for repeated cells
            cell_repeat = int(cell.get(_ATTR_TABLE_REPEAT_COLS, "1"))

            typed_value, display_text = _extract_cell_value(cell)

            # Extract annotations from cell
            cell_annotations = _extract_annotations(cell)
            all_annotations.extend(cell_annotations)

            # Add value for each repeated cell (but limit large repeats for empty cells)
            if typed_value is None and cell_repeat > 100:
                # Large repeat of empty cells - just add one to track column position
                row_values.append((None, ""))
            else:
                row_values.extend([(typed_value, display_text)] * cell_repeat)

        # Add row for each repeated row (but limit to avoid huge empty areas)
        if row_repeat > 100 and all(v[0] is None for v in row_values):
            # Large repeat of empty rows - add just one
            raw_rows.append(row_values)
        else:
            raw_rows.extend([row_values] * row_repeat)

        if row_values:
            max_cols = max(max_cols, len(row_values))

    # Trim trailing empty rows
    while raw_rows and all(v[0] is None for v in raw_rows[-1]):
        raw_rows.pop()

    # Trim trailing empty columns
    if raw_rows:
        # Find the last column with any data
        last_data_col = 0
        for row in raw_rows:
            for i in range(len(row) - 1, -1, -1):
                if row[i][0] is not None:
                    last_data_col = max(last_data_col, i + 1)
                    break
        max_cols = last_data_col

    # Build final data as list of lists, padding rows to have same column count
    rows_data: list[list[Any]] = []
    text_lines = []

    for row_values in raw_rows:
        # Build row with proper padding
        row_data = []
        row_texts = []
        for i in range(max_cols):
            if i < len(row_values):
                typed_value, display_text = row_values[i]
                row_data.append(typed_value)
                if display_text:
                    row_texts.append(display_text)
            else:
                row_data.append(None)

        rows_data.append(row_data)
        if row_texts:
            text_lines.append("\t".join(row_texts))

    sheet.data = rows_data
    sheet.text = "\n".join(text_lines)
    sheet.annotations = all_annotations
    sheet.images, image_counter = _extract_images(ctx, table, image_counter)

    return sheet, image_counter


def read_ods(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdsContent, Any, None]:
    """
    Extract all relevant content from an OpenDocument Spreadsheet (.ods) file.

    Primary entry point for ODS file extraction. Opens the ZIP archive,
    parses content.xml and meta.xml, and extracts sheet data with cell
    values and annotations.

    This function uses a generator pattern for API consistency with other
    extractors, even though ODS files contain exactly one workbook.

    Args:
        file_like: BytesIO object containing the complete ODS file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned OdsContent.metadata.

    Yields:
        OdsContent: Single OdsContent object containing:
            - metadata: OpenDocumentMetadata with title, creator, dates
            - sheets: List of OdsSheet objects with per-sheet data

    Raises:
        ValueError: If content.xml is missing or spreadsheet body not found.

    Example:
        >>> import io
        >>> with open("budget.ods", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for workbook in read_ods(data, path="budget.ods"):
        ...         for sheet in workbook.sheets:
        ...             print(f"Sheet: {sheet.name}")
        ...             for row in sheet.data[:5]:
        ...                 print(row)
    """
    try:
        file_like.seek(0)
        if is_odf_encrypted(file_like):
            raise ExtractionFileEncryptedError("ODS is encrypted or password-protected")

        ctx = _OdsContext(file_like)
        try:
            metadata = _extract_metadata(ctx.meta_root)

            content_root = ctx.content_root
            if content_root is None:
                raise ExtractionFailedError("Invalid ODS file: content.xml not found")

            body = content_root.find(".//office:body/office:spreadsheet", NS)
            if body is None:
                raise ExtractionFailedError(
                    "Invalid ODS file: spreadsheet body not found"
                )

            sheets: list[OdsSheet] = []
            image_counter = 0
            for sheet_num, table in enumerate(body.findall("table:table", NS), start=1):
                sheet, image_counter = _extract_sheet(
                    ctx, table, sheet_num, image_counter
                )
                sheets.append(sheet)
        finally:
            ctx.close()

        # Populate file metadata from path
        metadata.populate_from_path(path)

        yield OdsContent(
            metadata=metadata,
            sheets=sheets,
        )
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract ODS file", cause=exc) from exc
