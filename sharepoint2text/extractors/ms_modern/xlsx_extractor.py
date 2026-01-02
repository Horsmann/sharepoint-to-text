"""
XLSX Spreadsheet Extractor
==========================

Extracts text content and metadata from Microsoft Excel .xlsx files
(Office Open XML format, Excel 2007 and later).

This module uses the openpyxl library for parsing cells, sheets, and
metadata from XLSX files.

File Format Background
----------------------
The .xlsx format is a ZIP archive containing XML files following the Office
Open XML (OOXML) standard. Key components:

    xl/workbook.xml: Workbook properties and sheet list
    xl/worksheets/sheet1.xml, sheet2.xml, ...: Individual sheet data
    xl/sharedStrings.xml: Shared string table (for cell text)
    xl/styles.xml: Cell formatting and styles
    docProps/core.xml: Metadata (title, creator, dates)

XML Namespaces:
    - spreadsheetml: http://schemas.openxmlformats.org/spreadsheetml/2006/main
    - r: http://schemas.openxmlformats.org/officeDocument/2006/relationships

Dependencies
------------
openpyxl: https://github.com/theorchard/openpyxl
    pip install openpyxl

    Provides:
    - Cell value reading with type detection
    - Sheet enumeration
    - Row/column iteration
    - Core properties (metadata)
    - Date handling and formatting
    - Embedded image extraction

Data Representation
-------------------
Each sheet is represented in two forms:

1. Structured data (list of dicts):
   - First row is treated as headers
   - Each subsequent row becomes a dictionary
   - Keys are header values (empty headers get "Unnamed: N")
   - Native Python types preserved (int, float, str, datetime)

2. Text representation:
   - Formatted as aligned text table
   - Columns right-aligned with consistent spacing
   - Suitable for display or text search

Data Type Handling
------------------
Cell values are converted to appropriate Python types:
    - None: Empty cell
    - str: Text content
    - int/float: Numeric values (integers displayed without decimals)
    - datetime: Converted to ISO format strings for JSON compatibility

The extractor uses openpyxl's data_only=True mode, which returns
calculated values for formula cells rather than the formulas themselves.

Row/Column Trimming
-------------------
Empty trailing rows and columns are automatically trimmed to avoid
processing large numbers of empty cells in sparse spreadsheets.

Known Limitations
-----------------
- Formulas are not extracted (only calculated values)
- Charts are not extracted
- Conditional formatting is not applied to text output
- Pivot tables show only cached data
- Very large spreadsheets may use significant memory
- Password-protected files are not supported
- Merged cells may have unexpected behavior

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_modern.xlsx_extractor import read_xlsx
    >>>
    >>> with open("data.xlsx", "rb") as f:
    ...     for workbook in read_xlsx(io.BytesIO(f.read()), path="data.xlsx"):
    ...         print(f"Creator: {workbook.metadata.creator}")
    ...         for sheet in workbook.sheets:
    ...             print(f"Sheet: {sheet.name}")
    ...             print(f"Rows: {len(sheet.data)}")
    ...             print(sheet.text[:200])

See Also
--------
- openpyxl documentation: https://openpyxl.readthedocs.io/
- xls_extractor: For legacy .xls format

Maintenance Notes
-----------------
- read_only=True mode is used for memory efficiency
- data_only=True returns calculated values, not formulas
- First row is always treated as headers (no auto-detection)
- Empty sheets produce empty data and text output
"""

import datetime
import io
import logging
from typing import Any, Generator

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    XlsxContent,
    XlsxImage,
    XlsxMetadata,
    XlsxSheet,
)
from sharepoint2text.extractors.util.encryption import is_ooxml_encrypted
from sharepoint2text.extractors.util.ooxml_context import OOXMLZipContext
from sharepoint2text.extractors.util.zip_bomb import validate_zip_bytesio
from sharepoint2text.extractors.util.zip_utils import (
    parse_relationships,
)

logger = logging.getLogger(__name__)

# XML namespaces used in XLSX drawings (pre-computed for speed)
XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Pre-computed tag names for hot paths
XDR_ONE_CELL_ANCHOR = f"{{{XDR_NS}}}oneCellAnchor"
XDR_TWO_CELL_ANCHOR = f"{{{XDR_NS}}}twoCellAnchor"
XDR_ABSOLUTE_ANCHOR = f"{{{XDR_NS}}}absoluteAnchor"
XDR_PIC = f"{{{XDR_NS}}}pic"
XDR_EXT = f"{{{XDR_NS}}}ext"
XDR_NVPICPR = f"{{{XDR_NS}}}nvPicPr"
XDR_CNVPR = f"{{{XDR_NS}}}cNvPr"
XDR_BLIPFILL = f"{{{XDR_NS}}}blipFill"
A_BLIP = f"{{{A_NS}}}blip"
R_EMBED = f"{{{R_NS}}}embed"

# Anchor types tuple for iteration
ANCHOR_TYPES = (XDR_ONE_CELL_ANCHOR, XDR_TWO_CELL_ANCHOR, XDR_ABSOLUTE_ANCHOR)

# EMUs to pixels conversion factor (9525 EMUs = 1 pixel at 96 DPI)
EMU_PER_PIXEL = 9525

# Content type mapping by file extension (cached at module level)
_CONTENT_TYPE_MAP = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "tif": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
}


def _read_metadata(file_like: io.BytesIO) -> XlsxMetadata:
    """
    Extract document metadata from the XLSX file's core properties.

    Uses openpyxl to access document properties stored in docProps/core.xml.

    Args:
        file_like: BytesIO containing the XLSX file.

    Returns:
        XlsxMetadata object with:
        - title, description, creator, keywords, language
        - last_modified_by
        - created, modified (ISO format dates)
        - revision number

    Notes:
        - Workbook is opened in read_only mode for efficiency
        - Workbook is closed after metadata extraction
    """
    file_like.seek(0)
    wb = load_workbook(file_like, read_only=True, data_only=True)
    props = wb.properties

    metadata = XlsxMetadata(
        title=props.title or "",
        description=props.description or "",
        creator=props.creator or "",
        last_modified_by=props.lastModifiedBy or "",
        created=(
            props.created.isoformat()
            if isinstance(props.created, datetime.datetime)
            else ""
        ),
        modified=(
            props.modified.isoformat()
            if isinstance(props.modified, datetime.datetime)
            else ""
        ),
        keywords=props.keywords or "",
        language=props.language or "",
        revision=props.revision,
    )
    wb.close()
    return metadata


def _get_cell_value(cell_value: Any) -> Any:
    """
    Convert cell value to appropriate Python type for structured output.

    Handles datetime conversion to ISO format strings for JSON compatibility.

    Args:
        cell_value: Raw value from openpyxl cell.

    Returns:
        Converted value:
        - None for empty cells
        - ISO format string for datetime/date/time values
        - Original value for other types (str, int, float, bool)
    """
    if cell_value is None:
        return None
    if isinstance(cell_value, datetime.datetime):
        return cell_value.isoformat()
    if isinstance(cell_value, datetime.date):
        return cell_value.isoformat()
    if isinstance(cell_value, datetime.time):
        return cell_value.isoformat()
    return cell_value


def _format_value_for_display(value: Any) -> str:
    """
    Format a value as string for text table display.

    Handles special formatting for numeric values (integers without decimals).

    Args:
        value: Value to format (any type).

    Returns:
        String representation of the value:
        - Empty string for None
        - Integer format for whole number floats
        - String conversion for all other values
    """
    if value is None:
        return ""
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value)


def _is_cell_non_empty(val: Any) -> bool:
    """Check if a cell value is non-empty."""
    return val is not None and (not isinstance(val, str) or val.strip() != "")


def _find_last_data_column(rows: list[tuple]) -> int:
    """
    Find the index of the last column that contains any data.

    Scans all rows to find the rightmost column with non-empty content.
    Used for trimming empty trailing columns.

    Args:
        rows: List of row tuples from worksheet iteration.

    Returns:
        1-based column count (0 if no data found).
    """
    if not rows:
        return 0

    max_col = 0
    for row in rows:
        for i in range(len(row) - 1, -1, -1):
            if _is_cell_non_empty(row[i]):
                max_col = max(max_col, i + 1)
                break
    return max_col


def _is_meaningful_value(val: Any) -> bool:
    """Check if a value is meaningful (not None, not empty, not an 'Unnamed:' placeholder)."""
    if val is None:
        return False
    if isinstance(val, str):
        return bool(val.strip()) and not val.startswith("Unnamed: ")
    return True


def _is_table_name_row(row: list[Any]) -> bool:
    """
    Check if a row is just a table name (single non-empty cell).

    Table name rows in Excel often have a single cell with the table title
    and the rest empty. These should be excluded from structured data output.

    Note: After header processing, empty cells become "Unnamed: N" strings,
    so we treat those as empty when counting non-empty cells.

    Args:
        row: List of cell values (may be processed headers with "Unnamed: N").

    Returns:
        True if the row has exactly one meaningful cell, False otherwise.
    """
    non_empty = sum(1 for val in row if _is_meaningful_value(val))
    return non_empty == 1 and len(row) > 1


def _find_last_data_row(rows: list[tuple]) -> int:
    """
    Find the index of the last row that contains any data.

    Scans rows in reverse to find the last row with non-empty content.
    Used for trimming empty trailing rows.

    Args:
        rows: List of row tuples from worksheet iteration.

    Returns:
        1-based row count (0 if no data found).
    """
    if not rows:
        return 0

    for i in range(len(rows) - 1, -1, -1):
        if any(_is_cell_non_empty(val) for val in rows[i]):
            return i + 1
    return 0


def _read_sheet_data(ws: Worksheet) -> tuple[list[dict[str, Any]], list[list[Any]]]:
    """
    Read sheet data and return both structured and raw formats.

    Extracts all data from a worksheet, treating the first row as headers.
    Automatically trims empty trailing rows and columns.

    Args:
        ws: openpyxl Worksheet object.

    Returns:
        Tuple of (records, all_rows) where:
        - records: List of dicts with header keys and cell values
        - all_rows: List of lists including header row (for text formatting)

    Processing:
        1. Read all rows using iter_rows(values_only=True)
        2. Trim trailing empty rows and columns
        3. Use first row as headers (empty headers get "Unnamed: N")
        4. Convert remaining rows to dict format
    """
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []

    # Trim trailing empty rows and columns
    last_row = _find_last_data_row(rows)
    rows = rows[:last_row]
    if not rows:
        return [], []

    last_col = _find_last_data_column(rows)
    rows = [row[:last_col] for row in rows]

    # Generate headers from first row (empty headers get "Unnamed: N")
    headers = [
        (
            f"Unnamed: {i}"
            if val is None or (isinstance(val, str) and not val.strip())
            else str(val)
        )
        for i, val in enumerate(rows[0])
    ]

    # Convert remaining rows to records format
    records: list[dict[str, Any]] = []
    all_rows: list[list[Any]] = [headers]

    for row in rows[1:]:
        record = {
            headers[i]: _get_cell_value(val)
            for i, val in enumerate(row)
            if i < len(headers)
        }
        row_values = [_get_cell_value(val) for val in row]
        records.append(record)
        all_rows.append(row_values)

    return records, all_rows


def _format_sheet_as_text(all_rows: list[list[Any]]) -> str:
    """
    Format sheet data as an aligned text table.

    Creates a text representation similar to pandas DataFrame.to_string()
    with right-aligned columns and consistent spacing.

    Args:
        all_rows: List of rows including header row. Each row is a list
            of values (already converted to appropriate Python types).

    Returns:
        Formatted text table with:
        - Right-aligned columns
        - Single-space column separation
        - Header as first row
        - Empty string if no data
    """
    if not all_rows:
        return ""

    num_cols = max(len(row) for row in all_rows)
    col_widths = [0] * num_cols

    # Format all values and calculate column widths in one pass
    formatted_rows: list[list[str]] = []
    for row in all_rows:
        formatted_row = [
            _format_value_for_display(row[i] if i < len(row) else None)
            for i in range(num_cols)
        ]
        for i, val in enumerate(formatted_row):
            col_widths[i] = max(col_widths[i], len(val))
        formatted_rows.append(formatted_row)

    # Build text output with right-aligned columns
    return "\n".join(
        " ".join(val.rjust(col_widths[i]) for i, val in enumerate(row))
        for row in formatted_rows
    )


def _get_content_type(filename: str) -> str:
    """
    Determine MIME content type from image filename extension.

    Args:
        filename: Image filename with extension.

    Returns:
        MIME type string (e.g., 'image/png', 'image/jpeg').
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return _CONTENT_TYPE_MAP.get(ext, "image/unknown")


def _get_image_pixel_dimensions(
    image_data: bytes,
) -> tuple[int | None, int | None]:
    """Best-effort extraction of pixel dimensions from common raster formats."""
    if not image_data:
        return None, None

    # PNG
    if image_data.startswith(b"\x89PNG\r\n\x1a\n") and len(image_data) >= 24:
        width = int.from_bytes(image_data[16:20], "big")
        height = int.from_bytes(image_data[20:24], "big")
        return (width or None, height or None)

    # GIF
    if image_data[:6] in (b"GIF87a", b"GIF89a") and len(image_data) >= 10:
        width = int.from_bytes(image_data[6:8], "little")
        height = int.from_bytes(image_data[8:10], "little")
        return (width or None, height or None)

    # BMP
    if image_data[:2] == b"BM" and len(image_data) >= 26:
        width = int.from_bytes(image_data[18:22], "little", signed=True)
        height = int.from_bytes(image_data[22:26], "little", signed=True)
        return (abs(width) or None, abs(height) or None)

    # JPEG
    if image_data.startswith(b"\xff\xd8"):
        i = 2
        size = len(image_data)
        while i + 4 <= size:
            if image_data[i] != 0xFF:
                i += 1
                continue
            marker = image_data[i + 1]
            if marker in (0xD9, 0xDA):
                break
            length = int.from_bytes(image_data[i + 2 : i + 4], "big")
            if length < 2:
                break
            if marker in (
                0xC0,
                0xC1,
                0xC2,
                0xC3,
                0xC5,
                0xC6,
                0xC7,
                0xC9,
                0xCA,
                0xCB,
                0xCD,
                0xCE,
                0xCF,
            ):
                if i + 2 + length <= size:
                    height = int.from_bytes(image_data[i + 5 : i + 7], "big")
                    width = int.from_bytes(image_data[i + 7 : i + 9], "big")
                    return (width or None, height or None)
                break
            i += 2 + length

    return None, None


def _resolve_drawing_path(target: str) -> str:
    """Normalize drawing relationship targets to ZIP paths."""
    if target.startswith("/"):
        return target[1:]
    if target.startswith(".."):
        return "xl/" + target[3:]
    return "xl/worksheets/" + target


def _resolve_image_path(target: str) -> str:
    """Normalize image relationship targets to ZIP paths."""
    if target.startswith("/"):
        return target[1:]
    return "xl/media/" + target.rsplit("/", 1)[-1]


def _extract_images_from_zip(
    file_like: io.BytesIO, sheet_names: list[str]
) -> dict[int, list[XlsxImage]]:
    """
    Extract all images from an XLSX file by parsing the ZIP archive directly.

    XLSX files store images in xl/media/ and reference them via drawings.
    Each sheet may have a drawing (xl/drawings/drawingN.xml) that contains
    image references with metadata like name and description.

    Args:
        file_like: BytesIO containing the XLSX file.
        sheet_names: List of sheet names in order (to map drawing to sheet index).

    Returns:
        Dictionary mapping sheet_index to list of XlsxImage objects.
    """
    images_by_sheet: dict[int, list[XlsxImage]] = {}
    image_counter = 0

    ctx = OOXMLZipContext(file_like)
    try:
        namelist = ctx.namelist

        # Build mapping of sheet index to drawing file
        sheet_to_drawing: dict[int, str] = {}

        for sheet_idx in range(len(sheet_names)):
            rels_path = f"xl/worksheets/_rels/sheet{sheet_idx + 1}.xml.rels"
            if rels_path not in namelist:
                continue

            rels_root = ctx.read_xml_root(rels_path)
            for rel in parse_relationships(rels_root):
                if "drawing" in rel["type"]:
                    sheet_to_drawing[sheet_idx] = _resolve_drawing_path(rel["target"])
                    break

        # Process each drawing file
        for sheet_idx, drawing_path in sheet_to_drawing.items():
            if drawing_path not in namelist:
                continue

            # Parse drawing relationships to get image file paths
            drawing_rels_path = drawing_path.replace(
                "drawings/", "drawings/_rels/"
            ).replace(".xml", ".xml.rels")

            rid_to_image: dict[str, str] = {}
            if drawing_rels_path in namelist:
                for rel in parse_relationships(ctx.read_xml_root(drawing_rels_path)):
                    if "image" in rel["type"]:
                        rid_to_image[rel["id"]] = _resolve_image_path(rel["target"])

            # Parse the drawing XML to get image metadata
            drawing_root = ctx.read_xml_root(drawing_path)
            sheet_images: list[XlsxImage] = []

            for anchor_type in ANCHOR_TYPES:
                for anchor in drawing_root.iter(anchor_type):
                    pic = anchor.find(XDR_PIC)
                    if pic is None:
                        continue

                    try:
                        # Get dimensions from ext element
                        width, height = 0, 0
                        ext = anchor.find(XDR_EXT)
                        if ext is not None:
                            try:
                                width = int(ext.get("cx", "0")) // EMU_PER_PIXEL
                                height = int(ext.get("cy", "0")) // EMU_PER_PIXEL
                            except ValueError:
                                pass

                        # Get caption and description from non-visual properties
                        caption, description = "", ""
                        nvPicPr = pic.find(XDR_NVPICPR)
                        if nvPicPr is not None:
                            cNvPr = nvPicPr.find(XDR_CNVPR)
                            if cNvPr is not None:
                                caption = cNvPr.get("name", "")
                                description = cNvPr.get("descr", "")

                        # Get the blip reference to find the image file
                        blipFill = pic.find(XDR_BLIPFILL)
                        if blipFill is None:
                            continue

                        blip = blipFill.find(A_BLIP)
                        if blip is None:
                            continue

                        embed_rid = blip.get(R_EMBED, "")
                        if not embed_rid or embed_rid not in rid_to_image:
                            continue

                        image_path = rid_to_image[embed_rid]
                        if image_path not in namelist:
                            continue

                        # Read the image data
                        image_bytes = ctx.read_bytes(image_path)
                        filename = image_path.rsplit("/", 1)[-1]

                        if width <= 0 or height <= 0:
                            width, height = _get_image_pixel_dimensions(image_bytes)

                        image_counter += 1
                        sheet_images.append(
                            XlsxImage(
                                image_index=image_counter,
                                sheet_index=sheet_idx,
                                filename=filename,
                                content_type=_get_content_type(filename),
                                data=io.BytesIO(image_bytes),
                                size_bytes=len(image_bytes),
                                width=width,
                                height=height,
                                caption=caption,
                                description=description,
                            )
                        )

                    except Exception as e:
                        logger.debug(f"Failed to extract image from drawing: {e}")

            if sheet_images:
                images_by_sheet[sheet_idx] = sheet_images

    except Exception as e:
        logger.debug(f"Failed to extract images from XLSX: {e}")
    finally:
        ctx.close()

    return images_by_sheet


def _read_content(file_like: io.BytesIO) -> list[XlsxSheet]:
    """
    Read all sheets from an XLSX file and extract their content.

    Uses openpyxl in read_only mode for text/data extraction (memory efficient),
    then parses the ZIP archive directly for image extraction.

    Args:
        file_like: BytesIO containing the XLSX file.

    Returns:
        List of XlsxSheet objects, one per worksheet, each containing:
        - name: Sheet name
        - data: List of dicts (row data with header keys)
        - text: Formatted text table
        - images: List of XlsxImage objects

    Notes:
        - Workbook is opened in read_only and data_only mode for text
        - Images are extracted by parsing the ZIP archive directly
        - Empty sheets have empty data and text
        - Workbook is closed after extraction
    """
    file_like.seek(0)
    wb = load_workbook(file_like, read_only=True, data_only=True)

    sheet_names = list(wb.sheetnames)
    sheets: list[XlsxSheet] = []

    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        records, all_rows = _read_sheet_data(ws)
        text = _format_sheet_as_text(all_rows)

        # Skip first row if it's just a table name
        data_rows = (
            all_rows[1:] if all_rows and _is_table_name_row(all_rows[0]) else all_rows
        )

        sheets.append(
            XlsxSheet(
                name=str(sheet_name),
                data=data_rows,
                text=text,
                images=[],
            )
        )

    wb.close()

    # Extract images by parsing the ZIP archive directly
    images_by_sheet = _extract_images_from_zip(file_like, sheet_names)

    for sheet_idx, sheet_images in images_by_sheet.items():
        if sheet_idx < len(sheets):
            sheets[sheet_idx].images = sheet_images

    return sheets


def read_xlsx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsxContent, Any, None]:
    """
    Extract all relevant content from an Excel .xlsx file.

    Primary entry point for XLSX file extraction. Parses all sheets,
    extracts cell values, and retrieves document metadata using openpyxl.

    This function uses a generator pattern for API consistency with other
    extractors, even though XLSX files contain exactly one workbook.

    Args:
        file_like: BytesIO object containing the complete XLSX file data.
            The stream is read multiple times (for content and metadata).
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned XlsxContent.metadata.

    Yields:
        XlsxContent: Single XlsxContent object containing:
            - metadata: XlsxMetadata with title, creator, dates
            - sheets: List of XlsxSheet objects (name, data, text)

    Raises:
        ExtractionFileEncryptedError: If the XLSX is encrypted or password-protected.

    Example:
        >>> import io
        >>> with open("data.xlsx", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for workbook in read_xlsx(data, path="data.xlsx"):
        ...         print(f"Creator: {workbook.metadata.creator}")
        ...         print(f"Sheets: {len(workbook.sheets)}")
        ...         for sheet in workbook.sheets:
        ...             print(f"  {sheet.name}: {len(sheet.data)} rows")
        ...             # Access structured data
        ...             for row in sheet.data[:3]:
        ...                 print(f"    {row}")

    Performance Notes:
        - Uses read_only mode for memory efficiency
        - Large spreadsheets still load all data into memory
        - Consider streaming approaches for very large files
    """
    try:
        file_like.seek(0)
        if is_ooxml_encrypted(file_like):
            raise ExtractionFileEncryptedError(
                "XLSX is encrypted or password-protected"
            )

        validate_zip_bytesio(file_like, source="read_xlsx")

        sheets = _read_content(file_like)
        metadata = _read_metadata(file_like)
        metadata.populate_from_path(path)

        total_rows = sum(len(sheet.data) for sheet in sheets)
        total_images = sum(len(sheet.images) for sheet in sheets)
        logger.info(
            "Extracted XLSX: %d sheets, %d total rows, %d images",
            len(sheets),
            total_rows,
            total_images,
        )

        yield XlsxContent(metadata=metadata, sheets=sheets)
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract XLSX file", cause=exc) from exc
