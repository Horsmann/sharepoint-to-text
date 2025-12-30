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
import zipfile
from typing import Any, Dict, Generator, List
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from sharepoint2text.extractors.data_types import (
    XlsxContent,
    XlsxImage,
    XlsxMetadata,
    XlsxSheet,
)

logger = logging.getLogger(__name__)


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


def _find_last_data_column(rows: List[tuple]) -> int:
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
            val = row[i]
            if val is not None and (not isinstance(val, str) or val.strip() != ""):
                max_col = max(max_col, i + 1)
                break
    return max_col


def _is_table_name_row(row: List[Any]) -> bool:
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

    def is_meaningful(val: Any) -> bool:
        if val is None:
            return False
        if isinstance(val, str):
            # Treat "Unnamed: N" placeholders as empty
            if val.startswith("Unnamed: "):
                return False
            return bool(val.strip())
        return True

    non_empty = sum(1 for val in row if is_meaningful(val))
    return non_empty == 1 and len(row) > 1


def _find_last_data_row(rows: List[tuple]) -> int:
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
        row = rows[i]
        for val in row:
            if val is not None and (not isinstance(val, str) or val.strip() != ""):
                return i + 1
    return 0


def _read_sheet_data(ws: Worksheet) -> tuple[List[Dict[str, Any]], List[List[Any]]]:
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

    # First row is the header
    header_row = rows[0]

    # Generate column names, using "Unnamed: N" for empty headers
    headers = []
    for i, val in enumerate(header_row):
        if val is None or (isinstance(val, str) and val.strip() == ""):
            headers.append(f"Unnamed: {i}")
        else:
            headers.append(str(val))

    # Convert remaining rows to records format
    records = []
    all_rows = [headers]

    for row in rows[1:]:
        record = {}
        row_values = []
        for i, cell_value in enumerate(row):
            if i < len(headers):
                record[headers[i]] = _get_cell_value(cell_value)
            row_values.append(_get_cell_value(cell_value))
        records.append(record)
        all_rows.append(row_values)

    return records, all_rows


def _format_sheet_as_text(all_rows: List[List[Any]]) -> str:
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

    # Calculate column widths
    num_cols = max(len(row) for row in all_rows) if all_rows else 0
    col_widths = [0] * num_cols

    formatted_rows = []
    for row in all_rows:
        formatted_row = []
        for i in range(num_cols):
            val = row[i] if i < len(row) else None
            formatted_val = _format_value_for_display(val)
            formatted_row.append(formatted_val)
            col_widths[i] = max(col_widths[i], len(formatted_val))
        formatted_rows.append(formatted_row)

    # Build text output with right-aligned columns
    lines = []
    for formatted_row in formatted_rows:
        cells = []
        for i, val in enumerate(formatted_row):
            cells.append(val.rjust(col_widths[i]))
        lines.append(" ".join(cells))

    return "\n".join(lines)


def _get_content_type(filename: str) -> str:
    """
    Determine MIME content type from image filename extension.

    Args:
        filename: Image filename with extension.

    Returns:
        MIME type string (e.g., 'image/png', 'image/jpeg').
    """
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    content_types = {
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
    return content_types.get(ext, "image/unknown")


def _extract_images_from_zip(
    file_like: io.BytesIO, sheet_names: List[str]
) -> Dict[int, List[XlsxImage]]:
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
    # XML namespaces used in XLSX drawings
    NS = {
        "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    images_by_sheet: Dict[int, List[XlsxImage]] = {}
    image_counter = 0

    file_like.seek(0)
    try:
        with zipfile.ZipFile(file_like, "r") as zf:
            namelist = zf.namelist()

            # Build mapping of sheet index to drawing file
            # Parse xl/worksheets/_rels/sheetN.xml.rels to find drawing references
            sheet_to_drawing: Dict[int, str] = {}

            for sheet_idx in range(len(sheet_names)):
                sheet_num = sheet_idx + 1
                rels_path = f"xl/worksheets/_rels/sheet{sheet_num}.xml.rels"
                if rels_path not in namelist:
                    continue

                rels_xml = zf.read(rels_path).decode("utf-8")
                rels_root = ET.fromstring(rels_xml)

                # Find Relationship elements - try with and without namespace prefix
                relationships = rels_root.findall("rel:Relationship", NS)
                if not relationships:
                    relationships = rels_root.findall(
                        "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
                    )

                for rel in relationships:
                    rel_type = rel.get("Type", "")
                    if "drawing" in rel_type:
                        target = rel.get("Target", "")
                        # Target can be:
                        # - Absolute: "/xl/drawings/drawing1.xml"
                        # - Relative: "../drawings/drawing1.xml"
                        if target.startswith("/"):
                            drawing_path = target[1:]  # Remove leading /
                        elif target.startswith(".."):
                            drawing_path = "xl/" + target[3:]  # Remove "../"
                        else:
                            drawing_path = "xl/worksheets/" + target
                        sheet_to_drawing[sheet_idx] = drawing_path
                        break

            # Process each drawing file
            for sheet_idx, drawing_path in sheet_to_drawing.items():
                if drawing_path not in namelist:
                    continue

                # Parse drawing relationships to get image file paths
                drawing_rels_path = drawing_path.replace(
                    "drawings/", "drawings/_rels/"
                ).replace(".xml", ".xml.rels")

                rid_to_image: Dict[str, str] = {}
                if drawing_rels_path in namelist:
                    rels_xml = zf.read(drawing_rels_path).decode("utf-8")
                    rels_root = ET.fromstring(rels_xml)

                    # Find Relationship elements - try with and without namespace prefix
                    relationships = rels_root.findall("rel:Relationship", NS)
                    if not relationships:
                        relationships = rels_root.findall(
                            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
                        )

                    for rel in relationships:
                        rel_type = rel.get("Type", "")
                        if "image" in rel_type:
                            rid = rel.get("Id", "")
                            target = rel.get("Target", "")
                            # Target might be absolute (/xl/media/...) or relative
                            if target.startswith("/"):
                                image_path = target[1:]  # Remove leading /
                            else:
                                # Relative to drawings folder
                                image_path = "xl/media/" + target.rsplit("/", 1)[-1]
                            rid_to_image[rid] = image_path

                # Parse the drawing XML to get image metadata
                drawing_xml = zf.read(drawing_path).decode("utf-8")
                drawing_root = ET.fromstring(drawing_xml)

                sheet_images: List[XlsxImage] = []

                # EMUs to pixels conversion factor (9525 EMUs = 1 pixel at 96 DPI)
                EMU_PER_PIXEL = 9525

                # Find all anchor elements (oneCellAnchor, twoCellAnchor, absoluteAnchor)
                anchor_types = [
                    f"{{{NS['xdr']}}}oneCellAnchor",
                    f"{{{NS['xdr']}}}twoCellAnchor",
                    f"{{{NS['xdr']}}}absoluteAnchor",
                ]

                for anchor_type in anchor_types:
                    for anchor in drawing_root.iter(anchor_type):
                        # Find the pic element within this anchor
                        pic = anchor.find("xdr:pic", NS)
                        if pic is None:
                            continue

                        try:
                            # Get dimensions from ext element (sibling of pic)
                            width = 0
                            height = 0
                            ext = anchor.find("xdr:ext", NS)
                            if ext is not None:
                                cx = ext.get("cx", "0")
                                cy = ext.get("cy", "0")
                                try:
                                    width = int(cx) // EMU_PER_PIXEL
                                    height = int(cy) // EMU_PER_PIXEL
                                except ValueError:
                                    pass

                            # Get non-visual properties for name and description
                            nvPicPr = pic.find("xdr:nvPicPr", NS)
                            caption = ""
                            description = ""

                            if nvPicPr is not None:
                                cNvPr = nvPicPr.find("xdr:cNvPr", NS)
                                if cNvPr is not None:
                                    caption = cNvPr.get("name", "")
                                    description = cNvPr.get("descr", "")

                            # Get the blip reference to find the image file
                            blipFill = pic.find("xdr:blipFill", NS)
                            if blipFill is None:
                                continue

                            blip = blipFill.find("a:blip", NS)
                            if blip is None:
                                continue

                            # Get the relationship ID
                            embed_rid = blip.get(f"{{{NS['r']}}}embed", "")
                            if not embed_rid or embed_rid not in rid_to_image:
                                continue

                            image_path = rid_to_image[embed_rid]
                            if image_path not in namelist:
                                continue

                            # Read the image data
                            image_bytes = zf.read(image_path)
                            image_data = io.BytesIO(image_bytes)
                            size_bytes = len(image_bytes)

                            # Get filename from path
                            filename = image_path.rsplit("/", 1)[-1]
                            content_type = _get_content_type(filename)

                            sheet_images.append(
                                XlsxImage(
                                    image_index=image_counter,
                                    sheet_index=sheet_idx,
                                    filename=filename,
                                    content_type=content_type,
                                    data=image_data,
                                    size_bytes=size_bytes,
                                    width=width,
                                    height=height,
                                    caption=caption,
                                    description=description,
                                )
                            )
                            image_counter += 1

                        except Exception as e:
                            logger.warning(f"Failed to extract image from drawing: {e}")
                            continue

                if sheet_images:
                    images_by_sheet[sheet_idx] = sheet_images

    except Exception as e:
        logger.warning(f"Failed to extract images from XLSX: {e}")

    return images_by_sheet


def _read_content(file_like: io.BytesIO) -> List[XlsxSheet]:
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
    logger.debug("Reading content")
    file_like.seek(0)
    wb = load_workbook(file_like, read_only=True, data_only=True)

    sheet_names = list(wb.sheetnames)
    sheets = []
    for sheet_name in sheet_names:
        logger.debug(f"Reading sheet: [{sheet_name}]")
        ws = wb[sheet_name]
        records, all_rows = _read_sheet_data(ws)
        text = _format_sheet_as_text(all_rows)

        # Determine data rows: skip first row if it's just a table name
        if all_rows and _is_table_name_row(all_rows[0]):
            data_rows = all_rows[1:]
        else:
            data_rows = all_rows

        sheets.append(
            XlsxSheet(
                name=str(sheet_name),
                data=data_rows,
                text=text,
                images=[],  # Will be populated in the image extraction pass
            )
        )

    wb.close()

    # Extract images by parsing the ZIP archive directly
    images_by_sheet = _extract_images_from_zip(file_like, sheet_names)

    for sheet_idx, sheet_images in images_by_sheet.items():
        if sheet_idx < len(sheets):
            sheets[sheet_idx].images = sheet_images
            logger.debug(
                f"Extracted {len(sheet_images)} images from sheet: [{sheet_names[sheet_idx]}]"
            )

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
