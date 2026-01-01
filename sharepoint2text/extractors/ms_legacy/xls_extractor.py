"""
XLS Spreadsheet Extractor
=========================

Extracts text content and metadata from legacy Microsoft Excel .xls files
(Excel 97-2003 binary format, also known as BIFF8).

This module uses the xlrd library for cell and sheet parsing, and olefile
for metadata extraction from the OLE container.

File Format Background
----------------------
The .xls format stores spreadsheet data in a binary format called BIFF
(Binary Interchange File Format). BIFF8, used by Excel 97-2003, stores
data within an OLE2 compound document structure.

Key Components:
    - Workbook stream: Contains sheets, cells, formulas, formatting
    - Sheets: Individual worksheets with rows and columns
    - Cells: Individual data values (text, numbers, dates, formulas)
    - Styles: Cell formatting (fonts, colors, borders)

Cell Types (xlrd constants):
    - XL_CELL_EMPTY (0): Empty cell
    - XL_CELL_TEXT (1): Text string
    - XL_CELL_NUMBER (2): Number (float)
    - XL_CELL_DATE (3): Date (stored as float, needs conversion)
    - XL_CELL_BOOLEAN (4): Boolean (True/False)
    - XL_CELL_ERROR (5): Error value
    - XL_CELL_BLANK (6): Blank (formatted but empty)

Dependencies
------------
xlrd: https://github.com/python-excel/xlrd
    pip install xlrd

    Provides:
    - BIFF8 format parsing
    - Cell value extraction with type detection
    - Date conversion from Excel serial format
    - Sheet enumeration

    Note: xlrd 2.x only supports .xls, not .xlsx (use openpyxl for .xlsx)

olefile: https://github.com/decalage2/olefile
    pip install olefile

    Provides:
    - OLE compound document parsing
    - Metadata extraction from SummaryInformation

Known Limitations
-----------------
- Password-protected/encrypted files are not supported
- Macros (VBA) are not extracted
- Charts and images are not extracted
- External references may not resolve
- Very large spreadsheets may use significant memory
- Formulas are not evaluated (only stored results)
- Merged cells may have unexpected behavior

Data Representation
-------------------
Each sheet is represented in two forms:

1. Structured data (list of dicts):
   - First row is treated as headers
   - Each subsequent row becomes a dictionary
   - Keys are header values, values are cell contents
   - Native Python types are used (int, float, str, bool)

2. Text representation:
   - Formatted as a text table (similar to pandas to_string)
   - Columns are aligned and padded
   - Suitable for display or text search

Date Handling
-------------
Excel stores dates as floating-point numbers (days since 1899-12-30 or
1904-01-01 depending on workbook mode). The extractor converts these
to ISO format strings (YYYY-MM-DD or YYYY-MM-DD HH:MM:SS).

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_legacy.xls_extractor import read_xls
    >>>
    >>> with open("data.xls", "rb") as f:
    ...     for workbook in read_xls(io.BytesIO(f.read()), path="data.xls"):
    ...         print(f"Title: {workbook.metadata.title}")
    ...         for sheet in workbook.sheets:
    ...             print(f"Sheet: {sheet.name}")
    ...             print(f"Rows: {len(sheet.data)}")
    ...             print(sheet.text[:200])

See Also
--------
- MS-XLS specification: https://docs.microsoft.com/en-us/openspecs/office_file_formats/
- xlsx_extractor: For modern Excel .xlsx format
- doc_extractor: For Word documents
- ppt_extractor: For PowerPoint presentations

Maintenance Notes
-----------------
- xlrd is essentially unmaintained but stable for .xls
- Consider openpyxl for .xlsx files
- First row is always treated as headers (no detection)
- Empty sheets produce empty text output
"""

import hashlib
import io
import logging
import struct
from typing import Any, Dict, Generator, List

import olefile
import xlrd

from sharepoint2text.exceptions import ExtractionFileEncryptedError
from sharepoint2text.extractors.data_types import (
    XlsContent,
    XlsImage,
    XlsMetadata,
    XlsSheet,
)
from sharepoint2text.extractors.util.encryption import is_xls_encrypted
from sharepoint2text.extractors.util.image_utils import (
    BLIP_INSTANCE_JPEG_2,
    BLIP_INSTANCE_PNG_2,
    BLIP_TYPE_DIB,
    BLIP_TYPE_EMF,
    BLIP_TYPE_WMF,
    BLIP_TYPES,
    detect_image_type,
    get_image_dimensions,
    wrap_dib_as_bmp,
)

logger = logging.getLogger(__name__)


def _cell_value_to_str(cell: xlrd.sheet.Cell, workbook: xlrd.Book) -> str:
    """
    Convert a cell value to a string representation for display.

    Handles all Excel cell types and produces human-readable strings
    suitable for text output and display.

    Args:
        cell: xlrd Cell object containing value and type.
        workbook: xlrd Book object (needed for date conversion).

    Returns:
        String representation of the cell value:
        - Empty cells: ""
        - Text: The string value
        - Numbers: Integer if whole number, else float as string
        - Dates: ISO format (YYYY-MM-DD or YYYY-MM-DD HH:MM:SS)
        - Booleans: "True" or "False"
        - Errors: "#ERROR"

    Date Handling:
        Excel stores dates as floating-point numbers. The workbook's
        datemode (0 for 1900 system, 1 for 1904 system) is needed to
        correctly convert to a date tuple. Dates with no time component
        (00:00:00) are formatted as date-only.

    Note:
        Errors during date conversion fall back to the raw float value.
    """
    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return ""
    elif cell.ctype == xlrd.XL_CELL_TEXT:
        return str(cell.value)
    elif cell.ctype == xlrd.XL_CELL_NUMBER:
        # Check if it's an integer
        if cell.value == int(cell.value):
            return str(int(cell.value))
        return str(cell.value)
    elif cell.ctype == xlrd.XL_CELL_DATE:
        try:
            date_tuple = xlrd.xldate_as_tuple(cell.value, workbook.datemode)
            # Format as ISO date or datetime
            if date_tuple[3:] == (0, 0, 0):
                return f"{date_tuple[0]:04d}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
            return f"{date_tuple[0]:04d}-{date_tuple[1]:02d}-{date_tuple[2]:02d} {date_tuple[3]:02d}:{date_tuple[4]:02d}:{date_tuple[5]:02d}"
        except Exception:
            return str(cell.value)
    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return "True" if cell.value else "False"
    elif cell.ctype == xlrd.XL_CELL_ERROR:
        return "#ERROR"
    else:
        return str(cell.value)


def _get_cell_native_value(cell: xlrd.sheet.Cell, workbook: xlrd.Book) -> Any:
    """
    Get the native Python value of a cell for structured data output.

    Similar to _cell_value_to_str but returns native Python types
    suitable for data processing rather than display strings.

    Args:
        cell: xlrd Cell object containing value and type.
        workbook: xlrd Book object (needed for date conversion).

    Returns:
        Native Python value:
        - Empty cells: None
        - Text: str
        - Numbers: int if whole number, else float
        - Dates: ISO format string (for JSON compatibility)
        - Booleans: bool (True/False)
        - Errors: None

    Note:
        Dates are returned as strings rather than datetime objects
        for JSON serialization compatibility.
    """
    if cell.ctype == xlrd.XL_CELL_EMPTY:
        return None
    elif cell.ctype == xlrd.XL_CELL_TEXT:
        return cell.value
    elif cell.ctype == xlrd.XL_CELL_NUMBER:
        if cell.value == int(cell.value):
            return int(cell.value)
        return cell.value
    elif cell.ctype == xlrd.XL_CELL_DATE:
        try:
            date_tuple = xlrd.xldate_as_tuple(cell.value, workbook.datemode)
            if date_tuple[3:] == (0, 0, 0):
                return f"{date_tuple[0]:04d}-{date_tuple[1]:02d}-{date_tuple[2]:02d}"
            return f"{date_tuple[0]:04d}-{date_tuple[1]:02d}-{date_tuple[2]:02d} {date_tuple[3]:02d}:{date_tuple[4]:02d}:{date_tuple[5]:02d}"
        except Exception:
            return cell.value
    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    elif cell.ctype == xlrd.XL_CELL_ERROR:
        return None
    else:
        return cell.value


def _format_sheet_as_text(headers: List[str], rows: List[List[str]]) -> str:
    """
    Format sheet data as a text table with aligned columns.

    Creates a text representation similar to pandas DataFrame.to_string()
    output, with right-aligned columns and consistent spacing.

    Args:
        headers: List of header strings (first row of sheet).
        rows: List of lists of string values (data rows).

    Returns:
        Formatted text table with:
        - Right-aligned columns
        - Two-space column separation
        - Headers as first row
        - Empty string if no data

    Example Output:
        Name  Age  City
        John   30  NYC
        Jane   25  LA
    """
    if not headers and not rows:
        return ""

    # Calculate column widths
    all_rows = [headers] + rows if headers else rows
    if not all_rows:
        return ""

    num_cols = max(len(row) for row in all_rows) if all_rows else 0
    col_widths = [0] * num_cols

    for row in all_rows:
        for i, val in enumerate(row):
            col_widths[i] = max(col_widths[i], len(str(val)))

    # Build text output
    lines = []
    for row in all_rows:
        padded = []
        for i, val in enumerate(row):
            width = col_widths[i] if i < len(col_widths) else len(str(val))
            padded.append(str(val).rjust(width))
        lines.append("  ".join(padded))

    return "\n".join(lines)


def _read_content(file_like: io.BytesIO) -> List[XlsSheet]:
    """
    Read all sheets from an XLS file and extract their content.

    Uses xlrd to parse the workbook and extract cell values from each
    sheet. First row is treated as headers for the structured data output.

    Args:
        file_like: BytesIO containing the XLS file data.

    Returns:
        List of XlsSheet objects, one per worksheet, each containing:
        - name: Sheet name
        - data: List of dicts (row data with header keys)
        - text: Formatted text table

    Implementation Notes:
        - Empty sheets have empty data and text
        - First row becomes headers for all subsequent rows
        - Each row is converted to both dict form and text form
    """
    logger.debug("Reading content")
    workbook = xlrd.open_workbook(file_contents=file_like.read())

    sheets = []
    for sheet in workbook.sheets():
        logger.debug(f"Reading sheet: [{sheet.name}]")

        if sheet.nrows == 0:
            sheets.append(XlsSheet(name=sheet.name, data=[], text=""))
            continue

        # Get headers from first row
        headers: List[str] = []
        if sheet.nrows > 0:
            headers = [
                _cell_value_to_str(sheet.cell(0, col), workbook)
                for col in range(sheet.ncols)
            ]

        # Build data (list of dicts) and rows for text representation
        data: List[Dict[str, Any]] = []
        text_rows: List[List[str]] = []

        for row_idx in range(1, sheet.nrows):
            row_dict: Dict[str, Any] = {}
            row_text: List[str] = []

            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                header = (
                    headers[col_idx] if col_idx < len(headers) else f"col_{col_idx}"
                )

                row_dict[header] = _get_cell_native_value(cell, workbook)
                row_text.append(_cell_value_to_str(cell, workbook))

            data.append(row_dict)
            text_rows.append(row_text)

        text = _format_sheet_as_text(headers, text_rows)

        sheets.append(
            XlsSheet(
                name=sheet.name,
                data=data,
                text=text,
            )
        )

    return sheets


def _read_metadata(file_like: io.BytesIO) -> XlsMetadata:
    """
    Extract document metadata from the XLS file's OLE container.

    Uses olefile to access the SummaryInformation stream which
    contains standard document properties.

    Args:
        file_like: BytesIO containing the XLS file data.

    Returns:
        XlsMetadata object with:
        - title, author, subject, company
        - last_saved_by
        - created, modified (ISO format dates)

    Notes:
        - OleFileIO is opened and closed within this function
        - Bytes values are decoded as UTF-8
        - Missing properties result in empty strings
    """
    ole = olefile.OleFileIO(file_like)
    meta = ole.get_metadata()

    result = XlsMetadata(
        title=meta.title.decode("utf-8") if meta.title else "",
        author=meta.author.decode("utf-8") if meta.author else "",
        subject=meta.subject.decode("utf-8") if meta.subject else "",
        company=meta.company.decode("utf-8") if meta.company else "",
        last_saved_by=meta.last_saved_by.decode("utf-8") if meta.last_saved_by else "",
        created=meta.create_time.isoformat() if meta.create_time else "",
        modified=meta.last_saved_time.isoformat() if meta.last_saved_time else "",
    )
    ole.close()
    return result


def read_xls(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsContent, Any, None]:
    """
    Extract all relevant content from a legacy Excel .xls file.

    Primary entry point for XLS file extraction. Parses all sheets,
    extracts cell values, and retrieves document metadata.

    This function uses a generator pattern for API consistency with other
    extractors, even though XLS files contain exactly one workbook.

    Args:
        file_like: BytesIO object containing the complete XLS file data.
            The stream position is reset between reads for content and
            metadata extraction.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned XlsContent.metadata.

    Yields:
        XlsContent: Single XlsContent object containing:
            - metadata: XlsMetadata with title, author, dates
            - sheets: List of XlsSheet objects (name, data, text)
            - full_text: All sheet text concatenated with blank lines

    Example:
        >>> import io
        >>> with open("sales.xls", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for workbook in read_xls(data, path="sales.xls"):
        ...         print(f"Author: {workbook.metadata.author}")
        ...         print(f"Sheets: {len(workbook.sheets)}")
        ...         for sheet in workbook.sheets:
        ...             print(f"  {sheet.name}: {len(sheet.data)} rows")
        ...             # Access structured data
        ...             for row in sheet.data[:3]:
        ...                 print(f"    {row}")

    Implementation Notes:
        - File is read twice (once for content, once for metadata)
        - Stream position is reset between reads
        - Full text is sheets joined by double newlines
        - Empty sheets are included with empty data and text

    Performance Notes:
        - Entire file is loaded into memory by xlrd
        - Large spreadsheets with many cells may use significant memory
        - Consider streaming approaches for very large files
    """
    file_like.seek(0)
    if is_xls_encrypted(file_like):
        raise ExtractionFileEncryptedError("XLS is encrypted or password-protected")

    sheets = _read_content(file_like=file_like)
    file_like.seek(0)
    metadata = _read_metadata(file_like=file_like)
    metadata.populate_from_path(path)

    full_text = "\n\n".join(sheet.text for sheet in sheets)

    # Extract images from the Workbook stream
    file_like.seek(0)
    images = _extract_images_from_workbook(file_like)

    yield XlsContent(
        metadata=metadata,
        sheets=sheets,
        images=images,
        full_text=full_text,
    )


# =============================================================================
# Image Extraction from Workbook Stream
# =============================================================================


def _extract_images_from_workbook(file_like: io.BytesIO) -> List[XlsImage]:
    """
    Extract images from the Workbook stream of an XLS file.

    XLS files store images in the Workbook stream using OfficeArt BLIP
    (Binary Large Image Picture) format, the same format used by PPT files.

    Args:
        file_like: BytesIO containing the XLS file data.

    Returns:
        List of XlsImage objects with extracted image data.

    Notes:
        - Supports PNG, JPEG, GIF, BMP, and TIFF formats
        - EMF/WMF metafiles are extracted but may not be widely supported
        - Image dimensions are extracted where available
    """
    images: List[XlsImage] = []

    file_like.seek(0)
    if not olefile.isOleFile(file_like):
        return images

    file_like.seek(0)

    try:
        ole = olefile.OleFileIO(file_like)
        if not ole.exists("Workbook"):
            ole.close()
            return images

        workbook_stream = ole.openstream("Workbook")
        data = workbook_stream.read()
        ole.close()
    except Exception as e:
        logger.debug(f"Failed to read Workbook stream: {e}")
        return images

    if len(data) < 25:
        return images

    image_index = 0
    seen_hashes: set[str] = set()

    # Scan for BLIP records in the stream
    offset = 0
    while offset < len(data) - 8:
        # Check for OfficeArt record header
        try:
            rec_ver_instance = struct.unpack("<H", data[offset : offset + 2])[0]
            rec_type = struct.unpack("<H", data[offset + 2 : offset + 4])[0]
            rec_len = struct.unpack("<I", data[offset + 4 : offset + 8])[0]
        except struct.error:
            offset += 1
            continue

        # Validate record length
        if rec_len <= 0 or rec_len > len(data) - offset - 8:
            offset += 1
            continue

        # Check if this is a BLIP record
        is_blip = rec_type in BLIP_TYPES

        if is_blip:
            record_data = data[offset + 8 : offset + 8 + rec_len]

            if len(record_data) > 17:
                # BLIP structure:
                # - 16 bytes: rgbUid1 (MD4 hash of image data)
                # - 1 byte: tag
                # - remaining: image data

                blip_header_size = 17

                # Check for secondary UID
                rec_instance = (rec_ver_instance >> 4) & 0x0FFF
                if rec_instance in (BLIP_INSTANCE_PNG_2, BLIP_INSTANCE_JPEG_2):
                    blip_header_size = 33

                if blip_header_size < len(record_data):
                    image_data = record_data[blip_header_size:]

                    # Detect image type
                    detected = detect_image_type(image_data)

                    if detected is None:
                        # Handle metafiles
                        if rec_type == BLIP_TYPE_EMF:
                            detected = ("emf", "image/x-emf")
                        elif rec_type == BLIP_TYPE_WMF:
                            detected = ("wmf", "image/x-wmf")
                        elif rec_type == BLIP_TYPE_DIB:
                            image_data = wrap_dib_as_bmp(image_data)
                            if image_data:
                                detected = ("bmp", "image/bmp")

                    if detected and len(image_data) > 0:
                        # Deduplicate by hash
                        digest = hashlib.sha1(image_data).hexdigest()
                        if digest not in seen_hashes:
                            seen_hashes.add(digest)
                            image_index += 1

                            width, height = get_image_dimensions(
                                image_data, detected[0]
                            )

                            images.append(
                                XlsImage(
                                    image_index=image_index,
                                    content_type=detected[1],
                                    data=image_data,
                                    size_bytes=len(image_data),
                                    width=width,
                                    height=height,
                                )
                            )

            # Skip to next record
            offset += 8 + rec_len
        else:
            offset += 1

    return images
