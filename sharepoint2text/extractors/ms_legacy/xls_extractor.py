"""
XLS content extractor using xlrd and olefile libraries.
"""

import io
import logging
from typing import Any, Dict, Generator, List

import olefile
import xlrd

from sharepoint2text.extractors.data_types import XlsContent, XlsMetadata, XlsSheet

logger = logging.getLogger(__name__)


def _cell_value_to_str(cell: xlrd.sheet.Cell, workbook: xlrd.Book) -> str:
    """Convert a cell value to a string representation."""
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
    """Get the native Python value of a cell for the data dict."""
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
    """Format sheet data as a text table similar to pandas to_string."""
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
    Extract all relevant content from an XLS file.

    Args:
        file_like: A BytesIO object containing the XLS file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        MicrosoftXlsContent dataclass with all extracted content.
    """
    file_like.seek(0)
    sheets = _read_content(file_like=file_like)
    file_like.seek(0)
    metadata = _read_metadata(file_like=file_like)
    metadata.populate_from_path(path)

    full_text = "\n\n".join(sheet.text for sheet in sheets)

    yield XlsContent(
        metadata=metadata,
        sheets=sheets,
        full_text=full_text,
    )
