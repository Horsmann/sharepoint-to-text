"""
XLS Spreadsheet Extractor

Extracts text content and metadata from legacy Microsoft Excel .xls files
(Excel 97-2003 binary format, BIFF8).

Uses xlrd for cell/sheet parsing and olefile for OLE metadata extraction.
"""

import hashlib
import io
import logging
import struct
from typing import Any, Generator

import olefile
import xlrd

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFileEncryptedError,
    LegacyMicrosoftParsingError,
)
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

# =============================================================================
# Pre-compiled struct for record header parsing (performance optimization)
# =============================================================================

_RECORD_HEADER = struct.Struct("<HHI")  # ver_instance, type, length
_RECORD_HEADER_SIZE = 8

# =============================================================================
# Cell value handling
# =============================================================================

# Cell type constants for quick comparison
_CELL_EMPTY = xlrd.XL_CELL_EMPTY
_CELL_TEXT = xlrd.XL_CELL_TEXT
_CELL_NUMBER = xlrd.XL_CELL_NUMBER
_CELL_DATE = xlrd.XL_CELL_DATE
_CELL_BOOLEAN = xlrd.XL_CELL_BOOLEAN
_CELL_ERROR = xlrd.XL_CELL_ERROR


def _format_date_tuple(dt: tuple) -> str:
    """Format xlrd date tuple as ISO date/datetime string."""
    if dt[3:] == (0, 0, 0):
        return f"{dt[0]:04d}-{dt[1]:02d}-{dt[2]:02d}"
    return f"{dt[0]:04d}-{dt[1]:02d}-{dt[2]:02d} {dt[3]:02d}:{dt[4]:02d}:{dt[5]:02d}"


def _get_cell_value(
    cell: xlrd.sheet.Cell, workbook: xlrd.Book, as_string: bool = False
) -> Any:
    """
    Get cell value as native Python type or string representation.

    Args:
        cell: xlrd Cell object.
        workbook: xlrd Book object (for date conversion).
        as_string: If True, return string representation; else native type.

    Returns:
        Native Python value or string representation based on as_string flag.
    """
    ctype = cell.ctype
    value = cell.value

    if ctype == _CELL_EMPTY:
        return "" if as_string else None

    if ctype == _CELL_TEXT:
        return str(value) if as_string else value

    if ctype == _CELL_NUMBER:
        # Check if it's an integer
        int_val = int(value)
        if value == int_val:
            return str(int_val) if as_string else int_val
        return str(value) if as_string else value

    if ctype == _CELL_DATE:
        try:
            dt = xlrd.xldate_as_tuple(value, workbook.datemode)
            return _format_date_tuple(dt)
        except Exception:
            return str(value) if as_string else value

    if ctype == _CELL_BOOLEAN:
        if as_string:
            return "True" if value else "False"
        return bool(value)

    if ctype == _CELL_ERROR:
        return "#ERROR" if as_string else None

    return str(value) if as_string else value


# =============================================================================
# Sheet formatting
# =============================================================================


def _format_sheet_as_text(headers: list[str], rows: list[list[str]]) -> str:
    """Format sheet data as aligned text table."""
    if not headers and not rows:
        return ""

    all_rows = [headers] + rows if headers else rows
    if not all_rows:
        return ""

    # Calculate column widths in single pass
    num_cols = max(len(row) for row in all_rows)
    col_widths = [0] * num_cols

    for row in all_rows:
        for i, val in enumerate(row):
            width = len(val)
            if width > col_widths[i]:
                col_widths[i] = width

    # Build output with pre-allocated list
    lines = []
    for row in all_rows:
        parts = []
        for i, val in enumerate(row):
            width = col_widths[i] if i < num_cols else len(val)
            parts.append(val.rjust(width))
        lines.append("  ".join(parts))

    return "\n".join(lines)


# =============================================================================
# Content extraction
# =============================================================================


def _read_content(file_like: io.BytesIO) -> list[XlsSheet]:
    """Read all sheets from XLS file and extract content."""
    logger.debug("Reading content")
    workbook = xlrd.open_workbook(file_contents=file_like.read())

    sheets = []
    for sheet in workbook.sheets():
        logger.debug(f"Reading sheet: [{sheet.name}]")

        if sheet.nrows == 0:
            sheets.append(XlsSheet(name=sheet.name, data=[], text=""))
            continue

        nrows = sheet.nrows
        ncols = sheet.ncols

        # Get headers from first row
        headers = [
            _get_cell_value(sheet.cell(0, col), workbook, as_string=True)
            for col in range(ncols)
        ]

        # Build data and text rows
        data: list[dict[str, Any]] = []
        text_rows: list[list[str]] = []

        for row_idx in range(1, nrows):
            row_dict: dict[str, Any] = {}
            row_text: list[str] = []

            for col_idx in range(ncols):
                cell = sheet.cell(row_idx, col_idx)
                header = (
                    headers[col_idx] if col_idx < len(headers) else f"col_{col_idx}"
                )

                row_dict[header] = _get_cell_value(cell, workbook, as_string=False)
                row_text.append(_get_cell_value(cell, workbook, as_string=True))

            data.append(row_dict)
            text_rows.append(row_text)

        sheets.append(
            XlsSheet(
                name=sheet.name,
                data=data,
                text=_format_sheet_as_text(headers, text_rows),
            )
        )

    return sheets


# =============================================================================
# Metadata extraction
# =============================================================================


def _read_metadata(file_like: io.BytesIO) -> XlsMetadata:
    """Extract document metadata from OLE container."""
    with olefile.OleFileIO(file_like) as ole:
        meta = ole.get_metadata()

        def decode(val: bytes | None) -> str:
            return val.decode("utf-8") if val else ""

        return XlsMetadata(
            title=decode(meta.title),
            author=decode(meta.author),
            subject=decode(meta.subject),
            company=decode(meta.company),
            last_saved_by=decode(meta.last_saved_by),
            created=meta.create_time.isoformat() if meta.create_time else "",
            modified=meta.last_saved_time.isoformat() if meta.last_saved_time else "",
        )


# =============================================================================
# Main entry point
# =============================================================================


def read_xls(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[XlsContent, Any, None]:
    """
    Extract content from a legacy Excel .xls file.

    Uses a generator pattern for API consistency. XLS files yield exactly one
    XlsContent object containing sheets, metadata, and images.
    """
    try:
        file_like.seek(0)
        if is_xls_encrypted(file_like):
            raise ExtractionFileEncryptedError("XLS is encrypted or password-protected")

        sheets = _read_content(file_like)

        file_like.seek(0)
        metadata = _read_metadata(file_like)
        metadata.populate_from_path(path)

        file_like.seek(0)
        images = _extract_images_from_workbook(file_like)

        yield XlsContent(
            metadata=metadata,
            sheets=sheets,
            images=images,
            full_text="\n\n".join(sheet.text for sheet in sheets),
        )
    except ExtractionError:
        raise
    except Exception as exc:
        raise LegacyMicrosoftParsingError(
            "Failed to extract XLS file", cause=exc
        ) from exc


# =============================================================================
# Image extraction
# =============================================================================


def _extract_images_from_workbook(file_like: io.BytesIO) -> list[XlsImage]:
    """Extract images from the Workbook stream (BLIP format)."""
    file_like.seek(0)
    if not olefile.isOleFile(file_like):
        return []

    file_like.seek(0)

    try:
        with olefile.OleFileIO(file_like) as ole:
            if not ole.exists("Workbook"):
                return []
            data = ole.openstream("Workbook").read()
    except Exception as e:
        logger.debug(f"Failed to read Workbook stream: {e}")
        return []

    if len(data) < 25:
        return []

    images: list[XlsImage] = []
    seen_hashes: set[str] = set()
    image_index = 0
    offset = 0
    data_len = len(data)

    while offset <= data_len - _RECORD_HEADER_SIZE:
        try:
            ver_instance, rec_type, rec_len = _RECORD_HEADER.unpack_from(data, offset)
        except struct.error:
            offset += 1
            continue

        # Validate record length
        if rec_len <= 0 or rec_len > data_len - offset - _RECORD_HEADER_SIZE:
            offset += 1
            continue

        # Check if this is a BLIP record
        if rec_type not in BLIP_TYPES:
            offset += 1
            continue

        record_data = data[
            offset + _RECORD_HEADER_SIZE : offset + _RECORD_HEADER_SIZE + rec_len
        ]

        if len(record_data) <= 17:
            offset += _RECORD_HEADER_SIZE + rec_len
            continue

        # BLIP header size: 17 bytes normally, 33 with secondary UID
        header_size = 17
        rec_instance = (ver_instance >> 4) & 0x0FFF
        if rec_instance in (BLIP_INSTANCE_PNG_2, BLIP_INSTANCE_JPEG_2):
            header_size = 33

        if header_size >= len(record_data):
            offset += _RECORD_HEADER_SIZE + rec_len
            continue

        image_data = record_data[header_size:]
        detected = detect_image_type(image_data)

        # Handle metafiles and DIB
        if detected is None:
            if rec_type == BLIP_TYPE_EMF:
                detected = ("emf", "image/x-emf")
            elif rec_type == BLIP_TYPE_WMF:
                detected = ("wmf", "image/x-wmf")
            elif rec_type == BLIP_TYPE_DIB:
                image_data = wrap_dib_as_bmp(image_data)
                if image_data:
                    detected = ("bmp", "image/bmp")

        if not detected or not image_data:
            offset += _RECORD_HEADER_SIZE + rec_len
            continue

        # Deduplicate
        digest = hashlib.sha1(image_data).hexdigest()
        if digest not in seen_hashes:
            seen_hashes.add(digest)
            image_index += 1
            width, height = get_image_dimensions(image_data, detected[0])

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

        offset += _RECORD_HEADER_SIZE + rec_len

    return images
