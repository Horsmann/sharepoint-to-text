"""
XLSX/XLSB Spreadsheet Extractor

Extracts text content and metadata from Microsoft Excel spreadsheet files.

- .xlsx/.xlsm: Parsed with openpyxl.
- .xlsb: Parsed via a lightweight XLSB fallback path (and optional pyxlsb).
"""

import datetime
import io
import logging
import re
import zipfile
from typing import Any, Generator

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    XlsxContent,
    XlsxImage,
    XlsxMetadata,
    XlsxSheet,
)
from sharepoint2text.parsing.extractors.ms_modern.ooxml_shared import (
    OOXMLZipContext,
    get_image_content_type,
    get_image_pixel_dimensions,
)
from sharepoint2text.parsing.extractors.util.encryption import is_ooxml_encrypted
from sharepoint2text.parsing.extractors.util.zip_bomb import validate_zip_bytesio

logger = logging.getLogger(__name__)

# =============================================================================
# XML Namespaces and pre-computed tag names
# =============================================================================

XDR_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

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

ANCHOR_TYPES = (XDR_ONE_CELL_ANCHOR, XDR_TWO_CELL_ANCHOR, XDR_ABSOLUTE_ANCHOR)

# =============================================================================
# Constants
# =============================================================================

EMU_PER_PIXEL = 9525

# Datetime types for isinstance check
_DATETIME_TYPES = (datetime.datetime, datetime.date, datetime.time)
_XLSB_SST_ITEM_RECORD = 19


def _is_xlsb_path(path: str | None) -> bool:
    return bool(path and path.lower().endswith(".xlsb"))


def _iter_xlsb_records(data: bytes):
    """Yield (record_type, payload) from an XLSB stream."""
    i = 0
    n = len(data)
    while i < n:
        # Record type varint
        rt = 0
        shift = 0
        while True:
            if i >= n:
                return
            b = data[i]
            i += 1
            rt |= (b & 0x7F) << shift
            if (b & 0x80) == 0:
                break
            shift += 7
            if shift > 35:
                return

        # Payload length varint
        rl = 0
        shift = 0
        while True:
            if i >= n:
                return
            b = data[i]
            i += 1
            rl |= (b & 0x7F) << shift
            if (b & 0x80) == 0:
                break
            shift += 7
            if shift > 35:
                return

        if rl < 0 or i + rl > n:
            return
        yield rt, data[i : i + rl]
        i += rl


def _extract_xlsb_shared_strings(raw: bytes) -> list[str]:
    """Extract shared strings from xl/sharedStrings.bin in an XLSB ZIP."""
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        if "xl/sharedStrings.bin" not in zf.namelist():
            return []
        sst_data = zf.read("xl/sharedStrings.bin")

    out: list[str] = []
    for record_type, payload in _iter_xlsb_records(sst_data):
        if record_type != _XLSB_SST_ITEM_RECORD or len(payload) < 5:
            continue
        # XLWideString in BrtSSTItem: 1-byte flags + 4-byte char count + UTF-16LE chars
        cch = int.from_bytes(payload[1:5], "little", signed=False)
        if cch <= 0:
            continue
        byte_len = cch * 2
        start = 5
        end = start + byte_len
        if end > len(payload):
            continue
        text = payload[start:end].decode("utf-16le", errors="ignore").strip()
        if text:
            out.append(text)
    return out


def _extract_xlsb_sheet_names(raw: bytes) -> list[str]:
    """Best-effort extraction of sheet names from workbook.bin."""
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        worksheet_entries = [
            name
            for name in zf.namelist()
            if name.startswith("xl/worksheets/sheet") and name.endswith(".bin")
        ]
        default_names = [f"Sheet{i}" for i in range(1, len(worksheet_entries) + 1)]
        if "xl/workbook.bin" not in zf.namelist():
            return default_names or ["Sheet1"]
        workbook_bin = zf.read("xl/workbook.bin")

    # Sheet names are UTF-16LE; pick likely runs and deduplicate in order.
    decoded = workbook_bin.decode("utf-16le", errors="ignore")
    candidates = re.findall(r"[A-Za-z0-9 _().-]{1,64}", decoded)
    names: list[str] = []
    for candidate in candidates:
        name = candidate.strip()
        if not name:
            continue
        if name.startswith("Sheet") or name.lower().endswith("table"):
            if name not in names:
                names.append(name)

    if not names:
        return default_names or ["Sheet1"]
    if default_names and len(names) >= len(default_names):
        return names[: len(default_names)]
    return names


# =============================================================================
# Cell value handling
# =============================================================================


def _get_cell_value(cell_value: Any) -> Any:
    """Convert cell value to appropriate Python type (datetime -> ISO string)."""
    if cell_value is None:
        return None
    if isinstance(cell_value, _DATETIME_TYPES):
        return cell_value.isoformat()
    return cell_value


def _format_value_for_display(value: Any) -> str:
    """Format a value as string for text table display."""
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value)


def _is_cell_non_empty(val: Any) -> bool:
    """Check if a cell value is non-empty."""
    return val is not None and (not isinstance(val, str) or val.strip() != "")


def _is_meaningful_value(val: Any) -> bool:
    """Check if value is meaningful (not None, empty, or 'Unnamed:' placeholder)."""
    if val is None:
        return False
    if isinstance(val, str):
        return bool(val.strip()) and not val.startswith("Unnamed: ")
    return True


# =============================================================================
# Row/column trimming
# =============================================================================


def _find_last_data_column(rows: list[tuple]) -> int:
    """Find the 1-based index of the last column with data."""
    if not rows:
        return 0

    max_col = 0
    for row in rows:
        for i in range(len(row) - 1, -1, -1):
            if _is_cell_non_empty(row[i]):
                max_col = max(max_col, i + 1)
                break
    return max_col


def _find_last_data_row(rows: list[tuple]) -> int:
    """Find the 1-based index of the last row with data."""
    if not rows:
        return 0

    for i in range(len(rows) - 1, -1, -1):
        if any(_is_cell_non_empty(val) for val in rows[i]):
            return i + 1
    return 0


def _is_table_name_row(row: list[Any]) -> bool:
    """Check if row has exactly one meaningful cell (table name pattern)."""
    non_empty = sum(1 for val in row if _is_meaningful_value(val))
    return non_empty == 1 and len(row) > 1


# =============================================================================
# Sheet data extraction
# =============================================================================


def _read_sheet_data(ws: Worksheet) -> tuple[list[dict[str, Any]], list[list[Any]]]:
    """Read sheet data and return (records, all_rows) with trimmed empty cells."""
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

    # Generate headers from first row
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
        records.append(record)
        all_rows.append([_get_cell_value(val) for val in row])

    return records, all_rows


def _format_sheet_as_text(all_rows: list[list[Any]]) -> str:
    """Format sheet data as an aligned text table."""
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
            if len(val) > col_widths[i]:
                col_widths[i] = len(val)
        formatted_rows.append(formatted_row)

    return "\n".join(
        " ".join(val.rjust(col_widths[i]) for i, val in enumerate(row))
        for row in formatted_rows
    )


# =============================================================================
# Metadata extraction
# =============================================================================


def _extract_metadata_from_workbook(wb) -> XlsxMetadata:
    """Extract document metadata from an openpyxl workbook properties object."""
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
    return metadata


def _read_metadata(file_like: io.BytesIO) -> XlsxMetadata:
    """Extract document metadata from XLSX core properties."""
    file_like.seek(0)
    wb = load_workbook(file_like, read_only=True, data_only=True)
    try:
        return _extract_metadata_from_workbook(wb)
    finally:
        wb.close()


# =============================================================================
# Image extraction
# =============================================================================


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
    """Extract all images from XLSX by parsing the ZIP archive directly."""
    images_by_sheet: dict[int, list[XlsxImage]] = {}
    image_counter = 0

    ctx = OOXMLZipContext(file_like)
    try:
        # Build mapping of sheet index to drawing file
        sheet_to_drawing: dict[int, str] = {}
        for sheet_idx in range(len(sheet_names)):
            rels_path = f"xl/worksheets/_rels/sheet{sheet_idx + 1}.xml.rels"
            for rel in ctx.read_relationships_if_exists(rels_path):
                if "drawing" in rel["type"]:
                    sheet_to_drawing[sheet_idx] = _resolve_drawing_path(rel["target"])
                    break

        # Process each drawing file
        for sheet_idx, drawing_path in sheet_to_drawing.items():
            drawing_root = ctx.read_xml_root_if_exists(drawing_path)
            if drawing_root is None:
                continue

            # Parse drawing relationships to get image file paths
            drawing_rels_path = drawing_path.replace(
                "drawings/", "drawings/_rels/"
            ).replace(".xml", ".xml.rels")

            rid_to_image: dict[str, str] = {}
            for rel in ctx.read_relationships_if_exists(drawing_rels_path):
                if "image" in rel["type"]:
                    rid_to_image[rel["id"]] = _resolve_image_path(rel["target"])

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

                        # Get caption and description
                        caption, description = "", ""
                        nvPicPr = pic.find(XDR_NVPICPR)
                        if nvPicPr is not None:
                            cNvPr = nvPicPr.find(XDR_CNVPR)
                            if cNvPr is not None:
                                caption = cNvPr.get("name", "")
                                description = cNvPr.get("descr", "")

                        # Get the blip reference
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
                        image_bytes = ctx.read_bytes_if_exists(image_path)
                        if image_bytes is None:
                            continue
                        filename = image_path.rsplit("/", 1)[-1]

                        if width <= 0 or height <= 0:
                            width, height = get_image_pixel_dimensions(image_bytes)

                        image_counter += 1
                        sheet_images.append(
                            XlsxImage(
                                image_index=image_counter,
                                sheet_index=sheet_idx,
                                filename=filename,
                                content_type=get_image_content_type(filename),
                                data=io.BytesIO(image_bytes),
                                size_bytes=len(image_bytes),
                                width=width,
                                height=height,
                                caption=caption,
                                description=description,
                            )
                        )

                    except (KeyError, ValueError, OSError) as e:
                        logger.debug(f"Failed to extract image from drawing: {e}")

            if sheet_images:
                images_by_sheet[sheet_idx] = sheet_images

    except (zipfile.BadZipFile, KeyError, ValueError, OSError) as e:
        logger.debug(f"Failed to extract images from XLSX: {e}")
    finally:
        ctx.close()

    return images_by_sheet


# =============================================================================
# Content extraction
# =============================================================================


def _read_content(file_like: io.BytesIO) -> list[XlsxSheet]:
    """Read all sheets from XLSX file and extract content."""
    file_like.seek(0)
    raw = file_like.read()

    wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        sheet_names = list(wb.sheetnames)
        sheets = _read_content_from_workbook(wb, sheet_names)
    finally:
        wb.close()

    images_by_sheet = _extract_images_from_zip(io.BytesIO(raw), sheet_names)
    for sheet_idx, sheet_images in images_by_sheet.items():
        if sheet_idx < len(sheets):
            sheets[sheet_idx].images = sheet_images

    return sheets


def _read_content_from_workbook(wb, sheet_names: list[str]) -> list[XlsxSheet]:
    """Read all sheets from an openpyxl workbook and extract content."""
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
    return sheets


# =============================================================================
# Main entry point
# =============================================================================


def _read_xlsb(raw: bytes, path: str | None = None) -> XlsxContent:
    """Extract text from XLSB files with optional pyxlsb support and a fallback."""
    metadata = XlsxMetadata()
    metadata.populate_from_path(path)

    # Prefer pyxlsb when available for row-accurate extraction.
    try:
        import tempfile

        from pyxlsb import open_workbook  # type: ignore

        with tempfile.NamedTemporaryFile(suffix=".xlsb") as tmp:
            tmp.write(raw)
            tmp.flush()

            sheets: list[XlsxSheet] = []
            with open_workbook(tmp.name) as wb:
                sheet_names = list(getattr(wb, "sheets", []) or [])
                for sheet_index, sheet_name in enumerate(sheet_names, start=1):
                    all_rows: list[list[Any]] = []
                    with wb.get_sheet(sheet_index) as sheet:
                        for row in sheet.rows(sparse=False):
                            vals: list[Any] = []
                            for cell in row:
                                value = getattr(cell, "v", None)
                                vals.append(_get_cell_value(value))
                            if any(_is_cell_non_empty(v) for v in vals):
                                all_rows.append(vals)

                    text = _format_sheet_as_text(all_rows) if all_rows else ""
                    sheets.append(
                        XlsxSheet(name=str(sheet_name), data=all_rows, text=text)
                    )

            if sheets:
                return XlsxContent(metadata=metadata, sheets=sheets)
    except (
        ImportError,
        ModuleNotFoundError,
        ValueError,
        OSError,
        zipfile.BadZipFile,
    ) as exc:
        logger.debug("pyxlsb parsing unavailable or failed, using fallback: %s", exc)

    # Fallback path: shared strings extraction.
    sheet_names = _extract_xlsb_sheet_names(raw)
    shared_strings = _extract_xlsb_shared_strings(raw)
    fallback_text = "\n".join(shared_strings).strip()
    fallback_table = [[s] for s in shared_strings] if shared_strings else []
    first_sheet_name = sheet_names[0] if sheet_names else "Sheet1"
    return XlsxContent(
        metadata=metadata,
        sheets=[
            XlsxSheet(
                name=first_sheet_name,
                data=fallback_table,
                text=fallback_text,
                images=[],
            )
        ],
    )


def read_xlsx(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[XlsxContent, Any, None]:
    """
    Extract all relevant content from an Excel .xlsx file.

    Uses a generator pattern for API consistency. XLSX files yield exactly one
    XlsxContent object containing sheets, metadata, and images.

    Args:
        file_like: BytesIO object containing the XLSX file data.
        path: Optional path to the source file for metadata.
        ignore_images: If True, skip image extraction.
    """
    try:
        file_like.seek(0)
        if is_ooxml_encrypted(file_like):
            raise ExtractionFileEncryptedError(
                "XLSX is encrypted or password-protected"
            )

        raw = file_like.read()

        if _is_xlsb_path(path):
            validate_zip_bytesio(io.BytesIO(raw), source="read_xlsb")
            content = _read_xlsb(raw, path=path)
            logger.info(
                "Extracted XLSB: %d sheets, %d total rows",
                len(content.sheets),
                sum(len(sheet.data) for sheet in content.sheets),
            )
            yield content
            return

        validate_zip_bytesio(io.BytesIO(raw), source="read_xlsx")

        wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
        try:
            metadata = _extract_metadata_from_workbook(wb)
            sheet_names = list(wb.sheetnames)
            sheets = _read_content_from_workbook(wb, sheet_names)
        finally:
            wb.close()

        if not ignore_images:
            images_by_sheet = _extract_images_from_zip(io.BytesIO(raw), sheet_names)
            for sheet_idx, sheet_images in images_by_sheet.items():
                if sheet_idx < len(sheets):
                    sheets[sheet_idx].images = sheet_images

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
    except (zipfile.BadZipFile, KeyError, ValueError, OSError) as exc:
        raise ExtractionFailedError("Failed to extract XLSX file", cause=exc) from exc
