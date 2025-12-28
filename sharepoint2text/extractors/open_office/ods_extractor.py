"""
ODS (OpenDocument Spreadsheet) content extractor.

Extracts content from .ods files which are ZIP archives containing XML files.
"""

import io
import logging
import mimetypes
import zipfile
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.extractors.data_types import (
    OdsAnnotation,
    OdsContent,
    OdsImage,
    OdsMetadata,
    OdsSheet,
)

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


def _get_text_recursive(element: ET.Element) -> str:
    """Recursively extract all text from an element and its children."""
    parts = []
    if element.text:
        parts.append(element.text)

    for child in element:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "s":
            # Space element - get count attribute
            count = int(child.get(f"{{{NS['text']}}}c", "1"))
            parts.append(" " * count)
        elif tag == "tab":
            parts.append("\t")
        elif tag == "line-break":
            parts.append("\n")
        elif tag == "annotation":
            # Skip annotations in main text extraction
            pass
        else:
            parts.append(_get_text_recursive(child))

        if child.tail:
            parts.append(child.tail)

    return "".join(parts)


def _extract_metadata(z: zipfile.ZipFile) -> OdsMetadata:
    """Extract metadata from meta.xml."""
    logger.debug("Extracting ODS metadata")
    metadata = OdsMetadata()

    if "meta.xml" not in z.namelist():
        return metadata

    with z.open("meta.xml") as f:
        tree = ET.parse(f)
        root = tree.getroot()

    meta_elem = root.find(".//office:meta", NS)
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


def _extract_cell_value(cell: ET.Element) -> str:
    """Extract the value from a table cell."""
    # Check for value type attribute
    value_type = cell.get(f"{{{NS['office']}}}value-type", "")

    # For numeric values, get the value attribute
    if value_type in ("float", "currency", "percentage"):
        value = cell.get(f"{{{NS['office']}}}value", "")
        if value:
            return value

    # For date/time values
    if value_type == "date":
        value = cell.get(f"{{{NS['office']}}}date-value", "")
        if value:
            return value

    if value_type == "time":
        value = cell.get(f"{{{NS['office']}}}time-value", "")
        if value:
            return value

    # For boolean values
    if value_type == "boolean":
        value = cell.get(f"{{{NS['office']}}}boolean-value", "")
        if value:
            return value

    # For string values or fallback, get text from paragraphs
    text_parts = []
    for p in cell.findall(".//text:p", NS):
        text_parts.append(_get_text_recursive(p))

    return "\n".join(text_parts)


def _extract_annotations(cell: ET.Element) -> list[OdsAnnotation]:
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

        annotations.append(OdsAnnotation(creator=creator, date=date, text=text))

    return annotations


def _extract_images(z: zipfile.ZipFile, table: ET.Element) -> list[OdsImage]:
    """Extract images from a table/sheet."""
    images = []

    for frame in table.findall(".//draw:frame", NS):
        name = frame.get(f"{{{NS['draw']}}}name", "")
        width = frame.get(f"{{{NS['svg']}}}width")
        height = frame.get(f"{{{NS['svg']}}}height")

        image_elem = frame.find("draw:image", NS)
        if image_elem is None:
            continue

        href = image_elem.get(f"{{{NS['xlink']}}}href", "")
        if not href:
            continue

        if href.startswith("http"):
            # External image reference
            images.append(
                OdsImage(
                    href=href,
                    name=name,
                    width=width,
                    height=height,
                )
            )
        else:
            # Internal image reference
            try:
                if href in z.namelist():
                    with z.open(href) as img_file:
                        img_data = img_file.read()
                        content_type = (
                            mimetypes.guess_type(href)[0] or "application/octet-stream"
                        )
                        images.append(
                            OdsImage(
                                href=href,
                                name=name or href.split("/")[-1],
                                content_type=content_type,
                                data=io.BytesIO(img_data),
                                size_bytes=len(img_data),
                                width=width,
                                height=height,
                            )
                        )
            except Exception as e:
                logger.debug(f"Failed to extract image {href}: {e}")
                images.append(OdsImage(href=href, name=name, error=str(e)))

    return images


def _get_column_name(index: int) -> str:
    """Convert column index to Excel-style column name (A, B, ..., Z, AA, AB, ...)."""
    result = ""
    while index >= 0:
        result = chr(ord("A") + (index % 26)) + result
        index = index // 26 - 1
    return result


def _extract_sheet(z: zipfile.ZipFile, table: ET.Element) -> OdsSheet:
    """Extract content from a single sheet (table:table element)."""
    sheet = OdsSheet()

    # Get sheet name
    sheet.name = table.get(f"{{{NS['table']}}}name", "")

    # Extract data rows
    rows_data = []
    all_annotations = []
    text_lines = []

    for row in table.findall("table:table-row", NS):
        row_data = {}
        row_texts = []
        col_index = 0

        # Check for repeated rows
        row_repeat = int(row.get(f"{{{NS['table']}}}number-rows-repeated", "1"))

        for cell in row.findall("table:table-cell", NS):
            # Check for repeated cells
            cell_repeat = int(
                cell.get(f"{{{NS['table']}}}number-columns-repeated", "1")
            )

            cell_value = _extract_cell_value(cell)

            # Extract annotations from cell
            cell_annotations = _extract_annotations(cell)
            all_annotations.extend(cell_annotations)

            # Add value to row data for each repeated cell
            for _ in range(cell_repeat):
                if cell_value:
                    col_name = _get_column_name(col_index)
                    row_data[col_name] = cell_value
                    row_texts.append(cell_value)
                col_index += 1

        # Add row data for each repeated row (but limit to avoid huge empty areas)
        if row_data:
            actual_repeats = min(row_repeat, 1)  # Only add once for data rows
            for _ in range(actual_repeats):
                rows_data.append(row_data.copy())
            if row_texts:
                text_lines.append("\t".join(row_texts))

    sheet.data = rows_data
    sheet.text = "\n".join(text_lines)
    sheet.annotations = all_annotations
    sheet.images = _extract_images(z, table)

    return sheet


def read_ods(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdsContent, Any, None]:
    """
    Extract all relevant content from an ODS file.

    Args:
        file_like: A BytesIO object containing the ODS file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        OdsContent dataclass with all extracted content.
    """
    file_like.seek(0)

    with zipfile.ZipFile(file_like, "r") as z:
        # Extract metadata
        metadata = _extract_metadata(z)

        # Parse content.xml
        if "content.xml" not in z.namelist():
            raise ValueError("Invalid ODS file: content.xml not found")

        with z.open("content.xml") as f:
            content_tree = ET.parse(f)
            content_root = content_tree.getroot()

        # Find the spreadsheet body
        body = content_root.find(".//office:body/office:spreadsheet", NS)
        if body is None:
            raise ValueError("Invalid ODS file: spreadsheet body not found")

        # Extract sheets
        sheets = []
        for table in body.findall("table:table", NS):
            sheet = _extract_sheet(z, table)
            sheets.append(sheet)

    # Populate file metadata from path
    metadata.populate_from_path(path)

    yield OdsContent(
        metadata=metadata,
        sheets=sheets,
    )
