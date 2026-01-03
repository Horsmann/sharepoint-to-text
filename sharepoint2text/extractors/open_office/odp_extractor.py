"""
ODP Presentation Extractor
==========================

Extracts text content, metadata, and structure from OpenDocument Presentation
(.odp) files created by LibreOffice Impress, OpenOffice, and other ODF-compatible
applications.

File Format Background
----------------------
ODP files are ZIP archives containing XML files following the OASIS OpenDocument
specification (ISO/IEC 26300). Key components:

    content.xml: Presentation content (slides, frames, shapes)
    meta.xml: Metadata (title, author, dates)
    styles.xml: Style definitions and master pages
    Pictures/: Embedded images

Presentation Structure in content.xml:
    - office:document-content: Root element
    - office:body: Container for content
    - office:presentation: Presentation body
    - draw:page: Individual slides
    - draw:frame: Containers for text boxes, images, tables
    - draw:text-box: Text content container
    - presentation:notes: Speaker notes

Slide Content Model
-------------------
Each slide (draw:page) contains frames positioned by x/y coordinates.
Frames can contain:
    - draw:text-box: Text paragraphs and lists
    - draw:image: Embedded or linked images
    - table:table: Tables with rows and cells

Text is organized in paragraphs (text:p) within text boxes. Paragraph
styles indicate content type (title, body, subtitle, etc.).

Frame Ordering
--------------
Frames are sorted by position (top-to-bottom, left-to-right) to maintain
logical reading order. Position is determined by svg:y and svg:x attributes.

Dependencies
------------
Python Standard Library only:
    - zipfile: ZIP archive handling
    - xml.etree.ElementTree: XML parsing
    - mimetypes: Image content type detection

Extracted Content
-----------------
Per-slide content includes:
    - slide_number: 1-based slide index
    - name: Slide name attribute
    - title: Detected from title-style paragraphs
    - body_text: Content from body-style paragraphs
    - other_text: Text from non-standard frames
    - tables: Table data as nested lists
    - images: Embedded images with binary data
    - notes: Speaker notes
    - annotations: Comments with creator and date

Title Detection
---------------
Title detection uses paragraph style names containing "Title" or matching
"TitleText" exactly. The first qualifying paragraph at the top of the
slide is designated as the title.

Known Limitations
-----------------
- Master slide text is not separately extracted
- Grouped shapes may not extract all text
- Animations and transitions are ignored
- Embedded media (audio/video) is not extracted
- Math formulas are not converted
- Password-protected files are not supported

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.open_office.odp_extractor import read_odp
    >>>
    >>> with open("slides.odp", "rb") as f:
    ...     for ppt in read_odp(io.BytesIO(f.read()), path="slides.odp"):
    ...         print(f"Title: {ppt.metadata.title}")
    ...         for slide in ppt.slides:
    ...             print(f"Slide {slide.slide_number}: {slide.title}")
    ...             print(f"  Notes: {slide.notes}")

See Also
--------
- odt_extractor: For OpenDocument Text files
- ods_extractor: For OpenDocument Spreadsheet files
- pptx_extractor: For Microsoft PowerPoint files

Maintenance Notes
-----------------
- Frame position parsing handles cm/in/pt units
- Style-based title detection may need extension for custom templates
- Speaker notes are in presentation:notes child elements
- Images stored in Pictures/ folder or as external xlink:href references
"""

import io
import logging
import mimetypes
import re
from functools import lru_cache
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    OdpContent,
    OdpSlide,
    OpenDocumentAnnotation,
    OpenDocumentImage,
    OpenDocumentMetadata,
)
from sharepoint2text.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.extractors.util.zip_context import ZipContext

logger = logging.getLogger(__name__)

# ODF namespaces (same as ODT plus presentation namespace)
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
    "presentation": "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
}

_ODF_LENGTH_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*([a-zA-Z]+)?\s*$")

# Namespaced tags/attributes used frequently.
_TEXT_SPACE_TAG = f"{{{NS['text']}}}s"
_TEXT_TAB_TAG = f"{{{NS['text']}}}tab"
_TEXT_LINE_BREAK_TAG = f"{{{NS['text']}}}line-break"
_OFFICE_ANNOTATION_TAG = f"{{{NS['office']}}}annotation"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"
_ATTR_TEXT_STYLE_NAME = f"{{{NS['text']}}}style-name"
_ATTR_DRAW_NAME = f"{{{NS['draw']}}}name"
_ATTR_SVG_X = f"{{{NS['svg']}}}x"
_ATTR_SVG_Y = f"{{{NS['svg']}}}y"
_ATTR_SVG_WIDTH = f"{{{NS['svg']}}}width"
_ATTR_SVG_HEIGHT = f"{{{NS['svg']}}}height"
_ATTR_XLINK_HREF = f"{{{NS['xlink']}}}href"


@lru_cache(maxsize=512)
def _guess_content_type(path: str) -> str:
    return mimetypes.guess_type(path)[0] or "application/octet-stream"


def _parse_odf_length_to_px(value: str | None) -> float:
    """Convert an ODF length string into a comparable pixel float.

    This is used to sort frames into a consistent reading order.
    """
    if not value:
        return 0.0
    match = _ODF_LENGTH_RE.match(value)
    if not match:
        return 0.0

    number = float(match.group(1))
    unit = (match.group(2) or "px").lower()

    # https://www.w3.org/TR/css-values-3/#absolute-lengths (96 dpi)
    if unit == "px":
        return number
    if unit == "in":
        return number * 96.0
    if unit == "cm":
        return (number / 2.54) * 96.0
    if unit == "mm":
        return (number / 25.4) * 96.0
    if unit == "pt":
        return (number / 72.0) * 96.0
    if unit == "pc":  # pica = 12pt
        return ((number * 12.0) / 72.0) * 96.0

    return number


class _OdpContext(ZipContext):
    """Cached context for ODP extraction."""

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
    logger.debug("Extracting ODP metadata")
    metadata = OpenDocumentMetadata()

    # Find the office:meta element
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


def _extract_annotations(element: ET.Element) -> list[OpenDocumentAnnotation]:
    """Extract annotations/comments from an element."""
    annotations = []

    for annotation in element.findall(".//office:annotation", NS):
        creator_elem = annotation.find("dc:creator", NS)
        creator = (
            creator_elem.text if creator_elem is not None and creator_elem.text else ""
        )

        date_elem = annotation.find("dc:date", NS)
        date = date_elem.text if date_elem is not None and date_elem.text else ""

        # Get annotation text
        text_parts = []
        for p in annotation.findall(".//text:p", NS):
            text_parts.append(_get_text_recursive(p))
        text = "\n".join(text_parts)

        annotations.append(
            OpenDocumentAnnotation(creator=creator, date=date, text=text)
        )

    return annotations


def _extract_table(table_elem: ET.Element) -> list[list[str]]:
    """Extract table data from a table element."""
    rows: list[ET.Element] = []
    rows.extend(table_elem.findall("table:table-header-rows/table:table-row", NS))
    rows.extend(table_elem.findall("table:table-row", NS))

    table_data: list[list[str]] = []
    for row in rows:
        row_data: list[str] = []
        for cell in row.findall("table:table-cell", NS):
            cell_texts = [
                _get_text_recursive(p) for p in cell.iterfind(".//text:p", NS)
            ]
            row_data.append("\n".join(cell_texts))
        if row_data:
            table_data.append(row_data)
    return table_data


def _extract_image(
    ctx: _OdpContext,
    frame: ET.Element,
    slide_number: int,
    image_index: int,
) -> OpenDocumentImage | None:
    """Extract image data from a frame element.

    Extracts images with their metadata:
    - caption: Always empty (ODP slides don't have captions like ODT documents)
    - description: Combined from svg:title and svg:desc elements (with newline separator)
    - image_index: Sequential index of the image in the presentation
    - unit_index: The slide number where the image appears
    """
    # Get frame attributes
    name = frame.get(_ATTR_DRAW_NAME, "")
    width = frame.get(_ATTR_SVG_WIDTH)
    height = frame.get(_ATTR_SVG_HEIGHT)

    # Extract title and description from frame
    # ODF uses svg:title and svg:desc elements for accessibility
    # In ODP, we combine title and desc into description (no caption support)
    title_elem = frame.find("svg:title", NS)
    title = title_elem.text if title_elem is not None and title_elem.text else ""

    desc_elem = frame.find("svg:desc", NS)
    desc = desc_elem.text if desc_elem is not None and desc_elem.text else ""

    # Combine title and description with newline separator
    if title and desc:
        description = f"{title}\n{desc}"
    else:
        description = title or desc

    # ODP slides don't have captions like ODT documents
    caption = ""

    # Find image element
    image_elem = frame.find("draw:image", NS)
    if image_elem is None:
        return None

    href = image_elem.get(_ATTR_XLINK_HREF, "")
    if not href:
        return None

    if href.startswith("http"):
        # External image reference
        return OpenDocumentImage(
            href=href,
            name=name,
            width=width,
            height=height,
            image_index=image_index,
            caption=caption,
            description=description,
            unit_index=slide_number,
        )

    # Internal image reference
    try:
        if ctx.exists(href):
            img_data = ctx.read_bytes(href)
            return OpenDocumentImage(
                href=href,
                name=name or href.split("/")[-1],
                content_type=_guess_content_type(href),
                data=io.BytesIO(img_data),
                size_bytes=len(img_data),
                width=width,
                height=height,
                image_index=image_index,
                caption=caption,
                description=description,
                unit_index=slide_number,
            )
    except Exception as e:
        logger.debug("Failed to extract image %s: %s", href, e)
        return OpenDocumentImage(
            href=href,
            name=name or href,
            error=str(e),
            width=width,
            height=height,
            image_index=image_index,
            caption=caption,
            description=description,
            unit_index=slide_number,
        )

    return None


def _extract_slide(
    ctx: _OdpContext,
    page: ET.Element,
    slide_number: int,
    image_counter: int = 0,
) -> tuple[OdpSlide, int]:
    """Extract content from a single slide (draw:page element).

    Args:
        z: The open zipfile containing the presentation.
        page: The draw:page XML element for this slide.
        slide_number: The 1-based slide number.
        image_counter: The current global image counter across all slides.

    Returns:
        A tuple of (OdpSlide, updated_image_counter).
    """
    slide = OdpSlide(slide_number=slide_number)

    # Get slide name
    slide.name = page.get(_ATTR_DRAW_NAME, "")

    # Collect all frames with their positions for sorting
    frames_with_positions: list[tuple[float, float, ET.Element]] = []
    for frame in page.findall("draw:frame", NS):
        y_val = _parse_odf_length_to_px(frame.get(_ATTR_SVG_Y))
        x_val = _parse_odf_length_to_px(frame.get(_ATTR_SVG_X))
        frames_with_positions.append((y_val, x_val, frame))

    # Sort frames by position (top to bottom, then left to right)
    frames_with_positions.sort(key=lambda item: (item[0], item[1]))

    # Track if we've found a title (first text at top of slide)
    found_title = False

    for _, _, frame in frames_with_positions:
        # Check for text box
        text_box = frame.find("draw:text-box", NS)
        if text_box is not None:
            for p in text_box.findall(".//text:p", NS):
                text = _get_text_recursive(p).strip()
                if text:
                    # Check style to determine if it's a title
                    style_name = p.get(_ATTR_TEXT_STYLE_NAME, "")
                    if not found_title and (
                        "Title" in style_name or style_name == "TitleText"
                    ):
                        slide.title = text
                        found_title = True
                    elif "Body" in style_name or style_name == "BodyText":
                        slide.body_text.append(text)
                    else:
                        slide.other_text.append(text)

            # Extract annotations from text box
            annotations = _extract_annotations(text_box)
            slide.annotations.extend(annotations)

        # Check for table
        table = frame.find("table:table", NS)
        if table is not None:
            table_data = _extract_table(table)
            if table_data:
                slide.tables.append(table_data)

        # Check for image
        image = _extract_image(ctx, frame, slide_number, image_counter + 1)
        if image is not None:
            image_counter += 1
            slide.images.append(image)

    # Extract speaker notes
    notes_elem = page.find("presentation:notes", NS)
    if notes_elem is not None:
        for frame in notes_elem.findall(".//draw:frame", NS):
            text_box = frame.find("draw:text-box", NS)
            if text_box is not None:
                for p in text_box.findall(".//text:p", NS):
                    note_text = _get_text_recursive(p).strip()
                    if note_text:
                        slide.notes.append(note_text)

    return slide, image_counter


def read_odp(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdpContent, Any, None]:
    """
    Extract all relevant content from an OpenDocument Presentation (.odp) file.

    Primary entry point for ODP file extraction. Opens the ZIP archive,
    parses content.xml and meta.xml, and extracts slide content organized
    by slide number.

    This function uses a generator pattern for API consistency with other
    extractors, even though ODP files contain exactly one presentation.

    Args:
        file_like: BytesIO object containing the complete ODP file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned OdpContent.metadata.

    Yields:
        OdpContent: Single OdpContent object containing:
            - metadata: OpenDocumentMetadata with title, creator, dates
            - slides: List of OdpSlide objects with per-slide content

    Raises:
        ValueError: If content.xml is missing or presentation body not found.

    Example:
        >>> import io
        >>> with open("presentation.odp", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for ppt in read_odp(data, path="presentation.odp"):
        ...         print(f"Slides: {len(ppt.slides)}")
        ...         for slide in ppt.slides:
        ...             print(f"  {slide.slide_number}: {slide.title}")
    """
    try:
        file_like.seek(0)
        if is_odf_encrypted(file_like):
            raise ExtractionFileEncryptedError("ODP is encrypted or password-protected")

        ctx = _OdpContext(file_like)
        try:
            metadata = _extract_metadata(ctx.meta_root)

            content_root = ctx.content_root
            if content_root is None:
                raise ExtractionFailedError("Invalid ODP file: content.xml not found")

            body = content_root.find(".//office:body/office:presentation", NS)
            if body is None:
                raise ExtractionFailedError(
                    "Invalid ODP file: presentation body not found"
                )

            slides: list[OdpSlide] = []
            image_counter = 0
            for slide_num, page in enumerate(body.findall("draw:page", NS), start=1):
                slide, image_counter = _extract_slide(
                    ctx, page, slide_num, image_counter
                )
                slides.append(slide)
        finally:
            ctx.close()

        # Populate file metadata from path
        metadata.populate_from_path(path)

        yield OdpContent(
            metadata=metadata,
            slides=slides,
        )
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract ODP file", cause=exc) from exc
