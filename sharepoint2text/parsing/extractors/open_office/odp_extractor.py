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
    >>> from sharepoint2text.parsing.extractors.open_office.odp_extractor import read_odp
    >>>
    >>> with open("slides.odp", "rb") as f:
    ...     for ppt in read_odp(io.BytesIO(f.read()), path="slides.odp"):
    ...         print(f"Title: {ppt.metadata.title}")
    ...         for slide in ppt.units:
    ...             print(f"Slide {slide.number}: {slide.title}")
    ...             print(f"  Notes: {slide.properties['odp.notes']}")

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
import re
from typing import Any, Generator, cast

from sharepoint2text.parsing import _defused_xml as ET
from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors._model import source_metadata
from sharepoint2text.parsing.extractors.open_office._shared import (
    element_text,
    extract_odf_metadata,
    guess_content_type,
)
from sharepoint2text.parsing.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.parsing.extractors.util.zip_context import ZipContext
from sharepoint2text.parsing.models import (
    Annotation,
    CellValue,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    JsonValue,
    Table,
)

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
_TEXT_P_TAG = f"{{{NS['text']}}}p"
_DRAW_FRAME_TAG = f"{{{NS['draw']}}}frame"
_DRAW_TEXT_BOX_TAG = f"{{{NS['draw']}}}text-box"
_DRAW_IMAGE_TAG = f"{{{NS['draw']}}}image"
_SVG_TITLE_TAG = f"{{{NS['svg']}}}title"
_SVG_DESC_TAG = f"{{{NS['svg']}}}desc"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"
_ATTR_TEXT_STYLE_NAME = f"{{{NS['text']}}}style-name"
_ATTR_DRAW_NAME = f"{{{NS['draw']}}}name"
_ATTR_SVG_X = f"{{{NS['svg']}}}x"
_ATTR_SVG_Y = f"{{{NS['svg']}}}y"
_ATTR_SVG_WIDTH = f"{{{NS['svg']}}}width"
_ATTR_SVG_HEIGHT = f"{{{NS['svg']}}}height"
_ATTR_XLINK_HREF = f"{{{NS['xlink']}}}href"

_TEXT_SKIP_TAGS: set[str] = {_OFFICE_ANNOTATION_TAG}


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
        """Return the parsed root of the OpenDocument content part.

        Returns:
            Parsed content root element.
        """
        return self._content_root

    @property
    def meta_root(self) -> ET.Element | None:
        """Return the parsed root of the OpenDocument metadata part.

        Returns:
            Parsed metadata root element.
        """
        return self._meta_root


def _get_text_recursive(element: ET.Element) -> str:
    return element_text(
        element,
        text_space_tag=_TEXT_SPACE_TAG,
        text_tab_tag=_TEXT_TAB_TAG,
        text_line_break_tag=_TEXT_LINE_BREAK_TAG,
        attr_text_c=_ATTR_TEXT_C,
        skip_tags=_TEXT_SKIP_TAGS,
    )


def _extract_metadata(meta_root: ET.Element | None) -> DocumentMetadata:
    """Extract metadata from meta.xml."""
    return extract_odf_metadata(meta_root, NS)


def _extract_annotations(element: ET.Element) -> list[Annotation]:
    """Extract annotations/comments from an element."""
    annotations = []

    for annotation in element.iter(_OFFICE_ANNOTATION_TAG):
        creator_elem = annotation.find("dc:creator", NS)
        creator = (
            creator_elem.text if creator_elem is not None and creator_elem.text else ""
        )

        date_elem = annotation.find("dc:date", NS)
        date = date_elem.text if date_elem is not None and date_elem.text else ""

        # Get annotation text
        text_parts = []
        for p in annotation.iter(_TEXT_P_TAG):
            text_parts.append(_get_text_recursive(p))
        text = "\n".join(text_parts)

        annotations.append(
            Annotation(
                kind="comment",
                author=creator or None,
                text=text,
                properties={"odp.date": date} if date else {},
            )
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
            cell_texts = [_get_text_recursive(p) for p in cell.iter(_TEXT_P_TAG)]
            row_data.append("\n".join(cell_texts))
        if row_data:
            table_data.append(row_data)
    return table_data


def _extract_image(
    ctx: _OdpContext,
    frame: ET.Element,
    slide_number: int,
    image_index: int,
) -> ImageAsset | None:
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
    title_elem = frame.find(_SVG_TITLE_TAG)
    title = title_elem.text if title_elem is not None and title_elem.text else ""

    desc_elem = frame.find(_SVG_DESC_TAG)
    desc = desc_elem.text if desc_elem is not None and desc_elem.text else ""

    # Combine title and description with newline separator
    if title and desc:
        description = f"{title}\n{desc}"
    else:
        description = title or desc

    # Find image element
    image_elem = frame.find(_DRAW_IMAGE_TAG)
    if image_elem is None:
        return None

    href = image_elem.get(_ATTR_XLINK_HREF, "")
    if not href:
        return None

    if href.startswith("http"):
        # External image reference
        return ImageAsset(
            number=image_index,
            filename=name or href,
            width=int(round(_parse_odf_length_to_px(width))) or None,
            height=int(round(_parse_odf_length_to_px(height))) or None,
            description=description or None,
            properties={"odp.href": href, "odp.unit_number": slide_number},
        )

    # Internal image reference
    try:
        if ctx.exists(href):
            img_data = ctx.read_bytes(href)
            return ImageAsset(
                number=image_index,
                filename=name or href.split("/")[-1],
                media_type=guess_content_type(href),
                data=img_data,
                width=int(round(_parse_odf_length_to_px(width))) or None,
                height=int(round(_parse_odf_length_to_px(height))) or None,
                description=description or None,
                properties={
                    "odp.href": href,
                    "odp.size_bytes": len(img_data),
                    "odp.unit_number": slide_number,
                },
            )
    except (KeyError, OSError, ValueError) as e:
        logger.debug("Failed to extract image %s: %s", href, e)
        return ImageAsset(
            number=image_index,
            filename=name or href,
            width=int(round(_parse_odf_length_to_px(width))) or None,
            height=int(round(_parse_odf_length_to_px(height))) or None,
            description=description or None,
            properties={
                "odp.href": href,
                "odp.error": str(e),
                "odp.unit_number": slide_number,
            },
        )

    return None


def _extract_slide(
    ctx: _OdpContext,
    page: ET.Element,
    slide_number: int,
    image_counter: int = 0,
    ignore_images: bool = False,
) -> tuple[ContentUnit, int]:
    """Extract content from a single slide (draw:page element).

    Args:
        ctx: The cached ODP context.
        page: The draw:page XML element for this slide.
        slide_number: The 1-based slide number.
        image_counter: The current global image counter across all slides.
        ignore_images: If True, skip image extraction.

    Returns:
        A tuple of (OdpSlide, updated_image_counter).
    """
    slide_name = page.get(_ATTR_DRAW_NAME, "")
    title = ""
    body_text: list[str] = []
    other_text: list[str] = []
    tables: list[Table] = []
    annotations: list[Annotation] = []
    images: list[ImageAsset] = []

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
        text_box = frame.find(_DRAW_TEXT_BOX_TAG)
        if text_box is not None:
            for p in text_box.iter(_TEXT_P_TAG):
                text = _get_text_recursive(p).strip()
                if text:
                    # Check style to determine if it's a title
                    style_name = p.get(_ATTR_TEXT_STYLE_NAME, "")
                    if not found_title and (
                        "Title" in style_name
                        or style_name == "TitleText"
                        or (
                            style_name == ""
                            and not title
                            and not body_text
                            and not other_text
                        )
                    ):
                        title = text
                        found_title = True
                    elif "Body" in style_name or style_name == "BodyText":
                        body_text.append(text)
                    else:
                        other_text.append(text)

            # Extract annotations from text box
            annotations.extend(_extract_annotations(text_box))

        # Check for table
        table = frame.find("table:table", NS)
        if table is not None:
            table_data = _extract_table(table)
            if table_data:
                tables.append(Table(rows=cast(list[list[CellValue]], table_data)))

        # Check for image
        if not ignore_images:
            image = _extract_image(ctx, frame, slide_number, image_counter + 1)
            if image is not None:
                image_counter += 1
                images.append(image)

    # Extract speaker notes
    notes_elem = page.find("presentation:notes", NS)
    if notes_elem is not None:
        for frame in notes_elem.iter(_DRAW_FRAME_TAG):
            text_box = frame.find(_DRAW_TEXT_BOX_TAG)
            if text_box is not None:
                for p in text_box.iter(_TEXT_P_TAG):
                    note_text = _get_text_recursive(p).strip()
                    if note_text:
                        annotations.append(Annotation(kind="note", text=note_text))

    properties: dict[str, JsonValue] = {"odp.slide_number": slide_number}
    if title:
        properties["odp.location"] = [title]
    if slide_name:
        properties["odp.name"] = slide_name
    properties["odp.body_text"] = cast(JsonValue, body_text)
    properties["odp.other_text"] = cast(JsonValue, other_text)
    return (
        ContentUnit(
            number=slide_number,
            kind="slide",
            title=title or None,
            text="\n".join([*body_text, *other_text]),
            images=images,
            tables=tables,
            annotations=annotations,
            properties=properties,
        ),
        image_counter,
    )


def read_odp(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[ExtractedDocument, Any, None]:
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
            returned document source metadata.
        ignore_images: If True, skip image extraction (not applicable for this format).

    Yields:
        ExtractedDocument: Single canonical presentation document containing:
            - metadata: Canonical metadata with title, creator, and dates
            - units: Canonical content units with per-slide content

    Raises:
        ValueError: If content.xml is missing or presentation body not found.

    Example:
        >>> import io
        >>> with open("presentation.odp", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for ppt in read_odp(data, path="presentation.odp"):
        ...         print(f"Slides: {len(ppt.units)}")
        ...         for slide in ppt.units:
        ...             print(f"  {slide.number}: {slide.title}")
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

            slides: list[ContentUnit] = []
            image_counter = 0
            for slide_num, page in enumerate(body.findall("draw:page", NS), start=1):
                slide, image_counter = _extract_slide(
                    ctx, page, slide_num, image_counter, ignore_images
                )
                slides.append(slide)
        finally:
            ctx.close()

        # Populate file metadata from path
        logger.debug(
            "Extracted ODP: slides=%d, images=%d",
            len(slides),
            sum(len(slide.images) for slide in slides),
        )

        slide_texts = [
            "\n".join(part for part in (slide.title, slide.text) if part).strip()
            for slide in slides
        ]
        yield ExtractedDocument(
            format="odp",
            source=source_metadata(path),
            metadata=metadata,
            units=slides,
            properties={
                "document.full_text": "\n".join(text for text in slide_texts if text)
            },
        )
    except ExtractionError:
        raise
    except (KeyError, ET.ParseError, OSError, ValueError) as exc:
        raise ExtractionFailedError("Failed to extract ODP file", cause=exc) from exc
