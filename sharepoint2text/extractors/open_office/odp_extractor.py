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
import zipfile
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    OdpAnnotation,
    OdpContent,
    OdpImage,
    OdpMetadata,
    OdpSlide,
)
from sharepoint2text.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.extractors.util.zip_bomb import open_zipfile
from sharepoint2text.extractors.util.zip_utils import read_zip_xml_root

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


def _get_text_recursive(element: ET.Element) -> str:
    """Recursively extract all text from an element and its children."""
    parts = []
    if element.text:
        parts.append(element.text)

    for child in element:
        # Handle special elements
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


def _extract_metadata(z: zipfile.ZipFile) -> OdpMetadata:
    """Extract metadata from meta.xml."""
    logger.debug("Extracting ODP metadata")
    metadata = OdpMetadata()

    if "meta.xml" not in z.namelist():
        return metadata

    root = read_zip_xml_root(z, "meta.xml")

    # Find the office:meta element
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


def _extract_annotations(element: ET.Element) -> list[OdpAnnotation]:
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

        annotations.append(OdpAnnotation(creator=creator, date=date, text=text))

    return annotations


def _extract_table(table_elem: ET.Element) -> list[list[str]]:
    """Extract table data from a table element."""
    table_data = []
    for row in table_elem.findall(".//table:table-row", NS):
        row_data = []
        for cell in row.findall(".//table:table-cell", NS):
            # Get all text from paragraphs in the cell
            cell_texts = []
            for p in cell.findall(".//text:p", NS):
                cell_texts.append(_get_text_recursive(p))
            row_data.append("\n".join(cell_texts))
        if row_data:
            table_data.append(row_data)
    return table_data


def _extract_image(
    z: zipfile.ZipFile,
    frame: ET.Element,
    slide_number: int,
    image_index: int,
) -> OdpImage | None:
    """Extract image data from a frame element.

    Extracts images with their metadata:
    - caption: Always empty (ODP slides don't have captions like ODT documents)
    - description: Combined from svg:title and svg:desc elements (with newline separator)
    - image_index: Sequential index of the image in the presentation
    - unit_index: The slide number where the image appears
    """
    # Get frame attributes
    name = frame.get(f"{{{NS['draw']}}}name", "")
    width = frame.get(f"{{{NS['svg']}}}width")
    height = frame.get(f"{{{NS['svg']}}}height")

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

    href = image_elem.get(f"{{{NS['xlink']}}}href", "")
    if not href:
        return None

    if href.startswith("http"):
        # External image reference
        return OdpImage(
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
        if href in z.namelist():
            with z.open(href) as img_file:
                img_data = img_file.read()
                content_type = (
                    mimetypes.guess_type(href)[0] or "application/octet-stream"
                )
                return OdpImage(
                    href=href,
                    name=name or href.split("/")[-1],
                    content_type=content_type,
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
        logger.debug(f"Failed to extract image {href}: {e}")
        return OdpImage(href=href, name=name, error=str(e))

    return None


def _extract_slide(
    z: zipfile.ZipFile,
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
    slide.name = page.get(f"{{{NS['draw']}}}name", "")

    # Collect all frames with their positions for sorting
    frames_with_positions = []

    for frame in page.findall("draw:frame", NS):
        # Get position for ordering (top to bottom, left to right)
        y_pos = frame.get(f"{{{NS['svg']}}}y", "0cm")
        x_pos = frame.get(f"{{{NS['svg']}}}x", "0cm")

        # Parse position (simple parsing, handle cm/in units)
        try:
            y_val = float(y_pos.replace("cm", "").replace("in", "").replace("pt", ""))
        except ValueError:
            y_val = 0.0
        try:
            x_val = float(x_pos.replace("cm", "").replace("in", "").replace("pt", ""))
        except ValueError:
            x_val = 0.0

        frames_with_positions.append((y_val, x_val, frame))

    # Sort frames by position (top to bottom, then left to right)
    frames_with_positions.sort(key=lambda f: (f[0], f[1]))

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
                    style_name = p.get(f"{{{NS['text']}}}style-name", "")
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
        image = _extract_image(z, frame, slide_number, image_counter + 1)
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
            - metadata: OdpMetadata with title, creator, dates
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

        with open_zipfile(file_like, source="read_odp") as z:
            # Extract metadata
            metadata = _extract_metadata(z)

            # Parse content.xml
            if "content.xml" not in z.namelist():
                raise ExtractionFailedError("Invalid ODP file: content.xml not found")

            content_root = read_zip_xml_root(z, "content.xml")

            # Find the presentation body
            body = content_root.find(".//office:body/office:presentation", NS)
            if body is None:
                raise ExtractionFailedError(
                    "Invalid ODP file: presentation body not found"
                )

            # Extract slides
            slides = []
            image_counter = 0
            for slide_num, page in enumerate(body.findall("draw:page", NS), start=1):
                slide, image_counter = _extract_slide(z, page, slide_num, image_counter)
                slides.append(slide)

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
