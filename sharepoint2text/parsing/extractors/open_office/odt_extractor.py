"""
ODT Document Extractor
======================

Extracts text content, metadata, and structure from OpenDocument Text (.odt)
files created by LibreOffice, OpenOffice, and other ODF-compatible applications.

File Format Background
----------------------
ODT files are ZIP archives containing XML files following the OASIS OpenDocument
specification (ISO/IEC 26300). Key components:

    content.xml: Document body (paragraphs, tables, lists, drawings)
    meta.xml: Metadata (title, author, dates, statistics)
    styles.xml: Style definitions, master pages, headers/footers
    settings.xml: Application settings
    Pictures/: Embedded images

Document Structure in content.xml:
    - office:document-content: Root element
    - office:body: Container for document content
    - office:text: Text document body
    - text:p: Paragraphs
    - text:h: Headings (with outline-level attribute)
    - table:table: Tables with rows and cells
    - text:list: Ordered and unordered lists
    - draw:frame: Containers for images and text boxes

XML Namespaces
--------------
The module uses standard ODF namespaces:
    - office: Document structure
    - text: Text content elements
    - table: Table elements
    - draw: Drawing/image elements
    - style: Style definitions
    - meta: Metadata elements
    - dc: Dublin Core metadata
    - xlink: Hyperlink references
    - fo: XSL-FO compatible properties
    - svg: SVG compatible properties

Dependencies
------------
Python Standard Library only:
    - zipfile: ZIP archive handling
    - xml.etree.ElementTree: XML parsing
    - mimetypes: Image content type detection

Extracted Content
-----------------
The extractor retrieves:
    - paragraphs: Text paragraphs with style information and runs
    - tables: Table data as OdtTable objects
    - headers/footers: From styles.xml master pages
    - footnotes/endnotes: Note content with IDs
    - annotations: Comments with creator and date
    - hyperlinks: Link text and URLs
    - bookmarks: Named locations in document
    - images: Embedded images with binary data
    - styles: List of style names used
    - full_text: Complete text in reading order

Special Element Handling
------------------------
ODF uses special elements for whitespace preservation:
    - text:s: Space element (text:c attribute for count)
    - text:tab: Tab character
    - text:line-break: Soft line break

These are converted to appropriate characters during extraction.

Known Limitations
-----------------
- Tracked changes (revisions) are not separately reported
- Text boxes in drawings may not extract all content
- Math formulas are not converted (extracted as-is)
- Nested tables may not preserve complete structure
- Password-protected files are not supported
- Form controls are not extracted

Usage
-----
    >>> import io
    >>> from sharepoint2text.parsing.extractors.open_office.odt_extractor import read_odt
    >>>
    >>> with open("document.odt", "rb") as f:
    ...     for doc in read_odt(io.BytesIO(f.read()), path="document.odt"):
    ...         print(f"Title: {doc.metadata.title}")
    ...         print(f"Creator: {doc.metadata.creator}")
    ...         print(f"Paragraphs: {len(doc.paragraphs)}")
    ...         print(doc.full_text[:500])

See Also
--------
- odp_extractor: For OpenDocument Presentation files
- ods_extractor: For OpenDocument Spreadsheet files
- docx_extractor: For Microsoft Word files

Maintenance Notes
-----------------
- All extraction functions use the shared NS namespace dictionary
- _get_text_recursive handles special whitespace elements
- Headers/footers are in styles.xml, not content.xml
- Images are stored in Pictures/ folder within the ZIP
"""

import io
import logging
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
    odf_length_to_px,
)
from sharepoint2text.parsing.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.parsing.extractors.util.zip_context import ZipContext
from sharepoint2text.parsing.models import (
    Annotation,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    Table,
)

logger = logging.getLogger(__name__)


class _OdtContext(ZipContext):
    """
    Cached context for ODT extraction.

    Opens the ZIP file once and caches all parsed XML documents.
    This avoids repeatedly parsing the same XML files.
    """

    def __init__(self, file_like: io.BytesIO):
        """Initialize the ODT context and cache XML content."""
        super().__init__(file_like)

        # Cache for parsed XML roots
        self._content_root: ET.Element | None = None
        self._meta_root: ET.Element | None = None
        self._styles_root: ET.Element | None = None

        # Parse content.xml
        if "content.xml" in self.namelist:
            self._content_root = self.read_xml_root("content.xml")

        # Parse meta.xml
        if "meta.xml" in self.namelist:
            self._meta_root = self.read_xml_root("meta.xml")

        # Parse styles.xml
        if "styles.xml" in self.namelist:
            self._styles_root = self.read_xml_root("styles.xml")

    @property
    def content_root(self) -> ET.Element | None:
        """Get cached content.xml root.

        Returns:
            Parsed content root element.
        """
        return self._content_root

    @property
    def meta_root(self) -> ET.Element | None:
        """Get cached meta.xml root.

        Returns:
            Parsed metadata root element.
        """
        return self._meta_root

    @property
    def styles_root(self) -> ET.Element | None:
        """Get cached styles.xml root.

        Returns:
            Parsed styles root element.
        """
        return self._styles_root


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
}

_TEXT_SPACE_TAG = f"{{{NS['text']}}}s"
_TEXT_TAB_TAG = f"{{{NS['text']}}}tab"
_TEXT_LINE_BREAK_TAG = f"{{{NS['text']}}}line-break"
_TEXT_NOTE_TAG = f"{{{NS['text']}}}note"
_OFFICE_ANNOTATION_TAG = f"{{{NS['office']}}}annotation"

_TEXT_P_TAG = f"{{{NS['text']}}}p"
_TEXT_H_TAG = f"{{{NS['text']}}}h"
_TEXT_SPAN_TAG = f"{{{NS['text']}}}span"
_TEXT_A_TAG = f"{{{NS['text']}}}a"
_TEXT_SEQUENCE_TAG = f"{{{NS['text']}}}sequence"
_TABLE_TABLE_TAG = f"{{{NS['table']}}}table"
_TABLE_ROW_TAG = f"{{{NS['table']}}}table-row"
_TABLE_CELL_TAG = f"{{{NS['table']}}}table-cell"
_TEXT_LIST_TAG = f"{{{NS['text']}}}list"
_TEXT_LIST_ITEM_TAG = f"{{{NS['text']}}}list-item"
_TEXT_BOOKMARK_TAG = f"{{{NS['text']}}}bookmark"
_TEXT_BOOKMARK_START_TAG = f"{{{NS['text']}}}bookmark-start"

_DRAW_FRAME_TAG = f"{{{NS['draw']}}}frame"
_DRAW_TEXT_BOX_TAG = f"{{{NS['draw']}}}text-box"
_DRAW_IMAGE_TAG = f"{{{NS['draw']}}}image"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"
_ATTR_TEXT_STYLE_NAME = f"{{{NS['text']}}}style-name"
_ATTR_TEXT_OUTLINE_LEVEL = f"{{{NS['text']}}}outline-level"
_ATTR_TEXT_ID = f"{{{NS['text']}}}id"
_ATTR_TEXT_NOTE_CLASS = f"{{{NS['text']}}}note-class"
_ATTR_TEXT_NAME = f"{{{NS['text']}}}name"

_ATTR_XLINK_HREF = f"{{{NS['xlink']}}}href"
_ATTR_STYLE_NAME = f"{{{NS['style']}}}name"

_ATTR_DRAW_NAME = f"{{{NS['draw']}}}name"
_ATTR_SVG_WIDTH = f"{{{NS['svg']}}}width"
_ATTR_SVG_HEIGHT = f"{{{NS['svg']}}}height"

_TEXT_SKIP_TAGS: set[str] = {_TEXT_NOTE_TAG, _OFFICE_ANNOTATION_TAG}

_SVG_TITLE_TAG = f"{{{NS['svg']}}}title"
_SVG_DESC_TAG = f"{{{NS['svg']}}}desc"
_STYLE_STYLE_TAG = f"{{{NS['style']}}}style"


def _get_text_recursive(element: ET.Element) -> str:
    return element_text(
        element,
        text_space_tag=_TEXT_SPACE_TAG,
        text_tab_tag=_TEXT_TAB_TAG,
        text_line_break_tag=_TEXT_LINE_BREAK_TAG,
        attr_text_c=_ATTR_TEXT_C,
        skip_tags=_TEXT_SKIP_TAGS,
    )


def _extract_metadata_from_context(ctx: _OdtContext) -> DocumentMetadata:
    """Extract metadata from cached meta.xml root."""
    return extract_odf_metadata(ctx.meta_root, NS)


def _extract_paragraphs(
    body: ET.Element,
) -> list[tuple[str, int | None, str | None]]:
    """Extract paragraphs from the document body."""
    paragraphs = []

    # Find all paragraphs (text:p) and headings (text:h)
    for elem in body.iter():
        tag = elem.tag
        if tag in (_TEXT_P_TAG, _TEXT_H_TAG):
            text = _get_text_recursive(elem)
            style_name = elem.get(_ATTR_TEXT_STYLE_NAME)
            outline_level = None

            if tag == _TEXT_H_TAG:
                level = elem.get(_ATTR_TEXT_OUTLINE_LEVEL)
                if level:
                    try:
                        outline_level = int(level)
                    except ValueError:
                        pass

            paragraphs.append((text, outline_level, style_name))

    return paragraphs


def _extract_tables(body: ET.Element) -> list[Table]:
    """Extract tables from the document body."""
    tables = []

    for table in body.iter(_TABLE_TABLE_TAG):
        table_data: list[list[str]] = []
        for row in table.iter(_TABLE_ROW_TAG):
            row_data = []
            for cell in row.findall(_TABLE_CELL_TAG):
                cell_texts = [_get_text_recursive(p) for p in cell.iter(_TEXT_P_TAG)]
                row_data.append("\n".join(cell_texts))
            if row_data:
                table_data.append(row_data)
        if table_data:
            tables.append(Table(rows=cast(Any, table_data)))

    return tables


def _extract_hyperlinks(body: ET.Element) -> list[Annotation]:
    """Extract hyperlinks from the document."""
    hyperlinks = []

    for link in body.iter(_TEXT_A_TAG):
        href = link.get(_ATTR_XLINK_HREF, "")
        text = _get_text_recursive(link)
        if href:
            hyperlinks.append(Annotation(kind="hyperlink", text=text, target=href))

    return hyperlinks


def _extract_notes(body: ET.Element) -> tuple[list[Annotation], list[Annotation]]:
    """Extract footnotes and endnotes from the document."""
    footnotes = []
    endnotes = []

    for note in body.iter(_TEXT_NOTE_TAG):
        note_id = note.get(_ATTR_TEXT_ID, "")
        note_class = note.get(_ATTR_TEXT_NOTE_CLASS, "footnote")

        # Get note body text
        note_body = note.find("text:note-body", NS)
        text = ""
        if note_body is not None:
            text_parts = []
            for p in note_body.iter(_TEXT_P_TAG):
                text_parts.append(_get_text_recursive(p))
            text = "\n".join(text_parts)

        note_obj = Annotation(
            kind=note_class,
            text=text,
            properties={"odt.id": note_id},
        )

        if note_class == "endnote":
            endnotes.append(note_obj)
        else:
            footnotes.append(note_obj)

    return footnotes, endnotes


def _extract_annotations(body: ET.Element) -> list[Annotation]:
    """Extract annotations/comments from the document."""
    annotations = []

    for annotation in body.iter(_OFFICE_ANNOTATION_TAG):
        creator_elem = annotation.find("dc:creator", NS)
        creator = creator_elem.text if creator_elem is not None else ""

        date_elem = annotation.find("dc:date", NS)
        date = date_elem.text if date_elem is not None else ""

        # Get annotation text
        text_parts = []
        for p in annotation.iter(_TEXT_P_TAG):
            text_parts.append(_get_text_recursive(p))
        text = "\n".join(text_parts)

        annotations.append(
            Annotation(
                kind="comment",
                author=creator or "",
                text=text,
                properties={"odt.date": date or ""},
            )
        )

    return annotations


def _extract_bookmarks(body: ET.Element) -> list[Annotation]:
    """Extract bookmarks from the document."""
    bookmarks = []

    # Bookmark start elements
    for bookmark in body.iter(_TEXT_BOOKMARK_TAG):
        name = bookmark.get(_ATTR_TEXT_NAME, "")
        if name:
            bookmarks.append(Annotation(kind="bookmark", target=name))

    for bookmark in body.iter(_TEXT_BOOKMARK_START_TAG):
        name = bookmark.get(_ATTR_TEXT_NAME, "")
        if name:
            bookmarks.append(Annotation(kind="bookmark", target=name))

    return bookmarks


def _extract_caption_from_paragraph(para: ET.Element) -> str:
    """Extract caption text from a paragraph containing an image.

    In ODT files with image captions, the paragraph contains both the image frame
    and the caption text. This function extracts just the text content, properly
    handling text:sequence elements (used for auto-numbering like "Illustration 1").
    """
    parts = []

    # Get text before any child elements
    if para.text:
        parts.append(para.text)

    for child in para:
        tag = child.tag

        # Skip image frames - we only want the caption text
        if tag == _DRAW_FRAME_TAG:
            pass
        elif tag == _TEXT_SEQUENCE_TAG:
            # text:sequence elements contain auto-numbers like "1", "2"
            if child.text:
                parts.append(child.text)
        elif tag == _TEXT_SPACE_TAG:
            # Space element
            count = int(child.get(_ATTR_TEXT_C, "1"))
            parts.append(" " * count)
        elif tag == _TEXT_TAB_TAG:
            parts.append("\t")
        elif tag == _TEXT_LINE_BREAK_TAG:
            parts.append("\n")
        else:
            # Other elements - extract their text recursively
            parts.append(_get_text_recursive(child))

        # Get tail text after this element
        if child.tail:
            parts.append(child.tail)

    # Join and clean up whitespace
    caption = "".join(parts).strip()
    # Normalize internal whitespace
    caption = " ".join(caption.split())
    return caption


def _extract_images_from_context(
    ctx: _OdtContext, body: ET.Element
) -> list[ImageAsset]:
    """Extract images from the document using cached context.

    Extracts images with their metadata:
    - caption: From text-box paragraph text, svg:title element, or frame name
    - description: From svg:desc element (alt text)
    - image_index: Sequential index of the image in the document

    ODT files can have images in two formats:
    1. Simple: draw:frame > draw:image (caption from svg:title or frame name)
    2. Captioned: draw:frame > draw:text-box > text:p > draw:frame > draw:image
       (caption is the text content of the containing paragraph)
    """
    images: list[ImageAsset] = []
    image_counter = 0

    # Track which image hrefs we've already processed (to avoid duplicates)
    processed_hrefs: set[str] = set()

    # First, find images inside text-boxes (captioned images)
    for outer_frame in body.iter(_DRAW_FRAME_TAG):
        text_box = outer_frame.find(_DRAW_TEXT_BOX_TAG)
        if text_box is None:
            continue

        # Look for paragraphs in the text-box that contain images
        for para in text_box.iter(_TEXT_P_TAG):
            inner_frame = para.find(_DRAW_FRAME_TAG)
            if inner_frame is None:
                continue

            image_elem = inner_frame.find(_DRAW_IMAGE_TAG)
            if image_elem is None:
                continue

            # Extract image properties from the inner frame
            name = inner_frame.get(_ATTR_DRAW_NAME, "")
            width = inner_frame.get(_ATTR_SVG_WIDTH)
            height = inner_frame.get(_ATTR_SVG_HEIGHT)
            href = image_elem.get(_ATTR_XLINK_HREF, "")

            if not href or href.startswith("http"):
                continue

            # Mark as processed
            processed_hrefs.add(href)

            # Extract caption from the paragraph text
            caption = _extract_caption_from_paragraph(para)

            # Extract description from svg:desc if present
            desc_elem = inner_frame.find(_SVG_DESC_TAG)
            description = (
                desc_elem.text if desc_elem is not None and desc_elem.text else ""
            )

            try:
                if ctx.exists(href):
                    image_counter += 1
                    img_data = ctx.read_bytes(href)
                    images.append(
                        ImageAsset(
                            number=image_counter,
                            filename=name or href.split("/")[-1],
                            media_type=guess_content_type(href),
                            data=img_data,
                            width=odf_length_to_px(width),
                            height=odf_length_to_px(height),
                            caption=caption,
                            description=description,
                            properties={
                                "odt.href": href,
                                "odt.size_bytes": len(img_data),
                            },
                        )
                    )
            except (KeyError, OSError, ValueError) as e:
                logger.debug("Failed to extract image %s: %s", href, e)
                images.append(
                    ImageAsset(
                        number=image_counter + 1,
                        filename=name or href,
                        properties={"odt.href": href, "odt.error": str(e)},
                    )
                )

    # Then, find simple images (not in text-boxes)
    for frame in body.iter(_DRAW_FRAME_TAG):
        # Skip if this is a text-box frame
        if frame.find(_DRAW_TEXT_BOX_TAG) is not None:
            continue

        name = frame.get(_ATTR_DRAW_NAME, "")
        width = frame.get(_ATTR_SVG_WIDTH)
        height = frame.get(_ATTR_SVG_HEIGHT)

        # Extract title (caption) and description from frame
        title_elem = frame.find(_SVG_TITLE_TAG)
        caption = title_elem.text if title_elem is not None and title_elem.text else ""
        if not caption and name:
            caption = name

        desc_elem = frame.find(_SVG_DESC_TAG)
        description = desc_elem.text if desc_elem is not None and desc_elem.text else ""

        image_elem = frame.find(_DRAW_IMAGE_TAG)
        if image_elem is not None:
            href = image_elem.get(_ATTR_XLINK_HREF, "")

            # Skip if already processed
            if href in processed_hrefs:
                continue

            if href and not href.startswith("http"):
                try:
                    if ctx.exists(href):
                        image_counter += 1
                        img_data = ctx.read_bytes(href)
                        images.append(
                            ImageAsset(
                                number=image_counter,
                                filename=name or href.split("/")[-1],
                                media_type=guess_content_type(href),
                                data=img_data,
                                width=odf_length_to_px(width),
                                height=odf_length_to_px(height),
                                caption=caption,
                                description=description,
                                properties={
                                    "odt.href": href,
                                    "odt.size_bytes": len(img_data),
                                },
                            )
                        )
                        processed_hrefs.add(href)
                except (KeyError, OSError, ValueError) as e:
                    logger.debug("Failed to extract image %s: %s", href, e)
                    images.append(
                        ImageAsset(
                            number=image_counter + 1,
                            filename=name or href,
                            properties={"odt.href": href, "odt.error": str(e)},
                        )
                    )
            elif href:
                image_counter += 1
                images.append(
                    ImageAsset(
                        number=image_counter,
                        filename=name or None,
                        width=odf_length_to_px(width),
                        height=odf_length_to_px(height),
                        caption=caption,
                        description=description,
                        properties={"odt.href": href},
                    )
                )
                processed_hrefs.add(href)

    return images


def _extract_headers_footers_from_context(
    ctx: _OdtContext,
) -> tuple[list[Annotation], list[Annotation]]:
    """Extract headers and footers from cached styles.xml root."""
    headers: list[Annotation] = []
    footers: list[Annotation] = []

    root = ctx.styles_root
    if root is None:
        return headers, footers

    # Headers and footers are in master-styles
    master_styles = root.find(".//office:master-styles", NS)
    if master_styles is None:
        return headers, footers

    for master_page in master_styles.findall("style:master-page", NS):
        # Regular header
        header = master_page.find("style:header", NS)
        if header is not None:
            text = _get_text_recursive(header)
            if text.strip():
                headers.append(Annotation(kind="header", text=text))

        # Left header
        header_left = master_page.find("style:header-left", NS)
        if header_left is not None:
            text = _get_text_recursive(header_left)
            if text.strip():
                headers.append(
                    Annotation(
                        kind="header", text=text, properties={"odt.type": "left"}
                    )
                )

        # Regular footer
        footer = master_page.find("style:footer", NS)
        if footer is not None:
            text = _get_text_recursive(footer)
            if text.strip():
                footers.append(Annotation(kind="footer", text=text))

        # Left footer
        footer_left = master_page.find("style:footer-left", NS)
        if footer_left is not None:
            text = _get_text_recursive(footer_left)
            if text.strip():
                footers.append(
                    Annotation(
                        kind="footer", text=text, properties={"odt.type": "left"}
                    )
                )

    return headers, footers


def _extract_styles_from_context(ctx: _OdtContext) -> list[str]:
    """Extract style names from cached content.xml and styles.xml roots."""
    styles = set()

    # Extract from cached content.xml
    if ctx.content_root is not None:
        for style in ctx.content_root.iter(_STYLE_STYLE_TAG):
            name = style.get(_ATTR_STYLE_NAME)
            if name:
                styles.add(name)

    # Extract from cached styles.xml
    if ctx.styles_root is not None:
        for style in ctx.styles_root.iter(_STYLE_STYLE_TAG):
            name = style.get(_ATTR_STYLE_NAME)
            if name:
                styles.add(name)

    return list(styles)


def _append_full_text_from_element(elem: ET.Element, output: list[str]) -> None:
    """Append text from an element to output in document order."""
    tag = elem.tag

    if tag in (_TEXT_P_TAG, _TEXT_H_TAG):
        text = _get_text_recursive(elem)
        if text.strip():
            output.append(text)
        return

    if tag == _TABLE_TABLE_TAG:
        for row in elem.iter(_TABLE_ROW_TAG):
            for cell in row.findall(_TABLE_CELL_TAG):
                for p in cell.iter(_TEXT_P_TAG):
                    text = _get_text_recursive(p)
                    if text.strip():
                        output.append(text)
        return

    if tag == _TEXT_LIST_TAG:
        for item in elem.iter(_TEXT_LIST_ITEM_TAG):
            for p in item.iter(_TEXT_P_TAG):
                text = _get_text_recursive(p)
                if text.strip():
                    output.append(text)
        return

    for child in elem:
        _append_full_text_from_element(child, output)


def _extract_full_text(body: ET.Element) -> str:
    """Extract full text from the document body in reading order."""
    all_text: list[str] = []
    _append_full_text_from_element(body, all_text)
    return "\n".join(all_text)


def _build_units(
    paragraphs: list[tuple[str, int | None, str | None]],
    full_text: str,
    title: str | None,
    images: list[ImageAsset],
    tables: list[Table],
) -> list[ContentUnit]:
    """Build canonical sections from ODT outline levels."""
    units: list[ContentUnit] = []
    base_path = [title] if title else []
    heading_stack: list[tuple[int, str]] = []
    current_path: list[str] = []
    current_level: int | None = None
    lines: list[str] = []
    current_tables: list[Table] = []
    pending_tables: list[Table] = []
    table_index = 0
    in_table_block = False

    def flush() -> None:
        text = "\n".join(line for line in lines if line).strip()
        if not text and not current_tables:
            return
        heading_path = list(base_path)
        for heading in current_path:
            if not heading_path or heading_path[-1] != heading:
                heading_path.append(heading)
        units.append(
            ContentUnit(
                number=len(units) + 1,
                kind="section",
                text=text,
                title=current_path[-1] if current_path else title,
                heading_path=heading_path,
                tables=list(current_tables),
                properties={"odt.outline_level": current_level},
            )
        )

    for text, outline_level, style_name in paragraphs:
        normalized_style = (style_name or "").strip().lower()
        if outline_level is None and normalized_style.startswith(("title", "titel")):
            outline_level = 0
        if outline_level is None:
            is_table_text = (
                normalized_style.startswith("table") or "table_" in normalized_style
            )
            if is_table_text:
                if not in_table_block and table_index < len(tables):
                    pending_tables.append(tables[table_index])
                    table_index += 1
                in_table_block = True
                continue
            in_table_block = False
            if text.strip():
                lines.append(text.strip())
            continue
        flush()
        lines = []
        current_tables = []
        while heading_stack and heading_stack[-1][0] >= outline_level:
            heading_stack.pop()
        heading_stack.append((outline_level, text.strip()))
        current_path = [heading for _, heading in heading_stack if heading]
        current_level = outline_level
        if pending_tables:
            current_tables.extend(pending_tables)
            pending_tables = []
    current_tables.extend(pending_tables)
    flush()

    if not units:
        units = [
            ContentUnit(
                number=1,
                kind="document",
                text=full_text,
                title=title,
                heading_path=base_path,
            )
        ]
    if table_index < len(tables):
        units[-1].tables.extend(tables[table_index:])
    for image in images:
        owner = next(
            (
                unit
                for unit in units
                if (image.caption and image.caption in unit.text)
                or (image.description and image.description in unit.text)
            ),
            next(
                (
                    unit
                    for unit in reversed(units)
                    if unit.properties.get("odt.outline_level") in (1, None)
                ),
                units[-1],
            ),
        )
        image.properties["odt.unit_number"] = owner.number
        owner.images.append(image)
    return units


def read_odt(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[ExtractedDocument, Any, None]:
    """
    Extract all relevant content from an OpenDocument Text (.odt) file.

    Primary entry point for ODT file extraction. Opens the ZIP archive,
    parses content.xml and meta.xml, and extracts text, formatting,
    and embedded content.

    This function uses a generator pattern for API consistency with other
    extractors, even though ODT files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete ODT file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned document source metadata.
        ignore_images: If True, skip image extraction (not applicable for this format).

    Yields:
        ExtractedDocument: Single canonical text document containing:
            - metadata: title, creator, dates, and namespaced properties
            - units: heading-based canonical content units
            - tables: canonical tables
            - headers/footers: From master pages in styles.xml
            - images: canonical image assets with binary data
            - annotations: links, notes, comments, headers, and bookmarks
            - styles: List of style names
            - full_text: Complete document text

    Raises:
        ValueError: If content.xml is missing or document body not found.

    Example:
        >>> import io
        >>> with open("report.odt", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_odt(data, path="report.odt"):
        ...         print(f"Title: {doc.metadata.title}")
        ...         print(f"Tables: {len(doc.tables)}")
        ...         print(f"Images: {len(doc.images)}")

    Performance Notes:
        - ZIP file is opened once and all XML is cached
        - content.xml and styles.xml are parsed once and reused
    """
    try:
        file_like.seek(0)
        if is_odf_encrypted(file_like):
            raise ExtractionFileEncryptedError("ODT is encrypted or password-protected")

        # Create context and load all XML files once
        ctx = _OdtContext(file_like)
        try:
            # Validate content.xml exists
            if ctx.content_root is None:
                raise ExtractionFailedError("Invalid ODT file: content.xml not found")

            # Find the document body
            body = ctx.content_root.find(".//office:body/office:text", NS)
            if body is None:
                raise ExtractionFailedError("Invalid ODT file: document body not found")

            # Extract metadata from cached meta.xml
            metadata = _extract_metadata_from_context(ctx)

            # Extract content from body
            paragraphs = _extract_paragraphs(body)
            tables = _extract_tables(body)
            hyperlinks = _extract_hyperlinks(body)
            footnotes, endnotes = _extract_notes(body)
            annotations = _extract_annotations(body)
            bookmarks = _extract_bookmarks(body)
            images = [] if ignore_images else _extract_images_from_context(ctx, body)
            headers, footers = _extract_headers_footers_from_context(ctx)
            styles = _extract_styles_from_context(ctx)
            full_text = _extract_full_text(body)
        finally:
            ctx.close()

        logger.debug(
            "Extracted ODT: paragraphs=%d, tables=%d, images=%d",
            len(paragraphs),
            len(tables),
            len(images),
        )

        units = _build_units(paragraphs, full_text, metadata.title, images, tables)
        all_annotations = [
            *headers,
            *footers,
            *hyperlinks,
            *footnotes,
            *endnotes,
            *annotations,
            *bookmarks,
        ]
        # Assign annotations to the first unit
        if units and all_annotations:
            units[0].annotations.extend(all_annotations)

        yield ExtractedDocument(
            format="odt",
            source=source_metadata(path),
            metadata=metadata,
            units=units,
            properties={
                "odt.styles": cast(Any, styles),
                "document.full_text": full_text,
            },
        )
    except ExtractionError:
        raise
    except (KeyError, ET.ParseError, OSError, ValueError) as exc:
        raise ExtractionFailedError("Failed to extract ODT file", cause=exc) from exc
