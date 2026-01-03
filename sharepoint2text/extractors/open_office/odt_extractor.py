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
    >>> from sharepoint2text.extractors.open_office.odt_extractor import read_odt
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
import mimetypes
from functools import lru_cache
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    OdtBookmark,
    OdtContent,
    OdtHeaderFooter,
    OdtHyperlink,
    OdtNote,
    OdtParagraph,
    OdtRun,
    OdtTable,
    OpenDocumentAnnotation,
    OpenDocumentImage,
    OpenDocumentMetadata,
)
from sharepoint2text.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.extractors.util.zip_context import ZipContext

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
        """Get cached content.xml root."""
        return self._content_root

    @property
    def meta_root(self) -> ET.Element | None:
        """Get cached meta.xml root."""
        return self._meta_root

    @property
    def styles_root(self) -> ET.Element | None:
        """Get cached styles.xml root."""
        return self._styles_root

    def open_file(self, path: str) -> io.BufferedReader:
        """Open a file from the ZIP archive."""
        return self.open_stream(path)


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
_TEXT_SEQUENCE_TAG = f"{{{NS['text']}}}sequence"
_TABLE_TABLE_TAG = f"{{{NS['table']}}}table"
_TEXT_LIST_TAG = f"{{{NS['text']}}}list"

_DRAW_FRAME_TAG = f"{{{NS['draw']}}}frame"

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


@lru_cache(maxsize=512)
def _guess_content_type(path: str) -> str:
    return mimetypes.guess_type(path)[0] or "application/octet-stream"


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
        elif tag == _TEXT_NOTE_TAG:
            # Skip notes in main text extraction.
            pass
        elif tag == _OFFICE_ANNOTATION_TAG:
            # Skip annotations in main text extraction.
            pass
        else:
            parts.append(_get_text_recursive(child))

        tail = child.tail
        if tail:
            parts.append(tail)

    return "".join(parts)


def _extract_metadata_from_context(ctx: _OdtContext) -> OpenDocumentMetadata:
    """Extract metadata from cached meta.xml root."""
    logger.debug("Extracting ODT metadata")
    metadata = OpenDocumentMetadata()

    root = ctx.meta_root
    if root is None:
        return metadata

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


def _extract_paragraphs(body: ET.Element) -> list[OdtParagraph]:
    """Extract paragraphs from the document body."""
    logger.debug("Extracting ODT paragraphs")
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

            # Extract runs (text:span elements)
            runs = []
            for span in elem.iterfind(".//text:span", NS):
                span_text = _get_text_recursive(span)
                span_style = span.get(_ATTR_TEXT_STYLE_NAME)
                runs.append(OdtRun(text=span_text, style_name=span_style))

            paragraphs.append(
                OdtParagraph(
                    text=text,
                    style_name=style_name,
                    outline_level=outline_level,
                    runs=runs,
                )
            )

    return paragraphs


def _extract_tables(body: ET.Element) -> list[OdtTable]:
    """Extract tables from the document body."""
    logger.debug("Extracting ODT tables")
    tables = []

    for table in body.iterfind(".//table:table", NS):
        table_data: list[list[str]] = []
        for row in table.iterfind(".//table:table-row", NS):
            row_data = []
            for cell in row.iterfind(".//table:table-cell", NS):
                cell_texts = [
                    _get_text_recursive(p) for p in cell.iterfind(".//text:p", NS)
                ]
                row_data.append("\n".join(cell_texts))
            if row_data:
                table_data.append(row_data)
        if table_data:
            tables.append(OdtTable(data=table_data))

    return tables


def _extract_hyperlinks(body: ET.Element) -> list[OdtHyperlink]:
    """Extract hyperlinks from the document."""
    logger.debug("Extracting ODT hyperlinks")
    hyperlinks = []

    for link in body.iterfind(".//text:a", NS):
        href = link.get(_ATTR_XLINK_HREF, "")
        text = _get_text_recursive(link)
        if href:
            hyperlinks.append(OdtHyperlink(text=text, url=href))

    return hyperlinks


def _extract_notes(body: ET.Element) -> tuple[list[OdtNote], list[OdtNote]]:
    """Extract footnotes and endnotes from the document."""
    logger.debug("Extracting ODT notes")
    footnotes = []
    endnotes = []

    for note in body.iterfind(".//text:note", NS):
        note_id = note.get(_ATTR_TEXT_ID, "")
        note_class = note.get(_ATTR_TEXT_NOTE_CLASS, "footnote")

        # Get note body text
        note_body = note.find("text:note-body", NS)
        text = ""
        if note_body is not None:
            text_parts = []
            for p in note_body.iterfind(".//text:p", NS):
                text_parts.append(_get_text_recursive(p))
            text = "\n".join(text_parts)

        note_obj = OdtNote(id=note_id, note_class=note_class, text=text)

        if note_class == "endnote":
            endnotes.append(note_obj)
        else:
            footnotes.append(note_obj)

    return footnotes, endnotes


def _extract_annotations(body: ET.Element) -> list[OpenDocumentAnnotation]:
    """Extract annotations/comments from the document."""
    logger.debug("Extracting ODT annotations")
    annotations = []

    for annotation in body.iterfind(".//office:annotation", NS):
        creator_elem = annotation.find("dc:creator", NS)
        creator = creator_elem.text if creator_elem is not None else ""

        date_elem = annotation.find("dc:date", NS)
        date = date_elem.text if date_elem is not None else ""

        # Get annotation text
        text_parts = []
        for p in annotation.iterfind(".//text:p", NS):
            text_parts.append(_get_text_recursive(p))
        text = "\n".join(text_parts)

        annotations.append(
            OpenDocumentAnnotation(creator=creator, date=date, text=text)
        )

    return annotations


def _extract_bookmarks(body: ET.Element) -> list[OdtBookmark]:
    """Extract bookmarks from the document."""
    logger.debug("Extracting ODT bookmarks")
    bookmarks = []

    # Bookmark start elements
    for bookmark in body.iterfind(".//text:bookmark", NS):
        name = bookmark.get(_ATTR_TEXT_NAME, "")
        if name:
            bookmarks.append(OdtBookmark(name=name))

    for bookmark in body.iterfind(".//text:bookmark-start", NS):
        name = bookmark.get(_ATTR_TEXT_NAME, "")
        if name:
            bookmarks.append(OdtBookmark(name=name))

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
) -> list[OpenDocumentImage]:
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
    logger.debug("Extracting ODT images")
    images: list[OpenDocumentImage] = []
    image_counter = 0

    # Track which image hrefs we've already processed (to avoid duplicates)
    processed_hrefs = set()

    # First, find images inside text-boxes (captioned images)
    for outer_frame in body.iterfind(".//draw:frame", NS):
        text_box = outer_frame.find("draw:text-box", NS)
        if text_box is None:
            continue

        # Look for paragraphs in the text-box that contain images
        for para in text_box.iterfind(".//text:p", NS):
            inner_frame = para.find("draw:frame", NS)
            if inner_frame is None:
                continue

            image_elem = inner_frame.find("draw:image", NS)
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
            desc_elem = inner_frame.find("svg:desc", NS)
            description = (
                desc_elem.text if desc_elem is not None and desc_elem.text else ""
            )

            try:
                if ctx.exists(href):
                    image_counter += 1
                    img_data = ctx.read_bytes(href)
                    images.append(
                        OpenDocumentImage(
                            href=href,
                            name=name or href.split("/")[-1],
                            content_type=_guess_content_type(href),
                            data=io.BytesIO(img_data),
                            size_bytes=len(img_data),
                            width=width,
                            height=height,
                            image_index=image_counter,
                            caption=caption,
                            description=description,
                            unit_index=None,
                        )
                    )
            except Exception as e:
                logger.debug("Failed to extract image %s: %s", href, e)
                images.append(
                    OpenDocumentImage(href=href, name=name or href, error=str(e))
                )

    # Then, find simple images (not in text-boxes)
    for frame in body.iterfind(".//draw:frame", NS):
        # Skip if this is a text-box frame
        if frame.find("draw:text-box", NS) is not None:
            continue

        name = frame.get(_ATTR_DRAW_NAME, "")
        width = frame.get(_ATTR_SVG_WIDTH)
        height = frame.get(_ATTR_SVG_HEIGHT)

        # Extract title (caption) and description from frame
        title_elem = frame.find("svg:title", NS)
        caption = title_elem.text if title_elem is not None and title_elem.text else ""
        if not caption and name:
            caption = name

        desc_elem = frame.find("svg:desc", NS)
        description = desc_elem.text if desc_elem is not None and desc_elem.text else ""

        image_elem = frame.find("draw:image", NS)
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
                            OpenDocumentImage(
                                href=href,
                                name=name or href.split("/")[-1],
                                content_type=_guess_content_type(href),
                                data=io.BytesIO(img_data),
                                size_bytes=len(img_data),
                                width=width,
                                height=height,
                                image_index=image_counter,
                                caption=caption,
                                description=description,
                                unit_index=None,
                            )
                        )
                except Exception as e:
                    logger.debug("Failed to extract image %s: %s", href, e)
                    images.append(
                        OpenDocumentImage(href=href, name=name or href, error=str(e))
                    )
            elif href:
                image_counter += 1
                images.append(
                    OpenDocumentImage(
                        href=href,
                        name=name,
                        width=width,
                        height=height,
                        image_index=image_counter,
                        caption=caption,
                        description=description,
                        unit_index=None,
                    )
                )

    return images


def _extract_headers_footers_from_context(
    ctx: _OdtContext,
) -> tuple[list[OdtHeaderFooter], list[OdtHeaderFooter]]:
    """Extract headers and footers from cached styles.xml root."""
    logger.debug("Extracting ODT headers/footers")
    headers = []
    footers = []

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
                headers.append(OdtHeaderFooter(type="header", text=text))

        # Left header
        header_left = master_page.find("style:header-left", NS)
        if header_left is not None:
            text = _get_text_recursive(header_left)
            if text.strip():
                headers.append(OdtHeaderFooter(type="header-left", text=text))

        # Regular footer
        footer = master_page.find("style:footer", NS)
        if footer is not None:
            text = _get_text_recursive(footer)
            if text.strip():
                footers.append(OdtHeaderFooter(type="footer", text=text))

        # Left footer
        footer_left = master_page.find("style:footer-left", NS)
        if footer_left is not None:
            text = _get_text_recursive(footer_left)
            if text.strip():
                footers.append(OdtHeaderFooter(type="footer-left", text=text))

    return headers, footers


def _extract_styles_from_context(ctx: _OdtContext) -> list[str]:
    """Extract style names from cached content.xml and styles.xml roots."""
    logger.debug("Extracting ODT styles")
    styles = set()

    # Extract from cached content.xml
    if ctx.content_root is not None:
        for style in ctx.content_root.iterfind(".//style:style", NS):
            name = style.get(_ATTR_STYLE_NAME)
            if name:
                styles.add(name)

    # Extract from cached styles.xml
    if ctx.styles_root is not None:
        for style in ctx.styles_root.iterfind(".//style:style", NS):
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
        for row in elem.iterfind(".//table:table-row", NS):
            for cell in row.iterfind(".//table:table-cell", NS):
                for p in cell.iterfind(".//text:p", NS):
                    text = _get_text_recursive(p)
                    if text.strip():
                        output.append(text)
        return

    if tag == _TEXT_LIST_TAG:
        for item in elem.iterfind(".//text:list-item", NS):
            for p in item.iterfind(".//text:p", NS):
                text = _get_text_recursive(p)
                if text.strip():
                    output.append(text)
        return

    for child in elem:
        _append_full_text_from_element(child, output)


def _extract_full_text(body: ET.Element) -> str:
    """Extract full text from the document body in reading order."""
    logger.debug("Extracting ODT full text")
    all_text: list[str] = []
    _append_full_text_from_element(body, all_text)
    return "\n".join(all_text)


def read_odt(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdtContent, Any, None]:
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
            returned OdtContent.metadata.

    Yields:
        OdtContent: Single OdtContent object containing:
            - metadata: OpenDocumentMetadata with title, creator, dates, etc.
            - paragraphs: List of OdtParagraph with text and runs
            - tables: List of tables as OdtTable objects
            - headers/footers: From master pages in styles.xml
            - images: List of OpenDocumentImage with binary data
            - hyperlinks: List of OdtHyperlink with text and URL
            - footnotes/endnotes: OdtNote objects
            - annotations: OpenDocumentAnnotation objects with creator and date
            - bookmarks: OdtBookmark objects
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
            images = _extract_images_from_context(ctx, body)
            headers, footers = _extract_headers_footers_from_context(ctx)
            styles = _extract_styles_from_context(ctx)
            full_text = _extract_full_text(body)
        finally:
            ctx.close()

        # Populate file metadata from path
        metadata.populate_from_path(path)

        yield OdtContent(
            metadata=metadata,
            paragraphs=paragraphs,
            tables=tables,
            headers=headers,
            footers=footers,
            images=images,
            hyperlinks=hyperlinks,
            footnotes=footnotes,
            endnotes=endnotes,
            annotations=annotations,
            bookmarks=bookmarks,
            styles=styles,
            full_text=full_text,
        )
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract ODT file", cause=exc) from exc
