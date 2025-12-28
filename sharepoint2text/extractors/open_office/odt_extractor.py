"""
ODT (OpenDocument Text) content extractor.

Extracts content from .odt files which are ZIP archives containing XML files.
"""

import io
import logging
import mimetypes
import zipfile
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.extractors.data_types import (
    OdtAnnotation,
    OdtBookmark,
    OdtContent,
    OdtHeaderFooter,
    OdtHyperlink,
    OdtImage,
    OdtMetadata,
    OdtNote,
    OdtParagraph,
    OdtRun,
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
        elif tag == "note":
            # Skip notes in main text extraction
            pass
        elif tag == "annotation":
            # Skip annotations in main text extraction
            pass
        else:
            parts.append(_get_text_recursive(child))

        if child.tail:
            parts.append(child.tail)

    return "".join(parts)


def _extract_metadata(z: zipfile.ZipFile) -> OdtMetadata:
    """Extract metadata from meta.xml."""
    logger.debug("Extracting ODT metadata")
    metadata = OdtMetadata()

    if "meta.xml" not in z.namelist():
        return metadata

    with z.open("meta.xml") as f:
        tree = ET.parse(f)
        root = tree.getroot()

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
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

        if tag in ("p", "h"):
            text = _get_text_recursive(elem)
            style_name = elem.get(f"{{{NS['text']}}}style-name")
            outline_level = None

            if tag == "h":
                level = elem.get(f"{{{NS['text']}}}outline-level")
                if level:
                    try:
                        outline_level = int(level)
                    except ValueError:
                        pass

            # Extract runs (text:span elements)
            runs = []
            for span in elem.findall(".//text:span", NS):
                span_text = _get_text_recursive(span)
                span_style = span.get(f"{{{NS['text']}}}style-name")
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


def _extract_tables(body: ET.Element) -> list[list[list[str]]]:
    """Extract tables from the document body."""
    logger.debug("Extracting ODT tables")
    tables = []

    for table in body.findall(".//table:table", NS):
        table_data = []
        for row in table.findall(".//table:table-row", NS):
            row_data = []
            for cell in row.findall(".//table:table-cell", NS):
                # Get all text from paragraphs in the cell
                cell_texts = []
                for p in cell.findall(".//text:p", NS):
                    cell_texts.append(_get_text_recursive(p))
                row_data.append("\n".join(cell_texts))
            if row_data:
                table_data.append(row_data)
        if table_data:
            tables.append(table_data)

    return tables


def _extract_hyperlinks(body: ET.Element) -> list[OdtHyperlink]:
    """Extract hyperlinks from the document."""
    logger.debug("Extracting ODT hyperlinks")
    hyperlinks = []

    for link in body.findall(".//text:a", NS):
        href = link.get(f"{{{NS['xlink']}}}href", "")
        text = _get_text_recursive(link)
        if href:
            hyperlinks.append(OdtHyperlink(text=text, url=href))

    return hyperlinks


def _extract_notes(body: ET.Element) -> tuple[list[OdtNote], list[OdtNote]]:
    """Extract footnotes and endnotes from the document."""
    logger.debug("Extracting ODT notes")
    footnotes = []
    endnotes = []

    for note in body.findall(".//text:note", NS):
        note_id = note.get(f"{{{NS['text']}}}id", "")
        note_class = note.get(f"{{{NS['text']}}}note-class", "footnote")

        # Get note body text
        note_body = note.find("text:note-body", NS)
        text = ""
        if note_body is not None:
            text_parts = []
            for p in note_body.findall(".//text:p", NS):
                text_parts.append(_get_text_recursive(p))
            text = "\n".join(text_parts)

        note_obj = OdtNote(id=note_id, note_class=note_class, text=text)

        if note_class == "endnote":
            endnotes.append(note_obj)
        else:
            footnotes.append(note_obj)

    return footnotes, endnotes


def _extract_annotations(body: ET.Element) -> list[OdtAnnotation]:
    """Extract annotations/comments from the document."""
    logger.debug("Extracting ODT annotations")
    annotations = []

    for annotation in body.findall(".//office:annotation", NS):
        creator_elem = annotation.find("dc:creator", NS)
        creator = creator_elem.text if creator_elem is not None else ""

        date_elem = annotation.find("dc:date", NS)
        date = date_elem.text if date_elem is not None else ""

        # Get annotation text
        text_parts = []
        for p in annotation.findall(".//text:p", NS):
            text_parts.append(_get_text_recursive(p))
        text = "\n".join(text_parts)

        annotations.append(OdtAnnotation(creator=creator, date=date, text=text))

    return annotations


def _extract_bookmarks(body: ET.Element) -> list[OdtBookmark]:
    """Extract bookmarks from the document."""
    logger.debug("Extracting ODT bookmarks")
    bookmarks = []

    # Bookmark start elements
    for bookmark in body.findall(".//text:bookmark", NS):
        name = bookmark.get(f"{{{NS['text']}}}name", "")
        if name:
            bookmarks.append(OdtBookmark(name=name))

    for bookmark in body.findall(".//text:bookmark-start", NS):
        name = bookmark.get(f"{{{NS['text']}}}name", "")
        if name:
            bookmarks.append(OdtBookmark(name=name))

    return bookmarks


def _extract_images(z: zipfile.ZipFile, body: ET.Element) -> list[OdtImage]:
    """Extract images from the document."""
    logger.debug("Extracting ODT images")
    images = []

    for frame in body.findall(".//draw:frame", NS):
        name = frame.get(f"{{{NS['draw']}}}name", "")
        width = frame.get(f"{{{NS['svg']}}}width")
        height = frame.get(f"{{{NS['svg']}}}height")

        # Find image element
        image_elem = frame.find("draw:image", NS)
        if image_elem is not None:
            href = image_elem.get(f"{{{NS['xlink']}}}href", "")

            if href and not href.startswith("http"):
                # Internal image reference
                try:
                    if href in z.namelist():
                        with z.open(href) as img_file:
                            img_data = img_file.read()
                            content_type = (
                                mimetypes.guess_type(href)[0]
                                or "application/octet-stream"
                            )
                            images.append(
                                OdtImage(
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
                    images.append(OdtImage(href=href, name=name, error=str(e)))
            elif href:
                # External image reference
                images.append(
                    OdtImage(
                        href=href,
                        name=name,
                        width=width,
                        height=height,
                    )
                )

    return images


def _extract_headers_footers(
    z: zipfile.ZipFile,
) -> tuple[list[OdtHeaderFooter], list[OdtHeaderFooter]]:
    """Extract headers and footers from styles.xml."""
    logger.debug("Extracting ODT headers/footers")
    headers = []
    footers = []

    if "styles.xml" not in z.namelist():
        return headers, footers

    with z.open("styles.xml") as f:
        tree = ET.parse(f)
        root = tree.getroot()

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


def _extract_styles(z: zipfile.ZipFile) -> list[str]:
    """Extract style names used in the document."""
    logger.debug("Extracting ODT styles")
    styles = set()

    for xml_file in ["content.xml", "styles.xml"]:
        if xml_file not in z.namelist():
            continue

        with z.open(xml_file) as f:
            tree = ET.parse(f)
            root = tree.getroot()

        # Find all style definitions
        for style in root.findall(".//style:style", NS):
            name = style.get(f"{{{NS['style']}}}name")
            if name:
                styles.add(name)

    return list(styles)


def _extract_full_text(body: ET.Element) -> str:
    """Extract full text from the document body in reading order."""
    logger.debug("Extracting ODT full text")
    all_text = []

    def process_element(elem):
        """Process element and extract text in document order."""
        tag = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

        if tag in ("p", "h"):
            text = _get_text_recursive(elem)
            if text.strip():
                all_text.append(text)
        elif tag == "table":
            # Process table cells
            for row in elem.findall(".//table:table-row", NS):
                for cell in row.findall(".//table:table-cell", NS):
                    for p in cell.findall(".//text:p", NS):
                        text = _get_text_recursive(p)
                        if text.strip():
                            all_text.append(text)
        elif tag == "list":
            # Process list items
            for item in elem.findall(".//text:list-item", NS):
                for p in item.findall(".//text:p", NS):
                    text = _get_text_recursive(p)
                    if text.strip():
                        all_text.append(text)
        else:
            # Recurse for container elements
            for child in elem:
                process_element(child)

    process_element(body)
    return "\n".join(all_text)


def read_odt(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdtContent, Any, None]:
    """
    Extract all relevant content from an ODT file.

    Args:
        file_like: A BytesIO object containing the ODT file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        OdtContent dataclass with all extracted content.
    """
    file_like.seek(0)

    with zipfile.ZipFile(file_like, "r") as z:
        # Extract metadata
        metadata = _extract_metadata(z)

        # Parse content.xml
        if "content.xml" not in z.namelist():
            raise ValueError("Invalid ODT file: content.xml not found")

        with z.open("content.xml") as f:
            content_tree = ET.parse(f)
            content_root = content_tree.getroot()

        # Find the document body
        body = content_root.find(".//office:body/office:text", NS)
        if body is None:
            raise ValueError("Invalid ODT file: document body not found")

        # Extract content
        paragraphs = _extract_paragraphs(body)
        tables = _extract_tables(body)
        hyperlinks = _extract_hyperlinks(body)
        footnotes, endnotes = _extract_notes(body)
        annotations = _extract_annotations(body)
        bookmarks = _extract_bookmarks(body)
        images = _extract_images(z, body)
        headers, footers = _extract_headers_footers(z)
        styles = _extract_styles(z)
        full_text = _extract_full_text(body)

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
