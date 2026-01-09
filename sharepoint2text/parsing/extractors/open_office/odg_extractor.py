"""
ODG Drawing Extractor
====================

Extracts text content, metadata, and (basic) image information from OpenDocument
Graphics (.odg) files created by LibreOffice Draw, OpenOffice, and other
ODF-compatible applications.

ODG files are ZIP archives containing XML files following the OASIS OpenDocument
specification (ISO/IEC 26300). Key components:

    content.xml: Drawing content (pages, frames, shapes, text boxes)
    meta.xml: Metadata (title, author, dates)
    Pictures/: Embedded images
"""

import io
import logging
import mimetypes
from functools import lru_cache
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    OdgContent,
    OpenDocumentImage,
    OpenDocumentMetadata,
)
from sharepoint2text.parsing.extractors.util.encryption import is_odf_encrypted
from sharepoint2text.parsing.extractors.util.zip_context import ZipContext

logger = logging.getLogger(__name__)


NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
}

_TEXT_SPACE_TAG = f"{{{NS['text']}}}s"
_TEXT_TAB_TAG = f"{{{NS['text']}}}tab"
_TEXT_LINE_BREAK_TAG = f"{{{NS['text']}}}line-break"
_OFFICE_ANNOTATION_TAG = f"{{{NS['office']}}}annotation"

_TEXT_P_TAG = f"{{{NS['text']}}}p"
_TEXT_H_TAG = f"{{{NS['text']}}}h"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"
_ATTR_XLINK_HREF = f"{{{NS['xlink']}}}href"
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
    metadata = OpenDocumentMetadata()

    if meta_root is None:
        return metadata

    meta_elem = meta_root.find(".//office:meta", NS)
    if meta_elem is None:
        return metadata

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


def _extract_full_text(drawing_root: ET.Element) -> str:
    lines: list[str] = []
    for elem in drawing_root.iter():
        if elem.tag in (_TEXT_H_TAG, _TEXT_P_TAG):
            value = _get_text_recursive(elem).strip()
            if value:
                lines.append(value)
    return "\n".join(lines).strip()


def _extract_images(
    ctx: ZipContext, drawing_root: ET.Element
) -> list[OpenDocumentImage]:
    images: list[OpenDocumentImage] = []
    processed_hrefs: set[str] = set()
    image_counter = 0

    for frame in drawing_root.findall(".//draw:frame", NS):
        # Skip frames that are primarily text containers
        if frame.find("draw:text-box", NS) is not None:
            continue

        image_elem = frame.find("draw:image", NS)
        if image_elem is None:
            continue

        href = image_elem.get(_ATTR_XLINK_HREF, "")
        if not href or href in processed_hrefs:
            continue
        processed_hrefs.add(href)

        name = frame.get(_ATTR_DRAW_NAME, "")
        width = frame.get(_ATTR_SVG_WIDTH)
        height = frame.get(_ATTR_SVG_HEIGHT)

        title_elem = frame.find("svg:title", NS)
        caption = title_elem.text if title_elem is not None and title_elem.text else ""
        if not caption and name:
            caption = name

        desc_elem = frame.find("svg:desc", NS)
        description = desc_elem.text if desc_elem is not None and desc_elem.text else ""

        image_counter += 1
        if href.startswith("http"):
            images.append(
                OpenDocumentImage(
                    href=href,
                    name=name or href,
                    width=width,
                    height=height,
                    image_index=image_counter,
                    caption=caption,
                    description=description,
                    unit_name=None,
                )
            )
            continue

        try:
            if ctx.exists(href):
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
                        unit_name=None,
                    )
                )
            else:
                images.append(
                    OpenDocumentImage(
                        href=href,
                        name=name or href,
                        width=width,
                        height=height,
                        image_index=image_counter,
                        caption=caption,
                        description=description,
                        unit_name=None,
                    )
                )
        except Exception as exc:
            images.append(
                OpenDocumentImage(
                    href=href,
                    name=name or href,
                    error=str(exc),
                    image_index=image_counter,
                    caption=caption,
                    description=description,
                    unit_name=None,
                )
            )

    return images


def read_odg(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[OdgContent, Any, None]:
    """Extract text, metadata, and basic images from an ODG drawing file."""
    try:
        file_like.seek(0)
        if is_odf_encrypted(file_like):
            raise ExtractionFileEncryptedError("ODG is encrypted or password-protected")

        ctx = ZipContext(file_like)
        try:
            meta_root = (
                ctx.read_xml_root("meta.xml") if ctx.exists("meta.xml") else None
            )
            content_root = (
                ctx.read_xml_root("content.xml") if ctx.exists("content.xml") else None
            )
            if content_root is None:
                raise ExtractionFailedError("Invalid ODG file: content.xml not found")

            drawing = content_root.find(".//office:body/office:drawing", NS)
            if drawing is None:
                raise ExtractionFailedError("Invalid ODG file: drawing body not found")

            metadata = _extract_metadata(meta_root)
            full_text = _extract_full_text(drawing)
            images = _extract_images(ctx, drawing)
        finally:
            ctx.close()

        metadata.populate_from_path(path)
        yield OdgContent(metadata=metadata, full_text=full_text, images=images)
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract ODG file", cause=exc) from exc
