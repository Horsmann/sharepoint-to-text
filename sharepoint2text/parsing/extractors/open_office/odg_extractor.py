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
from typing import Any, Generator

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
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
)

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
_DRAW_FRAME_TAG = f"{{{NS['draw']}}}frame"
_DRAW_TEXT_BOX_TAG = f"{{{NS['draw']}}}text-box"
_DRAW_IMAGE_TAG = f"{{{NS['draw']}}}image"
_SVG_TITLE_TAG = f"{{{NS['svg']}}}title"
_SVG_DESC_TAG = f"{{{NS['svg']}}}desc"

_ATTR_TEXT_C = f"{{{NS['text']}}}c"
_ATTR_XLINK_HREF = f"{{{NS['xlink']}}}href"
_ATTR_DRAW_NAME = f"{{{NS['draw']}}}name"
_ATTR_SVG_WIDTH = f"{{{NS['svg']}}}width"
_ATTR_SVG_HEIGHT = f"{{{NS['svg']}}}height"

_TEXT_SKIP_TAGS: set[str] = {_OFFICE_ANNOTATION_TAG}


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


def _extract_full_text(drawing_root: ET.Element) -> str:
    lines: list[str] = []
    for elem in drawing_root.iter():
        if elem.tag in (_TEXT_H_TAG, _TEXT_P_TAG):
            value = _get_text_recursive(elem).strip()
            if value:
                lines.append(value)
    return "\n".join(lines).strip()


def _extract_images(ctx: ZipContext, drawing_root: ET.Element) -> list[ImageAsset]:
    images: list[ImageAsset] = []
    processed_hrefs: set[str] = set()
    image_counter = 0

    for frame in drawing_root.iter(_DRAW_FRAME_TAG):
        # Skip frames that are primarily text containers
        if frame.find(_DRAW_TEXT_BOX_TAG) is not None:
            continue

        image_elem = frame.find(_DRAW_IMAGE_TAG)
        if image_elem is None:
            continue

        href = image_elem.get(_ATTR_XLINK_HREF, "")
        if not href or href in processed_hrefs:
            continue
        processed_hrefs.add(href)

        name = frame.get(_ATTR_DRAW_NAME, "")
        width = frame.get(_ATTR_SVG_WIDTH)
        height = frame.get(_ATTR_SVG_HEIGHT)

        title_elem = frame.find(_SVG_TITLE_TAG)
        caption = title_elem.text if title_elem is not None and title_elem.text else ""
        if not caption and name:
            caption = name

        desc_elem = frame.find(_SVG_DESC_TAG)
        description = desc_elem.text if desc_elem is not None and desc_elem.text else ""

        image_counter += 1
        if href.startswith("http"):
            images.append(
                ImageAsset(
                    number=image_counter,
                    filename=name or href,
                    width=odf_length_to_px(width),
                    height=odf_length_to_px(height),
                    caption=caption or None,
                    description=description or None,
                    properties={"odg.href": href},
                )
            )
            continue

        try:
            if ctx.exists(href):
                img_data = ctx.read_bytes(href)
                images.append(
                    ImageAsset(
                        number=image_counter,
                        filename=name or href.split("/")[-1],
                        media_type=guess_content_type(href),
                        data=img_data,
                        width=odf_length_to_px(width),
                        height=odf_length_to_px(height),
                        caption=caption or None,
                        description=description or None,
                        properties={"odg.href": href, "odg.size_bytes": len(img_data)},
                    )
                )
            else:
                images.append(
                    ImageAsset(
                        number=image_counter,
                        filename=name or href,
                        width=odf_length_to_px(width),
                        height=odf_length_to_px(height),
                        caption=caption or None,
                        description=description or None,
                        properties={"odg.href": href},
                    )
                )
        except (KeyError, OSError, ValueError) as exc:
            images.append(
                ImageAsset(
                    number=image_counter,
                    filename=name or href,
                    caption=caption or None,
                    description=description or None,
                    properties={"odg.href": href, "odg.error": str(exc)},
                )
            )

    return images


def read_odg(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[ExtractedDocument, Any, None]:
    """
    Extract text, metadata, and basic images from an ODG drawing file.

    Args:
        ignore_images: If True, skip image extraction (not applicable for this format).
    """
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
            images = [] if ignore_images else _extract_images(ctx, drawing)
        finally:
            ctx.close()

        logger.debug(
            "Extracted ODG: characters=%d, images=%d", len(full_text), len(images)
        )
        yield ExtractedDocument(
            format="odg",
            source=source_metadata(path),
            metadata=metadata,
            units=[
                ContentUnit(
                    number=1,
                    kind="document",
                    text=full_text,
                    images=images,
                )
            ],
        )
    except ExtractionError:
        raise
    except (KeyError, ET.ParseError, OSError, ValueError) as exc:
        raise ExtractionFailedError("Failed to extract ODG file", cause=exc) from exc
