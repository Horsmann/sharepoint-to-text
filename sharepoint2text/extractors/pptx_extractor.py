"""
PPTX content extractor using python-pptx library.
"""

import io
import logging
from datetime import datetime
from typing import Any, Generator, List

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

from sharepoint2text.extractors.data_types import (
    PptxContent,
    PPTXImage,
    PptxMetadata,
    PPTXSlide,
)

logger = logging.getLogger(__name__)


def _dt_to_iso(dt: datetime | None) -> str:
    return dt.isoformat() if dt else ""


def read_pptx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PptxContent, Any, None]:
    """
    Extract all relevant content from a PPTX file.

    Args:
        file_like: A BytesIO object containing the PPTX file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        MicrosoftPptxContent dataclass with all extracted content.
    """
    logger.debug("Reading pptx")
    file_like.seek(0)
    prs = Presentation(file_like)

    cp = prs.core_properties
    metadata = PptxMetadata(
        title=cp.title or "",
        subject=cp.subject or "",
        author=cp.author or "",
        last_modified_by=cp.last_modified_by or "",
        created=_dt_to_iso(cp.created),
        modified=_dt_to_iso(cp.modified),
        keywords=cp.keywords or "",
        comments=cp.comments or "",
        category=cp.category or "",
        revision=cp.revision,
    )

    slides_result: List[PPTXSlide] = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        logger.debug(f"Processing slide [{slide_index}]")

        slide_title = ""
        slide_footer = ""
        content_placeholders: List[str] = []
        other_textboxes: List[str] = []
        images: List[PPTXImage] = []

        image_counter = 0

        for shape in slide.shapes:
            # ---------------------------
            # Image extraction
            # ---------------------------
            if shape.shape_type == shape.shape_type.PICTURE:
                try:
                    image = shape.image
                    image_counter += 1

                    images.append(
                        PPTXImage(
                            image_index=image_counter,
                            filename=image.filename,
                            content_type=image.content_type,
                            size_bytes=len(image.blob),
                            blob=image.blob,
                        )
                    )
                except Exception as e:
                    logger.error(e)
                    logger.exception(f"Failed to extract image on slide {slide_index}")
                continue

            # ---------------------------
            # Text extraction
            # ---------------------------
            if not shape.has_text_frame:
                continue

            text = shape.text.strip()
            if not text:
                continue

            if shape.is_placeholder:
                ptype = shape.placeholder_format.type

                if ptype in (
                    PP_PLACEHOLDER.TITLE,
                    PP_PLACEHOLDER.CENTER_TITLE,
                    PP_PLACEHOLDER.VERTICAL_TITLE,
                ):
                    slide_title = text

                elif ptype == PP_PLACEHOLDER.FOOTER:
                    slide_footer = text

                elif ptype in (
                    PP_PLACEHOLDER.BODY,
                    PP_PLACEHOLDER.SUBTITLE,
                    PP_PLACEHOLDER.OBJECT,
                    PP_PLACEHOLDER.VERTICAL_BODY,
                    PP_PLACEHOLDER.VERTICAL_OBJECT,
                    PP_PLACEHOLDER.TABLE,
                ):
                    content_placeholders.append(text)

                else:
                    other_textboxes.append(text)
            else:
                other_textboxes.append(text)

        # Build slide text
        slide_text_parts = []
        if slide_title:
            slide_text_parts.append(slide_title)
        slide_text_parts.extend(content_placeholders)
        slide_text_parts.extend(other_textboxes)
        slide_text = "\n".join(slide_text_parts)

        slides_result.append(
            PPTXSlide(
                slide_number=slide_index,
                title=slide_title,
                footer=slide_footer,
                content_placeholders=content_placeholders,
                other_textboxes=other_textboxes,
                images=images,
                text=slide_text,
            )
        )

    metadata.populate_from_path(path)
    yield PptxContent(metadata=metadata, slides=slides_result)
