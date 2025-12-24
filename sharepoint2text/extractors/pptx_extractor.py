import io
import logging

from pptx import Presentation
from pptx.enum.shapes import PP_PLACEHOLDER

logger = logging.getLogger(__name__)


def read_pptx(file_like: io.BytesIO) -> list[dict]:
    logger.debug("Reading pptx")
    prs = Presentation(file_like)

    result = []

    for slide_index, slide in enumerate(prs.slides, start=1):
        logger.debug(f"Processing slide [{slide_index}]")
        slide_data = {
            "slide_number": slide_index,
            "title": None,
            "footer": None,
            "content_placeholders": [],
            "other_textboxes": [],
        }

        for shape in slide.shapes:
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
                    slide_data["title"] = text

                elif ptype == PP_PLACEHOLDER.FOOTER:
                    slide_data["footer"] = text
                elif ptype in (
                    PP_PLACEHOLDER.BODY,
                    PP_PLACEHOLDER.SUBTITLE,
                    PP_PLACEHOLDER.OBJECT,
                    PP_PLACEHOLDER.VERTICAL_BODY,
                    PP_PLACEHOLDER.VERTICAL_OBJECT,
                    PP_PLACEHOLDER.TABLE,
                ):
                    slide_data["content_placeholders"].append(text)

                elif ptype in (
                    PP_PLACEHOLDER.PICTURE,
                    PP_PLACEHOLDER.CHART,
                    PP_PLACEHOLDER.MEDIA_CLIP,
                    PP_PLACEHOLDER.SLIDE_IMAGE,
                    PP_PLACEHOLDER.BITMAP,
                ):
                    logger.debug(f"Ignoring type [{ptype}]")

                else:
                    slide_data["other_textboxes"].append(text)

            else:
                slide_data["other_textboxes"].append(text)

        result.append(slide_data)

    return result
