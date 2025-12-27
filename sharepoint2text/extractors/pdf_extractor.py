import io
import logging
from typing import Any, Generator, List

from pypdf import PdfReader

from sharepoint2text.extractors.data_types import (
    PdfContent,
    PdfImage,
    PdfMetadata,
    PdfPage,
)

logger = logging.getLogger(__name__)


def read_pdf(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PdfContent, Any, None]:
    """
    Extract text and images from a PDF file.

    Args:
        file_like: a loaded binary of the pdf file as file-like object
        path: Optional file path to populate file metadata fields.

    Yields:
        PdfContent dataclass containing extracted content organized by page

    Limitations:
    The extraction assumes readable PDF files. If a PDF consist of images of scanned documents
    this function will not return any meaningful result.
    """
    file_like.seek(0)
    reader = PdfReader(file_like)

    pages = {}
    for page_num, page in enumerate(reader.pages, start=1):
        images = _extract_image_bytes(page)
        pages[page_num] = PdfPage(
            text=page.extract_text() or "",
            images=images,
        )

    metadata = PdfMetadata(total_pages=len(reader.pages))
    metadata.populate_from_path(path)

    yield PdfContent(
        pages=pages,
        metadata=metadata,
    )


def _extract_image_bytes(page) -> List[PdfImage]:
    found_images = []
    if "/XObject" in page.get("/Resources", {}):
        x_objects = page["/Resources"]["/XObject"].get_object()

        img_index = 0
        for obj_name in x_objects:
            obj = x_objects[obj_name]

            if obj.get("/Subtype") == "/Image":
                try:
                    image_data = _extract_image(obj, obj_name, img_index)
                    found_images.append(image_data)
                    img_index += 1
                except Exception as e:
                    logger.warning(
                        f"Silently ignoring - Failed to extract image [{obj_name}] [{img_index}]: %s",
                        e,
                    )
                    img_index += 1
    return found_images


def _extract_image(image_obj, name: str, index: int) -> PdfImage:
    """Extract image data from a PDF image object."""

    width = image_obj.get("/Width", 0)
    height = image_obj.get("/Height", 0)
    color_space = str(image_obj.get("/ColorSpace", "unknown"))
    bits = image_obj.get("/BitsPerComponent", 8)

    # Determine image format based on filter
    filter_type = image_obj.get("/Filter", "")
    if isinstance(filter_type, list):
        filter_type = filter_type[0] if filter_type else ""
    filter_type = str(filter_type)

    # Map filter to format
    format_map = {
        "/DCTDecode": "jpeg",
        "/JPXDecode": "jp2",
        "/FlateDecode": "png",
        "/CCITTFaxDecode": "tiff",
        "/JBIG2Decode": "jbig2",
        "/LZWDecode": "png",
    }
    img_format = format_map.get(filter_type, "raw")

    # Get raw image data
    try:
        data = image_obj.get_data()
    except Exception as e:
        logger.warning("Failed to extract image data: %s", e)
        data = image_obj._data if hasattr(image_obj, "_data") else b""

    return PdfImage(
        index=index,
        name=str(name),
        width=int(width),
        height=int(height),
        color_space=color_space,
        bits_per_component=int(bits),
        filter=filter_type,
        data=data,
        format=img_format,
    )
