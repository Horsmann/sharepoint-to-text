import io
import logging
import typing
from dataclasses import dataclass, field
from typing import Dict, List

from pypdf import PdfReader

from sharepoint2text.extractors.abstract_extractor import ExtractionInterface

logger = logging.getLogger(__name__)


@dataclass
class PdfImage:
    index: int = 0
    name: str = ""
    width: int = 0
    height: int = 0
    color_space: str = ""
    bits_per_component: int = 8
    filter: str = ""
    data: bytes = b""
    format: str = ""


@dataclass
class PdfPage:
    text: str = ""
    images: List[PdfImage] = field(default_factory=list)


@dataclass
class PdfMetadata:
    total_pages: int = 0


@dataclass
class PdfContent(ExtractionInterface):
    pages: Dict[int, PdfPage] = field(default_factory=dict)
    metadata: PdfMetadata = field(default_factory=PdfMetadata)

    def iterator(self) -> typing.Iterator[str]:
        for page_num in sorted(self.pages.keys()):
            yield self.pages[page_num].text

    def get_full_text(self) -> str:
        return "\n".join(self.iterator())


def read_pdf(file_like: io.BytesIO) -> PdfContent:
    """
    Extract text and images from a PDF file.

    Args:
        file_like: a loaded binary of the pdf file as file-like object
    Returns:
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

    return PdfContent(
        pages=pages,
        metadata=PdfMetadata(total_pages=len(reader.pages)),
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


# def save_images(extraction_result: dict, output_dir: str | Path) -> list[str]:
#     """
#     Save extracted images to disk.
#
#     Args:
#         extraction_result: Result from extract_pdf_content()
#         output_dir: Directory to save images
#
#     Returns:
#         List of saved file paths
#     """
#     output_dir = Path(output_dir)
#     output_dir.mkdir(parents=True, exist_ok=True)
#
#     saved_files = []
#     base_name = Path(extraction_result["filename"]).stem
#
#     for page_num, page_data in extraction_result["pages"].items():
#         for img in page_data["images"]:
#             if "error" in img or not img.get("data"):
#                 continue
#
#             ext = img.get("format", "bin")
#             # Use jpg extension for jpeg format
#             if ext == "jpeg":
#                 ext = "jpg"
#
#             filename = f"{base_name}_page{page_num}_img{img['index']}.{ext}"
#             filepath = output_dir / filename
#
#             with open(filepath, "wb") as f:
#                 f.write(img["data"])
#
#             saved_files.append(str(filepath))
#
#     return saved_files
