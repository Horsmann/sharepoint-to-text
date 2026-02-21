"""Shared helpers for OOXML extractors (DOCX/PPTX/XLSX)."""

from typing import Callable
from xml.etree.ElementTree import Element as XmlElement

from sharepoint2text.parsing.extractors.util.zip_context import ZipContext
from sharepoint2text.parsing.extractors.util.zip_utils import parse_relationships

IMAGE_CONTENT_TYPE_MAP = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "tif": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
}

JPEG_SOF_MARKERS = frozenset(
    {
        0xC0,
        0xC1,
        0xC2,
        0xC3,
        0xC5,
        0xC6,
        0xC7,
        0xC9,
        0xCA,
        0xCB,
        0xCD,
        0xCE,
        0xCF,
    }
)


def get_element_text(root: XmlElement | None, tag: str) -> str | None:
    """Extract text from an element if it exists and has text content."""
    if root is None:
        return None
    elem = root.find(tag)
    if elem is not None and elem.text:
        return elem.text
    return None


def get_image_content_type(
    filename: str,
    *,
    fallback_unknown: str = "image/unknown",
    fallback_to_extension: bool = False,
) -> str:
    """Guess MIME type from filename extension."""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext in IMAGE_CONTENT_TYPE_MAP:
        return IMAGE_CONTENT_TYPE_MAP[ext]
    if fallback_to_extension and ext:
        return f"image/{ext}"
    return fallback_unknown


def get_image_pixel_dimensions(image_data: bytes) -> tuple[int | None, int | None]:
    """Best-effort extraction of pixel dimensions from common raster formats."""
    if not image_data:
        return None, None

    # PNG
    if image_data.startswith(b"\x89PNG\r\n\x1a\n") and len(image_data) >= 24:
        width = int.from_bytes(image_data[16:20], "big")
        height = int.from_bytes(image_data[20:24], "big")
        return (width or None, height or None)

    # GIF
    if image_data[:6] in (b"GIF87a", b"GIF89a") and len(image_data) >= 10:
        width = int.from_bytes(image_data[6:8], "little")
        height = int.from_bytes(image_data[8:10], "little")
        return (width or None, height or None)

    # BMP
    if image_data[:2] == b"BM" and len(image_data) >= 26:
        width = int.from_bytes(image_data[18:22], "little", signed=True)
        height = int.from_bytes(image_data[22:26], "little", signed=True)
        return (abs(width) or None, abs(height) or None)

    # JPEG
    if image_data.startswith(b"\xff\xd8"):
        i = 2
        size = len(image_data)
        while i + 4 <= size:
            if image_data[i] != 0xFF:
                i += 1
                continue
            marker = image_data[i + 1]
            if marker in (0xD9, 0xDA):
                break
            length = int.from_bytes(image_data[i + 2 : i + 4], "big")
            if length < 2:
                break
            if marker in JPEG_SOF_MARKERS and i + 2 + length <= size:
                height = int.from_bytes(image_data[i + 5 : i + 7], "big")
                width = int.from_bytes(image_data[i + 7 : i + 9], "big")
                return (width or None, height or None)
            i += 2 + length

    return None, None


def extract_omml_formulas(
    root: XmlElement,
    *,
    omath_para_tag: str,
    omath_tag: str,
    converter: Callable[[XmlElement], str],
) -> list[tuple[str, bool]]:
    """
    Extract formulas from OMML XML.

    Returns tuples of (latex, is_display), where is_display=True means the
    formula came from an oMathPara container.
    """
    formulas: list[tuple[str, bool]] = []
    omath_in_para: set[int] = set()

    for omath_para in root.iter(omath_para_tag):
        omath = omath_para.find(omath_tag)
        if omath is None:
            continue
        omath_in_para.add(id(omath))
        latex = converter(omath)
        if latex.strip():
            formulas.append((latex, True))

    for omath in root.iter(omath_tag):
        if id(omath) in omath_in_para:
            continue
        latex = converter(omath)
        if latex.strip():
            formulas.append((latex, False))

    return formulas


class OOXMLZipContext(ZipContext):
    """Shared ZIP context for OOXML-based formats (DOCX, PPTX, XLSX)."""

    def read_xml_root_if_exists(self, path: str) -> XmlElement | None:
        if not self.exists(path):
            return None
        return self.read_xml_root(path)

    def read_bytes_if_exists(self, path: str) -> bytes | None:
        if not self.exists(path):
            return None
        return self.read_bytes(path)

    def read_text_if_exists(self, path: str) -> str | None:
        if not self.exists(path):
            return None
        return self.read_text(path)

    def read_relationships_if_exists(self, rels_path: str) -> list[dict[str, str]]:
        rels_root = self.read_xml_root_if_exists(rels_path)
        if rels_root is None:
            return []
        return parse_relationships(rels_root)
