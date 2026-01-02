"""
PPTX Presentation Extractor
===========================

Extracts text content, metadata, and structure from Microsoft PowerPoint .pptx
files (Office Open XML format, PowerPoint 2007 and later).

This module uses direct XML parsing of the pptx ZIP archive structure for all
content extraction, without requiring the python-pptx library.

File Format Background
----------------------
The .pptx format is a ZIP archive containing XML files following the Office
Open XML (OOXML) standard. Key components:

    ppt/presentation.xml: Presentation-level properties and slide ordering
    ppt/slides/slide1.xml, slide2.xml, ...: Individual slide content
    ppt/slides/_rels/slide1.xml.rels: Per-slide relationships (images, etc.)
    ppt/_rels/presentation.xml.rels: Presentation relationships
    ppt/slideLayouts/: Slide layout templates
    ppt/slideMasters/: Master slide definitions
    ppt/comments/comment1.xml, ...: Per-slide comments
    ppt/media/: Embedded images and media
    docProps/core.xml: Metadata (title, author, dates)

XML Namespaces:
    - p: http://schemas.openxmlformats.org/presentationml/2006/main
    - a: http://schemas.openxmlformats.org/drawingml/2006/main
    - m: http://schemas.openxmlformats.org/officeDocument/2006/math
    - r: http://schemas.openxmlformats.org/officeDocument/2006/relationships

Math Formula Handling
---------------------
PowerPoint can contain math formulas in OMML format (same as Word).
This module reuses the OMML-to-LaTeX converter from docx_extractor
to extract formulas as LaTeX notation.

Shape Types and Placeholders
----------------------------
PowerPoint shapes are categorized by type and placeholder function:

Placeholder Types (from p:ph type attribute):
    - title, ctrTitle: Slide titles
    - body, subTitle: Main content areas
    - ftr: Footer text
    - dt: Date/time placeholder
    - sldNum: Slide number placeholder

Text Ordering:
    Shapes are sorted by position (top-to-bottom, left-to-right) to
    maintain a logical reading order in the extracted text.

Extracted Content
-----------------
Per-slide content includes:
    - title: Slide title text
    - content_placeholders: Body text from content areas
    - other_textboxes: Text from non-placeholder shapes
    - tables: Table data as nested lists
    - images: Embedded images with metadata and binary data
    - formulas: Math formulas as LaTeX
    - comments: Slide comments with author and date
    - text: Complete slide text in reading order
    - base_text: Text without formulas/comments/captions

Known Limitations
-----------------
- SmartArt text extraction may be incomplete
- Chart data/labels are not extracted as text
- Grouped shapes may not extract all nested text
- Speaker notes are not currently extracted
- Audio/video content is not extracted
- Password-protected files are not supported
- Very large presentations may use significant memory

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_modern.pptx_extractor import read_pptx
    >>>
    >>> with open("slides.pptx", "rb") as f:
    ...     for ppt in read_pptx(io.BytesIO(f.read()), path="slides.pptx"):
    ...         print(f"Title: {ppt.metadata.title}")
    ...         for slide in ppt.slides:
    ...             print(f"Slide {slide.slide_number}: {slide.title}")
    ...             print(slide.text)

See Also
--------
- Office Open XML specification
- ppt_extractor: For legacy .ppt format

Maintenance Notes
-----------------
- Shape position sorting ensures consistent text order
- Comment extraction requires parsing per-slide comment XML files
- Formula extraction reuses docx_extractor's OMML converter
- Image alt text is extracted from cNvPr descr attribute
"""

import io
import logging
from typing import Any, Generator, List
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    PptxComment,
    PptxContent,
    PptxFormula,
    PptxImage,
    PptxMetadata,
    PptxSlide,
)
from sharepoint2text.extractors.util.encryption import is_ooxml_encrypted
from sharepoint2text.extractors.util.omml_to_latex import omml_to_latex
from sharepoint2text.extractors.util.ooxml_context import OOXMLZipContext
from sharepoint2text.extractors.util.zip_utils import (
    parse_relationships,
)

logger = logging.getLogger(__name__)

# XML Namespaces used in PPTX documents
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CP_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"
DCTERMS_NS = "{http://purl.org/dc/terms/}"

# Pre-computed tag names for hot paths (avoid repeated string concatenation)
P_SP = f"{P_NS}sp"
P_PIC = f"{P_NS}pic"
P_SPTREE = f"{P_NS}spTree"
P_NVSPPR = f"{P_NS}nvSpPr"
P_NVPR = f"{P_NS}nvPr"
P_PH = f"{P_NS}ph"
P_TXBODY = f"{P_NS}txBody"
P_GRAPHICFRAME = f"{P_NS}graphicFrame"
P_CNVPR = f"{P_NS}cNvPr"
P_CM = f"{P_NS}cm"
P_TEXT = f"{P_NS}text"
P_SPPR = f"{P_NS}spPr"
P_XFRM = f"{P_NS}xfrm"
P_SLDID = f"{P_NS}sldId"
P_SLDIDLST = f"{P_NS}sldIdLst"

A_P = f"{A_NS}p"
A_R = f"{A_NS}r"
A_T = f"{A_NS}t"
A_BR = f"{A_NS}br"
A_FLD = f"{A_NS}fld"
A_BLIP = f"{A_NS}blip"
A_XFRM = f"{A_NS}xfrm"
A_OFF = f"{A_NS}off"
A_TBL = f"{A_NS}tbl"
A_TR = f"{A_NS}tr"
A_TC = f"{A_NS}tc"
A_TXBODY = f"{A_NS}txBody"
A_GRAPHICDATA = f"{A_NS}graphicData"

M_OMATH = f"{M_NS}oMath"
M_OMATHPARA = f"{M_NS}oMathPara"

R_ID = f"{R_NS}id"
R_EMBED = f"{R_NS}embed"

# Pre-computed metadata tag names
_DC_TITLE = f"{DC_NS}title"
_DC_CREATOR = f"{DC_NS}creator"
_DC_SUBJECT = f"{DC_NS}subject"
_DC_DESCRIPTION = f"{DC_NS}description"
_CP_KEYWORDS = f"{CP_NS}keywords"
_CP_CATEGORY = f"{CP_NS}category"
_CP_LASTMODIFIEDBY = f"{CP_NS}lastModifiedBy"
_CP_REVISION = f"{CP_NS}revision"
_DCTERMS_CREATED = f"{DCTERMS_NS}created"
_DCTERMS_MODIFIED = f"{DCTERMS_NS}modified"

# Table graphic data URI for PPTX
TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table"

# Title placeholder types
TITLE_TYPES = frozenset({"title", "ctrTitle"})

# Body/content placeholder types
BODY_TYPES = frozenset({"body", "subTitle", "obj", "tbl"})

# Footer-related placeholder types
FOOTER_TYPES = frozenset({"ftr"})

# Placeholder types to skip (not useful for text extraction)
# Note: sldNum (slide number) is NOT skipped - it goes to other_textboxes
SKIP_TYPES = frozenset({"dt", "sldImg", "hdr"})

# Content type mapping by file extension (cached at module level)
_CONTENT_TYPE_MAP = {
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


def _get_image_pixel_dimensions(
    image_data: bytes,
) -> tuple[int | None, int | None]:
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
            if marker in (
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
            ):
                if i + 2 + length <= size:
                    height = int.from_bytes(image_data[i + 5 : i + 7], "big")
                    width = int.from_bytes(image_data[i + 7 : i + 9], "big")
                    return (width or None, height or None)
                break
            i += 2 + length

    return None, None


class _PptxContext(OOXMLZipContext):
    """
    Cached context for PPTX extraction.

    Opens the ZIP file once and caches all parsed XML documents and
    extracted data that is reused across multiple extraction functions.
    This avoids repeatedly opening the ZIP and parsing the same XML files.
    """

    def __init__(self, file_like: io.BytesIO):
        """Initialize the PPTX context and cache XML content."""
        super().__init__(file_like)

        # Cache for parsed XML roots
        self._core_root: ET.Element | None = None
        self._presentation_root: ET.Element | None = None
        self._presentation_rels_root: ET.Element | None = None

        # Cache for slide-related XML (keyed by path)
        self._slide_roots: dict[str, ET.Element] = {}
        self._slide_rels_roots: dict[str, ET.Element] = {}
        self._comment_roots: dict[str, ET.Element] = {}

        # Cache for extracted data
        self._slide_order: list[str] | None = None
        self._slide_relationships: dict[str, dict[str, dict[str, str]]] = {}

        self._load_xml_files()

    def _load_xml_files(self) -> None:
        """Load and parse all XML files from the ZIP at once."""
        # Core properties (metadata)
        if "docProps/core.xml" in self._namelist:
            self._core_root = self.read_xml_root("docProps/core.xml")

        # Presentation XML
        if "ppt/presentation.xml" in self._namelist:
            self._presentation_root = self.read_xml_root("ppt/presentation.xml")

        # Presentation relationships
        rels_path = "ppt/_rels/presentation.xml.rels"
        if rels_path in self._namelist:
            self._presentation_rels_root = self.read_xml_root(rels_path)

        # Pre-compute slide order so we know which slides to load
        self._slide_order = self._compute_slide_order()

        # Load all slide XML files
        for slide_path in self._slide_order:
            if slide_path in self._namelist:
                self._slide_roots[slide_path] = self.read_xml_root(slide_path)

                # Load slide relationships
                slide_dir = "/".join(slide_path.rsplit("/", 1)[:-1])
                slide_name = slide_path.rsplit("/", 1)[-1]
                rels_path = f"{slide_dir}/_rels/{slide_name}.rels"
                if rels_path in self._namelist:
                    self._slide_rels_roots[slide_path] = self.read_xml_root(rels_path)

        # Load all comment files
        for name in self._namelist:
            if name.startswith("ppt/comments/comment") and name.endswith(".xml"):
                self._comment_roots[name] = self.read_xml_root(name)

    def _compute_slide_order(self) -> list[str]:
        """Compute slide order from cached presentation XML."""
        if self._presentation_rels_root is None or self._presentation_root is None:
            return []

        # Build relationship map from rels
        rels_map: dict[str, str] = {}
        for rel in parse_relationships(self._presentation_rels_root):
            rel_id = rel["id"]
            target = rel["target"]
            rel_type = rel["type"].lower()
            if rel_id and target and "slide" in rel_type:
                if target.startswith("slides/"):
                    full_path = f"ppt/{target}"
                elif target.startswith("../"):
                    full_path = target.replace("../", "ppt/")
                else:
                    full_path = f"ppt/{target}"
                rels_map[rel_id] = full_path

        # Get slide order from presentation.xml
        slide_paths: list[str] = []
        sld_id_lst = self._presentation_root.find(f".//{P_SLDIDLST}")
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall(P_SLDID):
                if (r_id := sld_id.get(R_ID)) and r_id in rels_map:
                    slide_paths.append(rels_map[r_id])

        return slide_paths

    @property
    def slide_order(self) -> list[str]:
        """Get ordered list of slide paths."""
        if self._slide_order is None:
            self._slide_order = self._compute_slide_order()
        return self._slide_order

    def get_slide_root(self, slide_path: str) -> ET.Element | None:
        """Get cached slide XML root."""
        return self._slide_roots.get(slide_path)

    def get_slide_relationships(self, slide_path: str) -> dict[str, dict[str, str]]:
        """Get cached relationships for a slide."""
        if slide_path in self._slide_relationships:
            return self._slide_relationships[slide_path]

        relationships = {}
        rels_root = self._slide_rels_roots.get(slide_path)
        if rels_root is not None:
            for rel in parse_relationships(rels_root):
                rel_id = rel["id"]
                if not rel_id:
                    continue
                relationships[rel_id] = {
                    "type": rel["type"],
                    "target": rel["target"],
                }

        self._slide_relationships[slide_path] = relationships
        return relationships

    def get_comment_root(self, slide_number: int) -> ET.Element | None:
        """Get cached comment XML root for a slide."""
        comment_file = f"ppt/comments/comment{slide_number}.xml"
        return self._comment_roots.get(comment_file)

    def get_image_data(self, image_path: str) -> bytes | None:
        """Read image data from the ZIP file."""
        if image_path not in self._namelist:
            return None
        return self.read_bytes(image_path)


def _get_element_text(root: ET.Element | None, tag: str) -> str | None:
    """Extract text from an element if it exists and has text content."""
    if root is None:
        return None
    elem = root.find(tag)
    if elem is not None and elem.text:
        return elem.text
    return None


def _extract_metadata_from_context(ctx: _PptxContext) -> PptxMetadata:
    """
    Extract presentation metadata from cached core.xml root.

    Args:
        ctx: PptxContext with cached XML roots.

    Returns:
        PptxMetadata object with title, author, dates, revision, etc.
    """
    metadata = PptxMetadata()
    root = ctx._core_root
    if root is None:
        return metadata

    # Extract metadata fields using helper
    if text := _get_element_text(root, _DC_TITLE):
        metadata.title = text
    if text := _get_element_text(root, _DC_CREATOR):
        metadata.author = text
    if text := _get_element_text(root, _DC_SUBJECT):
        metadata.subject = text
    if text := _get_element_text(root, _CP_KEYWORDS):
        metadata.keywords = text
    if text := _get_element_text(root, _CP_CATEGORY):
        metadata.category = text
    if text := _get_element_text(root, _DC_DESCRIPTION):
        metadata.comments = text
    if text := _get_element_text(root, _DCTERMS_CREATED):
        metadata.created = text.rstrip("Z")
    if text := _get_element_text(root, _DCTERMS_MODIFIED):
        metadata.modified = text.rstrip("Z")
    if text := _get_element_text(root, _CP_LASTMODIFIEDBY):
        metadata.last_modified_by = text

    revision_elem = root.find(_CP_REVISION)
    if revision_elem is not None and revision_elem.text:
        try:
            metadata.revision = int(revision_elem.text)
        except ValueError:
            pass

    return metadata


def _extract_slide_comments_from_context(
    ctx: _PptxContext, slide_number: int
) -> list[PptxComment]:
    """
    Extract comments for a specific slide from cached comment XML.

    Args:
        ctx: PptxContext with cached XML roots.
        slide_number: 1-based slide number to extract comments for.

    Returns:
        List of PPTXComment objects with author, text, and date fields.
    """
    root = ctx.get_comment_root(slide_number)
    if root is None:
        return []

    comments: list[PptxComment] = []
    try:
        for cm in root.findall(f".//{P_CM}"):
            text_elem = cm.find(P_TEXT)
            comments.append(
                PptxComment(
                    author=cm.get("authorId", ""),
                    text=(
                        text_elem.text
                        if text_elem is not None and text_elem.text
                        else ""
                    ),
                    date=cm.get("dt", ""),
                )
            )
    except Exception as e:
        logger.debug(f"Failed to extract comments for slide {slide_number}: {e}")

    return comments


def _get_shape_position(shape_elem: ET.Element) -> tuple[int, int]:
    """
    Get the position of a shape element for sorting purposes.

    Returns position as (top, left) tuple so shapes can be sorted
    in reading order (top-to-bottom, then left-to-right).

    Args:
        shape_elem: XML Element representing a shape (p:sp or p:pic).

    Returns:
        Tuple of (top, left) coordinates. For shapes without explicit
        positions (like placeholders that inherit from master), returns
        a default position based on placeholder type:
        - Title: very top (0, 0)
        - Body/content: below title (1, 0)
        - Footer/slide number: at bottom (999999998, x)
        - Other: at bottom (999999999, x)
    """
    try:
        # First, try to get explicit position from xfrm
        sp_pr = shape_elem.find(f".//{P_SPPR}")
        if sp_pr is None:
            sp_pr = shape_elem.find(f".//{A_XFRM}")
        if sp_pr is None:
            sp_pr = shape_elem

        xfrm = sp_pr.find(A_XFRM)
        if xfrm is None:
            xfrm = sp_pr.find(P_XFRM)
        if xfrm is None:
            xfrm = shape_elem.find(f".//{A_XFRM}")
            if xfrm is None:
                xfrm = shape_elem.find(f".//{P_XFRM}")

        if xfrm is not None:
            off = xfrm.find(A_OFF)
            if off is not None:
                x = int(off.get("x", "0"))
                y = int(off.get("y", "0"))
                return (y, x)  # Sort by y (top) first, then x (left)

        # No explicit position - check if it's a placeholder and assign default
        nv_sp_pr = shape_elem.find(P_NVSPPR)
        if nv_sp_pr is not None:
            nv_pr = nv_sp_pr.find(P_NVPR)
            if nv_pr is not None:
                ph = nv_pr.find(P_PH)
                if ph is not None:
                    ph_type = ph.get("type", "")
                    ph_idx = ph.get("idx", "")

                    if ph_type in TITLE_TYPES:
                        return (0, 0)

                    if ph_type in BODY_TYPES or (not ph_type and ph_idx):
                        idx_num = int(ph_idx) if ph_idx.isdigit() else 0
                        return (1 + idx_num, 0)

                    if ph_type in FOOTER_TYPES or ph_type == "sldNum":
                        return (999999998, 0)

        return (999999999, 999999999)
    except Exception:
        return (999999999, 999999999)


def _extract_text_from_paragraphs(elem: ET.Element) -> str:
    """
    Extract all text from paragraph elements within an element.

    Handles special elements like line breaks (a:br) which are converted
    to vertical tab characters (\x0b) to match python-pptx behavior.

    Also handles field elements (a:fld) which contain dynamic content like
    slide numbers.

    Args:
        elem: XML element containing a:p (paragraph) elements.

    Returns:
        Combined text from all paragraphs, with newlines between paragraphs.
    """
    paragraphs: list[str] = []
    for p in elem.findall(f".//{A_P}"):
        texts: list[str] = []
        for child in p:
            tag = child.tag

            if tag == A_R or tag == A_FLD:  # Run or Field
                t = child.find(A_T)
                if t is not None and t.text:
                    texts.append(t.text)
            elif tag == A_BR:  # Line break
                texts.append("\x0b")
            elif tag == A_T:  # Direct text (less common)
                if child.text:
                    texts.append(child.text)

        paragraphs.append("".join(texts))
    return "\n".join(paragraphs)


def _extract_table_from_graphic_frame(elem: ET.Element) -> list[list[str]] | None:
    """
    Extract table data from a graphic frame if it contains a DrawingML table.

    Returns a list of rows (list of cell strings), or None if not a table.
    """
    graphic_data = elem.find(f".//{A_GRAPHICDATA}")
    if graphic_data is None or graphic_data.get("uri") != TABLE_URI:
        return None

    tbl = graphic_data.find(A_TBL)
    if tbl is None:
        return None

    table_data: list[list[str]] = []
    for tr in tbl.findall(A_TR):
        row_data: list[str] = []
        for tc in tr.findall(A_TC):
            tx_body = tc.find(A_TXBODY)
            cell_text = (
                _extract_text_from_paragraphs(tx_body).strip()
                if tx_body is not None
                else ""
            )
            row_data.append(cell_text)
        table_data.append(row_data)

    return table_data


def _extract_formulas_from_element(elem: ET.Element) -> list[tuple[str, bool]]:
    """
    Extract mathematical formulas from an element's XML content.

    Searches for OMML math elements and converts them to LaTeX.

    Args:
        elem: XML element that may contain math formulas.

    Returns:
        List of tuples (latex_string, is_display) where:
        - latex_string: LaTeX representation of the formula
        - is_display: True for display equations (oMathPara), False for inline
    """
    formulas: list[tuple[str, bool]] = []
    omath_in_para: set[int] = set()

    # First, find all oMathPara elements (display equations)
    for omath_para in elem.iter(M_OMATHPARA):
        omath = omath_para.find(M_OMATH)
        if omath is not None:
            omath_in_para.add(id(omath))
            latex = omml_to_latex(omath)
            if latex.strip():
                formulas.append((latex, True))

    # Then find inline oMath elements (not in oMathPara)
    for omath in elem.iter(M_OMATH):
        if id(omath) not in omath_in_para:
            latex = omml_to_latex(omath)
            if latex.strip():
                formulas.append((latex, False))

    return formulas


def _normalize_relative_path(base_dir: str, target: str) -> str:
    """Normalize a relative path by resolving .. segments."""
    if not target.startswith("../"):
        return f"{base_dir}/{target}"

    parts = f"{base_dir}/{target}".split("/")
    normalized: list[str] = []
    for part in parts:
        if part == "..":
            if normalized:
                normalized.pop()
        elif part:
            normalized.append(part)
    return "/".join(normalized)


def _process_slide_from_context(
    ctx: _PptxContext, slide_path: str, slide_number: int
) -> PptxSlide:
    """
    Process a single slide and extract all its content using cached XML.

    Args:
        ctx: PptxContext with cached XML roots.
        slide_path: Path to the slide XML file within the ZIP.
        slide_number: 1-based slide number.

    Returns:
        PPTXSlide object containing all extracted content.
    """
    slide_title = ""
    slide_footer = ""
    content_placeholders: list[str] = []
    other_textboxes: list[str] = []
    tables: list[list[list[str]]] = []
    images: list[PptxImage] = []
    formulas: list[PptxFormula] = []

    # Collect all content items with their positions for ordering
    ordered_content: list[tuple[tuple[int, int], str, str]] = []

    slide_rels = ctx.get_slide_relationships(slide_path)
    root = ctx.get_slide_root(slide_path)
    if root is None:
        return PptxSlide(slide_number=slide_number)

    sp_tree = root.find(f".//{P_SPTREE}")
    if sp_tree is None:
        return PptxSlide(slide_number=slide_number)

    # Collect all shapes with their positions
    shape_elements: list[tuple[str, ET.Element, tuple[int, int]]] = []
    for sp in sp_tree.findall(f".//{P_SP}"):
        shape_elements.append(("sp", sp, _get_shape_position(sp)))
    for pic in sp_tree.findall(f".//{P_PIC}"):
        shape_elements.append(("pic", pic, _get_shape_position(pic)))
    for frame in sp_tree.findall(f".//{P_GRAPHICFRAME}"):
        shape_elements.append(("graphicFrame", frame, _get_shape_position(frame)))

    shape_elements.sort(key=lambda x: x[2])

    image_counter = 0
    slide_dir = "/".join(slide_path.rsplit("/", 1)[:-1])

    for shape_type, elem, position in shape_elements:
        # Picture extraction
        if shape_type == "pic":
            try:
                blip = elem.find(f".//{A_BLIP}")
                if blip is None:
                    continue

                r_embed = blip.get(R_EMBED)
                if not r_embed or r_embed not in slide_rels:
                    continue

                target = slide_rels[r_embed].get("target", "")
                image_path = _normalize_relative_path(slide_dir, target)

                # Extract caption and description from cNvPr
                caption = ""
                description = ""
                if (cNvPr := elem.find(f".//{P_CNVPR}")) is not None:
                    caption = cNvPr.get("name", "")
                    description = cNvPr.get("descr", "")

                blob = ctx.get_image_data(image_path)
                if blob is not None:
                    image_counter += 1
                    ext = target.rsplit(".", 1)[-1].lower()
                    content_type = _CONTENT_TYPE_MAP.get(ext, f"image/{ext}")
                    width, height = _get_image_pixel_dimensions(blob)

                    images.append(
                        PptxImage(
                            image_index=image_counter,
                            filename=f"image.{ext}",
                            content_type=content_type,
                            size_bytes=len(blob),
                            blob=blob,
                            width=width,
                            height=height,
                            caption=caption,
                            description=description,
                            slide_number=slide_number,
                        )
                    )

                    if description:
                        ordered_content.append(
                            (position, "image_caption", f"[Image: {description}]")
                        )
            except Exception as e:
                logger.debug(f"Failed to extract image on slide {slide_number}: {e}")
            continue

        # Table extraction
        if shape_type == "graphicFrame":
            try:
                table_data = _extract_table_from_graphic_frame(elem)
                if table_data:
                    tables.append(table_data)
                    table_text = "\n".join("\t".join(row) for row in table_data).strip()
                    if table_text:
                        ordered_content.append((position, "table", table_text))
            except Exception as e:
                logger.debug(f"Failed to extract table on slide {slide_number}: {e}")
            continue

        # Shape (text) extraction
        nv_sp_pr = elem.find(P_NVSPPR)
        if nv_sp_pr is None:
            continue

        nv_pr = nv_sp_pr.find(P_NVPR)
        ph = nv_pr.find(P_PH) if nv_pr is not None else None

        # Extract formulas from shape
        for latex, is_display in _extract_formulas_from_element(elem):
            formulas.append(PptxFormula(latex=latex, is_display=is_display))
            formula_text = f"$${latex}$$" if is_display else f"${latex}$"
            ordered_content.append((position, "formula", formula_text))

        # Extract text
        tx_body = elem.find(P_TXBODY)
        if tx_body is None:
            continue

        text = _extract_text_from_paragraphs(tx_body).strip()
        if not text:
            continue

        # Determine placeholder type and categorize text
        if ph is not None:
            ph_type = ph.get("type", "")
            ph_idx = ph.get("idx", "")

            if ph_type in TITLE_TYPES:
                slide_title = text
                ordered_content.append((position, "title", text))
            elif ph_type in FOOTER_TYPES:
                slide_footer = text
            elif ph_type in SKIP_TYPES:
                pass
            elif ph_type in BODY_TYPES or (not ph_type and ph_idx):
                content_placeholders.append(text)
                ordered_content.append((position, "content", text))
            else:
                other_textboxes.append(text)
                ordered_content.append((position, "other", text))
        else:
            other_textboxes.append(text)
            ordered_content.append((position, "other", text))

    # Comment extraction
    comments = _extract_slide_comments_from_context(ctx, slide_number)
    for comment in comments:
        ordered_content.append(
            (
                (999999, 999999),
                "comment",
                f"[Comment: {comment.author}@{comment.date}: {comment.text}]",
            )
        )

    # Build slide text from ordered content
    ordered_content.sort(key=lambda x: x[0])
    slide_text = "\n".join(item[2] for item in ordered_content)

    # Build base text (without formulas, comments, image captions)
    base_content_types = frozenset({"title", "content", "other", "table"})
    base_text = "\n".join(
        item[2] for item in ordered_content if item[1] in base_content_types
    )

    return PptxSlide(
        slide_number=slide_number,
        title=slide_title,
        footer=slide_footer,
        content_placeholders=content_placeholders,
        other_textboxes=other_textboxes,
        tables=tables,
        images=images,
        formulas=formulas,
        comments=comments,
        text=slide_text,
        base_text=base_text,
    )


def read_pptx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[PptxContent, Any, None]:
    """
    Extract all relevant content from a PowerPoint .pptx file.

    Primary entry point for PPTX file extraction. Iterates through all slides,
    extracting text, images, formulas, and comments while maintaining shape
    ordering for consistent text output.

    This function uses a generator pattern for API consistency with other
    extractors, even though PPTX files contain exactly one presentation.

    Args:
        file_like: BytesIO object containing the complete PPTX file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned PptxContent.metadata.

    Yields:
        PptxContent: Single PptxContent object containing:
            - metadata: PptxMetadata with title, author, dates, revision
            - slides: List of PPTXSlide objects, each containing:
                - slide_number: 1-based slide index
                - title: Slide title text
                - content_placeholders: Body text from content areas
                - other_textboxes: Text from non-placeholder shapes
                - tables: List of tables as 2D lists of cell text
                - images: List of PPTXImage with binary data
                - formulas: List of PPTXFormula as LaTeX
                - comments: List of PPTXComment
                - text: Complete slide text with formulas and comments
                - base_text: Text without formulas/comments/captions

    Raises:
        ExtractionFileEncryptedError: If the PPTX is encrypted or password-protected.

    Processing Details:
        - Shapes are sorted by position (top-to-bottom, left-to-right)
        - Title placeholders are extracted separately from body content
        - Images include alt text/captions when available
        - Formulas are converted to LaTeX notation
        - Comments are appended at the end of slide content

    Example:
        >>> import io
        >>> with open("presentation.pptx", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for ppt in read_pptx(data, path="presentation.pptx"):
        ...         print(f"Title: {ppt.metadata.title}")
        ...         print(f"Slides: {len(ppt.slides)}")
        ...         for slide in ppt.slides:
        ...             print(f"  Slide {slide.slide_number}: {slide.title}")
        ...             print(f"    Images: {len(slide.images)}")

    Performance Notes:
        - ZIP file is opened once and all XML is cached
        - All XML documents are parsed once and reused
        - Images are loaded into memory as binary blobs
        - Large presentations with many images may use significant memory
    """
    try:
        logger.debug("Reading pptx")
        file_like.seek(0)
        if is_ooxml_encrypted(file_like):
            raise ExtractionFileEncryptedError(
                "PPTX is encrypted or password-protected"
            )

        # Create context that opens ZIP once and caches all parsed XML
        ctx = _PptxContext(file_like)
        try:
            # Extract metadata from cached XML
            metadata = _extract_metadata_from_context(ctx)

            # Get slide order from cached presentation.xml
            slide_paths = ctx.slide_order

            # Process each slide using cached XML
            slides_result: List[PptxSlide] = []
            for slide_index, slide_path in enumerate(slide_paths, start=1):
                slide = _process_slide_from_context(ctx, slide_path, slide_index)
                slides_result.append(slide)

            metadata.populate_from_path(path)

            total_images = sum(len(slide.images) for slide in slides_result)
            logger.info(
                "Extracted PPTX: %d slides, %d images",
                len(slides_result),
                total_images,
            )

            yield PptxContent(metadata=metadata, slides=slides_result)
        finally:
            ctx.close()
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract PPTX file", cause=exc) from exc
