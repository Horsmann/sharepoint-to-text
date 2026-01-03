"""
DOCX Document Extractor

Extracts text content, metadata, and structure from Microsoft Word .docx files
(Office Open XML format, Word 2007+).

Uses direct XML parsing of the docx ZIP archive structure for all content
extraction, without requiring the python-docx library.
"""

import io
import logging
from typing import Any, Generator
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.extractors.data_types import (
    DocxComment,
    DocxContent,
    DocxFormula,
    DocxHeaderFooter,
    DocxHyperlink,
    DocxImage,
    DocxMetadata,
    DocxNote,
    DocxParagraph,
    DocxRun,
    DocxSection,
)
from sharepoint2text.extractors.util.encryption import is_ooxml_encrypted
from sharepoint2text.extractors.util.omml_to_latex import omml_to_latex
from sharepoint2text.extractors.util.ooxml_context import OOXMLZipContext
from sharepoint2text.extractors.util.zip_utils import parse_relationships

logger = logging.getLogger(__name__)

# =============================================================================
# XML Namespaces
# =============================================================================

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}

# Namespace prefixes for element access
W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
CP_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"
DCTERMS_NS = "{http://purl.org/dc/terms/}"
PIC_NS = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
WPS_NS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"

# =============================================================================
# Pre-computed tag names (performance optimization)
# =============================================================================

W_T = f"{W_NS}t"
W_P = f"{W_NS}p"
W_R = f"{W_NS}r"
W_TBL = f"{W_NS}tbl"
W_TR = f"{W_NS}tr"
W_TC = f"{W_NS}tc"
W_PPR = f"{W_NS}pPr"
W_RPR = f"{W_NS}rPr"
W_PSTYLE = f"{W_NS}pStyle"
W_JC = f"{W_NS}jc"
W_VAL = f"{W_NS}val"
W_B = f"{W_NS}b"
W_I = f"{W_NS}i"
W_U = f"{W_NS}u"
W_SZ = f"{W_NS}sz"
W_COLOR = f"{W_NS}color"
W_RFONTS = f"{W_NS}rFonts"
W_DRAWING = f"{W_NS}drawing"
W_HYPERLINK = f"{W_NS}hyperlink"
W_FOOTNOTE = f"{W_NS}footnote"
W_ENDNOTE = f"{W_NS}endnote"
W_COMMENT = f"{W_NS}comment"
W_BODY = f"{W_NS}body"
W_SECTPR = f"{W_NS}sectPr"
W_BR = f"{W_NS}br"
W_TYPE = f"{W_NS}type"
W_LAST_RENDERED_PAGE_BREAK = f"{W_NS}lastRenderedPageBreak"
W_PGSZ = f"{W_NS}pgSz"
W_PGMAR = f"{W_NS}pgMar"
W_KEEPNEXT = f"{W_NS}keepNext"
W_STYLE = f"{W_NS}style"
W_STYLEID = f"{W_NS}styleId"
W_NAME = f"{W_NS}name"
W_ID = f"{W_NS}id"
W_AUTHOR = f"{W_NS}author"
W_DATE = f"{W_NS}date"
W_W = f"{W_NS}w"
W_H = f"{W_NS}h"
W_ORIENT = f"{W_NS}orient"
W_LEFT = f"{W_NS}left"
W_RIGHT = f"{W_NS}right"
W_TOP = f"{W_NS}top"
W_BOTTOM = f"{W_NS}bottom"
W_ASCII = f"{W_NS}ascii"
W_HANSI = f"{W_NS}hAnsi"
W_CS = f"{W_NS}cs"

M_OMATH = f"{M_NS}oMath"
M_OMATHPARA = f"{M_NS}oMathPara"
MC_CHOICE = f"{MC_NS}Choice"
R_ID = f"{R_NS}id"
R_EMBED = f"{R_NS}embed"
A_BLIP = f"{A_NS}blip"
PIC_CNVPR = f"{PIC_NS}cNvPr"
WPS_WSP = f"{WPS_NS}wsp"
WPS_TXBX = f"{WPS_NS}txbx"

# Metadata tag names
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

# =============================================================================
# Constants
# =============================================================================

# Unit conversions
EMU_PER_INCH = 914400
TWIPS_PER_INCH = 1440

# Caption style keywords
CAPTION_STYLE_KEYWORDS = ("caption", "bildunterschrift", "abbildung", "figure")

# Content type by extension
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

# Skip IDs for separator/continuation notes
_SKIP_NOTE_IDS = frozenset({"-1", "0"})

# JPEG SOF markers for dimension extraction
_JPEG_SOF_MARKERS = frozenset(
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


# =============================================================================
# Image dimension extraction
# =============================================================================


def _get_image_pixel_dimensions(image_data: bytes) -> tuple[int | None, int | None]:
    """Extract pixel dimensions from common raster formats."""
    if not image_data:
        return None, None

    # PNG
    if image_data.startswith(b"\x89PNG\r\n\x1a\n") and len(image_data) >= 24:
        w = int.from_bytes(image_data[16:20], "big")
        h = int.from_bytes(image_data[20:24], "big")
        return (w or None, h or None)

    # GIF
    if image_data[:6] in (b"GIF87a", b"GIF89a") and len(image_data) >= 10:
        w = int.from_bytes(image_data[6:8], "little")
        h = int.from_bytes(image_data[8:10], "little")
        return (w or None, h or None)

    # BMP
    if image_data[:2] == b"BM" and len(image_data) >= 26:
        w = int.from_bytes(image_data[18:22], "little", signed=True)
        h = int.from_bytes(image_data[22:26], "little", signed=True)
        return (abs(w) or None, abs(h) or None)

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
            if marker in _JPEG_SOF_MARKERS and i + 2 + length <= size:
                h = int.from_bytes(image_data[i + 5 : i + 7], "big")
                w = int.from_bytes(image_data[i + 7 : i + 9], "big")
                return (w or None, h or None)
            i += 2 + length

    return None, None


# =============================================================================
# Text extraction helpers
# =============================================================================


def _collect_text_from_element(element: ET.Element) -> str:
    """Extract all text from w:t elements within an element."""
    return "".join(t.text for t in element.iter(W_T) if t.text)


def _get_paragraph_style(para: ET.Element) -> str:
    """Return the paragraph style name (empty if absent)."""
    pPr = para.find(W_PPR)
    if pPr is not None:
        pStyle = pPr.find(W_PSTYLE)
        if pStyle is not None:
            return pStyle.get(W_VAL, "")
    return ""


def _has_keep_next(para: ET.Element) -> bool:
    """Return True when keepNext is enabled for the paragraph."""
    pPr = para.find(W_PPR)
    if pPr is None:
        return False
    keep_next = pPr.find(W_KEEPNEXT)
    if keep_next is None:
        return False
    val = keep_next.get(W_VAL, "true")
    return val.lower() not in ("false", "0")


def _is_caption_style(style_name: str) -> bool:
    """Return True when style name indicates a caption paragraph."""
    style_lower = style_name.lower()
    return any(kw in style_lower for kw in CAPTION_STYLE_KEYWORDS)


def _process_text_element(
    elem: ET.Element,
    parts: list[str],
    include_formulas: bool,
) -> None:
    """Append extracted text from a node, respecting AlternateContent and formulas."""
    tag = elem.tag

    if tag.endswith("}AlternateContent"):
        choice = elem.find(MC_CHOICE)
        if choice is not None:
            for child in choice:
                _process_text_element(child, parts, include_formulas)
        return

    if tag.endswith("}Fallback"):
        return

    if tag == W_R:
        for child in elem:
            if child.tag == W_T:
                if child.text:
                    parts.append(child.text)
            elif child.tag.endswith("}AlternateContent"):
                _process_text_element(child, parts, include_formulas)
        return

    if tag == M_OMATH:
        if include_formulas:
            latex = omml_to_latex(elem)
            if latex.strip():
                parts.append(f"${latex}$")
        return

    if tag == M_OMATHPARA:
        if include_formulas:
            omath = elem.find(M_OMATH)
            if omath is not None:
                latex = omml_to_latex(omath)
                if latex.strip():
                    parts.append(f"$${latex}$$")
        return

    for child in elem:
        _process_text_element(child, parts, include_formulas)


def _extract_paragraph_content(paragraph: ET.Element, include_formulas: bool) -> str:
    """Extract text from a paragraph, including inline and display equations."""
    parts: list[str] = []
    for child in paragraph:
        _process_text_element(child, parts, include_formulas)
    return "".join(parts)


def _extract_table_text(table: ET.Element, include_formulas: bool) -> list[str]:
    """Extract table text in row order, concatenating cell content."""
    texts: list[str] = []
    for row in table.iter(W_TR):
        for cell in row.iter(W_TC):
            cell_parts: list[str] = []
            for paragraph in cell.iter(W_P):
                text = _extract_paragraph_content(paragraph, include_formulas)
                if text.strip():
                    cell_parts.append(text)
            if cell_parts:
                texts.append(" ".join(cell_parts))
    return texts


def _extract_full_text_from_body(
    body: ET.Element | None, include_formulas: bool = True
) -> str:
    """Extract complete text content from a document body."""
    if body is None:
        return ""

    all_text: list[str] = []
    for element in body:
        if element.tag == W_P:
            text = _extract_paragraph_content(element, include_formulas)
            if text.strip():
                all_text.append(text)
        elif element.tag == W_TBL:
            all_text.extend(_extract_table_text(element, include_formulas))

    return "\n".join(all_text)


# =============================================================================
# DOCX Context (cached ZIP/XML access)
# =============================================================================


class _DocxContext(OOXMLZipContext):
    """Cached context for DOCX extraction - opens ZIP once and caches XML."""

    def __init__(self, file_like: io.BytesIO):
        super().__init__(file_like)

        # XML roots cache
        self._document_root: ET.Element | None = None
        self._core_root: ET.Element | None = None
        self._styles_root: ET.Element | None = None
        self._footnotes_root: ET.Element | None = None
        self._endnotes_root: ET.Element | None = None
        self._comments_root: ET.Element | None = None
        self._rels_root: ET.Element | None = None

        # Data cache
        self._relationships: dict[str, dict] | None = None
        self._styles: dict[str, str] | None = None
        self._header_footer_roots: dict[str, ET.Element] = {}

        self._load_xml_files()

    def _load_xml_files(self) -> None:
        """Load and parse all XML files from the ZIP at once."""
        xml_files = [
            ("word/document.xml", "_document_root"),
            ("docProps/core.xml", "_core_root"),
            ("word/styles.xml", "_styles_root"),
            ("word/footnotes.xml", "_footnotes_root"),
            ("word/endnotes.xml", "_endnotes_root"),
            ("word/comments.xml", "_comments_root"),
            ("word/_rels/document.xml.rels", "_rels_root"),
        ]

        for path, attr in xml_files:
            if path in self._namelist:
                setattr(self, attr, self.read_xml_root(path))

        # Pre-load header and footer files
        self._relationships = self._parse_relationships()
        for rel_info in self._relationships.values():
            rel_type = rel_info.get("type", "")
            target = rel_info.get("target", "")
            if "header" in rel_type.lower() or "footer" in rel_type.lower():
                hf_path = "word/" + target
                if hf_path in self._namelist:
                    self._header_footer_roots[hf_path] = self.read_xml_root(hf_path)

    def _parse_relationships(self) -> dict[str, dict]:
        """Parse relationships from cached rels root."""
        if self._rels_root is None:
            return {}

        relationships = {}
        for rel in parse_relationships(self._rels_root):
            rel_id = rel["id"]
            if rel_id:
                relationships[rel_id] = {
                    "type": rel["type"],
                    "target": rel["target"],
                    "target_mode": rel["target_mode"],
                }
        return relationships

    @property
    def document_body(self) -> ET.Element | None:
        """Get the document body element."""
        if self._document_root is None:
            return None
        return self._document_root.find(W_BODY)

    @property
    def relationships(self) -> dict[str, dict]:
        """Get cached relationships."""
        if self._relationships is None:
            self._relationships = self._parse_relationships()
        return self._relationships

    @property
    def styles(self) -> dict[str, str]:
        """Get cached style map (style_id -> style_name)."""
        if self._styles is None:
            self._styles = {}
            if self._styles_root is not None:
                for style in self._styles_root.findall(f".//{W_STYLE}"):
                    style_id = style.get(W_STYLEID) or ""
                    name_elem = style.find(W_NAME)
                    style_name = name_elem.get(W_VAL) if name_elem is not None else ""
                    if style_id:
                        self._styles[style_id] = style_name or style_id
        return self._styles

    def get_image_data(self, image_path: str) -> bytes | None:
        """Read image data from the ZIP file."""
        if image_path not in self._namelist:
            return None
        return self.read_bytes(image_path)


# =============================================================================
# Extraction functions
# =============================================================================


def _get_element_text(root: ET.Element | None, tag: str) -> str | None:
    """Extract text from an element if it exists and has text content."""
    if root is None:
        return None
    elem = root.find(tag)
    if elem is not None and elem.text:
        return elem.text
    return None


def _extract_metadata_from_context(ctx: _DocxContext) -> DocxMetadata:
    """Extract document metadata from cached core.xml root."""
    metadata = DocxMetadata()
    root = ctx._core_root
    if root is None:
        return metadata

    # Metadata field mappings: (tag, attribute)
    field_mappings = [
        (_DC_TITLE, "title"),
        (_DC_CREATOR, "author"),
        (_DC_SUBJECT, "subject"),
        (_CP_KEYWORDS, "keywords"),
        (_CP_CATEGORY, "category"),
        (_DC_DESCRIPTION, "comments"),
        (_DCTERMS_CREATED, "created"),
        (_DCTERMS_MODIFIED, "modified"),
        (_CP_LASTMODIFIEDBY, "last_modified_by"),
    ]

    for tag, attr in field_mappings:
        if text := _get_element_text(root, tag):
            setattr(metadata, attr, text)

    revision_elem = root.find(_CP_REVISION)
    if revision_elem is not None and revision_elem.text:
        try:
            metadata.revision = int(revision_elem.text)
        except ValueError:
            pass

    return metadata


def _extract_notes_from_root(root: ET.Element | None, note_tag: str) -> list[DocxNote]:
    """Extract notes (footnotes or endnotes) from an XML root element."""
    if root is None:
        return []

    return [
        DocxNote(id=note.get(W_ID) or "", text=_collect_text_from_element(note))
        for note in root.findall(f".//{note_tag}")
        if (note.get(W_ID) or "") not in _SKIP_NOTE_IDS
    ]


def _extract_footnotes_from_context(ctx: _DocxContext) -> list[DocxNote]:
    """Extract footnotes from cached footnotes.xml root."""
    return _extract_notes_from_root(ctx._footnotes_root, W_FOOTNOTE)


def _extract_endnotes_from_context(ctx: _DocxContext) -> list[DocxNote]:
    """Extract endnotes from cached endnotes.xml root."""
    return _extract_notes_from_root(ctx._endnotes_root, W_ENDNOTE)


def _extract_comments_from_context(ctx: _DocxContext) -> list[DocxComment]:
    """Extract comments from cached comments.xml root."""
    root = ctx._comments_root
    if root is None:
        return []

    return [
        DocxComment(
            id=comment.get(W_ID) or "",
            author=comment.get(W_AUTHOR) or "",
            date=comment.get(W_DATE) or "",
            text=_collect_text_from_element(comment),
        )
        for comment in root.findall(f".//{W_COMMENT}")
    ]


def _parse_twips_to_inches(value: str | None) -> float | None:
    """Convert twips string to inches, returning None on failure."""
    if not value:
        return None
    try:
        return int(value) / TWIPS_PER_INCH
    except ValueError:
        return None


def _extract_sections_from_context(ctx: _DocxContext) -> list[DocxSection]:
    """Extract section properties (page layout) from cached document body."""
    body = ctx.document_body
    if body is None:
        return []

    sect_pr_elements: list[ET.Element] = []

    # Sections in paragraphs
    for p in body.findall(f".//{W_P}"):
        ppr = p.find(W_PPR)
        if ppr is not None:
            sect_pr = ppr.find(W_SECTPR)
            if sect_pr is not None:
                sect_pr_elements.append(sect_pr)

    # Final section at end of body
    final_sect_pr = body.find(W_SECTPR)
    if final_sect_pr is not None:
        sect_pr_elements.append(final_sect_pr)

    sections: list[DocxSection] = []
    for sect_pr in sect_pr_elements:
        section = DocxSection()

        pg_sz = sect_pr.find(W_PGSZ)
        if pg_sz is not None:
            if inches := _parse_twips_to_inches(pg_sz.get(W_W)):
                section.page_width_inches = inches
            if inches := _parse_twips_to_inches(pg_sz.get(W_H)):
                section.page_height_inches = inches
            orient = pg_sz.get(W_ORIENT)
            if orient and orient != "portrait":
                section.orientation = orient

        pg_mar = sect_pr.find(W_PGMAR)
        if pg_mar is not None:
            for attr, tag in [
                ("left_margin_inches", W_LEFT),
                ("right_margin_inches", W_RIGHT),
                ("top_margin_inches", W_TOP),
                ("bottom_margin_inches", W_BOTTOM),
            ]:
                if inches := _parse_twips_to_inches(pg_mar.get(tag)):
                    setattr(section, attr, inches)

        sections.append(section)

    return sections


def _determine_hf_type(path: str, rel_type: str) -> str:
    """Determine header/footer type from path or relationship type."""
    path_lower = path.lower()
    rel_lower = rel_type.lower()
    if "first" in path_lower or "first" in rel_lower:
        return "first_page"
    if "even" in path_lower or "even" in rel_lower:
        return "even_page"
    return "default"


def _extract_header_footers_from_context(
    ctx: _DocxContext,
) -> tuple[list[DocxHeaderFooter], list[DocxHeaderFooter]]:
    """Extract headers and footers from cached header/footer XML roots."""
    headers: list[DocxHeaderFooter] = []
    footers: list[DocxHeaderFooter] = []

    for rel_info in ctx.relationships.values():
        rel_type = rel_info.get("type", "")
        target = rel_info.get("target", "")
        rel_type_lower = rel_type.lower()

        is_header = "header" in rel_type_lower
        is_footer = "footer" in rel_type_lower
        if not (is_header or is_footer):
            continue

        hf_path = "word/" + target
        root = ctx._header_footer_roots.get(hf_path)
        if root is None:
            continue

        text = _collect_text_from_element(root)
        if not text:
            continue

        hf_obj = DocxHeaderFooter(type=_determine_hf_type(hf_path, rel_type), text=text)
        if is_header:
            headers.append(hf_obj)
        else:
            footers.append(hf_obj)

    return headers, footers


def _parse_run_properties(
    rpr: ET.Element | None,
) -> tuple[bool | None, bool | None, bool | None, str | None, float | None, str | None]:
    """Parse run properties: (bold, italic, underline, font_name, font_size, font_color)."""
    if rpr is None:
        return None, None, None, None, None, None

    # Bold
    bold = None
    bold_elem = rpr.find(W_B)
    if bold_elem is not None:
        bold_val = bold_elem.get(W_VAL)
        bold = bold_val != "0" if bold_val else True

    # Italic
    italic = None
    italic_elem = rpr.find(W_I)
    if italic_elem is not None:
        italic_val = italic_elem.get(W_VAL)
        italic = italic_val != "0" if italic_val else True

    # Underline
    underline = None
    underline_elem = rpr.find(W_U)
    if underline_elem is not None:
        u_val = underline_elem.get(W_VAL)
        underline = u_val and u_val != "none"

    # Font name
    font_name = None
    rfonts = rpr.find(W_RFONTS)
    if rfonts is not None:
        font_name = rfonts.get(W_ASCII) or rfonts.get(W_HANSI) or rfonts.get(W_CS)

    # Font size (half-points to points)
    font_size = None
    sz = rpr.find(W_SZ)
    if sz is not None:
        sz_val = sz.get(W_VAL)
        if sz_val:
            try:
                font_size = int(sz_val) / 2
            except ValueError:
                pass

    # Font color
    font_color = None
    color = rpr.find(W_COLOR)
    if color is not None:
        font_color = color.get(W_VAL)

    return bold, italic, underline, font_name, font_size, font_color


def _extract_paragraphs_from_context(ctx: _DocxContext) -> list[DocxParagraph]:
    """Extract paragraphs with formatting and run information."""
    body = ctx.document_body
    if body is None:
        return []

    style_map = ctx.styles
    paragraphs: list[DocxParagraph] = []

    for p in body.findall(W_P):
        ppr = p.find(W_PPR)
        style_id = None
        alignment = None

        if ppr is not None:
            style_elem = ppr.find(W_PSTYLE)
            if style_elem is not None:
                style_id = style_elem.get(W_VAL)

            jc_elem = ppr.find(W_JC)
            if jc_elem is not None:
                alignment = jc_elem.get(W_VAL)

        style_name = style_map.get(style_id, style_id) if style_id else None

        has_page_break = any(
            br.get(W_TYPE) == "page" for br in p.iter(W_BR) if br is not None
        ) or (p.find(f".//{W_LAST_RENDERED_PAGE_BREAK}") is not None)

        # Extract runs
        runs: list[DocxRun] = []
        for r in p.findall(f".//{W_R}"):
            run_text = _collect_text_from_element(r)
            if not run_text:
                continue

            bold, italic, underline, font_name, font_size, font_color = (
                _parse_run_properties(r.find(W_RPR))
            )

            runs.append(
                DocxRun(
                    text=run_text,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    font_name=font_name,
                    font_size=font_size,
                    font_color=font_color,
                )
            )

        paragraphs.append(
            DocxParagraph(
                text="".join(run.text for run in runs),
                style=style_name,
                alignment=alignment,
                runs=runs,
                has_page_break=has_page_break,
            )
        )

    return paragraphs


def _extract_tables_from_context(
    ctx: _DocxContext,
) -> tuple[list[list[list[str]]], list[int]]:
    """Extract tables as lists of lists of cell text."""
    body = ctx.document_body
    if body is None:
        return [], []

    tables: list[list[list[str]]] = []
    table_anchor_paragraph_indices: list[int] = []
    current_paragraph_index = -1

    for child in list(body):
        if child.tag == W_P:
            current_paragraph_index += 1
            continue
        if child.tag != W_TBL:
            continue

        anchor = max(0, current_paragraph_index)

        for tbl in child.iter(W_TBL):
            table_data: list[list[str]] = []
            for tr in tbl.findall(W_TR):
                row_data = [
                    "\n".join(
                        _collect_text_from_element(p) for p in tc.findall(f".//{W_P}")
                    )
                    for tc in tr.findall(W_TC)
                ]
                table_data.append(row_data)
            tables.append(table_data)
            table_anchor_paragraph_indices.append(anchor)

    return tables, table_anchor_paragraph_indices


def _extract_images_from_context(ctx: _DocxContext) -> list[DocxImage]:
    """Extract images with captions and descriptions."""
    rels = ctx.relationships
    body = ctx.document_body

    image_metadata: dict[str, tuple[str, str]] = {}
    image_anchor_paragraph_indices: dict[str, set[int]] = {}

    if body is not None:
        paragraphs = list(body.findall(W_P))

        for para_idx, para in enumerate(paragraphs):
            for drawing in para.iter(W_DRAWING):
                caption = ""
                description = ""

                pic_cNvPr = drawing.find(f".//{PIC_CNVPR}")
                if pic_cNvPr is not None:
                    if descr := pic_cNvPr.get("descr", ""):
                        description = descr
                    if name := pic_cNvPr.get("name", ""):
                        caption = name

                for wsp in drawing.iter(WPS_WSP):
                    txbx = wsp.find(WPS_TXBX)
                    if txbx is not None:
                        if text := _collect_text_from_element(txbx):
                            caption = text
                            break

                # Check preceding/following paragraphs for caption
                preceding_caption = None
                if para_idx > 0:
                    prev_para = paragraphs[para_idx - 1]
                    prev_style = _get_paragraph_style(prev_para)
                    if _is_caption_style(prev_style) and _has_keep_next(prev_para):
                        if text := _collect_text_from_element(prev_para):
                            preceding_caption = text

                following_caption = None
                if para_idx + 1 < len(paragraphs):
                    next_para = paragraphs[para_idx + 1]
                    if _is_caption_style(_get_paragraph_style(next_para)):
                        if text := _collect_text_from_element(next_para):
                            following_caption = text

                if preceding_caption:
                    caption = preceding_caption
                elif following_caption:
                    caption = following_caption

                blip = drawing.find(f".//{A_BLIP}")
                if blip is not None:
                    if r_embed := blip.get(R_EMBED):
                        image_metadata[r_embed] = (caption, description)
                        image_anchor_paragraph_indices.setdefault(r_embed, set()).add(
                            para_idx
                        )

    # Build image list
    images: list[DocxImage] = []
    image_counter = 0

    for rel_id, rel_info in rels.items():
        rel_type = rel_info.get("type", "")
        target = rel_info.get("target", "")

        if "image" not in rel_type.lower():
            continue

        image_path = "word/" + target
        try:
            img_data = ctx.get_image_data(image_path)
            if img_data is None:
                continue

            image_counter += 1
            ext = target.rsplit(".", 1)[-1].lower()
            caption, description = image_metadata.get(rel_id, ("", ""))
            width, height = _get_image_pixel_dimensions(img_data)

            images.append(
                DocxImage(
                    rel_id=rel_id,
                    filename=target.rsplit("/", 1)[-1],
                    content_type=_CONTENT_TYPE_MAP.get(ext, f"image/{ext}"),
                    data=io.BytesIO(img_data),
                    size_bytes=len(img_data),
                    width=width,
                    height=height,
                    image_index=image_counter,
                    caption=caption,
                    description=description,
                    anchor_paragraph_indices=sorted(
                        image_anchor_paragraph_indices.get(rel_id, set())
                    ),
                )
            )
        except Exception as e:
            logger.debug(f"Image extraction failed for rel_id {rel_id} - {e}")
            images.append(DocxImage(rel_id=rel_id, error=str(e)))

    return images


def _extract_hyperlinks_from_context(ctx: _DocxContext) -> list[DocxHyperlink]:
    """Extract hyperlinks from the document."""
    body = ctx.document_body
    if body is None:
        return []

    rels = ctx.relationships
    hyperlinks: list[DocxHyperlink] = []

    for hyperlink in body.findall(f".//{W_HYPERLINK}"):
        r_id = hyperlink.get(R_ID)
        if r_id and r_id in rels:
            rel_info = rels[r_id]
            if "hyperlink" in rel_info.get("type", "").lower():
                hyperlinks.append(
                    DocxHyperlink(
                        text=_collect_text_from_element(hyperlink),
                        url=rel_info.get("target", ""),
                    )
                )

    return hyperlinks


def _extract_formulas_from_context(ctx: _DocxContext) -> list[DocxFormula]:
    """Extract all mathematical formulas from the document as LaTeX."""
    body = ctx.document_body
    if body is None:
        return []

    formulas: list[DocxFormula] = []
    omath_in_para: set[int] = set()

    # Display equations (oMathPara)
    for omath_para in body.iter(M_OMATHPARA):
        omath = omath_para.find(M_OMATH)
        if omath is not None:
            omath_in_para.add(id(omath))
            latex = omml_to_latex(omath)
            if latex.strip():
                formulas.append(DocxFormula(latex=latex, is_display=True))

    # Inline equations (not in oMathPara)
    for omath in body.iter(M_OMATH):
        if id(omath) not in omath_in_para:
            latex = omml_to_latex(omath)
            if latex.strip():
                formulas.append(DocxFormula(latex=latex, is_display=False))

    return formulas


# =============================================================================
# Main entry point
# =============================================================================


def read_docx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocxContent, Any, None]:
    """
    Extract all relevant content from a Word .docx file.

    Uses a generator pattern for API consistency. DOCX files yield exactly one
    DocxContent object containing paragraphs, tables, images, metadata, etc.
    """
    try:
        file_like.seek(0)
        if is_ooxml_encrypted(file_like):
            raise ExtractionFileEncryptedError(
                "DOCX is encrypted or password-protected"
            )

        ctx = _DocxContext(file_like)
        try:
            metadata = _extract_metadata_from_context(ctx)
            paragraphs = _extract_paragraphs_from_context(ctx)
            tables, table_anchor_paragraph_indices = _extract_tables_from_context(ctx)
            headers, footers = _extract_header_footers_from_context(ctx)
            images = _extract_images_from_context(ctx)
            hyperlinks = _extract_hyperlinks_from_context(ctx)
            footnotes = _extract_footnotes_from_context(ctx)
            endnotes = _extract_endnotes_from_context(ctx)
            formulas = _extract_formulas_from_context(ctx)
            comments = _extract_comments_from_context(ctx)
            sections = _extract_sections_from_context(ctx)
            styles = list({para.style for para in paragraphs if para.style})
            full_text = _extract_full_text_from_body(
                ctx.document_body, include_formulas=True
            )

            metadata.populate_from_path(path)

            logger.info(
                "Extracted DOCX: %d paragraphs, %d tables, %d images",
                len(paragraphs),
                len(tables),
                len(images),
            )

            yield DocxContent(
                metadata=metadata,
                paragraphs=paragraphs,
                tables=tables,
                table_anchor_paragraph_indices=table_anchor_paragraph_indices,
                headers=headers,
                footers=footers,
                images=images,
                hyperlinks=hyperlinks,
                footnotes=footnotes,
                endnotes=endnotes,
                comments=comments,
                sections=sections,
                styles=styles,
                formulas=formulas,
                full_text=full_text,
            )
        finally:
            ctx.close()
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract DOCX file", cause=exc) from exc
