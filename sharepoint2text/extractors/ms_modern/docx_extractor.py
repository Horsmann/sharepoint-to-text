"""
DOCX Document Extractor
=======================

Extracts text content, metadata, and structure from Microsoft Word .docx files
(Office Open XML format, Word 2007 and later).

This module uses direct XML parsing of the docx ZIP archive structure for all
content extraction, without requiring the python-docx library.

File Format Background
----------------------
The .docx format is a ZIP archive containing XML files following the Office
Open XML (OOXML) standard. Key components:

    word/document.xml: Main document body (paragraphs, tables)
    word/styles.xml: Style definitions
    word/footnotes.xml: Footnote content
    word/endnotes.xml: Endnote content
    word/comments.xml: Comment/annotation content
    word/header1.xml, footer1.xml: Header/footer content
    word/media/: Embedded images
    word/_rels/document.xml.rels: Relationships (images, hyperlinks)
    docProps/core.xml: Metadata (title, author, dates)

XML Namespaces:
    - w: http://schemas.openxmlformats.org/wordprocessingml/2006/main
    - m: http://schemas.openxmlformats.org/officeDocument/2006/math
    - mc: http://schemas.openxmlformats.org/markup-compatibility/2006
    - r: http://schemas.openxmlformats.org/officeDocument/2006/relationships
    - a: http://schemas.openxmlformats.org/drawingml/2006/main
    - cp: http://schemas.openxmlformats.org/package/2006/metadata/core-properties
    - dc: http://purl.org/dc/elements/1.1/
    - dcterms: http://purl.org/dc/terms/

Math Formula Handling
---------------------
Word documents store mathematical formulas in OMML (Office Math Markup Language).
This module converts OMML to LaTeX-like notation for text representation.

Supported OMML elements:
    - m:f (fraction) -> \\frac{num}{den}
    - m:sSup/m:sSub (super/subscript) -> base^{sup} / base_{sub}
    - m:rad (radical/root) -> \\sqrt{content}
    - m:nary (n-ary operators) -> \\sum, \\int, etc.
    - m:d (delimiter) -> parentheses, brackets
    - m:m (matrix) -> \\begin{matrix}...\\end{matrix}
    - m:func (functions) -> \\sin, \\cos, etc.
    - m:bar/m:acc (overline/accent) -> \\overline, \\hat, etc.

The OMML-to-LaTeX converter also handles:
    - Greek letters (α -> \\alpha, etc.)
    - Math symbols (∞ -> \\infty, etc.)
    - Malformed bracket placement in roots

AlternateContent Handling
-------------------------
Word uses mc:AlternateContent elements to provide fallback representations
for features like equations. This extractor processes only mc:Choice content
and skips mc:Fallback to avoid duplicate text extraction.

Extracted Content
-----------------
The extractor retrieves:
    - Main body text (paragraphs and tables in order)
    - Headers and footers (default, first page, even page)
    - Footnotes and endnotes
    - Comments with author and date
    - Images with metadata
    - Hyperlinks with URLs
    - Formulas as LaTeX
    - Section properties (page layout)
    - Style names used

Two text outputs are provided:
    - full_text: Complete text including formulas as LaTeX
    - base_full_text: Text without formula representations

Known Limitations
-----------------
- Embedded OLE objects are not extracted
- Complex SmartArt text may be incomplete
- Drawing canvas text may not extract properly
- Tracked changes are not separately reported
- Password-protected files are not supported
- Very large documents may use significant memory

Usage
-----
    >>> import io
    >>> from sharepoint2text.extractors.ms_modern.docx_extractor import read_docx
    >>>
    >>> with open("document.docx", "rb") as f:
    ...     for doc in read_docx(io.BytesIO(f.read()), path="document.docx"):
    ...         print(f"Title: {doc.metadata.title}")
    ...         print(f"Author: {doc.metadata.author}")
    ...         print(f"Paragraphs: {len(doc.paragraphs)}")
    ...         print(doc.full_text[:500])

See Also
--------
- OOXML WordprocessingML: https://docs.microsoft.com/en-us/openspecs/office_standards/
- doc_extractor: For legacy .doc format

Maintenance Notes
-----------------
- The OMML-to-LaTeX converter handles common cases but may need extension
- Direct XML parsing is used for all content extraction
- AlternateContent handling prevents duplicate formula text
- Greek letter and symbol mapping can be extended as needed
"""

import io
import logging
from typing import Any, Callable, Generator, Optional
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

# XML Namespaces used in OOXML documents
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
WP_NS = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
CP_NS = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
DC_NS = "{http://purl.org/dc/elements/1.1/}"
DCTERMS_NS = "{http://purl.org/dc/terms/}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
PIC_NS = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
WPS_NS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"

# Pre-computed tag names for hot paths (avoid repeated string concatenation)
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

# EMU (English Metric Units) conversion: 914400 EMU = 1 inch
EMU_PER_INCH = 914400
# Twips conversion: 1440 twips = 1 inch
TWIPS_PER_INCH = 1440

# Caption style keywords (case-insensitive matching)
CAPTION_STYLE_KEYWORDS = ("caption", "bildunterschrift", "abbildung", "figure")

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
) -> tuple[Optional[int], Optional[int]]:
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


def _collect_text_from_element(element: ET.Element) -> str:
    """Efficiently extract all text from w:t elements within an element."""
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
    formula_converter: Callable[[ET.Element], str],
) -> None:
    """Append extracted text from a node, respecting AlternateContent and formulas."""
    tag = elem.tag

    if tag.endswith("}AlternateContent"):
        choice = elem.find(MC_CHOICE)
        if choice is not None:
            for child in choice:
                _process_text_element(child, parts, include_formulas, formula_converter)
        return

    if tag.endswith("}Fallback"):
        return

    if tag == W_R:
        for child in elem:
            if child.tag == W_T:
                if child.text:
                    parts.append(child.text)
            elif child.tag.endswith("}AlternateContent"):
                _process_text_element(child, parts, include_formulas, formula_converter)
        return

    if tag == M_OMATH:
        if include_formulas:
            latex = formula_converter(elem)
            if latex.strip():
                parts.append(f"${latex}$")
        return

    if tag == M_OMATHPARA:
        if include_formulas:
            omath = elem.find(M_OMATH)
            if omath is not None:
                latex = formula_converter(omath)
                if latex.strip():
                    parts.append(f"$${latex}$$")
        return

    for child in elem:
        _process_text_element(child, parts, include_formulas, formula_converter)


def _extract_paragraph_content(
    paragraph: ET.Element,
    include_formulas: bool,
    formula_converter: Callable[[ET.Element], str],
) -> str:
    """Extract text from a paragraph, including inline and display equations."""
    parts: list[str] = []
    for child in paragraph:
        _process_text_element(child, parts, include_formulas, formula_converter)
    return "".join(parts)


def _extract_table_text(
    table: ET.Element,
    include_formulas: bool,
    formula_converter: Callable[[ET.Element], str],
) -> list[str]:
    """Extract table text in row order, concatenating cell content."""
    texts: list[str] = []
    for row in table.iter(W_TR):
        for cell in row.iter(W_TC):
            cell_parts: list[str] = []
            for paragraph in cell.iter(W_P):
                text = _extract_paragraph_content(
                    paragraph, include_formulas, formula_converter
                )
                if text.strip():
                    cell_parts.append(text)
            if cell_parts:
                texts.append(" ".join(cell_parts))
    return texts


def _extract_full_text_from_body(
    body: ET.Element | None, include_formulas: bool = True
) -> str:
    """
    Extract the complete text content from a pre-parsed document body.

    Combines all text from paragraphs, tables, and equations into a single
    string, preserving the document order.

    Args:
        body: Pre-parsed document body element (from cached context).
        include_formulas: Whether to include LaTeX formula representations
            in the output. If True, inline formulas are wrapped in $...$
            and display formulas in $$...$$. Default is True.

    Returns:
        Complete document text as a single string with newlines between
        paragraphs and table cells.
    """
    if body is None:
        return ""

    all_text: list[str] = []

    for element in body:
        tag = element.tag

        if tag == W_P:
            text = _extract_paragraph_content(element, include_formulas, omml_to_latex)
            if text.strip():
                all_text.append(text)

        elif tag == W_TBL:
            table_texts = _extract_table_text(element, include_formulas, omml_to_latex)
            all_text.extend(table_texts)

    return "\n".join(all_text)


class _DocxContext(OOXMLZipContext):
    """
    Cached context for DOCX extraction.

    Opens the ZIP file once and caches all parsed XML documents and
    extracted data that is reused across multiple extraction functions.
    This avoids repeatedly opening the ZIP and parsing the same XML files.
    """

    def __init__(self, file_like: io.BytesIO):
        """Initialize the DOCX context and cache XML content."""
        super().__init__(file_like)

        # Cache for parsed XML roots
        self._document_root: ET.Element | None = None
        self._core_root: ET.Element | None = None
        self._styles_root: ET.Element | None = None
        self._footnotes_root: ET.Element | None = None
        self._endnotes_root: ET.Element | None = None
        self._comments_root: ET.Element | None = None
        self._rels_root: ET.Element | None = None

        # Cache for extracted data
        self._relationships: dict[str, dict] | None = None
        self._styles: dict[str, str] | None = None

        # Cache for header/footer roots (keyed by path)
        self._header_footer_roots: dict[str, ET.Element] = {}

        self._load_xml_files()

    def _load_xml_files(self) -> None:
        """Load and parse all XML files from the ZIP at once."""
        # Main document
        if "word/document.xml" in self._namelist:
            self._document_root = self.read_xml_root("word/document.xml")

        # Core properties (metadata)
        if "docProps/core.xml" in self._namelist:
            self._core_root = self.read_xml_root("docProps/core.xml")

        # Styles
        if "word/styles.xml" in self._namelist:
            self._styles_root = self.read_xml_root("word/styles.xml")

        # Footnotes
        if "word/footnotes.xml" in self._namelist:
            self._footnotes_root = self.read_xml_root("word/footnotes.xml")

        # Endnotes
        if "word/endnotes.xml" in self._namelist:
            self._endnotes_root = self.read_xml_root("word/endnotes.xml")

        # Comments
        if "word/comments.xml" in self._namelist:
            self._comments_root = self.read_xml_root("word/comments.xml")

        # Relationships
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in self._namelist:
            self._rels_root = self.read_xml_root(rels_path)

        # Pre-load header and footer files
        self._relationships = self._parse_relationships()
        for rel_id, rel_info in self._relationships.items():
            rel_type = rel_info.get("type", "")
            target = rel_info.get("target", "")
            if "header" in rel_type.lower() or "footer" in rel_type.lower():
                hf_path = "word/" + target
                if hf_path in self._namelist:
                    self._header_footer_roots[hf_path] = self.read_xml_root(hf_path)

    def _parse_relationships(self) -> dict[str, dict]:
        """Parse relationships from cached rels root."""
        relationships = {}
        if self._rels_root is None:
            return relationships

        for rel in parse_relationships(self._rels_root):
            rel_id = rel["id"]
            if not rel_id:
                continue
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


def _get_element_text(root: ET.Element | None, tag: str) -> str | None:
    """Extract text from an element if it exists and has text content."""
    if root is None:
        return None
    elem = root.find(tag)
    if elem is not None and elem.text:
        return elem.text
    return None


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


def _extract_metadata_from_context(ctx: _DocxContext) -> DocxMetadata:
    """
    Extract document metadata from cached core.xml root.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        DocxMetadata object with title, author, dates, revision, etc.
    """
    metadata = DocxMetadata()
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
        metadata.created = text
    if text := _get_element_text(root, _DCTERMS_MODIFIED):
        metadata.modified = text
    if text := _get_element_text(root, _CP_LASTMODIFIEDBY):
        metadata.last_modified_by = text

    revision_elem = root.find(_CP_REVISION)
    if revision_elem is not None and revision_elem.text:
        try:
            metadata.revision = int(revision_elem.text)
        except ValueError:
            pass

    return metadata


# Skip IDs for separator and continuation notes
_SKIP_NOTE_IDS = frozenset(["-1", "0"])


def _extract_notes_from_root(root: ET.Element | None, note_tag: str) -> list[DocxNote]:
    """
    Extract notes (footnotes or endnotes) from an XML root element.

    Args:
        root: XML root element containing notes.
        note_tag: Tag name for notes (W_FOOTNOTE or W_ENDNOTE).

    Returns:
        List of DocxNote objects, excluding separator and continuation notes.
    """
    if root is None:
        return []

    notes: list[DocxNote] = []
    for note in root.findall(f".//{note_tag}"):
        note_id = note.get(W_ID) or ""
        if note_id not in _SKIP_NOTE_IDS:
            notes.append(DocxNote(id=note_id, text=_collect_text_from_element(note)))
    return notes


def _extract_footnotes_from_context(ctx: _DocxContext) -> list[DocxNote]:
    """Extract footnotes from cached footnotes.xml root."""
    return _extract_notes_from_root(ctx._footnotes_root, W_FOOTNOTE)


def _extract_comments_from_context(ctx: _DocxContext) -> list[DocxComment]:
    """
    Extract comments/annotations from cached comments.xml root.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxComment objects with id, author, date, and text fields.
    """
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


def _extract_endnotes_from_context(ctx: _DocxContext) -> list[DocxNote]:
    """Extract endnotes from cached endnotes.xml root."""
    return _extract_notes_from_root(ctx._endnotes_root, W_ENDNOTE)


def _parse_twips_to_inches(value: str | None) -> float | None:
    """Convert twips string to inches, returning None on failure."""
    if not value:
        return None
    try:
        return int(value) / TWIPS_PER_INCH
    except ValueError:
        return None


def _extract_sections_from_context(ctx: _DocxContext) -> list[DocxSection]:
    """
    Extract section properties (page layout) from cached document body.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxSection objects with page dimensions and margins in inches.
    """
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

        # Page size
        pg_sz = sect_pr.find(W_PGSZ)
        if pg_sz is not None:
            if inches := _parse_twips_to_inches(pg_sz.get(W_W)):
                section.page_width_inches = inches
            if inches := _parse_twips_to_inches(pg_sz.get(W_H)):
                section.page_height_inches = inches
            orient = pg_sz.get(W_ORIENT)
            if orient and orient != "portrait":
                section.orientation = orient

        # Page margins
        pg_mar = sect_pr.find(W_PGMAR)
        if pg_mar is not None:
            if inches := _parse_twips_to_inches(pg_mar.get(W_LEFT)):
                section.left_margin_inches = inches
            if inches := _parse_twips_to_inches(pg_mar.get(W_RIGHT)):
                section.right_margin_inches = inches
            if inches := _parse_twips_to_inches(pg_mar.get(W_TOP)):
                section.top_margin_inches = inches
            if inches := _parse_twips_to_inches(pg_mar.get(W_BOTTOM)):
                section.bottom_margin_inches = inches

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
    """
    Extract headers and footers from cached header/footer XML roots.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        Tuple of (headers_list, footers_list) where each list contains
        DocxHeaderFooter objects with type and text fields.
    """
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

        hf_type = _determine_hf_type(hf_path, rel_type)
        hf_obj = DocxHeaderFooter(type=hf_type, text=text)

        if is_header:
            headers.append(hf_obj)
        else:
            footers.append(hf_obj)

    return headers, footers


def _parse_run_properties(
    rpr: ET.Element | None,
) -> tuple[bool | None, bool | None, bool | None, str | None, float | None, str | None]:
    """Parse run properties and return (bold, italic, underline, font_name, font_size, font_color)."""
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

    # Font size (in half-points)
    font_size = None
    sz = rpr.find(W_SZ)
    if sz is not None:
        sz_val = sz.get(W_VAL)
        if sz_val:
            try:
                font_size = int(sz_val) / 2  # Convert half-points to points
            except ValueError:
                pass

    # Font color
    font_color = None
    color = rpr.find(W_COLOR)
    if color is not None:
        font_color = color.get(W_VAL)

    return bold, italic, underline, font_name, font_size, font_color


def _extract_paragraphs_from_context(ctx: _DocxContext) -> list[DocxParagraph]:
    """
    Extract paragraphs with their formatting and run information.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxParagraph objects containing text, style, alignment,
        and a list of DocxRun objects with formatting details.
    """
    body = ctx.document_body
    if body is None:
        return []

    style_map = ctx.styles
    paragraphs: list[DocxParagraph] = []

    # Only iterate through direct children of body to get top-level paragraphs
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

        para_text = "".join(run.text for run in runs)
        paragraphs.append(
            DocxParagraph(
                text=para_text,
                style=style_name,
                alignment=alignment,
                runs=runs,
            )
        )

    return paragraphs


def _extract_tables_from_context(ctx: _DocxContext) -> list[list[list[str]]]:
    """
    Extract tables as lists of lists of cell text.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of tables, where each table is a list of rows, and each row
        is a list of cell text strings.
    """
    body = ctx.document_body
    if body is None:
        return []

    tables: list[list[list[str]]] = []
    for tbl in body.findall(f".//{W_TBL}"):
        table_data: list[list[str]] = []
        for tr in tbl.findall(W_TR):
            row_data: list[str] = []
            for tc in tr.findall(W_TC):
                # Collect text from each paragraph in the cell
                cell_text_parts = [
                    _collect_text_from_element(p) for p in tc.findall(f".//{W_P}")
                ]
                row_data.append("\n".join(cell_text_parts))
            table_data.append(row_data)
        tables.append(table_data)

    return tables


def _extract_images_from_context(ctx: _DocxContext) -> list[DocxImage]:
    """
    Extract images from the document.

    Parses the document body to find drawing elements and extracts images
    with their captions and descriptions (alt text).

    Caption is extracted from (in order of priority):
    1. A preceding paragraph with caption-like style AND keepNext attribute
       (indicates it should stay with the following element - the image)
    2. A following paragraph with caption-like style
    3. Text boxes associated with the image (wps:wsp with wps:txbx)
    4. Falls back to the name attribute of pic:cNvPr

    Description (alt text) is extracted from:
    - The descr attribute of pic:cNvPr (the picture's non-visual properties)

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxImage objects with binary data, metadata, captions,
        and descriptions.
    """
    rels = ctx.relationships
    body = ctx.document_body

    # Map of rel_id -> (caption, description) from document drawings
    image_metadata: dict[str, tuple[str, str]] = {}

    if body is not None:
        paragraphs = list(body.findall(W_P))

        for para_idx, para in enumerate(paragraphs):
            for drawing in para.iter(W_DRAWING):
                caption = ""
                description = ""

                # Extract alt text and name from pic:cNvPr
                pic_cNvPr = drawing.find(f".//{PIC_CNVPR}")
                if pic_cNvPr is not None:
                    if descr := pic_cNvPr.get("descr", ""):
                        description = descr
                    if name := pic_cNvPr.get("name", ""):
                        caption = name

                # Check text boxes for caption
                for wsp in drawing.iter(WPS_WSP):
                    txbx = wsp.find(WPS_TXBX)
                    if txbx is not None:
                        text = _collect_text_from_element(txbx)
                        if text:
                            caption = text
                            break

                # Check preceding paragraph for caption
                preceding_caption = None
                if para_idx > 0:
                    prev_para = paragraphs[para_idx - 1]
                    prev_style = _get_paragraph_style(prev_para)
                    if _is_caption_style(prev_style) and _has_keep_next(prev_para):
                        if text := _collect_text_from_element(prev_para):
                            preceding_caption = text

                # Check following paragraph for caption
                following_caption = None
                if para_idx + 1 < len(paragraphs):
                    next_para = paragraphs[para_idx + 1]
                    if _is_caption_style(_get_paragraph_style(next_para)):
                        if text := _collect_text_from_element(next_para):
                            following_caption = text

                # Apply caption priority
                if preceding_caption:
                    caption = preceding_caption
                elif following_caption:
                    caption = following_caption

                # Find relationship ID for the image
                blip = drawing.find(f".//{A_BLIP}")
                if blip is not None:
                    if r_embed := blip.get(R_EMBED):
                        image_metadata[r_embed] = (caption, description)

    # Extract images using relationships
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
            content_type = _CONTENT_TYPE_MAP.get(ext, f"image/{ext}")
            caption, description = image_metadata.get(rel_id, ("", ""))
            width, height = _get_image_pixel_dimensions(img_data)

            images.append(
                DocxImage(
                    rel_id=rel_id,
                    filename=target.rsplit("/", 1)[-1],
                    content_type=content_type,
                    data=io.BytesIO(img_data),
                    size_bytes=len(img_data),
                    width=width,
                    height=height,
                    image_index=image_counter,
                    caption=caption,
                    description=description,
                )
            )
        except Exception as e:
            logger.debug(f"Image extraction failed for rel_id {rel_id} - {e}")
            images.append(DocxImage(rel_id=rel_id, error=str(e)))

    return images


def _extract_hyperlinks_from_context(ctx: _DocxContext) -> list[DocxHyperlink]:
    """
    Extract hyperlinks from the document.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxHyperlink objects with text and URL.
    """
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
    """
    Extract all mathematical formulas from the document as LaTeX.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxFormula objects with:
        - latex: LaTeX representation of the formula
        - is_display: True for display equations (oMathPara), False for inline
    """
    body = ctx.document_body
    if body is None:
        return []

    formulas: list[DocxFormula] = []
    omath_in_para: set[int] = set()

    # First, find all oMathPara elements and their child oMath
    for omath_para in body.iter(M_OMATHPARA):
        omath = omath_para.find(M_OMATH)
        if omath is not None:
            omath_in_para.add(id(omath))
            latex = omml_to_latex(omath)
            if latex.strip():
                formulas.append(DocxFormula(latex=latex, is_display=True))

    # Then find inline oMath elements (not in oMathPara)
    for omath in body.iter(M_OMATH):
        if id(omath) not in omath_in_para:
            latex = omml_to_latex(omath)
            if latex.strip():
                formulas.append(DocxFormula(latex=latex, is_display=False))

    return formulas


def read_docx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocxContent, Any, None]:
    """
    Extract all relevant content from a Word .docx file.

    Primary entry point for DOCX file extraction. Parses the document structure,
    extracts text, formatting, and metadata using direct XML parsing of the
    docx ZIP archive.

    This function uses a generator pattern for API consistency with other
    extractors, even though DOCX files contain exactly one document.

    Args:
        file_like: BytesIO object containing the complete DOCX file data.
            The stream position is reset to the beginning before reading.
        path: Optional filesystem path to the source file. If provided,
            populates file metadata (filename, extension, folder) in the
            returned DocxContent.metadata.

    Yields:
        DocxContent: Single DocxContent object containing:
            - metadata: DocxMetadata with title, author, dates, revision
            - paragraphs: List of DocxParagraph with text and runs
            - tables: List of tables as 2D lists of cell text
            - headers, footers: Header/footer content by type
            - images: List of DocxImage with binary data
            - hyperlinks: List of DocxHyperlink with text and URL
            - footnotes, endnotes: Note content
            - comments: Comment content with author and date
            - sections: Page layout information
            - styles: List of style names used
            - formulas: List of DocxFormula as LaTeX
            - full_text: Complete text including formulas
            - base_full_text: Complete text without formulas

    Raises:
        ExtractionFileEncryptedError: If the DOCX is encrypted or password-protected.

    Example:
        >>> import io
        >>> with open("report.docx", "rb") as f:
        ...     data = io.BytesIO(f.read())
        ...     for doc in read_docx(data, path="report.docx"):
        ...         print(f"Title: {doc.metadata.title}")
        ...         print(f"Tables: {len(doc.tables)}")
        ...         print(f"Images: {len(doc.images)}")
        ...         print(doc.full_text[:500])

    Performance Notes:
        - ZIP file is opened once and all XML is cached
        - All XML documents are parsed once and reused
        - Images are loaded into memory as BytesIO objects
        - Large documents may use significant memory
    """
    try:
        file_like.seek(0)
        if is_ooxml_encrypted(file_like):
            raise ExtractionFileEncryptedError(
                "DOCX is encrypted or password-protected"
            )

        # Create context that opens ZIP once and caches all parsed XML
        ctx = _DocxContext(file_like)
        try:
            # === Core Properties (Metadata) ===
            metadata = _extract_metadata_from_context(ctx)

            # === Paragraphs ===
            paragraphs = _extract_paragraphs_from_context(ctx)

            # === Tables ===
            tables = _extract_tables_from_context(ctx)

            # === Headers and Footers ===
            headers, footers = _extract_header_footers_from_context(ctx)

            # === Images ===
            images = _extract_images_from_context(ctx)

            # === Hyperlinks ===
            hyperlinks = _extract_hyperlinks_from_context(ctx)

            # === Footnotes ===
            footnotes = _extract_footnotes_from_context(ctx)

            # === Endnotes ===
            endnotes = _extract_endnotes_from_context(ctx)

            # === Formulas ===
            formulas = _extract_formulas_from_context(ctx)

            # === Comments ===
            comments = _extract_comments_from_context(ctx)

            # === Sections (page layout) ===
            sections = _extract_sections_from_context(ctx)

            # === Styles used ===
            styles = list({para.style for para in paragraphs if para.style})

            # === Full text (convenience) - use cached body for both ===
            body = ctx.document_body
            full_text = _extract_full_text_from_body(body, include_formulas=True)
            base_full_text = _extract_full_text_from_body(body, include_formulas=False)

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
                base_full_text=base_full_text,
            )
        finally:
            ctx.close()
    except ExtractionError:
        raise
    except Exception as exc:
        raise ExtractionFailedError("Failed to extract DOCX file", cause=exc) from exc
