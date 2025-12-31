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
import zipfile
from typing import Any, Generator, Optional
from xml.etree import ElementTree as ET

from sharepoint2text.exceptions import ExtractionFileEncryptedError
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
from sharepoint2text.extractors.util.omml_to_latex import (
    GREEK_TO_LATEX,
    convert_greek_and_symbols,
    omml_to_latex,
)

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

# EMU (English Metric Units) conversion: 914400 EMU = 1 inch
EMU_PER_INCH = 914400
# Twips conversion: 1440 twips = 1 inch
TWIPS_PER_INCH = 1440


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


class _DocxFullTextExtractor:
    """
    Extracts a complete text representation from a DOCX file.

    This class handles the complex task of extracting text from Word documents
    while preserving the order of paragraphs, tables, and mathematical formulas.
    It processes the raw XML structure of the document.

    Key Features:
        - Preserves document element order (paragraphs, tables, formulas)
        - Converts OMML math formulas to LaTeX notation
        - Handles AlternateContent elements correctly (avoids duplicates)
        - Supports both inline ($...$) and display ($$...$$) math

    Class Attributes:
        GREEK_TO_LATEX: Mapping of Greek letters and math symbols to LaTeX
            (imported from omml_to_latex module)

    Usage:
        >>> text = _DocxFullTextExtractor.extract_full_text(file_like)
        >>> latex = _DocxFullTextExtractor.omml_to_latex(omath_element)
    """

    # Re-export from omml_to_latex module for backwards compatibility
    GREEK_TO_LATEX = GREEK_TO_LATEX

    @classmethod
    def _convert_greek_and_symbols(cls, text: str) -> str:
        """
        Convert Greek letters and math symbols to LaTeX equivalents.

        Args:
            text: Input string potentially containing Unicode Greek/math chars.

        Returns:
            String with Greek letters and symbols replaced by LaTeX commands.

        Example:
            >>> _DocxFullTextExtractor._convert_greek_and_symbols("αβγ")
            '\\alpha\\beta\\gamma'
        """
        return convert_greek_and_symbols(text)

    @classmethod
    def omml_to_latex(cls, omath_element) -> str:
        """
        Convert an OMML (Office Math Markup Language) element to LaTeX notation.

        This method delegates to the omml_to_latex module for the actual conversion.

        Args:
            omath_element: An ElementTree Element representing an m:oMath or m:oMathPara
                element from the document XML.

        Returns:
            LaTeX string representation of the mathematical expression.

        See Also:
            sharepoint2text.extractors.omml_to_latex.omml_to_latex for full documentation.
        """
        return omml_to_latex(omath_element)

    @classmethod
    def extract_full_text_from_body(
        cls, body: ET.Element | None, include_formulas: bool = True
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
        logger.debug("Extracting document full text")

        if body is None:
            return ""

        all_text = []

        def process_element(elem, parts: list):
            """Recursively process element, handling AlternateContent properly.

            Only processes mc:Choice content and skips mc:Fallback to avoid
            extracting duplicate content from fallback representations.
            """
            tag = elem.tag.split("}")[-1]

            # Handle AlternateContent - only use Choice, skip Fallback
            if tag == "AlternateContent":
                choice = elem.find(f"{MC_NS}Choice")
                if choice is not None:
                    for child in choice:
                        process_element(child, parts)
                return

            # Skip Fallback elements entirely to avoid duplicate content
            if tag == "Fallback":
                return

            # Regular run of text
            if tag == "r":
                for child in elem:
                    child_tag = child.tag.split("}")[-1]
                    if child_tag == "t":
                        if child.text:
                            parts.append(child.text)
                    elif child_tag == "AlternateContent":
                        process_element(child, parts)
                return

            # Inline equation
            if tag == "oMath":
                if include_formulas:
                    latex = cls.omml_to_latex(elem)
                    if latex.strip():
                        parts.append(f"${latex}$")
                return

            # Display equation
            if tag == "oMathPara":
                if include_formulas:
                    omath = elem.find(f"{M_NS}oMath")
                    if omath is not None:
                        latex = cls.omml_to_latex(omath)
                        if latex.strip():
                            parts.append(f"$${latex}$$")
                return

            # Recurse into other elements
            for child in elem:
                process_element(child, parts)

        def extract_paragraph_content(p_element) -> str:
            """Extract text from paragraph including inline and display equations."""
            parts = []
            for child in p_element:
                process_element(child, parts)
            return "".join(parts)

        def extract_table_text(tbl_element) -> list[str]:
            """Extract text from table in row order."""
            texts = []
            for row in tbl_element.iter(f"{W_NS}tr"):
                for cell in row.iter(f"{W_NS}tc"):
                    cell_parts = []
                    for p in cell.iter(f"{W_NS}p"):
                        text = extract_paragraph_content(p)
                        if text.strip():
                            cell_parts.append(text)
                    if cell_parts:
                        texts.append(" ".join(cell_parts))
            return texts

        # Iterate through body elements in document order
        for element in body:
            tag = element.tag.split("}")[-1]

            if tag == "p":  # Paragraph (may contain oMathPara)
                text = extract_paragraph_content(element)
                if text.strip():
                    all_text.append(text)

            elif tag == "tbl":  # Table
                table_texts = extract_table_text(element)
                all_text.extend(table_texts)

        return "\n".join(all_text)


class _DocxContext:
    """
    Cached context for DOCX extraction.

    Opens the ZIP file once and caches all parsed XML documents and
    extracted data that is reused across multiple extraction functions.
    This avoids repeatedly opening the ZIP and parsing the same XML files.
    """

    def __init__(self, file_like: io.BytesIO):
        self.file_like = file_like
        file_like.seek(0)

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
        self._namelist: set[str] | None = None

        # Cache for header/footer roots (keyed by path)
        self._header_footer_roots: dict[str, ET.Element] = {}

        # Open ZIP once and read all needed files
        with zipfile.ZipFile(file_like, "r") as z:
            self._namelist = set(z.namelist())
            self._load_xml_files(z)

    def _load_xml_files(self, z: zipfile.ZipFile) -> None:
        """Load and parse all XML files from the ZIP at once."""
        # Main document
        if "word/document.xml" in self._namelist:
            with z.open("word/document.xml") as f:
                self._document_root = ET.parse(f).getroot()

        # Core properties (metadata)
        if "docProps/core.xml" in self._namelist:
            with z.open("docProps/core.xml") as f:
                self._core_root = ET.parse(f).getroot()

        # Styles
        if "word/styles.xml" in self._namelist:
            with z.open("word/styles.xml") as f:
                self._styles_root = ET.parse(f).getroot()

        # Footnotes
        if "word/footnotes.xml" in self._namelist:
            with z.open("word/footnotes.xml") as f:
                self._footnotes_root = ET.parse(f).getroot()

        # Endnotes
        if "word/endnotes.xml" in self._namelist:
            with z.open("word/endnotes.xml") as f:
                self._endnotes_root = ET.parse(f).getroot()

        # Comments
        if "word/comments.xml" in self._namelist:
            with z.open("word/comments.xml") as f:
                self._comments_root = ET.parse(f).getroot()

        # Relationships
        rels_path = "word/_rels/document.xml.rels"
        if rels_path in self._namelist:
            with z.open(rels_path) as f:
                self._rels_root = ET.parse(f).getroot()

        # Pre-load header and footer files
        self._relationships = self._parse_relationships()
        for rel_id, rel_info in self._relationships.items():
            rel_type = rel_info.get("type", "")
            target = rel_info.get("target", "")
            if "header" in rel_type.lower() or "footer" in rel_type.lower():
                hf_path = "word/" + target
                if hf_path in self._namelist:
                    with z.open(hf_path) as f:
                        self._header_footer_roots[hf_path] = ET.parse(f).getroot()

    def _parse_relationships(self) -> dict[str, dict]:
        """Parse relationships from cached rels root."""
        relationships = {}
        if self._rels_root is None:
            return relationships

        for rel in self._rels_root.findall(f".//{REL_NS}Relationship"):
            rel_id = rel.get("Id") or ""
            rel_type = rel.get("Type") or ""
            rel_target = rel.get("Target") or ""
            target_mode = rel.get("TargetMode") or ""
            relationships[rel_id] = {
                "type": rel_type,
                "target": rel_target,
                "target_mode": target_mode,
            }
        return relationships

    @property
    def document_body(self) -> ET.Element | None:
        """Get the document body element."""
        if self._document_root is None:
            return None
        return self._document_root.find(f"{W_NS}body")

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
                for style in self._styles_root.findall(f".//{W_NS}style"):
                    style_id = style.get(f"{W_NS}styleId") or ""
                    name_elem = style.find(f"{W_NS}name")
                    style_name = (
                        name_elem.get(f"{W_NS}val") if name_elem is not None else ""
                    )
                    if style_id:
                        self._styles[style_id] = style_name or style_id
        return self._styles

    def get_image_data(self, image_path: str) -> bytes | None:
        """Read image data from the ZIP file."""
        if image_path not in self._namelist:
            return None
        self.file_like.seek(0)
        with zipfile.ZipFile(self.file_like, "r") as z:
            with z.open(image_path) as img_file:
                return img_file.read()


def _extract_metadata_from_context(ctx: _DocxContext) -> DocxMetadata:
    """
    Extract document metadata from cached core.xml root.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        DocxMetadata object with title, author, dates, revision, etc.
    """
    logger.debug("Extracting metadata")
    metadata = DocxMetadata()

    root = ctx._core_root
    if root is None:
        return metadata

    # Extract metadata fields
    title_elem = root.find(f"{DC_NS}title")
    if title_elem is not None and title_elem.text:
        metadata.title = title_elem.text

    creator_elem = root.find(f"{DC_NS}creator")
    if creator_elem is not None and creator_elem.text:
        metadata.author = creator_elem.text

    subject_elem = root.find(f"{DC_NS}subject")
    if subject_elem is not None and subject_elem.text:
        metadata.subject = subject_elem.text

    # Keywords - may be in cp:keywords or dc:subject
    keywords_elem = root.find(f"{CP_NS}keywords")
    if keywords_elem is not None and keywords_elem.text:
        metadata.keywords = keywords_elem.text

    category_elem = root.find(f"{CP_NS}category")
    if category_elem is not None and category_elem.text:
        metadata.category = category_elem.text

    description_elem = root.find(f"{DC_NS}description")
    if description_elem is not None and description_elem.text:
        metadata.comments = description_elem.text

    created_elem = root.find(f"{DCTERMS_NS}created")
    if created_elem is not None and created_elem.text:
        metadata.created = created_elem.text

    modified_elem = root.find(f"{DCTERMS_NS}modified")
    if modified_elem is not None and modified_elem.text:
        metadata.modified = modified_elem.text

    last_modified_by_elem = root.find(f"{CP_NS}lastModifiedBy")
    if last_modified_by_elem is not None and last_modified_by_elem.text:
        metadata.last_modified_by = last_modified_by_elem.text

    revision_elem = root.find(f"{CP_NS}revision")
    if revision_elem is not None and revision_elem.text:
        try:
            metadata.revision = int(revision_elem.text)
        except ValueError:
            pass

    return metadata


def _extract_footnotes_from_context(ctx: _DocxContext) -> list[DocxNote]:
    """
    Extract footnotes from cached footnotes.xml root.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxNote objects with id and text fields.
        Separator (-1) and continuation (0) footnotes are filtered out.
    """
    logger.debug("Extracting footnotes")
    footnotes = []

    root = ctx._footnotes_root
    if root is None:
        return footnotes

    for fn in root.findall(f".//{W_NS}footnote"):
        fn_id = fn.get(f"{W_NS}id") or ""
        if fn_id not in ["-1", "0"]:  # Skip separator and continuation footnotes
            text_parts = []
            for t in fn.findall(f".//{W_NS}t"):
                if t.text:
                    text_parts.append(t.text)
            footnotes.append(DocxNote(id=fn_id, text="".join(text_parts)))

    return footnotes


def _extract_comments_from_context(ctx: _DocxContext) -> list[DocxComment]:
    """
    Extract comments/annotations from cached comments.xml root.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxComment objects with id, author, date, and text fields.
    """
    logger.debug("Extracting comments")
    comments = []

    root = ctx._comments_root
    if root is None:
        return comments

    for comment in root.findall(f".//{W_NS}comment"):
        text_parts = []
        for t in comment.findall(f".//{W_NS}t"):
            if t.text:
                text_parts.append(t.text)
        comments.append(
            DocxComment(
                id=comment.get(f"{W_NS}id") or "",
                author=comment.get(f"{W_NS}author") or "",
                date=comment.get(f"{W_NS}date") or "",
                text="".join(text_parts),
            )
        )

    return comments


def _extract_endnotes_from_context(ctx: _DocxContext) -> list[DocxNote]:
    """
    Extract endnotes from cached endnotes.xml root.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxNote objects with id and text fields.
        Separator (-1) and continuation (0) endnotes are filtered out.
    """
    logger.debug("Extracting endnotes")
    endnotes = []

    root = ctx._endnotes_root
    if root is None:
        return endnotes

    for en in root.findall(f".//{W_NS}endnote"):
        en_id = en.get(f"{W_NS}id") or ""
        if en_id not in ["-1", "0"]:  # Skip separator and continuation endnotes
            text_parts = []
            for t in en.findall(f".//{W_NS}t"):
                if t.text:
                    text_parts.append(t.text)
            endnotes.append(DocxNote(id=en_id, text="".join(text_parts)))

    return endnotes


def _extract_sections_from_context(ctx: _DocxContext) -> list[DocxSection]:
    """
    Extract section properties (page layout) from cached document body.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxSection objects with page dimensions and margins in inches.
    """
    logger.debug("Extracting sections")
    sections = []

    body = ctx.document_body
    if body is None:
        return sections

    # Find all sectPr elements (in paragraphs and at end of body)
    sect_pr_elements = []

    # Sections in paragraphs
    for p in body.findall(f".//{W_NS}p"):
        ppr = p.find(f"{W_NS}pPr")
        if ppr is not None:
            sect_pr = ppr.find(f"{W_NS}sectPr")
            if sect_pr is not None:
                sect_pr_elements.append(sect_pr)

    # Final section at end of body
    final_sect_pr = body.find(f"{W_NS}sectPr")
    if final_sect_pr is not None:
        sect_pr_elements.append(final_sect_pr)

    for sect_pr in sect_pr_elements:
        section = DocxSection()

        # Page size
        pg_sz = sect_pr.find(f"{W_NS}pgSz")
        if pg_sz is not None:
            w_val = pg_sz.get(f"{W_NS}w")
            h_val = pg_sz.get(f"{W_NS}h")
            orient = pg_sz.get(f"{W_NS}orient")

            if w_val:
                try:
                    section.page_width_inches = int(w_val) / TWIPS_PER_INCH
                except ValueError:
                    pass
            if h_val:
                try:
                    section.page_height_inches = int(h_val) / TWIPS_PER_INCH
                except ValueError:
                    pass
            # Only set orientation for non-default (landscape)
            # Portrait is the default and should remain as None
            if orient and orient != "portrait":
                section.orientation = orient

        # Page margins
        pg_mar = sect_pr.find(f"{W_NS}pgMar")
        if pg_mar is not None:
            left = pg_mar.get(f"{W_NS}left")
            right = pg_mar.get(f"{W_NS}right")
            top = pg_mar.get(f"{W_NS}top")
            bottom = pg_mar.get(f"{W_NS}bottom")

            if left:
                try:
                    section.left_margin_inches = int(left) / TWIPS_PER_INCH
                except ValueError:
                    pass
            if right:
                try:
                    section.right_margin_inches = int(right) / TWIPS_PER_INCH
                except ValueError:
                    pass
            if top:
                try:
                    section.top_margin_inches = int(top) / TWIPS_PER_INCH
                except ValueError:
                    pass
            if bottom:
                try:
                    section.bottom_margin_inches = int(bottom) / TWIPS_PER_INCH
                except ValueError:
                    pass

        sections.append(section)

    return sections


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
    logger.debug("Extracting header/footer")
    headers = []
    footers = []

    rels = ctx.relationships

    # Find header and footer files
    header_files = []
    footer_files = []

    for rel_id, rel_info in rels.items():
        rel_type = rel_info.get("type", "")
        target = rel_info.get("target", "")

        if "header" in rel_type.lower():
            header_files.append(("word/" + target, rel_type))
        elif "footer" in rel_type.lower():
            footer_files.append(("word/" + target, rel_type))

    # Extract text from header files
    for header_path, rel_type in header_files:
        root = ctx._header_footer_roots.get(header_path)
        if root is not None:
            text_parts = []
            for t in root.findall(f".//{W_NS}t"):
                if t.text:
                    text_parts.append(t.text)

            if text_parts:
                # Determine type from filename or relationship
                hdr_type = "default"
                if "first" in header_path.lower() or "first" in rel_type.lower():
                    hdr_type = "first_page"
                elif "even" in header_path.lower() or "even" in rel_type.lower():
                    hdr_type = "even_page"

                headers.append(
                    DocxHeaderFooter(type=hdr_type, text="".join(text_parts))
                )

    # Extract text from footer files
    for footer_path, rel_type in footer_files:
        root = ctx._header_footer_roots.get(footer_path)
        if root is not None:
            text_parts = []
            for t in root.findall(f".//{W_NS}t"):
                if t.text:
                    text_parts.append(t.text)

            if text_parts:
                # Determine type from filename or relationship
                ftr_type = "default"
                if "first" in footer_path.lower() or "first" in rel_type.lower():
                    ftr_type = "first_page"
                elif "even" in footer_path.lower() or "even" in rel_type.lower():
                    ftr_type = "even_page"

                footers.append(
                    DocxHeaderFooter(type=ftr_type, text="".join(text_parts))
                )

    return headers, footers


def _extract_paragraphs_from_context(ctx: _DocxContext) -> list[DocxParagraph]:
    """
    Extract paragraphs with their formatting and run information.

    Args:
        ctx: DocxContext with cached XML roots.

    Returns:
        List of DocxParagraph objects containing text, style, alignment,
        and a list of DocxRun objects with formatting details.
    """
    logger.debug("Extracting paragraphs")
    paragraphs = []

    body = ctx.document_body
    if body is None:
        return paragraphs

    style_map = ctx.styles

    # Only iterate through direct children of body to get top-level paragraphs
    # This excludes paragraphs nested inside tables (which are extracted separately)
    for p in body.findall(f"{W_NS}p"):
        # Get paragraph properties
        ppr = p.find(f"{W_NS}pPr")
        style_id = None
        alignment = None

        if ppr is not None:
            style_elem = ppr.find(f"{W_NS}pStyle")
            if style_elem is not None:
                style_id = style_elem.get(f"{W_NS}val")

            jc_elem = ppr.find(f"{W_NS}jc")
            if jc_elem is not None:
                alignment = jc_elem.get(f"{W_NS}val")

        # Get style name from style ID
        style_name = style_map.get(style_id, style_id) if style_id else None

        # Extract runs
        runs = []
        for r in p.findall(f".//{W_NS}r"):
            run_text_parts = []
            for t in r.findall(f".//{W_NS}t"):
                if t.text:
                    run_text_parts.append(t.text)

            run_text = "".join(run_text_parts)
            if not run_text:
                continue

            # Get run properties
            rpr = r.find(f"{W_NS}rPr")
            bold = None
            italic = None
            underline = None
            font_name = None
            font_size = None
            font_color = None

            if rpr is not None:
                bold_elem = rpr.find(f"{W_NS}b")
                if bold_elem is not None:
                    bold_val = bold_elem.get(f"{W_NS}val")
                    bold = bold_val != "0" if bold_val else True

                italic_elem = rpr.find(f"{W_NS}i")
                if italic_elem is not None:
                    italic_val = italic_elem.get(f"{W_NS}val")
                    italic = italic_val != "0" if italic_val else True

                underline_elem = rpr.find(f"{W_NS}u")
                if underline_elem is not None:
                    u_val = underline_elem.get(f"{W_NS}val")
                    underline = u_val and u_val != "none"

                # Font name from rFonts
                rfonts = rpr.find(f"{W_NS}rFonts")
                if rfonts is not None:
                    font_name = (
                        rfonts.get(f"{W_NS}ascii")
                        or rfonts.get(f"{W_NS}hAnsi")
                        or rfonts.get(f"{W_NS}cs")
                    )

                # Font size (in half-points)
                sz = rpr.find(f"{W_NS}sz")
                if sz is not None:
                    sz_val = sz.get(f"{W_NS}val")
                    if sz_val:
                        try:
                            font_size = int(sz_val) / 2  # Convert half-points to points
                        except ValueError:
                            pass

                # Font color
                color = rpr.find(f"{W_NS}color")
                if color is not None:
                    font_color = color.get(f"{W_NS}val")

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

        # Get full paragraph text
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
    logger.debug("Extracting tables")
    tables = []

    body = ctx.document_body
    if body is None:
        return tables

    for tbl in body.findall(f".//{W_NS}tbl"):
        table_data = []
        for tr in tbl.findall(f"{W_NS}tr"):
            row_data = []
            for tc in tr.findall(f"{W_NS}tc"):
                cell_text_parts = []
                for p in tc.findall(f".//{W_NS}p"):
                    para_text_parts = []
                    for t in p.findall(f".//{W_NS}t"):
                        if t.text:
                            para_text_parts.append(t.text)
                    cell_text_parts.append("".join(para_text_parts))
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
    logger.debug("Extracting images")
    images = []
    rels = ctx.relationships
    body = ctx.document_body

    # Map of rel_id -> (caption, description) from document drawings
    image_metadata: dict[str, tuple[str, str]] = {}

    # Namespace for picture elements
    PIC_NS = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
    # Namespace for WordprocessingML shapes
    WPS_NS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"

    # Caption style keywords (case-insensitive matching)
    CAPTION_STYLE_KEYWORDS = ("caption", "bildunterschrift", "abbildung", "figure")

    def _get_paragraph_style(para) -> str:
        """Get the style name from a paragraph."""
        pPr = para.find(f"{W_NS}pPr")
        if pPr is not None:
            pStyle = pPr.find(f"{W_NS}pStyle")
            if pStyle is not None:
                return pStyle.get(f"{W_NS}val", "")
        return ""

    def _has_keep_next(para) -> bool:
        """Check if a paragraph has the keepNext property."""
        pPr = para.find(f"{W_NS}pPr")
        if pPr is not None:
            keep_next = pPr.find(f"{W_NS}keepNext")
            # keepNext is present and not explicitly set to false
            if keep_next is not None:
                val = keep_next.get(f"{W_NS}val", "true")
                return val.lower() not in ("false", "0")
        return False

    def _extract_paragraph_text(para) -> str:
        """Extract all text from a paragraph."""
        text_parts = []
        for t in para.iter(f"{W_NS}t"):
            if t.text:
                text_parts.append(t.text)
        return "".join(text_parts)

    def _is_caption_style(style_name: str) -> bool:
        """Check if a style name indicates a caption paragraph."""
        style_lower = style_name.lower()
        return any(kw in style_lower for kw in CAPTION_STYLE_KEYWORDS)

    if body is not None:
        # Get all paragraphs as a list for sibling access
        paragraphs = list(body.findall(f"{W_NS}p"))

        for para_idx, para in enumerate(paragraphs):
            # Find drawings in this paragraph
            for drawing in para.iter(f"{W_NS}drawing"):
                caption = ""
                description = ""

                # Look for picture element (pic:pic) which contains the actual image
                # The pic:cNvPr element has the image's name and description (alt text)
                pic_cNvPr = drawing.find(f".//{PIC_NS}cNvPr")
                if pic_cNvPr is not None:
                    # descr attribute is the alt text / description for accessibility
                    descr = pic_cNvPr.get("descr", "")
                    if descr:
                        description = descr
                    # name attribute can be a fallback for caption
                    name = pic_cNvPr.get("name", "")
                    if name:
                        caption = name

                # Look for caption text in associated text boxes (wps:wsp with wps:txbx)
                # This is used for image captions/subtitles within the drawing group
                for wsp in drawing.iter(f"{WPS_NS}wsp"):
                    txbx = wsp.find(f"{WPS_NS}txbx")
                    if txbx is not None:
                        # Extract text from the text box content
                        text_parts = []
                        for t in txbx.iter(f"{W_NS}t"):
                            if t.text:
                                text_parts.append(t.text)
                        if text_parts:
                            caption = "".join(text_parts)
                            break  # Use the first text box as caption

                # Check for caption in PRECEDING paragraph with caption-like style
                # AND keepNext attribute (indicates it should stay with the image)
                # This is the most reliable indicator of a title caption
                preceding_caption = None
                if para_idx > 0:
                    prev_para = paragraphs[para_idx - 1]
                    prev_style = _get_paragraph_style(prev_para)
                    if _is_caption_style(prev_style) and _has_keep_next(prev_para):
                        caption_text = _extract_paragraph_text(prev_para)
                        if caption_text:
                            preceding_caption = caption_text

                # Check for caption in FOLLOWING paragraph with caption-like style
                following_caption = None
                if para_idx + 1 < len(paragraphs):
                    next_para = paragraphs[para_idx + 1]
                    next_style = _get_paragraph_style(next_para)
                    if _is_caption_style(next_style):
                        caption_text = _extract_paragraph_text(next_para)
                        if caption_text:
                            following_caption = caption_text

                # Priority: preceding caption (with keepNext) > following caption
                # > text box caption > name attribute
                if preceding_caption:
                    caption = preceding_caption
                elif following_caption:
                    caption = following_caption

                # Find the relationship ID for the image
                # Path: a:graphic > a:graphicData > pic:pic > pic:blipFill
                #       > a:blip[@r:embed]
                blip = drawing.find(f".//{A_NS}blip")
                if blip is not None:
                    r_embed = blip.get(f"{R_NS}embed")
                    if r_embed:
                        image_metadata[r_embed] = (caption, description)

    # Now extract images using relationships, with metadata from drawings
    image_counter = 0
    for rel_id, rel_info in rels.items():
        rel_type = rel_info.get("type", "")
        target = rel_info.get("target", "")

        if "image" in rel_type.lower():
            image_path = "word/" + target
            try:
                img_data = ctx.get_image_data(image_path)
                if img_data is None:
                    continue

                image_counter += 1

                # Determine content type from extension
                ext = target.split(".")[-1].lower()
                content_type_map = {
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
                content_type = content_type_map.get(ext, f"image/{ext}")

                # Get caption and description from document drawings
                caption, description = image_metadata.get(rel_id, ("", ""))
                width, height = _get_image_pixel_dimensions(img_data)

                images.append(
                    DocxImage(
                        rel_id=rel_id,
                        filename=target.split("/")[-1],
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
    logger.debug("Extracting hyperlinks")
    hyperlinks = []

    body = ctx.document_body
    if body is None:
        return hyperlinks

    rels = ctx.relationships

    for hyperlink in body.findall(f".//{W_NS}hyperlink"):
        r_id = hyperlink.get(f"{R_NS}id")
        if r_id and r_id in rels:
            rel_info = rels[r_id]
            if "hyperlink" in rel_info.get("type", "").lower():
                text_parts = []
                for t in hyperlink.findall(f".//{W_NS}t"):
                    if t.text:
                        text_parts.append(t.text)
                hyperlinks.append(
                    DocxHyperlink(
                        text="".join(text_parts), url=rel_info.get("target", "")
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
    logger.debug("Extracting formulas")
    formulas = []

    body = ctx.document_body
    if body is None:
        return formulas

    # Track oMath elements that are inside oMathPara to avoid duplicates
    omath_in_para = set()

    # First, find all oMathPara elements and their child oMath
    for omath_para in body.iter(f"{M_NS}oMathPara"):
        omath = omath_para.find(f"{M_NS}oMath")
        if omath is not None:
            omath_in_para.add(id(omath))
            latex = _DocxFullTextExtractor.omml_to_latex(omath)
            if latex.strip():
                formulas.append(DocxFormula(latex=latex, is_display=True))

    # Then find inline oMath elements (not in oMathPara)
    for omath in body.iter(f"{M_NS}oMath"):
        if id(omath) not in omath_in_para:
            latex = _DocxFullTextExtractor.omml_to_latex(omath)
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
    file_like.seek(0)
    if is_ooxml_encrypted(file_like):
        raise ExtractionFileEncryptedError("DOCX is encrypted or password-protected")

    # Create context that opens ZIP once and caches all parsed XML
    ctx = _DocxContext(file_like)

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
    styles_set = set()
    for para in paragraphs:
        if para.style:
            styles_set.add(para.style)
    styles = list(styles_set)

    # === Full text (convenience) - use cached body for both ===
    body = ctx.document_body
    full_text = _DocxFullTextExtractor.extract_full_text_from_body(
        body=body, include_formulas=True
    )
    base_full_text = _DocxFullTextExtractor.extract_full_text_from_body(
        body=body, include_formulas=False
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
