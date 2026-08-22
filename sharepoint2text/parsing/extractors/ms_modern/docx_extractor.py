"""
DOCX Document Extractor

Extracts text content, metadata, and structure from Microsoft Word .docx files
(Office Open XML format, Word 2007+).

Uses direct XML parsing of the docx ZIP archive structure for all content
extraction, without requiring the python-docx library.
"""

import io
import logging
from dataclasses import dataclass
from typing import Any, Generator

from sharepoint2text.parsing import _defused_xml as ET
from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors._records import (
    DocxComment,
    DocxFormula,
    DocxHeaderFooter,
    DocxHyperlink,
    DocxImage,
    DocxMetadata,
    DocxNote,
    DocxParagraph,
    DocxParserOutput,
    DocxRun,
    DocxSection,
)
from sharepoint2text.parsing.extractors.ms_modern.omml_to_latex import omml_to_latex
from sharepoint2text.parsing.extractors.ms_modern.ooxml_namespaces import (
    _CP_CATEGORY,
    _CP_KEYWORDS,
    _CP_LASTMODIFIEDBY,
    _CP_REVISION,
    _DC_CREATOR,
    _DC_DESCRIPTION,
    _DC_SUBJECT,
    _DC_TITLE,
    _DCTERMS_CREATED,
    _DCTERMS_MODIFIED,
    A_BLIP,
    CAPTION_STYLE_KEYWORDS,
    DOCX_NAMESPACES,
    M_OMATH,
    M_OMATHPARA,
    PIC_CNVPR,
    R_EMBED,
    R_ID,
    SKIP_NOTE_IDS,
    W_ASCII,
    W_AUTHOR,
    W_B,
    W_BODY,
    W_BOTTOM,
    W_BR,
    W_COLOR,
    W_COMMENT,
    W_CS,
    W_DATE,
    W_DRAWING,
    W_ENDNOTE,
    W_FOOTNOTE,
    W_H,
    W_HANSI,
    W_HYPERLINK,
    W_I,
    W_ID,
    W_JC,
    W_KEEPNEXT,
    W_LAST_RENDERED_PAGE_BREAK,
    W_LEFT,
    W_NAME,
    W_ORIENT,
    W_P,
    W_PGMAR,
    W_PGSZ,
    W_PPR,
    W_PSTYLE,
    W_R,
    W_RFONTS,
    W_RIGHT,
    W_RPR,
    W_SECTPR,
    W_STYLE,
    W_STYLEID,
    W_SZ,
    W_TBL,
    W_TC,
    W_TOP,
    W_TR,
    W_TYPE,
    W_U,
    W_VAL,
    W_W,
    WPS_TXBX,
    WPS_WSP,
)
from sharepoint2text.parsing.extractors.ms_modern.ooxml_shared import (
    OOXMLZipContext,
    extract_omml_formulas,
    get_element_text,
    get_image_content_type,
    get_image_pixel_dimensions,
)
from sharepoint2text.parsing.extractors.ms_modern.ooxml_text_processing import (
    collect_text_from_element,
    extract_text_with_formulas,
    get_first_attribute,
    half_points_to_points,
    parse_boolean_element,
    twips_to_inches,
)
from sharepoint2text.parsing.extractors.util.encryption import is_ooxml_encrypted

logger = logging.getLogger(__name__)

# Re-export NAMESPACES for any code that might import it from here
NAMESPACES = DOCX_NAMESPACES


# =============================================================================
# Text extraction helpers
# =============================================================================


def _extract_paragraph_content(paragraph: ET.Element, include_formulas: bool) -> str:
    """Extract text from a paragraph, including inline and display equations."""
    return extract_text_with_formulas(
        paragraph,
        include_formulas=include_formulas,
        omath_tag=M_OMATH,
        omath_para_tag=M_OMATHPARA,
        formula_formatter=omml_to_latex,
    )


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
    return texts  # pragma: no cover


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
        ]

        for path, attr in xml_files:
            setattr(self, attr, self.read_xml_root_if_exists(path))

        # Pre-load header and footer files
        self._relationships = self._parse_relationships()
        for rel_info in self._relationships.values():
            rel_type = rel_info.get("type", "")
            target = rel_info.get("target", "")
            if "header" in rel_type.lower() or "footer" in rel_type.lower():
                hf_path = "word/" + target
                root = self.read_xml_root_if_exists(hf_path)
                if root is not None:
                    self._header_footer_roots[hf_path] = root

    def _parse_relationships(self) -> dict[str, dict]:
        """Parse relationships from cached rels root."""
        relationships = {}
        for rel in self.read_relationships_if_exists("word/_rels/document.xml.rels"):
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
        """Get the document body element.

        Returns:
            Document body element.
        """
        if self._document_root is None:
            return None
        return self._document_root.find(W_BODY)

    @property
    def relationships(self) -> dict[str, dict]:
        """Get cached relationships.

        Returns:
            Relationship mapping for the main document.
        """
        if self._relationships is None:
            self._relationships = self._parse_relationships()
        return self._relationships

    @property
    def styles(self) -> dict[str, str]:
        """Get cached style map (style_id -> style_name).

        Returns:
            Style identifier to display-name mapping.
        """
        if self._styles is None:
            self._styles = {}
            if self._styles_root is not None:
                for style in self._styles_root.iter(W_STYLE):
                    style_id = style.get(W_STYLEID) or ""
                    name_elem = style.find(W_NAME)
                    style_name = name_elem.get(W_VAL) if name_elem is not None else ""
                    if style_id:
                        self._styles[style_id] = style_name or style_id
        return self._styles


# =============================================================================
# Shared body analysis
# =============================================================================


@dataclass
class _DocxBodyAnalysis:
    paragraph_elements: list[ET.Element]
    paragraphs: list[DocxParagraph]
    tables: list[list[list[str]]]
    table_anchor_paragraph_indices: list[int]
    hyperlinks: list[DocxHyperlink]
    formulas: list[DocxFormula]
    sections: list[DocxSection]
    full_text: str


def _build_paragraph(
    paragraph: ET.Element,
    style_map: dict[str, str],
) -> DocxParagraph:
    """Extract a top-level paragraph with formatting and run information."""
    ppr = paragraph.find(W_PPR)
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

    has_page_break = any(br.get(W_TYPE) == "page" for br in paragraph.iter(W_BR)) or (
        next(paragraph.iter(W_LAST_RENDERED_PAGE_BREAK), None) is not None
    )

    runs: list[DocxRun] = []
    paragraph_text_parts: list[str] = []
    for run in paragraph.iter(W_R):
        run_text = collect_text_from_element(run)
        if not run_text:
            continue
        paragraph_text_parts.append(run_text)

        bold, italic, underline, font_name, font_size, font_color = (
            _parse_run_properties(run.find(W_RPR))
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

    return DocxParagraph(
        text="".join(paragraph_text_parts),
        style=style_name,
        alignment=alignment,
        runs=runs,
        has_page_break=has_page_break,
    )


def _extract_table_data(table: ET.Element) -> list[list[str]]:
    """Extract table data while preserving the existing cell-text semantics."""
    table_data: list[list[str]] = []
    for tr in table.findall(W_TR):
        row_data: list[str] = []
        for tc in tr.findall(W_TC):
            cell_paragraphs = [collect_text_from_element(p) for p in tc.iter(W_P)]
            row_data.append("\n".join(cell_paragraphs))
        table_data.append(row_data)
    return table_data


def _extract_hyperlinks_from_element(
    element: ET.Element,
    rels: dict[str, dict],
) -> list[DocxHyperlink]:
    """Extract hyperlinks from a body subtree."""
    hyperlinks: list[DocxHyperlink] = []
    for hyperlink in element.iter(W_HYPERLINK):
        r_id = hyperlink.get(R_ID)
        if r_id and r_id in rels:
            rel_info = rels[r_id]
            if "hyperlink" in rel_info.get("type", "").lower():
                hyperlinks.append(
                    DocxHyperlink(
                        text=collect_text_from_element(hyperlink),
                        url=rel_info.get("target", ""),
                    )
                )
    return hyperlinks


def _extract_formulas_from_element(element: ET.Element) -> list[DocxFormula]:
    """Extract formulas from a body subtree."""
    return [
        DocxFormula(latex=latex, is_display=is_display)
        for latex, is_display in extract_omml_formulas(
            element,
            omath_para_tag=M_OMATHPARA,
            omath_tag=M_OMATH,
            converter=omml_to_latex,
        )
    ]


def _collect_section_properties(
    element: ET.Element,
    section_properties: list[ET.Element],
) -> None:
    """Collect section property elements from a body subtree."""
    for paragraph in element.iter(W_P):
        ppr = paragraph.find(W_PPR)
        if ppr is None:
            continue
        sect_pr = ppr.find(W_SECTPR)
        if sect_pr is not None:
            section_properties.append(sect_pr)


def _build_sections(section_properties: list[ET.Element]) -> list[DocxSection]:
    """Build section objects from collected section property elements."""
    sections: list[DocxSection] = []
    for sect_pr in section_properties:
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


def _analyze_document_body(ctx: _DocxContext) -> _DocxBodyAnalysis:
    """Walk the document body once and reuse the extracted structures downstream."""
    body = ctx.document_body
    if body is None:
        return _DocxBodyAnalysis(
            paragraph_elements=[],
            paragraphs=[],
            tables=[],
            table_anchor_paragraph_indices=[],
            hyperlinks=[],
            formulas=[],
            sections=[],
            full_text="",
        )

    style_map = ctx.styles
    rels = ctx.relationships

    paragraph_elements: list[ET.Element] = []
    paragraphs: list[DocxParagraph] = []
    tables: list[list[list[str]]] = []
    table_anchor_paragraph_indices: list[int] = []
    hyperlinks: list[DocxHyperlink] = []
    formulas: list[DocxFormula] = []
    full_text_parts: list[str] = []
    section_properties: list[ET.Element] = []

    current_paragraph_index = -1

    for child in body:
        if child.tag == W_P:
            paragraph_elements.append(child)
            paragraphs.append(_build_paragraph(child, style_map))
            current_paragraph_index += 1

            text = _extract_paragraph_content(child, include_formulas=True)
            if text.strip():
                full_text_parts.append(text)

            hyperlinks.extend(_extract_hyperlinks_from_element(child, rels))
            formulas.extend(_extract_formulas_from_element(child))
            _collect_section_properties(child, section_properties)
            continue

        if child.tag != W_TBL:
            continue

        anchor = max(0, current_paragraph_index)
        for table in child.iter(W_TBL):
            tables.append(_extract_table_data(table))
            table_anchor_paragraph_indices.append(anchor)

        full_text_parts.extend(_extract_table_text(child, include_formulas=True))
        hyperlinks.extend(_extract_hyperlinks_from_element(child, rels))
        formulas.extend(_extract_formulas_from_element(child))
        _collect_section_properties(child, section_properties)

    final_sect_pr = body.find(W_SECTPR)
    if final_sect_pr is not None:
        section_properties.append(final_sect_pr)

    return _DocxBodyAnalysis(
        paragraph_elements=paragraph_elements,
        paragraphs=paragraphs,
        tables=tables,
        table_anchor_paragraph_indices=table_anchor_paragraph_indices,
        hyperlinks=hyperlinks,
        formulas=formulas,
        sections=_build_sections(section_properties),
        full_text="\n".join(full_text_parts),
    )


# =============================================================================
# Extraction functions
# =============================================================================


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
        if text := get_element_text(root, tag):
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
        DocxNote(id=note.get(W_ID) or "", text=collect_text_from_element(note))
        for note in root.iter(note_tag)
        if (note.get(W_ID) or "") not in SKIP_NOTE_IDS
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
            text=collect_text_from_element(comment),
        )
        for comment in root.iter(W_COMMENT)
    ]


def _parse_twips_to_inches(value: str | None) -> float | None:
    """Convert twips string to inches, returning None on failure."""
    return twips_to_inches(value)


def _extract_sections_from_context(ctx: _DocxContext) -> list[DocxSection]:
    """Extract section properties (page layout) from cached document body."""
    body = ctx.document_body
    if body is None:
        return []

    sect_pr_elements: list[ET.Element] = []

    # Sections in paragraphs
    for p in body.iter(W_P):
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

        text = collect_text_from_element(root)
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
    bold = parse_boolean_element(rpr.find(W_B), W_VAL)

    # Italic
    italic = parse_boolean_element(rpr.find(W_I), W_VAL)

    # Underline
    underline = None
    u_elem = rpr.find(W_U)
    if u_elem is not None:
        u_val = u_elem.get(W_VAL)
        underline = bool(u_val and u_val != "none")

    # Font name (with fallback chain: ascii -> hAnsi -> cs)
    rfonts = rpr.find(W_RFONTS)
    font_name = get_first_attribute(rfonts, W_ASCII, W_HANSI, W_CS)

    # Font size (half-points to points)
    sz = rpr.find(W_SZ)
    font_size = half_points_to_points(sz.get(W_VAL) if sz is not None else None)

    # Font color
    color = rpr.find(W_COLOR)
    font_color = color.get(W_VAL) if color is not None else None

    return (bold, italic, underline, font_name, font_size, font_color)


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

        has_page_break = any(br.get(W_TYPE) == "page" for br in p.iter(W_BR)) or (
            next(p.iter(W_LAST_RENDERED_PAGE_BREAK), None) is not None
        )

        # Extract runs
        runs: list[DocxRun] = []
        paragraph_text_parts: list[str] = []
        for r in p.iter(W_R):
            run_text = collect_text_from_element(r)
            if not run_text:
                continue
            paragraph_text_parts.append(run_text)

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
                text="".join(paragraph_text_parts),
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
                row_data: list[str] = []
                for tc in tr.findall(W_TC):
                    cell_paragraphs = [
                        collect_text_from_element(p) for p in tc.iter(W_P)
                    ]
                    row_data.append("\n".join(cell_paragraphs))
                table_data.append(row_data)
            tables.append(table_data)
            table_anchor_paragraph_indices.append(anchor)

    return tables, table_anchor_paragraph_indices


def _extract_images_from_context(
    ctx: _DocxContext,
    paragraphs: list[ET.Element] | None = None,
) -> list[DocxImage]:
    """Extract images with captions and descriptions."""
    rels = ctx.relationships
    body = ctx.document_body

    image_metadata: dict[str, tuple[str, str]] = {}
    image_anchor_paragraph_indices: dict[str, set[int]] = {}

    if body is not None:
        if paragraphs is None:
            paragraphs = list(body.findall(W_P))

        for para_idx, para in enumerate(paragraphs):
            for drawing in para.iter(W_DRAWING):
                caption = ""
                description = ""

                pic_cNvPr = next(drawing.iter(PIC_CNVPR), None)
                if pic_cNvPr is not None:
                    if descr := pic_cNvPr.get("descr", ""):
                        description = descr
                    if name := pic_cNvPr.get("name", ""):
                        caption = name

                for wsp in drawing.iter(WPS_WSP):
                    txbx = wsp.find(WPS_TXBX)
                    if txbx is not None:
                        if text := collect_text_from_element(txbx):
                            caption = text
                            break

                # Check preceding/following paragraphs for caption
                preceding_caption = None
                if para_idx > 0:
                    prev_para = paragraphs[para_idx - 1]
                    prev_style = _get_paragraph_style(prev_para)
                    if _is_caption_style(prev_style) and _has_keep_next(prev_para):
                        if text := collect_text_from_element(prev_para):
                            preceding_caption = text

                following_caption = None
                if para_idx + 1 < len(paragraphs):
                    next_para = paragraphs[para_idx + 1]
                    if _is_caption_style(_get_paragraph_style(next_para)):
                        if text := collect_text_from_element(next_para):
                            following_caption = text

                if preceding_caption:
                    caption = preceding_caption
                elif following_caption:
                    caption = following_caption

                blip = next(drawing.iter(A_BLIP), None)
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

        # Handle both absolute paths (starting with /) and relative paths
        if target.startswith("/"):
            image_path = target[1:]  # Remove leading /
        else:
            image_path = "word/" + target
        try:
            img_data = ctx.read_bytes_if_exists(image_path)
            if img_data is None:
                continue

            image_counter += 1
            caption, description = image_metadata.get(rel_id, ("", ""))
            width, height = get_image_pixel_dimensions(img_data)

            images.append(
                DocxImage(
                    rel_id=rel_id,
                    filename=target.rsplit("/", 1)[-1],
                    content_type=get_image_content_type(
                        target, fallback_to_extension=True
                    ),
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
        except (KeyError, ValueError, OSError, UnicodeDecodeError) as e:
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

    for hyperlink in body.iter(W_HYPERLINK):
        r_id = hyperlink.get(R_ID)
        if r_id and r_id in rels:
            rel_info = rels[r_id]
            if "hyperlink" in rel_info.get("type", "").lower():
                hyperlinks.append(
                    DocxHyperlink(
                        text=collect_text_from_element(hyperlink),
                        url=rel_info.get("target", ""),
                    )
                )

    return hyperlinks


def _extract_formulas_from_context(ctx: _DocxContext) -> list[DocxFormula]:
    """Extract all mathematical formulas from the document as LaTeX."""
    body = ctx.document_body
    if body is None:
        return []

    return [
        DocxFormula(latex=latex, is_display=is_display)
        for latex, is_display in extract_omml_formulas(
            body,
            omath_para_tag=M_OMATHPARA,
            omath_tag=M_OMATH,
            converter=omml_to_latex,
        )
    ]


# =============================================================================
# Main entry point
# =============================================================================


def read_docx(
    file_like: io.BytesIO, path: str | None = None, *, ignore_images: bool = False
) -> Generator[DocxParserOutput, Any, None]:
    """
    Extract all relevant content from a Word .docx file.

    Uses a generator pattern for API consistency. DOCX files yield exactly one
    DocxParserOutput object containing paragraphs, tables, images, metadata, etc.

    Args:
        file_like: BytesIO object containing the DOCX file data.
        path: Optional path to the source file for metadata.
        ignore_images: If True, skip image extraction.
    """
    source_path = path or "<in-memory>"
    logger.info("Entering DOCX extraction: %s", source_path)
    try:
        file_like.seek(0)
        if is_ooxml_encrypted(file_like):
            raise ExtractionFileEncryptedError(
                "DOCX is encrypted or password-protected"
            )

        ctx = _DocxContext(file_like)
        try:
            metadata = _extract_metadata_from_context(ctx)
            body_analysis = _analyze_document_body(ctx)
            paragraphs = body_analysis.paragraphs
            tables = body_analysis.tables
            table_anchor_paragraph_indices = (
                body_analysis.table_anchor_paragraph_indices
            )
            headers, footers = _extract_header_footers_from_context(ctx)
            images = (
                []
                if ignore_images
                else _extract_images_from_context(
                    ctx, paragraphs=body_analysis.paragraph_elements
                )
            )
            hyperlinks = body_analysis.hyperlinks
            footnotes = _extract_footnotes_from_context(ctx)
            endnotes = _extract_endnotes_from_context(ctx)
            formulas = body_analysis.formulas
            comments = _extract_comments_from_context(ctx)
            sections = body_analysis.sections
            styles = list({para.style for para in paragraphs if para.style})
            full_text = body_analysis.full_text

            metadata.populate_from_path(path)

            logger.debug(
                "Extracted DOCX: %d paragraphs, %d tables, %d images",
                len(paragraphs),
                len(tables),
                len(images),
            )

            yield DocxParserOutput(
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
    except (KeyError, ET.ParseError, OSError, ValueError, UnicodeDecodeError) as exc:
        raise ExtractionFailedError("Failed to extract DOCX file", cause=exc) from exc
    finally:
        logger.info("Leaving DOCX extraction: %s", source_path)
