"""
DOCX Document Extractor

Extracts text content, metadata, and structure from Microsoft Word .docx files
(Office Open XML format, Word 2007+).

Uses direct XML parsing of the docx ZIP archive structure for all content
extraction, without requiring the python-docx library.
"""

import io
import logging
import re
from dataclasses import dataclass
from typing import Any, Generator, cast

from sharepoint2text.parsing import _defused_xml as ET
from sharepoint2text.parsing.exceptions import (
    ExtractionError,
    ExtractionFailedError,
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors._model import source_metadata
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
    W_RFONTS,
    W_RIGHT,
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
from sharepoint2text.parsing.models import (
    Annotation,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    JsonValue,
    Table,
)

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
class _ParagraphData:
    """Store transient paragraph data used to build canonical content units."""

    text: str
    style: str | None
    has_page_break: bool


@dataclass
class _DocxBodyAnalysis:
    """Store the results of the single document-body traversal."""

    paragraph_elements: list[ET.Element]
    paragraphs: list[_ParagraphData]
    tables: list[Table]
    table_anchor_paragraph_indices: list[int]
    hyperlinks: list[Annotation]
    formulas: list[Annotation]
    sections: list[dict[str, JsonValue]]
    full_text: str


def _build_paragraph(
    paragraph: ET.Element,
    style_map: dict[str, str],
) -> _ParagraphData:
    """Extract transient text and structural data from a top-level paragraph."""
    ppr = paragraph.find(W_PPR)
    style_id = None
    if ppr is not None:
        style_elem = ppr.find(W_PSTYLE)
        if style_elem is not None:
            style_id = style_elem.get(W_VAL)

    style_name = style_map.get(style_id, style_id) if style_id else None

    has_page_break = any(br.get(W_TYPE) == "page" for br in paragraph.iter(W_BR)) or (
        next(paragraph.iter(W_LAST_RENDERED_PAGE_BREAK), None) is not None
    )

    return _ParagraphData(
        text=collect_text_from_element(paragraph),
        style=style_name,
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
) -> list[Annotation]:
    """Extract hyperlinks from a body subtree."""
    hyperlinks: list[Annotation] = []
    for hyperlink in element.iter(W_HYPERLINK):
        r_id = hyperlink.get(R_ID)
        if r_id and r_id in rels:
            rel_info = rels[r_id]
            if "hyperlink" in rel_info.get("type", "").lower():
                hyperlinks.append(
                    Annotation(
                        kind="hyperlink",
                        text=collect_text_from_element(hyperlink),
                        target=rel_info.get("target", ""),
                    )
                )
    return hyperlinks


def _extract_formulas_from_element(element: ET.Element) -> list[Annotation]:
    """Extract formulas from a body subtree."""
    return [
        Annotation(
            kind="formula",
            text=latex,
            properties={"docx.is_display": is_display},
        )
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


def _build_sections(
    section_properties: list[ET.Element],
) -> list[dict[str, JsonValue]]:
    """Convert Word section layout elements to JSON-compatible properties."""
    sections: list[dict[str, JsonValue]] = []
    for section_element in section_properties:
        section: dict[str, JsonValue] = {}
        page_size = section_element.find(W_PGSZ)
        if page_size is not None:
            section["page_width_inches"] = _parse_twips_to_inches(page_size.get(W_W))
            section["page_height_inches"] = _parse_twips_to_inches(page_size.get(W_H))
            orientation = page_size.get(W_ORIENT)
            section["orientation"] = (
                orientation if orientation and orientation != "portrait" else None
            )
        margins = section_element.find(W_PGMAR)
        if margins is not None:
            for name, tag in (
                ("left_margin_inches", W_LEFT),
                ("right_margin_inches", W_RIGHT),
                ("top_margin_inches", W_TOP),
                ("bottom_margin_inches", W_BOTTOM),
            ):
                section[name] = _parse_twips_to_inches(margins.get(tag))
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
    paragraphs: list[_ParagraphData] = []
    tables: list[Table] = []
    table_anchor_paragraph_indices: list[int] = []
    hyperlinks: list[Annotation] = []
    formulas: list[Annotation] = []
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
            tables.append(Table(rows=cast(Any, _extract_table_data(table))))
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


def _extract_metadata_from_context(ctx: _DocxContext) -> DocumentMetadata:
    """Extract document metadata from cached core.xml root."""
    metadata = DocumentMetadata()
    root = ctx._core_root
    if root is None:
        return metadata

    # Metadata field mappings: (tag, attribute)
    field_mappings = [
        (_DC_TITLE, "title"),
        (_DC_CREATOR, "author"),
        (_DC_SUBJECT, "subject"),
        (_DCTERMS_CREATED, "created"),
        (_DCTERMS_MODIFIED, "modified"),
    ]

    for tag, attr in field_mappings:
        if text := get_element_text(root, tag):
            setattr(metadata, attr, text)

    if keywords := get_element_text(root, _CP_KEYWORDS):
        metadata.keywords = [
            value.strip() for value in keywords.split(",") if value.strip()
        ]

    for tag, name in [
        (_CP_CATEGORY, "category"),
        (_DC_DESCRIPTION, "comments"),
        (_CP_LASTMODIFIEDBY, "last_modified_by"),
    ]:
        if value := get_element_text(root, tag):
            metadata.properties[f"docx.{name}"] = value

    revision_elem = root.find(_CP_REVISION)
    if revision_elem is not None and revision_elem.text:
        try:
            metadata.properties["docx.revision"] = int(revision_elem.text)
        except ValueError:
            pass

    return metadata


def _extract_notes_from_root(
    root: ET.Element | None, note_tag: str, kind: str
) -> list[Annotation]:
    """Extract notes (footnotes or endnotes) from an XML root element."""
    if root is None:
        return []

    return [
        Annotation(
            kind=kind,
            text=collect_text_from_element(note),
            properties={"docx.id": note.get(W_ID) or ""},
        )
        for note in root.iter(note_tag)
        if (note.get(W_ID) or "") not in SKIP_NOTE_IDS
    ]


def _extract_footnotes_from_context(ctx: _DocxContext) -> list[Annotation]:
    """Extract footnotes from cached footnotes.xml root."""
    return _extract_notes_from_root(ctx._footnotes_root, W_FOOTNOTE, "footnote")


def _extract_endnotes_from_context(ctx: _DocxContext) -> list[Annotation]:
    """Extract endnotes from cached endnotes.xml root."""
    return _extract_notes_from_root(ctx._endnotes_root, W_ENDNOTE, "endnote")


def _extract_comments_from_context(ctx: _DocxContext) -> list[Annotation]:
    """Extract comments from cached comments.xml root."""
    root = ctx._comments_root
    if root is None:
        return []

    return [
        Annotation(
            kind="comment",
            author=comment.get(W_AUTHOR) or "",
            text=collect_text_from_element(comment),
            properties={
                "docx.id": comment.get(W_ID) or "",
                "docx.date": comment.get(W_DATE) or "",
            },
        )
        for comment in root.iter(W_COMMENT)
    ]


def _parse_twips_to_inches(value: str | None) -> float | None:
    """Convert twips string to inches, returning None on failure."""
    return twips_to_inches(value)


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
) -> tuple[list[Annotation], list[Annotation]]:
    """Extract headers and footers from cached header/footer XML roots."""
    headers: list[Annotation] = []
    footers: list[Annotation] = []

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

        hf_obj = Annotation(
            kind="header" if is_header else "footer",
            text=text,
            properties={"docx.type": _determine_hf_type(hf_path, rel_type)},
        )
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


def _extract_images_from_context(
    ctx: _DocxContext,
    paragraphs: list[ET.Element] | None = None,
) -> list[tuple[ImageAsset, list[int]]]:
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
    images: list[tuple[ImageAsset, list[int]]] = []
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
                (
                    ImageAsset(
                        number=image_counter,
                        filename=target.rsplit("/", 1)[-1],
                        media_type=get_image_content_type(
                            target, fallback_to_extension=True
                        ),
                        data=img_data,
                        width=width,
                        height=height,
                        caption=caption,
                        description=description,
                        properties={
                            "docx.relationship_id": rel_id,
                            "docx.size_bytes": len(img_data),
                        },
                    ),
                    sorted(image_anchor_paragraph_indices.get(rel_id, set())),
                )
            )
        except (KeyError, ValueError, OSError, UnicodeDecodeError) as e:
            logger.debug("Failed to extract DOCX image %s: %s", rel_id, e)
            image_counter += 1
            images.append(
                (
                    ImageAsset(
                        number=image_counter,
                        properties={
                            "docx.relationship_id": rel_id,
                            "docx.error": str(e),
                        },
                    ),
                    [],
                )
            )

    return images


def _heading_level(style: str | None) -> int | None:
    """Return the structural heading level encoded by a Word style name."""
    if not style:
        return None
    if re.match(r"^(title|titel)\b", style.strip(), flags=re.IGNORECASE):
        return 0
    match = re.match(
        r"^(heading|überschrift)\s*(\d+)?\b", style.strip(), flags=re.IGNORECASE
    )
    if not match:
        return None
    return int(match.group(2) or "1")


def _build_content_units(
    analysis: _DocxBodyAnalysis,
    images_with_anchors: list[tuple[ImageAsset, list[int]]],
    title: str | None,
) -> list[ContentUnit]:
    """Build canonical units by grouping paragraphs under Word headings."""
    images_by_paragraph: dict[int, list[ImageAsset]] = {}
    for image, anchors in images_with_anchors:
        for anchor in anchors:
            images_by_paragraph.setdefault(anchor, []).append(image)
    tables_by_paragraph: dict[int, list[Table]] = {}
    anchors = analysis.table_anchor_paragraph_indices
    if len(anchors) != len(analysis.tables):
        anchors = [0] * len(analysis.tables)
    for table, anchor in zip(analysis.tables, anchors):
        tables_by_paragraph.setdefault(anchor, []).append(table)

    heading_indices = [
        index
        for index, paragraph in enumerate(analysis.paragraphs)
        if _heading_level(paragraph.style) is not None
    ]
    if not heading_indices:
        return [
            ContentUnit(
                number=1,
                kind="document",
                text=analysis.full_text,
                title=title,
                images=[image for image, _ in images_with_anchors],
                tables=list(analysis.tables),
            )
        ]

    heading_index_set = set(heading_indices)
    next_heading_for_index: list[int | None] = [None] * len(analysis.paragraphs)
    next_heading: int | None = None
    for index in range(len(analysis.paragraphs) - 1, -1, -1):
        next_heading_for_index[index] = next_heading
        if index in heading_index_set:
            next_heading = index

    heading_has_payload: dict[int, bool] = {}
    for position, heading_index in enumerate(heading_indices):
        end_index = (
            heading_indices[position + 1] - 1
            if position + 1 < len(heading_indices)
            else len(analysis.paragraphs) - 1
        )
        heading_has_payload[heading_index] = any(
            analysis.paragraphs[index].text.strip()
            or images_by_paragraph.get(index)
            or tables_by_paragraph.get(index)
            for index in range(heading_index + 1, end_index + 1)
        )

    units: list[ContentUnit] = []
    heading_stack: list[tuple[int, str]] = []
    current_path: list[str] = []
    current_level: int | None = None
    current_lines: list[str] = []
    current_start: int | None = None
    current_has_payload = False

    def flush_current(end_index: int, next_heading_level: int | None = None) -> None:
        """Append the current heading unit when its structure should be retained."""
        if not current_path or current_start is None:
            return
        unit_images: list[ImageAsset] = []
        unit_tables: list[Table] = []
        for paragraph_index in range(current_start, end_index + 1):
            unit_images.extend(images_by_paragraph.get(paragraph_index, []))
            unit_tables.extend(tables_by_paragraph.get(paragraph_index, []))
        text = "\n".join(line for line in current_lines if line.strip()).strip()
        if (
            not text
            and not unit_images
            and not unit_tables
            and next_heading_level is not None
            and current_level is not None
            and next_heading_level > current_level
        ):
            return
        units.append(
            ContentUnit(
                number=len(units) + 1,
                kind="section",
                text=text,
                title=current_path[-1],
                heading_path=list(current_path),
                images=unit_images,
                tables=unit_tables,
                properties={"docx.heading_level": current_level},
            )
        )

    for paragraph_index, paragraph in enumerate(analysis.paragraphs):
        level = _heading_level(paragraph.style)
        if level is not None:
            flush_current(paragraph_index - 1, level)
            while heading_stack and heading_stack[-1][0] >= level:
                heading_stack.pop()
            heading_stack.append((level, paragraph.text.strip()))
            current_path = [text for _, text in heading_stack if text]
            current_level = level
            current_lines = []
            current_start = paragraph_index
            current_has_payload = bool(
                images_by_paragraph.get(paragraph_index)
                or tables_by_paragraph.get(paragraph_index)
            )
            continue

        if images_by_paragraph.get(paragraph_index) or tables_by_paragraph.get(
            paragraph_index
        ):
            current_has_payload = True
        if current_path and not current_has_payload and paragraph.has_page_break:
            next_heading_index = next_heading_for_index[paragraph_index]
            if next_heading_index is not None and heading_has_payload.get(
                next_heading_index, False
            ):
                flush_current(paragraph_index)
                current_start = paragraph_index + 1
                current_lines = []
                current_has_payload = False
                continue
        if paragraph.text.strip():
            current_lines.append(paragraph.text.strip())
            current_has_payload = True

    if analysis.paragraphs:
        flush_current(len(analysis.paragraphs) - 1)
    return units


# =============================================================================
# Main entry point
# =============================================================================


def read_docx(
    file_like: io.BytesIO,
    path: str | None = None,
    *,
    ignore_images: bool = False,
    extract_annotations: bool = False,
) -> Generator[ExtractedDocument, Any, None]:
    """
    Extract all relevant content from a Word .docx file.

    Uses a generator pattern for API consistency. DOCX files yield exactly one
    canonical extraction result containing units, tables, images, and metadata.

    Args:
        file_like: BytesIO object containing the DOCX file data.
        path: Optional path to the source file for metadata.
        ignore_images: If True, skip image extraction.
        extract_annotations: If True, include comments in full_text and unit.text.
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
            body_analysis = _analyze_document_body(ctx)
            paragraphs = body_analysis.paragraphs
            tables = body_analysis.tables
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
            logger.debug(
                "Extracted DOCX: paragraphs=%d, tables=%d, images=%d",
                len(paragraphs),
                len(tables),
                len(images),
            )

            annotations = [
                *headers,
                *footers,
                *hyperlinks,
                *footnotes,
                *endnotes,
                *comments,
                *formulas,
            ]
            units = _build_content_units(body_analysis, images, metadata.title)

            # When extract_annotations=True, append comments to the last unit's text
            # and let full_text be computed from unit texts for consistency
            if extract_annotations and comments:
                comment_lines = [
                    f"[Comment: {c.author}@{c.properties.get('docx.date', '')}: {c.text}]"
                    for c in comments
                ]
                if units:
                    last_unit = units[-1]
                    if last_unit.text:
                        units[-1] = ContentUnit(
                            number=last_unit.number,
                            kind=last_unit.kind,
                            text=last_unit.text + "\n" + "\n".join(comment_lines),
                            title=last_unit.title,
                            heading_path=list(last_unit.heading_path),
                            images=list(last_unit.images),
                            tables=list(last_unit.tables),
                            annotations=list(last_unit.annotations),
                            properties=dict(last_unit.properties),
                        )
                    else:
                        units[-1] = ContentUnit(
                            number=last_unit.number,
                            kind=last_unit.kind,
                            text="\n".join(comment_lines),
                            title=last_unit.title,
                            heading_path=list(last_unit.heading_path),
                            images=list(last_unit.images),
                            tables=list(last_unit.tables),
                            annotations=list(last_unit.annotations),
                            properties=dict(last_unit.properties),
                        )

            owned_image_ids = {id(image) for unit in units for image in unit.images}
            owned_table_ids = {id(table) for unit in units for table in unit.tables}

            # Assign unowned images/tables and annotations to the first unit
            unowned_images = [
                image for image, _ in images if id(image) not in owned_image_ids
            ]
            unowned_tables = [
                table for table in tables if id(table) not in owned_table_ids
            ]
            if units:
                units[0].images.extend(unowned_images)
                units[0].tables.extend(unowned_tables)
                units[0].annotations.extend(annotations)

            # When extract_annotations=True, don't set document.full_text so
            # full_text is computed from unit texts, ensuring consistency
            properties: dict[str, JsonValue] = {
                "docx.paragraph_count": len(paragraphs),
                "docx.sections": cast(JsonValue, body_analysis.sections),
            }
            if not extract_annotations:
                properties["document.full_text"] = body_analysis.full_text

            yield ExtractedDocument(
                format="docx",
                source=source_metadata(path),
                metadata=metadata,
                units=units,
                properties=properties,
            )
        finally:
            ctx.close()
    except ExtractionError:
        raise
    except (KeyError, ET.ParseError, OSError, ValueError, UnicodeDecodeError) as exc:
        raise ExtractionFailedError("Failed to extract DOCX file", cause=exc) from exc
