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
import zipfile
from typing import Any, Generator, List, Tuple
from xml.etree import ElementTree as ET

from sharepoint2text.extractors.data_types import (
    PPTXComment,
    PptxContent,
    PPTXFormula,
    PPTXImage,
    PptxMetadata,
    PPTXSlide,
)
from sharepoint2text.extractors.ms_modern.docx_extractor import _DocxFullTextExtractor

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

# Title placeholder types
TITLE_TYPES = {"title", "ctrTitle"}

# Body/content placeholder types
BODY_TYPES = {"body", "subTitle", "obj", "tbl"}

# Footer-related placeholder types
FOOTER_TYPES = {"ftr"}

# Placeholder types to skip (not useful for text extraction)
# Note: sldNum (slide number) is NOT skipped - it goes to other_textboxes
SKIP_TYPES = {"dt", "sldImg", "hdr"}


def _extract_metadata(file_like: io.BytesIO) -> PptxMetadata:
    """
    Extract presentation metadata from docProps/core.xml.

    Args:
        file_like: BytesIO containing the PPTX file.

    Returns:
        PptxMetadata object with title, author, dates, revision, etc.
    """
    logger.debug("Extracting metadata")
    file_like.seek(0)

    metadata = PptxMetadata()

    with zipfile.ZipFile(file_like, "r") as z:
        if "docProps/core.xml" not in z.namelist():
            return metadata

        with z.open("docProps/core.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

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
        # Remove 'Z' suffix for consistency with existing format
        metadata.created = created_elem.text.rstrip("Z")

    modified_elem = root.find(f"{DCTERMS_NS}modified")
    if modified_elem is not None and modified_elem.text:
        # Remove 'Z' suffix for consistency with existing format
        metadata.modified = modified_elem.text.rstrip("Z")

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


def _get_slide_order(file_like: io.BytesIO) -> List[str]:
    """
    Get the ordered list of slide paths from presentation.xml.

    The slide order is determined by the order of p:sldId elements
    in the p:sldIdLst element of presentation.xml.

    Args:
        file_like: BytesIO containing the PPTX file.

    Returns:
        List of slide XML file paths in presentation order.
    """
    file_like.seek(0)
    slide_paths = []

    with zipfile.ZipFile(file_like, "r") as z:
        if "ppt/presentation.xml" not in z.namelist():
            return slide_paths

        # First, read the relationships to map rId to slide paths
        rels_map = {}
        rels_path = "ppt/_rels/presentation.xml.rels"
        if rels_path in z.namelist():
            with z.open(rels_path) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for rel in root.findall(f".//{REL_NS}Relationship"):
                    rel_id = rel.get("Id")
                    target = rel.get("Target")
                    rel_type = rel.get("Type") or ""
                    if rel_id and target and "slide" in rel_type.lower():
                        # Convert relative path to full path
                        if target.startswith("slides/"):
                            full_path = f"ppt/{target}"
                        elif target.startswith("../"):
                            full_path = target.replace("../", "ppt/")
                        else:
                            full_path = f"ppt/{target}"
                        rels_map[rel_id] = full_path

        # Read presentation.xml to get slide order
        with z.open("ppt/presentation.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()

        # Find sldIdLst and iterate through sldId elements
        sld_id_lst = root.find(f".//{P_NS}sldIdLst")
        if sld_id_lst is not None:
            for sld_id in sld_id_lst.findall(f"{P_NS}sldId"):
                r_id = sld_id.get(f"{R_NS}id")
                if r_id and r_id in rels_map:
                    slide_paths.append(rels_map[r_id])

    return slide_paths


def _extract_slide_comments(
    file_like: io.BytesIO, slide_number: int
) -> List[PPTXComment]:
    """
    Extract comments for a specific slide by parsing the comments XML.

    Comments in PPTX files are stored in separate XML files, one per slide
    that has comments (ppt/comments/comment{n}.xml).

    Args:
        file_like: BytesIO containing the PPTX file.
        slide_number: 1-based slide number to extract comments for.

    Returns:
        List of PPTXComment objects with author, text, and date fields.
        Returns empty list if no comment file exists for the slide.

    Notes:
        - Author is stored as authorId (numeric), not the actual name
        - Date is in ISO format from the XML
    """
    comments = []
    file_like.seek(0)

    try:
        with zipfile.ZipFile(file_like, "r") as z:
            comment_file = f"ppt/comments/comment{slide_number}.xml"
            if comment_file not in z.namelist():
                return comments

            with z.open(comment_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

            for cm in root.findall(f".//{P_NS}cm"):
                author_id = cm.get("authorId", "")
                text_elem = cm.find(f"{P_NS}text")
                text = text_elem.text if text_elem is not None else ""
                dt = cm.get("dt", "")
                comments.append(
                    PPTXComment(
                        author=author_id,
                        text=text or "",
                        date=dt,
                    )
                )
    except Exception as e:
        logger.debug(f"Failed to extract comments for slide {slide_number}: {e}")

    return comments


def _get_slide_relationships(
    file_like: io.BytesIO, slide_path: str
) -> dict[str, dict[str, str]]:
    """
    Get relationships for a specific slide.

    Args:
        file_like: BytesIO containing the PPTX file.
        slide_path: Path to the slide XML file.

    Returns:
        Dictionary mapping relationship IDs to {type, target} dicts.
    """
    file_like.seek(0)
    relationships = {}

    # Convert slide path to relationships path
    # e.g., ppt/slides/slide1.xml -> ppt/slides/_rels/slide1.xml.rels
    slide_dir = "/".join(slide_path.rsplit("/", 1)[:-1])
    slide_name = slide_path.rsplit("/", 1)[-1]
    rels_path = f"{slide_dir}/_rels/{slide_name}.rels"

    with zipfile.ZipFile(file_like, "r") as z:
        if rels_path not in z.namelist():
            return relationships

        with z.open(rels_path) as f:
            tree = ET.parse(f)
            root = tree.getroot()

        for rel in root.findall(f".//{REL_NS}Relationship"):
            rel_id = rel.get("Id") or ""
            rel_type = rel.get("Type") or ""
            rel_target = rel.get("Target") or ""

            relationships[rel_id] = {"type": rel_type, "target": rel_target}

    return relationships


def _get_shape_position(shape_elem) -> Tuple[int, int]:
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
        sp_pr = shape_elem.find(f".//{P_NS}spPr") or shape_elem.find(f".//{A_NS}xfrm")
        if sp_pr is None:
            sp_pr = shape_elem

        xfrm = sp_pr.find(f"{A_NS}xfrm")
        if xfrm is None:
            xfrm = shape_elem.find(f".//{A_NS}xfrm")

        if xfrm is not None:
            off = xfrm.find(f"{A_NS}off")
            if off is not None:
                x = int(off.get("x", "0"))
                y = int(off.get("y", "0"))
                return (y, x)  # Sort by y (top) first, then x (left)

        # No explicit position - check if it's a placeholder and assign default
        nv_sp_pr = shape_elem.find(f"{P_NS}nvSpPr")
        if nv_sp_pr is not None:
            nv_pr = nv_sp_pr.find(f"{P_NS}nvPr")
            if nv_pr is not None:
                ph = nv_pr.find(f"{P_NS}ph")
                if ph is not None:
                    ph_type = ph.get("type", "")
                    ph_idx = ph.get("idx", "")

                    # Title placeholders go at the very top
                    if ph_type in TITLE_TYPES:
                        return (0, 0)

                    # Body placeholders go after title
                    if ph_type in BODY_TYPES or (not ph_type and ph_idx):
                        # Use idx to order multiple body placeholders
                        idx_num = int(ph_idx) if ph_idx.isdigit() else 0
                        return (1 + idx_num, 0)

                    # Footer-related placeholders go at the bottom
                    if ph_type in FOOTER_TYPES or ph_type == "sldNum":
                        return (999999998, 0)

        # Default for shapes without position info - place at bottom
        return (999999999, 999999999)
    except Exception:
        return (999999999, 999999999)


def _extract_text_from_paragraphs(elem) -> str:
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
    paragraphs = []
    for p in elem.findall(f".//{A_NS}p"):
        texts = []
        # Process all child elements in order to handle text and breaks
        for child in p:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag == "r":  # Run - contains text
                t = child.find(f"{A_NS}t")
                if t is not None and t.text:
                    texts.append(t.text)
            elif tag == "fld":  # Field - contains dynamic content (slide number, etc.)
                t = child.find(f"{A_NS}t")
                if t is not None and t.text:
                    texts.append(t.text)
            elif tag == "br":  # Line break - represented as vertical tab
                texts.append("\x0b")
            elif tag == "t":  # Direct text (less common)
                if child.text:
                    texts.append(child.text)
            # Skip pPr (paragraph properties), endParaRPr, etc.

        para_text = "".join(texts)
        paragraphs.append(para_text)
    return "\n".join(paragraphs)


def _extract_formulas_from_element(elem) -> List[Tuple[str, bool]]:
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
    formulas = []

    # Track oMath elements inside oMathPara to avoid duplicates
    omath_in_para = set()

    # First, find all oMathPara elements (display equations)
    for omath_para in elem.iter(f"{M_NS}oMathPara"):
        omath = omath_para.find(f"{M_NS}oMath")
        if omath is not None:
            omath_in_para.add(omath)
            latex = _DocxFullTextExtractor.omml_to_latex(omath)
            if latex.strip():
                formulas.append((latex, True))

    # Then find inline oMath elements (not in oMathPara)
    for omath in elem.iter(f"{M_NS}oMath"):
        if omath not in omath_in_para:
            latex = _DocxFullTextExtractor.omml_to_latex(omath)
            if latex.strip():
                formulas.append((latex, False))

    return formulas


def _process_slide(
    file_like: io.BytesIO, slide_path: str, slide_number: int
) -> PPTXSlide:
    """
    Process a single slide and extract all its content.

    Args:
        file_like: BytesIO containing the PPTX file.
        slide_path: Path to the slide XML file within the ZIP.
        slide_number: 1-based slide number.

    Returns:
        PPTXSlide object containing all extracted content.
    """
    logger.debug(f"Processing slide [{slide_number}]: {slide_path}")

    file_like.seek(0)

    slide_title = ""
    slide_footer = ""
    content_placeholders: List[str] = []
    other_textboxes: List[str] = []
    images: List[PPTXImage] = []
    formulas: List[PPTXFormula] = []

    # Collect all content items with their positions for ordering
    # Each item: (position, content_type, content_text)
    ordered_content: List[Tuple[Tuple[int, int], str, str]] = []

    # Get slide relationships for images
    slide_rels = _get_slide_relationships(file_like, slide_path)
    file_like.seek(0)

    with zipfile.ZipFile(file_like, "r") as z:
        if slide_path not in z.namelist():
            logger.warning(f"Slide not found: {slide_path}")
            return PPTXSlide(slide_number=slide_number)

        with z.open(slide_path) as f:
            tree = ET.parse(f)
            root = tree.getroot()

        # Find the shape tree
        sp_tree = root.find(f".//{P_NS}spTree")
        if sp_tree is None:
            return PPTXSlide(slide_number=slide_number)

        # Collect all shapes and pictures with their positions
        shape_elements = []

        # Regular shapes (p:sp)
        for sp in sp_tree.findall(f".//{P_NS}sp"):
            shape_elements.append(("sp", sp, _get_shape_position(sp)))

        # Pictures (p:pic)
        for pic in sp_tree.findall(f".//{P_NS}pic"):
            shape_elements.append(("pic", pic, _get_shape_position(pic)))

        # Sort by position (top to bottom, left to right)
        shape_elements.sort(key=lambda x: x[2])

        image_counter = 0

        for shape_type, elem, position in shape_elements:
            # ---------------------------
            # Picture extraction
            # ---------------------------
            if shape_type == "pic":
                try:
                    image_counter += 1

                    # Get image relationship ID
                    blip = elem.find(f".//{A_NS}blip")
                    if blip is None:
                        continue

                    r_embed = blip.get(f"{R_NS}embed")
                    if not r_embed or r_embed not in slide_rels:
                        continue

                    rel_info = slide_rels[r_embed]
                    target = rel_info.get("target", "")

                    # Build full image path
                    slide_dir = "/".join(slide_path.rsplit("/", 1)[:-1])
                    if target.startswith("../"):
                        image_path = f"{slide_dir}/{target}"
                        # Normalize path to resolve .. segments
                        parts = image_path.split("/")
                        normalized = []
                        for part in parts:
                            if part == "..":
                                if normalized:
                                    normalized.pop()
                            elif part:  # Skip empty parts
                                normalized.append(part)
                        image_path = "/".join(normalized)
                    else:
                        image_path = f"{slide_dir}/{target}"

                    # Extract alt text / description
                    caption = ""
                    cNvPr = elem.find(f".//{P_NS}cNvPr")
                    if cNvPr is not None:
                        descr = cNvPr.get("descr", "")
                        if descr:
                            caption = descr

                    # Read image data
                    if image_path in z.namelist():
                        with z.open(image_path) as img_file:
                            blob = img_file.read()

                        # Determine content type and filename from extension
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

                        # Generate generic filename based on extension
                        # (matches python-pptx behavior)
                        generic_filename = f"image.{ext}"

                        images.append(
                            PPTXImage(
                                image_index=image_counter,
                                filename=generic_filename,
                                content_type=content_type,
                                size_bytes=len(blob),
                                blob=blob,
                                caption=caption,
                            )
                        )

                        # Add caption to ordered content if present
                        if caption:
                            ordered_content.append(
                                (position, "image_caption", f"[Image: {caption}]")
                            )
                except Exception as e:
                    logger.error(e)
                    logger.exception(f"Failed to extract image on slide {slide_number}")
                continue

            # ---------------------------
            # Shape (text) extraction
            # ---------------------------
            # Get placeholder info
            nv_sp_pr = elem.find(f"{P_NS}nvSpPr")
            if nv_sp_pr is None:
                continue

            nv_pr = nv_sp_pr.find(f"{P_NS}nvPr")
            ph = nv_pr.find(f"{P_NS}ph") if nv_pr is not None else None

            # Extract formulas from shape
            shape_formulas = _extract_formulas_from_element(elem)
            for latex, is_display in shape_formulas:
                formula = PPTXFormula(latex=latex, is_display=is_display)
                formulas.append(formula)
                if is_display:
                    ordered_content.append((position, "formula", f"$${latex}$$"))
                else:
                    ordered_content.append((position, "formula", f"${latex}$"))

            # Extract text
            tx_body = elem.find(f"{P_NS}txBody")
            if tx_body is None:
                continue

            text = _extract_text_from_paragraphs(tx_body).strip()
            if not text:
                continue

            # Determine placeholder type
            if ph is not None:
                ph_type = ph.get("type", "")
                ph_idx = ph.get("idx", "")

                if ph_type in TITLE_TYPES:
                    slide_title = text
                    ordered_content.append((position, "title", text))
                elif ph_type in FOOTER_TYPES:
                    slide_footer = text
                    # Footer is typically not included in main text
                elif ph_type in SKIP_TYPES:
                    # Skip date, slide number, etc.
                    pass
                elif ph_type in BODY_TYPES or (not ph_type and ph_idx):
                    # Body placeholder or indexed placeholder without type
                    content_placeholders.append(text)
                    ordered_content.append((position, "content", text))
                else:
                    other_textboxes.append(text)
                    ordered_content.append((position, "other", text))
            else:
                # Non-placeholder shape
                other_textboxes.append(text)
                ordered_content.append((position, "other", text))

    # ---------------------------
    # Comment extraction
    # ---------------------------
    comments = _extract_slide_comments(file_like, slide_number)
    # Add comments at the end of the slide content
    for comment in comments:
        ordered_content.append(
            (
                (999999, 999999),
                "comment",
                f"[Comment: {comment.author}@{comment.date}: {comment.text}]",
            )
        )

    # Build slide text from ordered content
    # Sort by position (already sorted but ensure stability)
    ordered_content.sort(key=lambda x: x[0])
    slide_text_parts = [item[2] for item in ordered_content]
    slide_text = "\n".join(slide_text_parts)

    # Build base text (without formulas, comments, image captions)
    base_content_types = {"title", "content", "other"}
    base_text_parts = [
        item[2] for item in ordered_content if item[1] in base_content_types
    ]
    base_text = "\n".join(base_text_parts)

    return PPTXSlide(
        slide_number=slide_number,
        title=slide_title,
        footer=slide_footer,
        content_placeholders=content_placeholders,
        other_textboxes=other_textboxes,
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
                - images: List of PPTXImage with binary data
                - formulas: List of PPTXFormula as LaTeX
                - comments: List of PPTXComment
                - text: Complete slide text with formulas and comments
                - base_text: Text without formulas/comments/captions

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
        - Images are loaded into memory as binary blobs
        - Large presentations with many images may use significant memory
        - Comments require a separate ZIP file read per slide
    """
    logger.debug("Reading pptx")
    file_like.seek(0)

    # Extract metadata
    metadata = _extract_metadata(file_like)

    # Get slide order from presentation.xml
    slide_paths = _get_slide_order(file_like)

    # Process each slide
    slides_result: List[PPTXSlide] = []
    for slide_index, slide_path in enumerate(slide_paths, start=1):
        slide = _process_slide(file_like, slide_path, slide_index)
        slides_result.append(slide)

    metadata.populate_from_path(path)

    total_images = sum(len(slide.images) for slide in slides_result)
    logger.info(
        "Extracted PPTX: %d slides, %d images",
        len(slides_result),
        total_images,
    )

    yield PptxContent(metadata=metadata, slides=slides_result)
