"""
DOCX content extractor using python-docx library.
"""

import datetime
import io
import logging
import zipfile
from typing import Any, Generator
from xml.etree import ElementTree as ET

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from sharepoint2text.extractors.data_types import (
    DocxComment,
    DocxContent,
    DocxHeaderFooter,
    DocxHyperlink,
    DocxImage,
    DocxMetadata,
    DocxNote,
    DocxParagraph,
    DocxRun,
    DocxSection,
)

logger = logging.getLogger(__name__)


def _extract_full_text(file_like: io.BytesIO) -> str:
    """Combines the full text of the docx file into a single text.
    Paragraphs and tables are kept in the order of occurrence."""
    logger.debug("Extracting document full text")
    file_like.seek(0)
    doc = Document(file_like)
    all_text = []

    # Iterate through body elements in document order
    for element in doc.element.body:
        tag = element.tag.split("}")[-1]  # Strip namespace

        if tag == "p":  # Paragraph
            para = Paragraph(element, doc)
            if para.text.strip():
                all_text.append(para.text)

        elif tag == "tbl":  # Table
            table = Table(element, doc)
            for row in table.rows:
                for cell in row.cells:
                    text = " ".join(p.text for p in cell.paragraphs if p.text.strip())
                    if text:
                        all_text.append(text)

    return "\n".join(all_text)


def _extract_footnotes(file_like: io.BytesIO) -> list[DocxNote]:
    logger.debug("Extracting footnotes")
    footnotes = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    with zipfile.ZipFile(file_like, "r") as z:
        if "word/footnotes.xml" not in z.namelist():
            return footnotes

        with z.open("word/footnotes.xml") as f:
            tree = ET.parse(f)

        for fn in tree.findall(".//w:footnote", ns):
            fn_id = fn.get(f"{w_ns}id") or ""
            if fn_id not in ["-1", "0"]:  # Skip separator and continuation footnotes
                footnotes.append(
                    DocxNote(
                        id=fn_id,
                        text="".join(t.text or "" for t in fn.findall(".//w:t", ns)),
                    )
                )

    return footnotes


def _extract_comments(file_like: io.BytesIO) -> list[DocxComment]:
    logger.debug("Extracting comments")
    file_like.seek(0)
    comments = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    with zipfile.ZipFile(file_like, "r") as z:
        if "word/comments.xml" not in z.namelist():
            return comments

        with z.open("word/comments.xml") as f:
            tree = ET.parse(f)

        for comment in tree.findall(".//w:comment", ns):
            comments.append(
                DocxComment(
                    id=comment.get(f"{w_ns}id") or "",
                    author=comment.get(f"{w_ns}author") or "",
                    date=comment.get(f"{w_ns}date") or "",
                    text="".join(t.text or "" for t in comment.findall(".//w:t", ns)),
                )
            )

    return comments


def _extract_sections(doc: DocumentObject) -> list[DocxSection]:
    logger.debug("Extracting sections")
    sections = []
    for section in doc.sections:
        sections.append(
            DocxSection(
                page_width_inches=(
                    section.page_width.inches if section.page_width else None
                ),
                page_height_inches=(
                    section.page_height.inches if section.page_height else None
                ),
                left_margin_inches=(
                    section.left_margin.inches if section.left_margin else None
                ),
                right_margin_inches=(
                    section.right_margin.inches if section.right_margin else None
                ),
                top_margin_inches=(
                    section.top_margin.inches if section.top_margin else None
                ),
                bottom_margin_inches=(
                    section.bottom_margin.inches if section.bottom_margin else None
                ),
                orientation=(str(section.orientation) if section.orientation else None),
            )
        )
    return sections


def _extract_header_footers(
    doc: DocumentObject,
) -> tuple[list[DocxHeaderFooter], list[DocxHeaderFooter]]:
    logger.debug("Extracting header/footer")
    headers = []
    footers = []
    for section in doc.sections:
        # Default
        if section.header and section.header.paragraphs:
            text = "\n".join(p.text for p in section.header.paragraphs)
            if text.strip():
                headers.append(DocxHeaderFooter(type="default", text=text))
        if section.footer and section.footer.paragraphs:
            text = "\n".join(p.text for p in section.footer.paragraphs)
            if text.strip():
                footers.append(DocxHeaderFooter(type="default", text=text))

        # First page
        if section.first_page_header and section.first_page_header.paragraphs:
            text = "\n".join(p.text for p in section.first_page_header.paragraphs)
            if text.strip():
                headers.append(DocxHeaderFooter(type="first_page", text=text))
        if section.first_page_footer and section.first_page_footer.paragraphs:
            text = "\n".join(p.text for p in section.first_page_footer.paragraphs)
            if text.strip():
                footers.append(DocxHeaderFooter(type="first_page", text=text))

        # Even page
        if section.even_page_header and section.even_page_header.paragraphs:
            text = "\n".join(p.text for p in section.even_page_header.paragraphs)
            if text.strip():
                headers.append(DocxHeaderFooter(type="even_page", text=text))
        if section.even_page_footer and section.even_page_footer.paragraphs:
            text = "\n".join(p.text for p in section.even_page_footer.paragraphs)
            if text.strip():
                footers.append(DocxHeaderFooter(type="even_page", text=text))
    return headers, footers


def _extract_paragraphs(doc: DocumentObject) -> list[DocxParagraph]:
    logger.debug("Extracting paragraphs")
    paragraphs = []
    for para in doc.paragraphs:
        runs = []
        for run in para.runs:
            runs.append(
                DocxRun(
                    text=run.text,
                    bold=run.bold,
                    italic=run.italic,
                    underline=run.underline,
                    font_name=run.font.name,
                    font_size=run.font.size.pt if run.font.size else None,
                    font_color=(
                        str(run.font.color.rgb)
                        if run.font.color and run.font.color.rgb
                        else None
                    ),
                )
            )
        paragraphs.append(
            DocxParagraph(
                text=para.text,
                style=para.style.name if para.style else None,
                alignment=str(para.alignment) if para.alignment else None,
                runs=runs,
            )
        )
    return paragraphs


def _extract_tables(doc: DocumentObject):
    logger.debug("Extracting tables")
    tables = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                cell_text = "\n".join(p.text for p in cell.paragraphs)
                row_data.append(cell_text)
            table_data.append(row_data)
        tables.append(table_data)
    return tables


def _extract_endnotes(file_like: io.BytesIO) -> list[DocxNote]:
    """Extracts endnotes which are like footnotes just that they are
    either forced to appear at the end of a section or at the end of a document"""
    logger.debug("Extracting endnotes")
    file_like.seek(0)
    endnotes = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    w_ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    with zipfile.ZipFile(file_like, "r") as z:
        if "word/endnotes.xml" not in z.namelist():
            return endnotes

        with z.open("word/endnotes.xml") as f:
            tree = ET.parse(f)

        for en in tree.findall(".//w:endnote", ns):
            en_id = en.get(f"{w_ns}id") or ""
            if en_id not in ["-1", "0"]:  # Skip separator and continuation endnotes
                endnotes.append(
                    DocxNote(
                        id=en_id,
                        text="".join(t.text or "" for t in en.findall(".//w:t", ns)),
                    )
                )

    return endnotes


def read_docx(
    file_like: io.BytesIO, path: str | None = None
) -> Generator[DocxContent, Any, None]:
    """
    Extract all relevant content from a DOCX file.

    Args:
        file_like: A BytesIO object containing the DOCX file data.
        path: Optional file path to populate file metadata fields.

    Yields:
        MicrosoftDocxContent dataclass with all extracted content.
    """
    file_like.seek(0)
    doc = Document(file_like)

    # === Core Properties (Metadata) ===
    props = doc.core_properties
    metadata = DocxMetadata(
        title=props.title or "",
        author=props.author or "",
        subject=props.subject or "",
        keywords=props.keywords or "",
        category=props.category or "",
        comments=props.comments or "",
        created=(
            props.created.isoformat()
            if isinstance(props.created, datetime.datetime)
            else ""
        ),
        modified=(
            props.modified.isoformat()
            if isinstance(props.modified, datetime.datetime)
            else ""
        ),
        last_modified_by=props.last_modified_by or "",
        revision=props.revision,
    )

    # === Paragraphs ===
    paragraphs = _extract_paragraphs(doc=doc)

    # === Tables ===
    tables = _extract_tables(doc=doc)

    # === Headers and Footers ===
    headers, footers = _extract_header_footers(doc=doc)

    # === Images ===
    images = []
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            try:
                image_part = rel.target_part
                images.append(
                    DocxImage(
                        rel_id=rel_id,
                        filename=image_part.partname.split("/")[-1],
                        content_type=image_part.content_type,
                        data=io.BytesIO(image_part.blob),
                        size_bytes=len(image_part.blob),
                    )
                )
            except Exception as e:
                logger.debug(f"Image extraction failed for rel_id {rel_id} - {e}")
                images.append(DocxImage(rel_id=rel_id, error=str(e)))

    # === Hyperlinks ===
    hyperlinks = []
    rels = doc.part.rels
    for para in doc.paragraphs:
        for hyperlink in para._element.findall(
            ".//w:hyperlink",
            {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
        ):
            r_id = hyperlink.get(qn("r:id"))
            if r_id and r_id in rels and "hyperlink" in rels[r_id].reltype:
                text = "".join(
                    t.text or ""
                    for t in hyperlink.findall(
                        ".//w:t",
                        {
                            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        },
                    )
                )
                hyperlinks.append(DocxHyperlink(text=text, url=rels[r_id].target_ref))

    # === Footnotes ===
    footnotes = _extract_footnotes(file_like=file_like)

    # === Endnotes ===
    endnotes = _extract_endnotes(file_like=file_like)
    # === Comments ===
    comments = _extract_comments(file_like=file_like)

    # === Sections (page layout) ===
    sections = _extract_sections(doc=doc)

    # === Styles used ===
    styles_set = set()
    for para in doc.paragraphs:
        if para.style:
            styles_set.add(para.style.name)
    styles = list(styles_set)

    # === Full text (convenience) ===
    full_text = _extract_full_text(file_like=file_like)

    metadata.populate_from_path(path)

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
        full_text=full_text,
    )
