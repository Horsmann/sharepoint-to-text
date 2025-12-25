"""
DOCX content extractor using python-docx library.
"""

import datetime
import io
import logging
import typing
from dataclasses import dataclass, field
from typing import List, Optional

from docx import Document
from docx.oxml.ns import qn

from sharepoint2text.extractors.abstract_extractor import ExtractionInterface

logging.basicConfig(level=logging.DEBUG)

logger = logging.getLogger(__name__)


@dataclass
class MicrosoftDocxMetadata:
    title: str = ""
    author: str = ""
    subject: str = ""
    keywords: str = ""
    category: str = ""
    comments: str = ""
    created: str = ""
    modified: str = ""
    last_modified_by: str = ""
    revision: Optional[int] = None


@dataclass
class MicrosoftDocxRun:
    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_color: Optional[str] = None


@dataclass
class MicrosoftDocxParagraph:
    text: str = ""
    style: Optional[str] = None
    alignment: Optional[str] = None
    runs: List[MicrosoftDocxRun] = field(default_factory=list)


@dataclass
class MicrosoftDocxHeaderFooter:
    type: str = ""
    text: str = ""


@dataclass
class MicrosoftDocxImage:
    rel_id: str = ""
    filename: str = ""
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    error: Optional[str] = None


@dataclass
class MicrosoftDocxHyperlink:
    text: str = ""
    url: str = ""


@dataclass
class MicrosoftDocxNote:
    id: str = ""
    text: str = ""


@dataclass
class MicrosoftDocxComment:
    id: str = ""
    author: str = ""
    date: str = ""
    text: str = ""


@dataclass
class MicrosoftDocxSection:
    page_width_inches: Optional[float] = None
    page_height_inches: Optional[float] = None
    left_margin_inches: Optional[float] = None
    right_margin_inches: Optional[float] = None
    top_margin_inches: Optional[float] = None
    bottom_margin_inches: Optional[float] = None
    orientation: Optional[str] = None


@dataclass
class MicrosoftDocxContent(ExtractionInterface):
    metadata: MicrosoftDocxMetadata = field(default_factory=MicrosoftDocxMetadata)
    paragraphs: List[MicrosoftDocxParagraph] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    headers: List[MicrosoftDocxHeaderFooter] = field(default_factory=list)
    footers: List[MicrosoftDocxHeaderFooter] = field(default_factory=list)
    images: List[MicrosoftDocxImage] = field(default_factory=list)
    hyperlinks: List[MicrosoftDocxHyperlink] = field(default_factory=list)
    footnotes: List[MicrosoftDocxNote] = field(default_factory=list)
    endnotes: List[MicrosoftDocxNote] = field(default_factory=list)
    comments: List[MicrosoftDocxComment] = field(default_factory=list)
    sections: List[MicrosoftDocxSection] = field(default_factory=list)
    styles: List[str] = field(default_factory=list)
    full_text: str = ""

    def iterator(self) -> typing.Iterator[str]:
        for text in [self.full_text]:
            yield text


def read_docx(file_like: io.BytesIO) -> MicrosoftDocxContent:
    """
    Extract all relevant content from a DOCX file.

    Args:
        file_like: A BytesIO object containing the DOCX file data.

    Returns:
        MicrosoftDocxContent dataclass with all extracted content.
    """
    file_like.seek(0)
    doc = Document(file_like)

    # === Core Properties (Metadata) ===
    props = doc.core_properties
    metadata = MicrosoftDocxMetadata(
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
    paragraphs = []
    for para in doc.paragraphs:
        runs = []
        for run in para.runs:
            runs.append(
                MicrosoftDocxRun(
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
            MicrosoftDocxParagraph(
                text=para.text,
                style=para.style.name if para.style else None,
                alignment=str(para.alignment) if para.alignment else None,
                runs=runs,
            )
        )

    # === Tables ===
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

    # === Headers and Footers ===
    headers = []
    footers = []
    for section in doc.sections:
        # Default
        if section.header and section.header.paragraphs:
            text = "\n".join(p.text for p in section.header.paragraphs)
            if text.strip():
                headers.append(MicrosoftDocxHeaderFooter(type="default", text=text))
        if section.footer and section.footer.paragraphs:
            text = "\n".join(p.text for p in section.footer.paragraphs)
            if text.strip():
                footers.append(MicrosoftDocxHeaderFooter(type="default", text=text))

        # First page
        if section.first_page_header and section.first_page_header.paragraphs:
            text = "\n".join(p.text for p in section.first_page_header.paragraphs)
            if text.strip():
                headers.append(MicrosoftDocxHeaderFooter(type="first_page", text=text))
        if section.first_page_footer and section.first_page_footer.paragraphs:
            text = "\n".join(p.text for p in section.first_page_footer.paragraphs)
            if text.strip():
                footers.append(MicrosoftDocxHeaderFooter(type="first_page", text=text))

        # Even page
        if section.even_page_header and section.even_page_header.paragraphs:
            text = "\n".join(p.text for p in section.even_page_header.paragraphs)
            if text.strip():
                headers.append(MicrosoftDocxHeaderFooter(type="even_page", text=text))
        if section.even_page_footer and section.even_page_footer.paragraphs:
            text = "\n".join(p.text for p in section.even_page_footer.paragraphs)
            if text.strip():
                footers.append(MicrosoftDocxHeaderFooter(type="even_page", text=text))

    # === Images ===
    images = []
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            try:
                image_part = rel.target_part
                images.append(
                    MicrosoftDocxImage(
                        rel_id=rel_id,
                        filename=image_part.partname.split("/")[-1],
                        content_type=image_part.content_type,
                        data=io.BytesIO(image_part.blob),
                        size_bytes=len(image_part.blob),
                    )
                )
            except Exception as e:
                logger.debug(f"Image extraction failed for rel_id {rel_id} - {e}")
                images.append(MicrosoftDocxImage(rel_id=rel_id, error=str(e)))

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
                hyperlinks.append(
                    MicrosoftDocxHyperlink(text=text, url=rels[r_id].target_ref)
                )

    # === Footnotes ===
    footnotes = []
    try:
        if doc.part.footnotes_part:
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for fn in doc.part.footnotes_part.element.findall(".//w:footnote", ns):
                fn_id = fn.get(qn("w:id"))
                if fn_id not in ["-1", "0"]:
                    text = "".join(t.text or "" for t in fn.findall(".//w:t", ns))
                    footnotes.append(MicrosoftDocxNote(id=fn_id, text=text))
    except AttributeError as e:
        logger.debug(f"Silently ignoring footnote extraction error {e}")

    # === Endnotes ===
    endnotes = []
    try:
        if doc.part.endnotes_part:
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for en in doc.part.endnotes_part.element.findall(".//w:endnote", ns):
                en_id = en.get(qn("w:id"))
                if en_id not in ["-1", "0"]:
                    text = "".join(t.text or "" for t in en.findall(".//w:t", ns))
                    endnotes.append(MicrosoftDocxNote(id=en_id, text=text))
    except AttributeError as e:
        logger.debug(f"Silently ignoring endnote extraction error {e}")

    # === Comments ===
    comments = []
    try:
        if doc.part.comments_part:
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            for comment in doc.part.comments_part.element.findall(".//w:comment", ns):
                comments.append(
                    MicrosoftDocxComment(
                        id=comment.get(qn("w:id")) or "",
                        author=comment.get(qn("w:author")) or "",
                        date=comment.get(qn("w:date")) or "",
                        text="".join(
                            t.text or "" for t in comment.findall(".//w:t", ns)
                        ),
                    )
                )
    except AttributeError as e:
        logger.debug(f"Silently ignoring comments extraction error {e}")

    # === Sections (page layout) ===
    sections = []
    for section in doc.sections:
        sections.append(
            MicrosoftDocxSection(
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

    # === Styles used ===
    styles_set = set()
    for para in doc.paragraphs:
        if para.style:
            styles_set.add(para.style.name)
    styles = list(styles_set)

    # === Full text (convenience) ===
    all_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            all_text.append(para.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = " ".join(p.text for p in cell.paragraphs if p.text.strip())
                if text:
                    all_text.append(text)
    full_text = "\n".join(all_text)

    return MicrosoftDocxContent(
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
