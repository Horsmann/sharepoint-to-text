import io
import typing
from abc import abstractmethod
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Protocol


@dataclass
class FileMetadataInterface:
    filename: str | None = None
    file_extension: str | None = None
    file_path: str | None = None
    folder_path: str | None = None

    def populate_from_path(self, path: str | Path | None) -> None:
        """Populate file metadata fields from a path."""
        if path is None:
            return
        p = Path(path)
        self.filename = p.name
        self.file_extension = p.suffix
        self.file_path = str(p.resolve()) if p.exists() else str(p)
        self.folder_path = (
            str(p.parent.resolve()) if p.parent.exists() else str(p.parent)
        )

    def to_dict(self) -> dict:
        return asdict(self)


class ExtractionInterface(Protocol):
    @abstractmethod
    def iterator(self) -> typing.Iterator[str]:
        """
        Returns an iterator over the extracted text i.e., the main text body of a file.
        Additional text areas may be missing if they are not part of the main text body of the file.
        This greatly depends on the underlying data source.
        A PDF returns text per pages, PowerPoint files return slides as units.
        Excel files return sheets.
        Content of footnotes, headers or alike is not part of this iterator's return values.
        The legacy and modern Word documents have no per-page representation in the files, they return only a single unit which is the full text.
        """
        ...

    @abstractmethod
    def get_full_text(self) -> str:
        """Full text of the slide deck as one single block of text"""
        ...

    @abstractmethod
    def get_metadata(self) -> FileMetadataInterface:
        """Returns the metadata of the extracted file"""
        ...


@dataclass
class EmailAddress:
    name: str = ""
    address: str = ""


@dataclass
class EmailMetadata(FileMetadataInterface):
    date: str = ""
    message_id: str = ""


@dataclass
class EmailContent(ExtractionInterface):
    from_email: EmailAddress
    subject: str = ""
    in_reply_to: str = ""
    reply_to: List[EmailAddress] = field(default_factory=list)
    to_emails: List[EmailAddress] = field(default_factory=list)
    to_cc: List[EmailAddress] = field(default_factory=list)
    to_bcc: List[EmailAddress] = field(default_factory=list)
    body_plain: str = ""
    body_html: str = ""
    metadata: EmailMetadata = field(default_factory=EmailMetadata)

    def __post_init__(self):
        self.subject = self.subject.strip()
        self.body_plain = self.body_plain.strip()

    def iterator(self) -> typing.Iterator[str]:
        yield (
            self.body_plain
            if self.body_plain
            else self.body_html if self.body_html else ""
        )

    def get_full_text(self) -> str:
        return "\n".join(self.iterator())

    def get_metadata(self) -> EmailMetadata:
        return self.metadata


############
# legacy doc
#############


@dataclass
class DocMetadata(FileMetadataInterface):
    title: str = ""
    author: str = ""
    subject: str = ""
    keywords: str = ""
    last_saved_by: str = ""
    create_time: str = None
    last_saved_time: str = None
    num_pages: int = 0
    num_words: int = 0
    num_chars: int = 0


@dataclass
class DocContent(ExtractionInterface):
    main_text: str = ""
    footnotes: str = ""
    headers_footers: str = ""
    annotations: str = ""
    metadata: DocMetadata = field(default_factory=DocMetadata)

    def iterator(self) -> typing.Iterator[str]:
        for text in [self.main_text]:
            yield text

    def get_full_text(self) -> str:
        """The full text of the document including a document title from the metadata if any are provided"""
        return (self.metadata.title + "\n" + "\n".join(self.iterator())).strip()

    def get_metadata(self) -> FileMetadataInterface:
        return self.metadata


##############
# modern docx
###############


@dataclass
class DocxMetadata(FileMetadataInterface):
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
class DocxRun:
    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_color: Optional[str] = None


@dataclass
class DocxParagraph:
    text: str = ""
    style: Optional[str] = None
    alignment: Optional[str] = None
    runs: List[DocxRun] = field(default_factory=list)


@dataclass
class DocxHeaderFooter:
    type: str = ""
    text: str = ""


@dataclass
class DocxImage:
    rel_id: str = ""
    filename: str = ""
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    error: Optional[str] = None


@dataclass
class DocxHyperlink:
    text: str = ""
    url: str = ""


@dataclass
class DocxNote:
    id: str = ""
    text: str = ""


@dataclass
class DocxComment:
    id: str = ""
    author: str = ""
    date: str = ""
    text: str = ""


@dataclass
class DocxSection:
    page_width_inches: Optional[float] = None
    page_height_inches: Optional[float] = None
    left_margin_inches: Optional[float] = None
    right_margin_inches: Optional[float] = None
    top_margin_inches: Optional[float] = None
    bottom_margin_inches: Optional[float] = None
    orientation: Optional[str] = None


@dataclass
class DocxContent(ExtractionInterface):
    metadata: DocxMetadata = field(default_factory=DocxMetadata)
    paragraphs: List[DocxParagraph] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    headers: List[DocxHeaderFooter] = field(default_factory=list)
    footers: List[DocxHeaderFooter] = field(default_factory=list)
    images: List[DocxImage] = field(default_factory=list)
    hyperlinks: List[DocxHyperlink] = field(default_factory=list)
    footnotes: List[DocxNote] = field(default_factory=list)
    endnotes: List[DocxNote] = field(default_factory=list)
    comments: List[DocxComment] = field(default_factory=list)
    sections: List[DocxSection] = field(default_factory=list)
    styles: List[str] = field(default_factory=list)
    full_text: str = ""

    def iterator(self) -> typing.Iterator[str]:
        for text in [self.full_text]:
            yield text

    def get_full_text(self) -> str:
        return "\n".join(self.iterator())

    def get_metadata(self) -> DocxMetadata:
        return self.metadata


######
# PDF
######


@dataclass
class PdfImage:
    index: int = 0
    name: str = ""
    width: int = 0
    height: int = 0
    color_space: str = ""
    bits_per_component: int = 8
    filter: str = ""
    data: bytes = b""
    format: str = ""


@dataclass
class PdfPage:
    text: str = ""
    images: List[PdfImage] = field(default_factory=list)


@dataclass
class PdfMetadata(FileMetadataInterface):
    total_pages: int = 0


@dataclass
class PdfContent(ExtractionInterface):
    pages: Dict[int, PdfPage] = field(default_factory=dict)
    metadata: PdfMetadata = field(default_factory=PdfMetadata)

    def iterator(self) -> typing.Iterator[str]:
        for page_num in sorted(self.pages.keys()):
            yield self.pages[page_num].text

    def get_full_text(self) -> str:
        return "\n".join(self.iterator())

    def get_metadata(self) -> PdfMetadata:
        return self.metadata


#########
# Plain
#########


@dataclass
class PlainTextContent(ExtractionInterface):
    content: str = ""
    metadata: FileMetadataInterface = field(default_factory=FileMetadataInterface)

    def iterator(self) -> typing.Iterator[str]:
        yield self.content

    def get_full_text(self) -> str:
        return self.content

    def get_metadata(self) -> FileMetadataInterface:
        return self.metadata


#############
# legacy PPT
##############

# Text placeholder types (from TextHeaderAtom)
PPT_TEXT_TYPE_TITLE = 0  # Title
PPT_TEXT_TYPE_BODY = 1  # Body
PPT_TEXT_TYPE_NOTES = 2  # Notes
PPT_TEXT_TYPE_OTHER = 4  # Other (not title/body/notes)
PPT_TEXT_TYPE_CENTER_BODY = 5  # Center body (subtitle)
PPT_TEXT_TYPE_CENTER_TITLE = 6  # Center title
PPT_TEXT_TYPE_HALF_BODY = 7  # Half body
PPT_TEXT_TYPE_QUARTER_BODY = 8  # Quarter body


@dataclass
class PptMetadata(FileMetadataInterface):
    """Metadata extracted from a PPT file."""

    title: str = ""
    subject: str = ""
    author: str = ""
    keywords: str = ""
    comments: str = ""
    last_saved_by: str = ""
    created: str = ""
    modified: str = ""
    revision_number: str = ""
    category: str = ""
    company: str = ""
    manager: str = ""
    creating_application: str = ""
    num_slides: int = 0
    num_notes: int = 0
    num_hidden_slides: int = 0


@dataclass
class PptTextBlock:
    """Represents a block of text with its type and context."""

    text: str
    text_type: int | None = None  # From TextHeaderAtom
    is_title: bool = False
    is_body: bool = False
    is_notes: bool = False

    @property
    def type_name(self) -> str:
        """Human-readable text type name."""
        type_names = {
            PPT_TEXT_TYPE_TITLE: "title",
            PPT_TEXT_TYPE_BODY: "body",
            PPT_TEXT_TYPE_NOTES: "notes",
            PPT_TEXT_TYPE_OTHER: "other",
            PPT_TEXT_TYPE_CENTER_BODY: "subtitle",
            PPT_TEXT_TYPE_CENTER_TITLE: "center_title",
            PPT_TEXT_TYPE_HALF_BODY: "half_body",
            PPT_TEXT_TYPE_QUARTER_BODY: "quarter_body",
        }
        return type_names.get(self.text_type, "unknown")


@dataclass
class PptSlideContent:
    """Represents the content of a single slide."""

    slide_number: int
    title: str | None = None
    body_text: list[str] = field(default_factory=list)
    other_text: list[str] = field(default_factory=list)
    all_text: list[PptTextBlock] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)

    @property
    def text_combined(self) -> str:
        """All text from this slide combined."""
        parts = []
        if self.title:
            parts.append(self.title)
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)

    def to_dict(self) -> dict[str, typing.Any]:
        """Convert to dictionary representation."""
        return {
            "slide_number": self.slide_number,
            "title": self.title,
            "body_text": self.body_text,
            "other_text": self.other_text,
            # 'all_text': [{'text': tb.text, 'type': tb.type_name} for tb in self.all_text],
            "notes": self.notes,
            # 'text_combined': self.text_combined,
        }


@dataclass
class PptContent(ExtractionInterface):
    """Complete extracted content from a PPT file."""

    metadata: PptMetadata = field(default_factory=PptMetadata)
    slides: list[PptSlideContent] = field(default_factory=list)
    master_text: list[str] = field(default_factory=list)  # Text from master slides
    all_text: list[str] = field(default_factory=list)
    streams: list[list[str]] = field(default_factory=list)

    def iterator(self) -> typing.Iterator[str]:
        """Iterate over slide text, yielding combined text per slide."""
        for slide in self.slides:
            yield slide.text_combined

    def get_full_text(self) -> str:
        """Full text of the slide deck as one single block of text"""
        return "\n".join(self.iterator())

    def get_metadata(self) -> PptMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    @property
    def slide_count(self) -> int:
        """Number of slides extracted."""
        return len(self.slides)


##############
# Modern PPTX
##############


@dataclass
class PptxMetadata(FileMetadataInterface):
    title: str = ""
    subject: str = ""
    author: str = ""
    last_modified_by: str = ""
    created: str = ""
    modified: str = ""
    keywords: str = ""
    comments: str = ""
    category: str = ""
    revision: Optional[int] = None


@dataclass
class PPTXImage:
    image_index: int = 0
    filename: str = ""
    content_type: str = ""
    size_bytes: int = 0
    blob: Optional[bytes] = None


@dataclass
class PPTXSlide:
    slide_number: int = 0
    title: str = ""
    footer: str = ""
    content_placeholders: List[str] = field(default_factory=list)
    other_textboxes: List[str] = field(default_factory=list)
    images: List[PPTXImage] = field(default_factory=list)
    text: str = ""


@dataclass
class PptxContent(ExtractionInterface):
    metadata: PptxMetadata = field(default_factory=PptxMetadata)
    slides: List[PPTXSlide] = field(default_factory=list)

    def iterator(self) -> typing.Iterator[str]:
        for slide in self.slides:
            yield slide.text.strip()

    def get_full_text(self) -> str:
        return "\n".join(list(self.iterator()))

    def get_metadata(self) -> PptxMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata


#############
# Legacy XLS
#############


@dataclass
class XlsMetadata(FileMetadataInterface):
    title: str = ""
    author: str = ""
    subject: str = ""
    company: str = ""
    last_saved_by: str = ""
    created: str = ""
    modified: str = ""


@dataclass
class XlsSheet:
    name: str = ""
    data: List[Dict[str, typing.Any]] = field(default_factory=list)
    text: str = ""


@dataclass
class XlsContent(ExtractionInterface):
    metadata: XlsMetadata = field(default_factory=XlsMetadata)
    sheets: List[XlsSheet] = field(default_factory=list)
    full_text: str = ""

    def iterator(self) -> typing.Iterator[str]:
        for sheet in self.sheets:
            yield sheet.text

    def get_full_text(self) -> str:
        return self.full_text

    def get_metadata(self) -> XlsMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata


##############
# Modern XLSX
##############


@dataclass
class XlsxMetadata(FileMetadataInterface):
    title: str = ""
    description: str = ""
    creator: str = ""
    last_modified_by: str = ""
    created: str = ""
    modified: str = ""
    keywords: str = ""
    language: str = ""
    revision: Optional[str] = None


@dataclass
class XlsxSheet:
    name: str = ""
    data: List[Dict[str, typing.Any]] = field(default_factory=list)
    text: str = ""


@dataclass
class XlsxContent(ExtractionInterface):
    metadata: XlsxMetadata = field(default_factory=XlsxMetadata)
    sheets: List[XlsxSheet] = field(default_factory=list)

    def iterator(self) -> typing.Iterator[str]:
        for sheet in self.sheets:
            yield sheet.name + "\n" + sheet.text.strip()

    def get_full_text(self) -> str:
        return "\n".join(list(self.iterator()))

    def get_metadata(self) -> XlsxMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata
