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


@dataclass
class TableInterface(Protocol):
    @abstractmethod
    def get_table(self) -> list[list[str]]:
        """Return the table data as a list of rows.

        The outer list contains rows, and each inner list contains the
        string values for a single row. This format is compatible with
        pandas and polars DataFrame constructors.
        """
        pass


@dataclass
class ImageMetadata:
    # the index of the unit where this image occurs
    # will be zero for formats with no page/slide units e.g. word
    unit_index: int = 0
    # A sequential index which shows which nth image this is
    image_index: int = 0
    content_type: str = ""


class ImageInterface(Protocol):

    @abstractmethod
    def get_bytes(self) -> io.BytesIO:
        """Returns the bytes of the image as a BytesIO object."""
        pass

    @abstractmethod
    def get_content_type(self) -> str:
        """Returns the content type of the image as a string."""
        pass

    @abstractmethod
    def get_caption(self) -> str:
        """Returns the caption of the image as a string."""
        pass

    @abstractmethod
    def get_description(self) -> str:
        """Returns the descriptive text of the image as a string."""
        pass

    @abstractmethod
    def get_metadata(self) -> ImageMetadata:
        pass


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
class DocxImage(ImageInterface):
    rel_id: str = ""
    filename: str = ""
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    error: Optional[str] = None
    image_index: int = 0
    caption: str = ""  # Title/name of the image shape
    description: str = ""  # Alt text / description for accessibility

    def get_bytes(self) -> io.BytesIO:
        """Returns the bytes of the image as a BytesIO object."""
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Returns the content type of the image as a string."""
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Returns the caption of the image as a string."""
        return self.caption.strip()

    def get_description(self) -> str:
        """Returns the descriptive text of the image as a string."""
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Returns the metadata of the image."""
        return ImageMetadata(
            image_index=self.image_index,
            content_type=self.content_type,
            unit_index=0,  # DOCX has no page/slide units
        )


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
class DocxFormula:
    latex: str = ""
    is_display: bool = (
        False  # True for display equations ($$...$$), False for inline ($...$)
    )


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
    formulas: List[DocxFormula] = field(default_factory=list)
    full_text: str = ""  # Full text including formulas
    base_full_text: str = ""  # Full text without formulas

    def iterator(
        self, include_formulas: bool = False, include_comments: bool = False
    ) -> typing.Iterator[str]:
        text = self.full_text if include_formulas else self.base_full_text
        yield text

        if include_comments:
            for comment in self.comments:
                yield f"[Comment: {comment.author}@{comment.date}: {comment.text}]"

    def get_full_text(
        self, include_formulas: bool = False, include_comments: bool = False
    ) -> str:
        """Get full text of the document.

        Args:
            include_formulas: Include LaTeX formulas in output (default: False)
            include_comments: Include document comments in output (default: False)
        """
        return "\n".join(self.iterator(include_formulas, include_comments))

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
        yield self.content.strip()

    def get_full_text(self) -> str:
        return "\n".join(list(self.iterator()))

    def get_metadata(self) -> FileMetadataInterface:
        return self.metadata


########
# HTML
########


@dataclass
class HtmlMetadata(FileMetadataInterface):
    title: str = ""
    language: str = ""
    charset: str = ""
    description: str = ""
    keywords: str = ""
    author: str = ""


@dataclass
class HtmlContent(ExtractionInterface):
    content: str = ""
    tables: List[List[List[str]]] = field(default_factory=list)
    headings: List[Dict[str, str]] = field(
        default_factory=list
    )  # List of {level: "h1", text: "..."}
    links: List[Dict[str, str]] = field(
        default_factory=list
    )  # List of {text: "...", href: "..."}
    metadata: HtmlMetadata = field(default_factory=HtmlMetadata)

    def iterator(self) -> typing.Iterator[str]:
        yield self.content.strip()

    def get_full_text(self) -> str:
        return "\n".join(list(self.iterator()))

    def get_metadata(self) -> HtmlMetadata:
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
class PptxImage(ImageInterface):
    image_index: int = 0
    filename: str = ""
    content_type: str = ""
    size_bytes: int = 0
    blob: Optional[bytes] = None
    caption: str = ""  # Title/name of the image shape
    description: str = ""  # Alt text / description for accessibility
    slide_number: int = 0

    def get_bytes(self) -> io.BytesIO:
        fl = io.BytesIO(self.blob)
        fl.seek(0)
        return fl

    def get_content_type(self) -> str:
        return self.content_type

    def get_metadata(self) -> ImageMetadata:
        return ImageMetadata(
            image_index=self.image_index,
            content_type=self.content_type,
            unit_index=self.slide_number,
        )

    def get_caption(self) -> str:
        return self.caption

    def get_description(self) -> str:
        return self.description


@dataclass
class PptxFormula:
    latex: str = ""
    is_display: bool = False  # True for display equations, False for inline


@dataclass
class PptxComment:
    author: str = ""
    text: str = ""
    date: str = ""


@dataclass
class PptxSlide:
    slide_number: int = 0
    title: str = ""
    footer: str = ""
    content_placeholders: List[str] = field(default_factory=list)
    other_textboxes: List[str] = field(default_factory=list)
    images: List[PptxImage] = field(default_factory=list)
    formulas: List[PptxFormula] = field(default_factory=list)
    comments: List[PptxComment] = field(default_factory=list)
    text: str = ""  # Full text including formulas, comments, captions
    base_text: str = ""  # Text without formulas, comments, captions

    def get_text(
        self,
        include_formulas: bool = False,
        include_comments: bool = False,
        include_image_captions: bool = False,
    ) -> str:
        """Get slide text with optional inclusion of formulas, comments, and image captions."""
        parts = [self.base_text] if self.base_text else []

        if include_formulas:
            for formula in self.formulas:
                if formula.is_display:
                    parts.append(f"$${formula.latex}$$")
                else:
                    parts.append(f"${formula.latex}$")

        if include_image_captions:
            for image in self.images:
                if image.description:
                    parts.append(f"[Image: {image.description}]")

        if include_comments:
            for comment in self.comments:
                parts.append(
                    f"[Comment: {comment.author}@{comment.date}: {comment.text}]"
                )

        return "\n".join(parts)


@dataclass
class PptxContent(ExtractionInterface):
    metadata: PptxMetadata = field(default_factory=PptxMetadata)
    slides: List[PptxSlide] = field(default_factory=list)

    def iterator(
        self,
        include_formulas: bool = False,
        include_comments: bool = False,
        include_image_captions: bool = False,
    ) -> typing.Iterator[str]:
        for slide in self.slides:
            yield slide.get_text(
                include_formulas=include_formulas,
                include_comments=include_comments,
                include_image_captions=include_image_captions,
            ).strip()

    def get_full_text(
        self,
        include_formulas: bool = False,
        include_comments: bool = False,
        include_image_captions: bool = False,
    ) -> str:
        """Get full text of all slides.

        Args:
            include_formulas: Include LaTeX formulas in output (default: False)
            include_comments: Include slide comments in output (default: False)
            include_image_captions: Include image captions/alt text in output (default: False)
        """
        return "\n".join(
            list(
                self.iterator(
                    include_formulas=include_formulas,
                    include_comments=include_comments,
                    include_image_captions=include_image_captions,
                )
            )
        )

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
            yield sheet.text.strip()

    def get_full_text(self) -> str:
        return self.full_text.strip()

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
class XlsxImage(ImageInterface):
    image_index: int = 0
    sheet_index: int = 0  # 0-based index of the sheet containing this image
    filename: str = ""
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    width: int = 0
    height: int = 0
    caption: str = ""  # Title/name of the image
    description: str = ""  # Alt text / description for accessibility

    def get_bytes(self) -> io.BytesIO:
        """Returns the bytes of the image as a BytesIO object."""
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Returns the content type of the image as a string."""
        return self.content_type

    def get_caption(self) -> str:
        """Returns the caption of the image as a string."""
        return self.caption

    def get_description(self) -> str:
        """Returns the descriptive text of the image as a string."""
        return self.description

    def get_metadata(self) -> ImageMetadata:
        """Returns the metadata of the image."""
        return ImageMetadata(
            image_index=self.image_index,
            content_type=self.content_type,
            unit_index=self.sheet_index,
        )


@dataclass
class XlsxSheet(TableInterface):
    name: str = ""
    data: List[List[typing.Any]] = field(default_factory=list)
    text: str = ""
    images: List[XlsxImage] = field(default_factory=list)

    def get_table(self) -> list[list[str]]:
        return self.data


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


###############
# OpenDocument Shared Types
###############


@dataclass
class OpenDocumentMetadata(FileMetadataInterface):
    """
    Base metadata class for OpenDocument formats (ODT, ODS, ODP).

    OpenDocument files share a common metadata structure defined by the
    ODF (Open Document Format) specification. This base class captures
    the standard metadata fields found in the meta.xml file within
    ODF archives.
    """

    title: str = ""
    description: str = ""
    subject: str = ""
    creator: str = ""
    keywords: str = ""
    initial_creator: str = ""
    creation_date: str = ""
    date: str = ""  # Last modified date
    language: str = ""
    editing_cycles: int = 0
    editing_duration: str = ""
    generator: str = ""  # Application that created the document


@dataclass
class OpenDocumentAnnotation:
    """
    Represents an annotation/comment in an OpenDocument file.

    Annotations follow the same structure across all ODF formats.
    """

    creator: str = ""
    date: str = ""
    text: str = ""


@dataclass
class OpenDocumentImage(ImageInterface):
    """
    Represents an embedded image in an OpenDocument file.

    Images are stored in the Pictures/ directory within the ODF archive
    and referenced via href attributes in the content.xml.

    Implements ImageInterface for consistent image handling across formats.
    """

    href: str = ""
    name: str = ""
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    width: Optional[str] = None
    height: Optional[str] = None
    error: Optional[str] = None
    image_index: int = 0
    caption: str = ""  # From svg:title or frame name
    description: str = ""  # From svg:desc (alt text)
    unit_index: int = 0  # Page/slide number (0 for ODT documents)

    def get_bytes(self) -> io.BytesIO:
        """Returns the bytes of the image as a BytesIO object."""
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Returns the content type of the image as a string."""
        return self.content_type

    def get_caption(self) -> str:
        """Returns the caption of the image as a string."""
        return self.caption

    def get_description(self) -> str:
        """Returns the descriptive text of the image as a string."""
        return self.description

    def get_metadata(self) -> ImageMetadata:
        """Returns the metadata of the image."""
        return ImageMetadata(
            image_index=self.image_index,
            content_type=self.content_type,
            unit_index=self.unit_index,
        )


# Type aliases for backwards compatibility and semantic clarity
OdpMetadata = OpenDocumentMetadata
OdsMetadata = OpenDocumentMetadata
OdtMetadata = OpenDocumentMetadata
OdpAnnotation = OpenDocumentAnnotation
OdsAnnotation = OpenDocumentAnnotation
OdtAnnotation = OpenDocumentAnnotation
OdpImage = OpenDocumentImage
OdsImage = OpenDocumentImage
OdtImage = OpenDocumentImage


###############
# OpenDocument ODP (Presentation)
###############


@dataclass
class OdpSlide:
    """Represents a single slide in the presentation."""

    slide_number: int = 0
    name: str = ""
    title: str = ""
    body_text: List[str] = field(default_factory=list)
    other_text: List[str] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    annotations: List[OdpAnnotation] = field(default_factory=list)
    images: List[OdpImage] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)  # Speaker notes

    @property
    def text_combined(self) -> str:
        """All text from this slide combined."""
        parts = []
        if self.title:
            parts.append(self.title)
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)


@dataclass
class OdpContent(ExtractionInterface):
    """Complete extracted content from an ODP file."""

    metadata: OdpMetadata = field(default_factory=OdpMetadata)
    slides: List[OdpSlide] = field(default_factory=list)

    def iterator(
        self, include_annotations: bool = False, include_notes: bool = False
    ) -> typing.Iterator[str]:
        """Iterate over slides, yielding combined text per slide.

        Args:
            include_annotations: Include annotations/comments in output
            include_notes: Include speaker notes in output
        """
        for slide in self.slides:
            parts = [slide.text_combined]

            if include_annotations:
                for annotation in slide.annotations:
                    parts.append(
                        f"[Annotation: {annotation.creator}@{annotation.date}: {annotation.text}]"
                    )

            if include_notes:
                for note in slide.notes:
                    parts.append(f"[Note: {note}]")

            yield "\n".join(parts)

    def get_full_text(
        self, include_annotations: bool = False, include_notes: bool = False
    ) -> str:
        """Get full text of all slides.

        Args:
            include_annotations: Include annotations/comments in output (default: False)
            include_notes: Include speaker notes in output (default: False)
        """
        return "\n".join(
            list(
                self.iterator(
                    include_annotations=include_annotations, include_notes=include_notes
                )
            )
        )

    def get_metadata(self) -> OdpMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    @property
    def slide_count(self) -> int:
        """Number of slides extracted."""
        return len(self.slides)


###############
# OpenDocument ODS (Spreadsheet)
###############


@dataclass
class OdsSheet(TableInterface):
    """Represents a single sheet in the spreadsheet."""

    name: str = ""
    data: List[List[typing.Any]] = field(default_factory=list)
    text: str = ""
    annotations: List[OdsAnnotation] = field(default_factory=list)
    images: List[OdsImage] = field(default_factory=list)

    def get_table(self) -> list[list[str]]:
        return self.data


@dataclass
class OdsContent(ExtractionInterface):
    """Complete extracted content from an ODS file."""

    metadata: OdsMetadata = field(default_factory=OdsMetadata)
    sheets: List[OdsSheet] = field(default_factory=list)

    def iterator(self) -> typing.Iterator[str]:
        """Iterate over sheets, yielding text per sheet."""
        for sheet in self.sheets:
            yield (sheet.name + "\n" + sheet.text.strip()).strip()

    def get_full_text(self) -> str:
        """Get full text of all sheets."""
        return "\n".join(list(self.iterator()))

    def get_metadata(self) -> OdsMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    @property
    def sheet_count(self) -> int:
        """Number of sheets extracted."""
        return len(self.sheets)


###############
# OpenDocument ODT (Text Document)
###############


@dataclass
class OdtRun:
    """Represents a span of text with formatting."""

    text: str = ""
    style_name: Optional[str] = None
    font_name: Optional[str] = None
    font_size: Optional[str] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color: Optional[str] = None


@dataclass
class OdtParagraph:
    """Represents a paragraph in the document."""

    text: str = ""
    style_name: Optional[str] = None
    outline_level: Optional[int] = None  # For headings
    runs: List["OdtRun"] = field(default_factory=list)


@dataclass
class OdtHeaderFooter:
    """Represents a header or footer."""

    type: str = ""  # header, footer, header-left, footer-left
    text: str = ""


@dataclass
class OdtHyperlink:
    """Represents a hyperlink."""

    text: str = ""
    url: str = ""


@dataclass
class OdtNote:
    """Represents a footnote or endnote."""

    id: str = ""
    note_class: str = ""  # footnote or endnote
    text: str = ""


@dataclass
class OdtBookmark:
    """Represents a bookmark."""

    name: str = ""


@dataclass
class OdtContent(ExtractionInterface):
    """Complete extracted content from an ODT file."""

    metadata: OdtMetadata = field(default_factory=OdtMetadata)
    paragraphs: List[OdtParagraph] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    headers: List[OdtHeaderFooter] = field(default_factory=list)
    footers: List[OdtHeaderFooter] = field(default_factory=list)
    images: List[OdtImage] = field(default_factory=list)
    hyperlinks: List[OdtHyperlink] = field(default_factory=list)
    footnotes: List[OdtNote] = field(default_factory=list)
    endnotes: List[OdtNote] = field(default_factory=list)
    annotations: List[OdtAnnotation] = field(default_factory=list)
    bookmarks: List[OdtBookmark] = field(default_factory=list)
    styles: List[str] = field(default_factory=list)
    full_text: str = ""

    def iterator(self, include_annotations: bool = False) -> typing.Iterator[str]:
        """Iterate over document text.

        Args:
            include_annotations: Include annotations/comments in output
        """
        yield self.full_text

        if include_annotations:
            for annotation in self.annotations:
                yield f"[Annotation: {annotation.creator}@{annotation.date}: {annotation.text}]"

    def get_full_text(self, include_annotations: bool = False) -> str:
        """Get full text of the document.

        Args:
            include_annotations: Include annotations/comments in output (default: False)
        """
        return "\n".join(self.iterator(include_annotations))

    def get_metadata(self) -> OdtMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata


#######
# RTF
#######


@dataclass
class RtfFont:
    """Represents a font definition in an RTF document."""

    font_id: int = 0
    font_family: str = ""  # e.g., roman, swiss, modern, script, decor, tech
    font_name: str = ""
    charset: int = 0
    pitch: int = 0  # 0=default, 1=fixed, 2=variable


@dataclass
class RtfColor:
    """Represents a color in the RTF color table."""

    index: int = 0
    red: int = 0
    green: int = 0
    blue: int = 0

    @property
    def hex_color(self) -> str:
        """Return color as hex string (#RRGGBB)."""
        return f"#{self.red:02x}{self.green:02x}{self.blue:02x}"


@dataclass
class RtfStyle:
    """Represents a paragraph or character style."""

    style_id: int = 0
    style_type: str = ""  # paragraph, character, table
    style_name: str = ""
    based_on: Optional[int] = None
    next_style: Optional[int] = None


@dataclass
class RtfMetadata(FileMetadataInterface):
    """Metadata extracted from an RTF file."""

    title: str = ""
    subject: str = ""
    author: str = ""
    keywords: str = ""
    comments: str = ""
    operator: str = ""  # Last editor
    category: str = ""
    manager: str = ""
    company: str = ""
    doc_comment: str = ""  # \doccomm
    version: int = 0
    revision: int = 0
    created: str = ""
    modified: str = ""
    num_pages: int = 0
    num_words: int = 0
    num_chars: int = 0
    num_chars_with_spaces: int = 0


@dataclass
class RtfParagraph:
    """Represents a paragraph of text with formatting information."""

    text: str = ""
    style_name: Optional[str] = None
    alignment: Optional[str] = None  # left, right, center, justify
    first_line_indent: int = 0  # in twips
    left_indent: int = 0
    right_indent: int = 0
    space_before: int = 0
    space_after: int = 0
    is_bold: bool = False
    is_italic: bool = False
    is_underline: bool = False
    font_size: Optional[float] = None  # in points


@dataclass
class RtfHeaderFooter:
    """Represents a header or footer."""

    type: str = (
        ""  # header, footer, headerl, headerr, footerl, footerr, headerf, footerf
    )
    text: str = ""


@dataclass
class RtfHyperlink:
    """Represents a hyperlink in the document."""

    text: str = ""
    url: str = ""


@dataclass
class RtfBookmark:
    """Represents a bookmark in the document."""

    name: str = ""
    text: str = ""


@dataclass
class RtfField:
    """Represents a field (e.g., page number, date, STYLEREF)."""

    field_type: str = ""
    field_instruction: str = ""
    field_result: str = ""


@dataclass
class RtfImage:
    """Represents an embedded image."""

    image_type: str = ""  # pngblip, jpegblip, emfblip, wmetafile
    width: int = 0  # in twips
    height: int = 0  # in twips
    data: Optional[bytes] = None  # Binary image data


@dataclass
class RtfFootnote:
    """Represents a footnote."""

    id: int = 0
    text: str = ""


@dataclass
class RtfAnnotation:
    """Represents an annotation/comment."""

    id: str = ""
    author: str = ""
    date: str = ""
    text: str = ""


@dataclass
class RtfContent(ExtractionInterface):
    """Complete extracted content from an RTF file."""

    metadata: RtfMetadata = field(default_factory=RtfMetadata)
    fonts: List[RtfFont] = field(default_factory=list)
    colors: List[RtfColor] = field(default_factory=list)
    styles: List[RtfStyle] = field(default_factory=list)
    paragraphs: List[RtfParagraph] = field(default_factory=list)
    headers: List[RtfHeaderFooter] = field(default_factory=list)
    footers: List[RtfHeaderFooter] = field(default_factory=list)
    hyperlinks: List[RtfHyperlink] = field(default_factory=list)
    bookmarks: List[RtfBookmark] = field(default_factory=list)
    fields: List[RtfField] = field(default_factory=list)
    images: List[RtfImage] = field(default_factory=list)
    footnotes: List[RtfFootnote] = field(default_factory=list)
    annotations: List[RtfAnnotation] = field(default_factory=list)
    pages: List[str] = field(default_factory=list)  # Text per page (split on \page)
    full_text: str = ""
    raw_text_blocks: List[str] = field(default_factory=list)

    def iterator(self) -> typing.Iterator[str]:
        """Iterate over pages, yielding text per page.

        RTF documents are split on explicit page breaks (\\page).
        If no page breaks exist, yields the full document as a single unit.
        """
        if self.pages:
            for page in self.pages:
                if page.strip():
                    yield page
        elif self.full_text:
            yield self.full_text
        else:
            # Fallback: combine all paragraphs
            combined = "\n".join(p.text for p in self.paragraphs if p.text.strip())
            if combined:
                yield combined

    def get_full_text(self) -> str:
        """Full text of the RTF document as one single block of text."""
        if self.full_text:
            return self.full_text
        return "\n".join(self.iterator())

    def get_metadata(self) -> RtfMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata
