from __future__ import annotations

import io
import logging
import re
import typing
from abc import abstractmethod
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Protocol

from sharepoint2text.extractors.serialization import (
    deserialize_extraction,
    serialize_extraction,
)

logger = logging.getLogger(__name__)


_ODF_LENGTH_RE = re.compile(r"^\s*(\d+(?:\.\d+)?)\s*([a-zA-Z]+)?\s*$")


def _odf_length_to_px(length: str | None) -> int | None:
    """Convert an ODF length (e.g. '10.5cm') to pixels using 96 DPI."""
    if not length:
        return None
    match = _ODF_LENGTH_RE.match(length)
    if not match:
        return None
    value = float(match.group(1))
    unit = (match.group(2) or "px").lower()

    # https://www.w3.org/TR/css-values-3/#absolute-lengths
    if unit == "px":
        return int(round(value))
    if unit == "in":
        return int(round(value * 96.0))
    if unit == "cm":
        return int(round((value / 2.54) * 96.0))
    if unit == "mm":
        return int(round((value / 25.4) * 96.0))
    if unit == "pt":
        return int(round((value / 72.0) * 96.0))
    if unit == "pc":  # pica = 12pt
        return int(round(((value * 12.0) / 72.0) * 96.0))

    return None


##############
# Interfaces #
##############
class ExtractionInterface(Protocol):
    @abstractmethod
    def iterate_units(self) -> typing.Iterator[UnitInterface]:
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
    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        """Iterates over the extracted images"""
        ...

    @abstractmethod
    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        """Iterates over the extracted tables"""
        ...

    @abstractmethod
    def get_full_text(self) -> str:
        """Convenience full-text representation as a single string.

        Most implementations return a newline-joined representation of the
        primary text units from `iterate_units()`. Some content types may:
        - prepend a title or other metadata
        - omit optional content by default (e.g., formulas, comments, notes)
        - expose flags on `get_full_text(...)` to include that optional content

        See `README.md` ("Format-Specific Notes on `get_full_text()`") for
        format-specific details.
        """
        ...

    @abstractmethod
    def get_metadata(self) -> FileMetadataInterface:
        """Returns the metadata of the extracted file"""
        ...

    @abstractmethod
    def to_json(self) -> dict:
        """Returns a JSON-serializable dictionary representation."""
        ...

    @classmethod
    def from_json(cls, data: dict) -> ExtractionInterface:
        """
        Deserialize a JSON dictionary back to an ExtractionInterface instance.

        This is the inverse of to_json(). It reconstructs the original
        dataclass hierarchy from the serialized JSON representation.

        Args:
            data: A dictionary produced by to_json() or serialize_extraction()

        Returns:
            An instance of the appropriate ExtractionInterface subclass

        Example:
            >>> content = read_file("document.docx")
            >>> json_data = content.to_json()
            >>> restored = ExtractionInterface.from_json(json_data)
            >>> assert restored.get_full_text() == content.get_full_text()
        """
        return deserialize_extraction(data)


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
    def get_table(self) -> list[list[typing.Any]]:
        """Return the table data as a list of rows.

        The outer list contains rows, and each inner list contains the
        values for a single row. This format is compatible with pandas
        and polars DataFrame constructors.
        """
        pass

    @abstractmethod
    def get_dim(self) -> TableDim:
        """Return the table dimensions (rows, columns)."""
        pass


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


@dataclass
class UnitMetadataInterface(Protocol):
    unit_number: int


class UnitInterface(Protocol):

    @abstractmethod
    def get_text(self) -> str:
        """Returns the text of the units as a string."""
        ...

    @abstractmethod
    def get_images(self) -> list[ImageInterface]:
        """Returns the images of the units as a list."""
        ...

    @abstractmethod
    def get_tables(self) -> list[TableData]:
        """Returns the images of the units as a list."""
        ...

    @abstractmethod
    def get_metadata(self) -> UnitMetadataInterface:
        """Returns (additional) metadata of a unit."""
        ...


@dataclass
class TableDim:
    rows: int = 0
    columns: int = 0


@dataclass
class TableData(TableInterface):
    data: list[list[typing.Any]] = field(default_factory=list)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, TableDim):
            return self.get_dim() == other
        return super().__eq__(other)

    def get_table(self) -> list[list[typing.Any]]:
        """Return table data as a list of rows.

        For tables originating from PDF extraction, rows are produced by
        heuristics that infer columns from whitespace and numeric tokens.
        The approach assumes visually aligned columns, consistent row
        spacing, and row labels followed by numeric values. It may break
        or split/merge rows when PDFs use multi-line labels, multi-column
        layouts, irregular spacing, or when numbers and labels are
        interleaved out of order by the content stream.
        """
        return self.data

    def get_dim(self) -> TableDim:
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class ImageMetadata(dict):
    # the number of the unit where this image occurs (1-based for pages/slides)
    # None for formats without pages/slides (e.g. docx, odt, ods, xlsx)
    unit_number: Optional[int] = None
    # A sequential number which shows which nth image this is. The first image has value 1
    image_number: int = 0
    content_type: str = ""
    # Pixel dimensions of the image when available
    width: Optional[int] = None
    height: Optional[int] = None

    def __post_init__(self) -> None:
        dict.__init__(
            self,
            unit_number=self.unit_number,
            image_number=self.image_number,
            content_type=self.content_type,
            width=self.width,
            height=self.height,
        )

    def __setattr__(self, name: str, value: typing.Any) -> None:
        super().__setattr__(name, value)
        if name in getattr(self, "__dataclass_fields__", {}):
            dict.__setitem__(self, name, value)

    def __setitem__(self, key: str, value: typing.Any) -> None:
        dict.__setitem__(self, key, value)
        if key in getattr(self, "__dataclass_fields__", {}):
            super().__setattr__(key, value)

    def to_dict(self) -> dict:
        return asdict(self)

    def to_json(self) -> dict:
        return serialize_extraction(self)

    @property
    def unit_index(self) -> Optional[int]:
        return self.unit_number

    @unit_index.setter
    def unit_index(self, value: Optional[int]) -> None:
        self.unit_number = value

    @property
    def image_index(self) -> int:
        return self.image_number

    @image_index.setter
    def image_index(self, value: int) -> None:
        self.image_number = value


def _join_unit_text(units: typing.Iterable[UnitInterface]) -> str:
    return "\n".join(unit.get_text() for unit in units)


#########
# Email #
#########
@dataclass
class EmailUnitMetadata(UnitMetadataInterface):
    body_type: str


@dataclass
class EmailUnit(UnitInterface):
    text: str
    body_type: str = ""  # plain|html|empty

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return []

    def get_tables(self) -> list[TableData]:
        return []

    def get_metadata(self) -> UnitMetadataInterface:
        return EmailUnitMetadata(unit_number=1, body_type=self.body_type)


@dataclass
class EmailAddress:
    name: str = ""
    address: str = ""


@dataclass
class EmailMetadata(FileMetadataInterface):
    date: str = ""
    message_id: str = ""


@dataclass
class EmailAttachment:
    filename: str
    mime_type: str
    data: io.BytesIO
    is_supported_mime_type: bool = False


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
    attachments: List[EmailAttachment] = field(default_factory=list)
    metadata: EmailMetadata = field(default_factory=EmailMetadata)

    def __post_init__(self):
        self.subject = self.subject.strip()
        self.body_plain = self.body_plain.strip()

    def iterate_units(self) -> typing.Iterator[EmailUnit]:
        if self.body_plain:
            yield EmailUnit(text=self.body_plain, body_type="plain")
            return
        if self.body_html:
            yield EmailUnit(text=self.body_html, body_type="html")
            return
        yield EmailUnit(text="", body_type="empty")

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        # not supported
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        yield from ()
        return

    def iterate_supported_attachments(
        self,
    ) -> typing.Generator[ExtractionInterface, None, None]:
        """Iterates over the attachments. If the file type is supported an extracted object is returned.
        Not supported attachments are silently skipped. The attachments are extracted at call-time.
        """
        from sharepoint2text.exceptions import (
            ExtractionFileEncryptedError,
            ExtractionFileFormatNotSupportedError,
        )
        from sharepoint2text.mime_types import MIME_TYPE_MAPPING
        from sharepoint2text.router import get_extractor

        for attachment in self.attachments:
            if not attachment.is_supported_mime_type:
                logger.debug(
                    "Skipping unsupported attachment: %s (mime=%s)",
                    attachment.filename,
                    attachment.mime_type,
                )
                continue

            try:
                extractor = get_extractor(attachment.filename)
            except ExtractionFileFormatNotSupportedError:
                file_type = MIME_TYPE_MAPPING.get(attachment.mime_type)
                if not file_type:
                    logger.debug(
                        "Skipping attachment with unknown type: %s (mime=%s)",
                        attachment.filename,
                        attachment.mime_type,
                    )
                    continue
                extractor = get_extractor(f"attachment.{file_type}")

            attachment.data.seek(0)
            try:
                yield from extractor(attachment.data, attachment.filename)
            except ExtractionFileEncryptedError:
                raise
            except Exception as exc:
                logger.debug(
                    "Failed to extract attachment: %s (mime=%s) error=%s",
                    attachment.filename,
                    attachment.mime_type,
                    exc,
                )
            finally:
                attachment.data.seek(0)

    def get_full_text(self) -> str:
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> EmailMetadata:
        return self.metadata

    def to_json(self) -> dict:
        return serialize_extraction(self)


############
# legacy doc
#############


@dataclass
class DocUnit(UnitInterface):
    text: str
    unit_number: int = 1
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    images: list[DocImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> DocUnitMeta:
        return DocUnitMeta(
            unit_number=self.unit_number,
            location=list(self.location),
            heading_level=self.heading_level,
            heading_path=list(self.heading_path),
        )


@dataclass
class DocUnitMeta(UnitMetadataInterface):
    unit_number: int = 1
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)


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
class DocImage(ImageInterface):
    image_number: int
    content_type: str
    data: bytes = b""
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None
    caption: str = ""
    unit_number: Optional[int] = None

    def get_bytes(self) -> io.BytesIO:
        fl = io.BytesIO(self.data)
        fl.seek(0)
        return fl

    def get_content_type(self) -> str:
        return self.content_type.strip()

    def get_caption(self) -> str:
        return self.caption.strip()

    def get_description(self) -> str:
        return ""

    def get_metadata(self) -> ImageMetadata:
        return ImageMetadata(
            image_number=self.image_number,
            content_type=self.content_type,
            unit_number=self.unit_number,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


@dataclass
class DocContent(ExtractionInterface):
    main_text: str = ""
    footnotes: str = ""
    headers_footers: str = ""
    annotations: str = ""
    images: List[DocImage] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    metadata: DocMetadata = field(default_factory=DocMetadata)

    def iterate_units(self) -> typing.Iterator[DocUnit]:
        lines = [line.rstrip() for line in (self.main_text or "").splitlines()]
        if not lines:
            yield DocUnit(text="", unit_number=1, location=[])
            return

        base_location = [self.metadata.title] if self.metadata.title else []

        table_index = 0
        pending_tables: list[TableData] = []

        def consume_table_if_present(line: str) -> bool:
            nonlocal table_index
            if table_index >= len(self.tables):
                return False
            tokens = [t for t in line.split() if t]
            if not tokens:
                return False
            flat_table = [cell for row in self.tables[table_index] for cell in row]
            if tokens != flat_table:
                return False
            pending_tables.append(TableData(data=self.tables[table_index]))
            table_index += 1
            return True

        def heading_level_for(line: str) -> int | None:
            text = line.strip()
            if not text:
                return None
            lowered = text.lower()
            if lowered.startswith("subsection"):
                return 2
            if lowered.startswith("chapter") or lowered == "intro":
                return 1
            return None

        units: list[DocUnit] = []
        heading_stack: list[tuple[int, str]] = []
        current_heading_level: int | None = None
        current_heading_path: list[str] = []
        current_lines: list[str] = []
        current_tables: list[TableData] = []
        unit_index = 1
        any_headings = False

        def flush_current() -> None:
            nonlocal unit_index, current_lines, current_tables
            text = "\n".join(line for line in current_lines if line).strip()
            if not (text or current_tables):
                current_lines = []
                current_tables = []
                return
            units.append(
                DocUnit(
                    text=text,
                    unit_number=unit_index,
                    location=base_location + list(current_heading_path),
                    heading_level=current_heading_level,
                    heading_path=list(current_heading_path),
                    tables=list(current_tables),
                )
            )
            unit_index += 1
            current_lines = []
            current_tables = []

        for line in lines:
            if consume_table_if_present(line):
                continue

            level = heading_level_for(line)
            if level is not None:
                heading_text = line.strip()
                if heading_text:
                    any_headings = True
                    flush_current()
                    while heading_stack and heading_stack[-1][0] >= level:
                        heading_stack.pop()
                    heading_stack.append((level, heading_text))
                    current_heading_level = level
                    current_heading_path = [t for _, t in heading_stack if t]
                    if pending_tables:
                        current_tables.extend(pending_tables)
                        pending_tables = []
                continue

            text = line.strip()
            if not text:
                continue
            current_lines.append(text)

        if pending_tables:
            current_tables.extend(pending_tables)
            pending_tables = []
        flush_current()

        if not any_headings:
            yield DocUnit(
                text=self.main_text.strip(),
                unit_number=1,
                location=base_location,
                images=[],
                tables=[TableData(data=table) for table in self.tables],
            )
            return

        # Attach unassigned images (no stable anchors in legacy DOC extraction).
        for image in self.images:
            matched_unit: DocUnit | None = None
            if image.caption:
                for unit in units:
                    if image.caption in unit.text:
                        matched_unit = unit
                        break
            if matched_unit is None:
                matched_unit = next(
                    (u for u in reversed(units) if u.heading_level == 1),
                    units[-1],
                )
            matched_unit.images.append(
                DocImage(
                    image_number=image.image_number,
                    content_type=image.content_type,
                    data=image.data,
                    size_bytes=image.size_bytes,
                    width=image.width,
                    height=image.height,
                    caption=image.caption,
                    unit_number=matched_unit.unit_number,
                )
            )

        for unit in units:
            yield unit

    def get_full_text(self) -> str:
        """The full text of the document including a document title from the metadata if any are provided"""
        return (
            self.metadata.title + "\n" + _join_unit_text(self.iterate_units())
        ).strip()

    def get_metadata(self) -> FileMetadataInterface:
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for table in self.tables:
            yield TableData(data=table)

    def to_json(self) -> dict:
        return serialize_extraction(self)


##############
# modern docx
###############


@dataclass
class DocxUnit(UnitInterface):
    text: str
    unit_number: int = 1
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    images: list[DocxImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[DocxImage]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> DocxUnitMetadata:
        return DocxUnitMetadata(
            unit_number=self.unit_number,
            location=list(self.location),
            heading_level=self.heading_level,
            heading_path=list(self.heading_path),
        )


@dataclass
class DocxUnitMetadata(UnitMetadataInterface):
    unit_number: int
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)


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
    has_page_break: bool = False


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
    width: Optional[int] = None
    height: Optional[int] = None
    error: Optional[str] = None
    image_index: int = 0
    caption: str = ""  # Title/name of the image shape
    description: str = ""  # Alt text / description for accessibility
    anchor_paragraph_indices: list[int] = field(default_factory=list)

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
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=None,  # DOCX has no page/slide units
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
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
    table_anchor_paragraph_indices: list[int] = field(default_factory=list)

    def iterate_units(self) -> typing.Iterator[DocxUnit]:
        heading_re = re.compile(r"^heading\s*(\d+)\b", flags=re.IGNORECASE)

        def heading_level(style: str | None) -> int | None:
            if not style:
                return None
            match = heading_re.match(style.strip())
            if not match:
                return None
            try:
                return int(match.group(1))
            except ValueError:
                return None

        any_headings = False
        unit_index = 0
        heading_stack: list[tuple[int, str]] = []
        current_heading_level: int | None = None
        current_heading_path: list[str] = []
        current_lines: list[str] = []
        current_heading_start_paragraph_index: int | None = None
        current_has_payload: bool = False

        # Pre-index images and tables by their anchor paragraph indices so we can
        # attach them to heading-based units.
        images_by_paragraph: dict[int, list[DocxImage]] = {}
        for img in self.images:
            for para_idx in img.anchor_paragraph_indices:
                images_by_paragraph.setdefault(para_idx, []).append(img)

        table_anchors = self.table_anchor_paragraph_indices
        if len(table_anchors) != len(self.tables):
            table_anchors = [0 for _ in self.tables]
        tables_by_paragraph: dict[int, list[TableData]] = {}
        for table, para_idx in zip(self.tables, table_anchors):
            tables_by_paragraph.setdefault(para_idx, []).append(TableData(data=table))

        heading_indices: list[int] = [
            idx
            for idx, paragraph in enumerate(self.paragraphs)
            if heading_level(paragraph.style) is not None
        ]
        heading_index_set = set(heading_indices)
        next_heading_for_index: list[int | None] = [None] * len(self.paragraphs)
        next_heading: int | None = None
        for idx in range(len(self.paragraphs) - 1, -1, -1):
            next_heading_for_index[idx] = next_heading
            if idx in heading_index_set:
                next_heading = idx

        heading_has_payload: dict[int, bool] = {}
        for idx, heading_idx in enumerate(heading_indices):
            end_idx = (
                heading_indices[idx + 1] - 1
                if idx + 1 < len(heading_indices)
                else len(self.paragraphs) - 1
            )
            has_payload = False
            for para_idx in range(heading_idx + 1, end_idx + 1):
                paragraph = self.paragraphs[para_idx]
                if paragraph.text.strip():
                    has_payload = True
                    break
                if images_by_paragraph.get(para_idx) or tables_by_paragraph.get(
                    para_idx
                ):
                    has_payload = True
                    break
            heading_has_payload[heading_idx] = has_payload

        def flush_current(
            *,
            end_paragraph_index: int,
            next_heading_level: int | None = None,
        ) -> typing.Iterator[DocxUnit]:
            nonlocal unit_index
            if not current_heading_path:
                return iter(())

            text = "\n".join(line for line in current_lines if line.strip()).strip()
            start_paragraph_index = current_heading_start_paragraph_index
            if start_paragraph_index is None:
                return iter(())

            unit_images: list[DocxImage] = []
            unit_tables: list[TableData] = []
            for para_idx in range(start_paragraph_index, end_paragraph_index + 1):
                unit_images.extend(images_by_paragraph.get(para_idx, ()))
                unit_tables.extend(tables_by_paragraph.get(para_idx, ()))

            if (
                not text
                and not unit_images
                and not unit_tables
                and next_heading_level is not None
                and current_heading_level is not None
                and next_heading_level > current_heading_level
            ):
                return iter(())

            unit_index += 1
            return iter(
                [
                    DocxUnit(
                        text=text,
                        unit_number=unit_index,
                        location=list(current_heading_path),
                        heading_level=current_heading_level,
                        heading_path=list(current_heading_path),
                        images=unit_images,
                        tables=unit_tables,
                    )
                ]
            )

        for paragraph_index, paragraph in enumerate(self.paragraphs):
            level = heading_level(paragraph.style)
            if level is not None:
                any_headings = True
                yield from flush_current(
                    end_paragraph_index=paragraph_index - 1, next_heading_level=level
                )

                heading_text = paragraph.text.strip()
                while heading_stack and heading_stack[-1][0] >= level:
                    heading_stack.pop()
                heading_stack.append((level, heading_text))

                current_heading_level = level
                current_heading_path = [t for _, t in heading_stack if t]
                current_lines = []
                current_heading_start_paragraph_index = paragraph_index
                current_has_payload = bool(
                    images_by_paragraph.get(paragraph_index)
                    or tables_by_paragraph.get(paragraph_index)
                )
                continue

            if images_by_paragraph.get(paragraph_index) or tables_by_paragraph.get(
                paragraph_index
            ):
                current_has_payload = True

            if (
                current_heading_path
                and not current_has_payload
                and paragraph.has_page_break
            ):
                next_heading_index = next_heading_for_index[paragraph_index]
                if next_heading_index is not None and heading_has_payload.get(
                    next_heading_index, False
                ):
                    yield from flush_current(end_paragraph_index=paragraph_index)
                    current_heading_start_paragraph_index = paragraph_index + 1
                    current_lines = []
                    current_has_payload = False
                    continue

            text = paragraph.text.strip()
            if text:
                current_lines.append(text)
                current_has_payload = True

        if self.paragraphs:
            yield from flush_current(end_paragraph_index=len(self.paragraphs) - 1)

        if any_headings:
            return

        yield DocxUnit(
            text=self.full_text,
            unit_number=1,
            location=[self.metadata.title] if self.metadata.title else [],
            heading_level=None,
            heading_path=[],
            images=list(self.images),
            tables=[TableData(data=table) for table in self.tables],
        )

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for table in self.tables:
            yield TableData(data=table)

    def get_full_text(self) -> str:
        """Get full text of the document."""
        return self.full_text

    def get_metadata(self) -> DocxMetadata:
        return self.metadata

    def to_json(self) -> dict:
        return serialize_extraction(self)


######
# PDF
######


@dataclass
class PdfUnitMetadata(UnitMetadataInterface):
    """PDF unit metadata"""

    pass


@dataclass
class PdfUnit(UnitInterface):
    page_number: int
    text: str
    images: list[ImageInterface] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> PdfUnitMetadata:
        return PdfUnitMetadata(unit_number=self.page_number)


@dataclass
class PdfImage(ImageInterface):
    index: int = 0
    name: str = ""
    caption: str = ""
    width: int = 0
    height: int = 0
    color_space: str = ""
    bits_per_component: int = 8
    filter: str = ""
    data: bytes = b""
    format: str = ""
    content_type: str = ""
    unit_index: Optional[int] = None

    def get_bytes(self) -> io.BytesIO:
        """Returns the bytes of the image as a BytesIO object."""
        fl = io.BytesIO(self.data)
        fl.seek(0)
        return fl

    def get_content_type(self) -> str:
        """Returns the content type of the image as a string."""
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Returns the caption of the image as a string."""
        return self.caption.strip()

    def get_description(self) -> str:
        """Returns the descriptive text of the image as a string."""
        return self.name

    def get_metadata(self) -> ImageMetadata:
        return ImageMetadata(
            image_number=self.index,
            content_type=self.get_content_type(),
            unit_number=self.unit_index,
            width=self.width if self.width > 0 else None,
            height=self.height if self.height > 0 else None,
        )


@dataclass
class PdfPage:
    text: str = ""
    images: List[PdfImage] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)


@dataclass
class PdfMetadata(FileMetadataInterface):
    total_pages: int = 0


@dataclass
class PdfContent(ExtractionInterface):
    pages: List[PdfPage] = field(default_factory=list)
    metadata: PdfMetadata = field(default_factory=PdfMetadata)

    def iterate_units(self) -> typing.Iterator[PdfUnit]:
        for page_number, page in enumerate(self.pages, start=1):
            yield PdfUnit(
                page_number=page_number,
                text=page.text,
                images=list(page.images),
                tables=[TableData(data=table) for table in page.tables],
            )

    def get_full_text(self) -> str:
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> PdfMetadata:
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for page in self.pages:
            for img in page.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        """Yield tables extracted from PDF pages.

        PDF tables are inferred from text layout heuristics, not explicit
        table structures. Extraction assumes consistent column alignment,
        row spacing, and labels followed by numeric values. Results may be
        incomplete or fragmented for multi-column pages, multi-line labels,
        merged cells, or when the PDF content stream interleaves text out
        of visual order.
        """
        for page in self.pages:
            for table in page.tables:
                yield TableData(data=table)

    def to_json(self) -> dict:
        return serialize_extraction(self)


#########
# Plain
#########


@dataclass
class PlainUnitMetadata(UnitMetadataInterface):
    """Plain Unit Metadata"""

    pass


@dataclass
class PlainTextUnit(UnitInterface):
    text: str

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return []

    def get_tables(self) -> list[TableData]:
        return []

    def get_metadata(self) -> PlainUnitMetadata:
        return PlainUnitMetadata(unit_number=1)


@dataclass
class PlainTextContent(ExtractionInterface):
    content: str = ""
    metadata: FileMetadataInterface = field(default_factory=FileMetadataInterface)

    def iterate_units(self) -> typing.Iterator[PlainTextUnit]:
        yield PlainTextUnit(text=self.content.strip())

    def get_full_text(self) -> str:
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> FileMetadataInterface:
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        yield from ()
        return

    def __post_init__(self):
        self.content = self.content.strip()

    def to_json(self) -> dict:
        return serialize_extraction(self)


########
# HTML
########


@dataclass
class HtmlUnitMetadata(UnitMetadataInterface):
    """Html Unit Metadata"""

    pass


@dataclass
class HtmlUnit(UnitInterface):
    text: str

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return []

    def get_tables(self) -> list[TableData]:
        return []

    def get_metadata(self) -> HtmlUnitMetadata:
        return HtmlUnitMetadata(unit_number=1)


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

    def iterate_units(self) -> typing.Iterator[HtmlUnit]:
        yield HtmlUnit(text=self.content.strip())

    def get_full_text(self) -> str:
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> HtmlMetadata:
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for table in self.tables:
            yield TableData(data=table)

    def to_json(self) -> dict:
        return serialize_extraction(self)


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
class PptUnitMetadata(UnitMetadataInterface):
    """Ppt Unit Metadata"""

    ...


@dataclass
class PptUnit(UnitInterface):
    slide_number: int
    text: str

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return []

    def get_tables(self) -> list[TableData]:
        return []

    def get_metadata(self) -> PptUnitMetadata:
        return PptUnitMetadata(unit_number=self.slide_number)


@dataclass
class PptImage(ImageInterface):
    """Represents an embedded image in a legacy PPT file."""

    image_index: int = 0
    content_type: str = ""
    data: bytes = b""
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None
    slide_number: int = 0

    def get_bytes(self) -> io.BytesIO:
        fl = io.BytesIO(self.data)
        fl.seek(0)
        return fl

    def get_content_type(self) -> str:
        return self.content_type.strip()

    def get_caption(self) -> str:
        return ""

    def get_description(self) -> str:
        return ""

    def get_metadata(self) -> ImageMetadata:
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.slide_number if self.slide_number > 0 else None,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


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
    images: list["PptImage"] = field(default_factory=list)

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

    def iterate_units(self) -> typing.Iterator[PptUnit]:
        """Iterate over slide text, yielding combined text per slide."""
        for slide in self.slides:
            yield PptUnit(slide_number=slide.slide_number, text=slide.text_combined)

    def get_full_text(self) -> str:
        """Full text of the slide deck as one single block of text"""
        texts = [unit.get_text().strip() for unit in self.iterate_units()]
        return "\n".join(text for text in texts if text)

    def get_metadata(self) -> PptMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    @property
    def slide_count(self) -> int:
        """Number of slides extracted."""
        return len(self.slides)

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        """Iterate over images from all slides."""
        for slide in self.slides:
            for img in slide.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        yield from ()
        return

    def to_json(self) -> dict:
        return serialize_extraction(self)


##############
# Modern PPTX
##############


@dataclass
class PptxUnitMetadata(UnitMetadataInterface):
    """Pptx Unit Metadata"""

    pass


@dataclass
class PptxUnit(UnitInterface):
    slide_number: int
    text: str
    images: list[PptxImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> PptxUnitMetadata:
        return PptxUnitMetadata(unit_number=self.slide_number)


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
    width: Optional[int] = None
    height: Optional[int] = None
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
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.slide_number,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
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
    tables: List[List[List[str]]] = field(default_factory=list)
    images: List[PptxImage] = field(default_factory=list)
    formulas: List[PptxFormula] = field(default_factory=list)
    comments: List[PptxComment] = field(default_factory=list)
    text: str = ""  # Full text including formulas, comments, captions
    base_text: str = ""  # Text without formulas, comments, captions

    def get_text(
        self,
        include_image_captions: bool = False,
    ) -> str:
        """Get slide text with formulas included and optional image captions."""
        parts = [self.base_text] if self.base_text else []

        for formula in self.formulas:
            if formula.is_display:
                parts.append(f"$${formula.latex}$$")
            else:
                parts.append(f"${formula.latex}$")

        if include_image_captions:
            for image in self.images:
                if image.description:
                    parts.append(f"[Image: {image.description}]")

        return "\n".join(parts)


@dataclass
class PptxContent(ExtractionInterface):
    metadata: PptxMetadata = field(default_factory=PptxMetadata)
    slides: List[PptxSlide] = field(default_factory=list)

    def iterate_units(
        self,
        include_image_captions: bool = False,
    ) -> typing.Iterator[PptxUnit]:
        for slide in self.slides:
            yield PptxUnit(
                slide_number=slide.slide_number,
                images=list(slide.images),
                tables=[TableData(data=table) for table in slide.tables],
                text=slide.get_text(
                    include_image_captions=include_image_captions,
                ).strip(),
            )

    def get_full_text(
        self,
        include_image_captions: bool = False,
    ) -> str:
        """Get full text of all slides.

        Args:
            include_image_captions: Include image captions/alt text in output (default: False)
        """
        return _join_unit_text(
            self.iterate_units(include_image_captions=include_image_captions)
        )

    def get_metadata(self) -> PptxMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for slide in self.slides:
            for img in slide.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for slide in self.slides:
            for table in slide.tables:
                yield TableData(data=table)

    def to_json(self) -> dict:
        return serialize_extraction(self)


#############
# Legacy XLS
#############


@dataclass
class XlsUnitMetadata(UnitMetadataInterface):
    sheet_name: str


@dataclass
class XlsUnit(UnitInterface):
    sheet_number: int
    sheet_name: str
    text: str
    tables: list[TableData] = field(default_factory=list)
    images: list[XlsImage] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> XlsUnitMetadata:
        return XlsUnitMetadata(
            unit_number=self.sheet_number, sheet_name=self.sheet_name
        )


@dataclass
class XlsImage(ImageInterface):
    """Represents an embedded image in a legacy XLS file."""

    image_index: int = 0
    content_type: str = ""
    data: bytes = b""
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None

    def get_bytes(self) -> io.BytesIO:
        fl = io.BytesIO(self.data)
        fl.seek(0)
        return fl

    def get_content_type(self) -> str:
        return self.content_type.strip()

    def get_caption(self) -> str:
        return ""

    def get_description(self) -> str:
        return ""

    def get_metadata(self) -> ImageMetadata:
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=None,  # XLS images are workbook-level, not sheet-level
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


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
class XlsSheet(TableInterface):
    name: str = ""
    data: List[Dict[str, typing.Any]] = field(default_factory=list)
    text: str = ""

    def get_table(self) -> list[list[typing.Any]]:
        if not self.data:
            return []
        headers = list(self.data[0].keys())
        rows = [headers]
        for row in self.data:
            rows.append([row.get(header) for header in headers])
        return rows

    def get_dim(self) -> TableDim:
        table = self.get_table()
        rows = len(table)
        columns = max((len(row) for row in table), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class XlsContent(ExtractionInterface):
    metadata: XlsMetadata = field(default_factory=XlsMetadata)
    sheets: List[XlsSheet] = field(default_factory=list)
    images: List[XlsImage] = field(default_factory=list)
    full_text: str = ""

    def iterate_units(self) -> typing.Iterator[XlsUnit]:
        for sheet_index, sheet in enumerate(self.sheets, start=1):
            table = sheet.get_table()
            normalized_table = (
                [
                    [str(cell) if cell is not None else None for cell in row]
                    for row in table
                ]
                if table
                else []
            )
            yield XlsUnit(
                sheet_number=sheet_index,
                sheet_name=sheet.name,
                tables=[TableData(data=normalized_table)] if normalized_table else [],
                images=list(self.images) if sheet_index == 1 else [],
                text=sheet.text.strip(),
            )

    def get_full_text(self) -> str:
        return self.full_text.strip()

    def get_metadata(self) -> XlsMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        """Iterate over images from the workbook."""
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for sheet in self.sheets:
            yield sheet

    def to_json(self) -> dict:
        return serialize_extraction(self)


##############
# Modern XLSX
##############


@dataclass
class XlsxUnitMetadata(UnitMetadataInterface):
    sheet_number: int
    sheet_name: str


@dataclass
class XlsxUnit(UnitInterface):
    sheet_index: int
    sheet_name: str
    text: str
    images: list[XlsxImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> XlsxUnitMetadata:
        return XlsxUnitMetadata(
            unit_number=self.sheet_index,
            sheet_name=self.sheet_name,
            sheet_number=self.sheet_index,
        )


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
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=None,  # XLSX has sheets, not pages/slides
            width=self.width if self.width > 0 else None,
            height=self.height if self.height > 0 else None,
        )


@dataclass
class XlsxSheet(TableInterface):
    name: str = ""
    data: List[List[typing.Any]] = field(default_factory=list)
    text: str = ""
    images: List[XlsxImage] = field(default_factory=list)

    def get_table(self) -> list[list[typing.Any]]:
        return self.data

    def get_dim(self) -> TableDim:
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class XlsxContent(ExtractionInterface):
    metadata: XlsxMetadata = field(default_factory=XlsxMetadata)
    sheets: List[XlsxSheet] = field(default_factory=list)

    def iterate_units(self) -> typing.Iterator[XlsxUnit]:
        for sheet_index, sheet in enumerate(self.sheets, start=1):
            yield XlsxUnit(
                sheet_index=sheet_index,
                sheet_name=sheet.name,
                images=list(sheet.images),
                tables=[TableData(data=sheet.data)] if sheet.data else [],
                text=sheet.name + "\n" + sheet.text.strip(),
            )

    def get_full_text(self) -> str:
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> XlsxMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for sheet in self.sheets:
            for img in sheet.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        """A single sheet is considered a full table"""
        for sheet in self.sheets:
            yield sheet

    def to_json(self) -> dict:
        return serialize_extraction(self)


#############################
# OpenDocument Shared Types #
#############################


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
    unit_index: Optional[int] = None  # Page/slide number (None for ODT/ODS)

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
        width_px = _odf_length_to_px(self.width)
        height_px = _odf_length_to_px(self.height)
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.unit_index,
            width=width_px if width_px and width_px > 0 else None,
            height=height_px if height_px and height_px > 0 else None,
        )


###############
# OpenDocument ODP (Presentation)
###############


@dataclass
class OdpUnit(UnitInterface):
    slide_number: int
    text: str
    location: list[str] = field(default_factory=list)
    images: list[OpenDocumentImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> OdpUnitMetadata:
        return OdpUnitMetadata(
            unit_number=self.slide_number,
            location=list(self.location),
            slide_number=self.slide_number,
        )


@dataclass
class OdpUnitMetadata(UnitMetadataInterface):
    unit_number: int
    location: list[str] = field(default_factory=list)
    slide_number: int = 1


@dataclass
class OdpSlide:
    """Represents a single slide in the presentation."""

    slide_number: int = 0
    name: str = ""
    title: str = ""
    body_text: List[str] = field(default_factory=list)
    other_text: List[str] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    annotations: List[OpenDocumentAnnotation] = field(default_factory=list)
    images: List[OpenDocumentImage] = field(default_factory=list)
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

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    slides: List[OdpSlide] = field(default_factory=list)

    def iterate_units(self) -> typing.Iterator[OdpUnit]:
        """Iterate over slides, yielding combined text per slide."""
        for slide in self.slides:
            parts = [slide.text_combined]

            yield OdpUnit(
                slide_number=slide.slide_number,
                text="\n".join(parts),
                location=[slide.title] if slide.title else [],
                images=list(slide.images),
                tables=[TableData(data=table) for table in slide.tables],
            )

    def get_full_text(self) -> str:
        """Get full text of all slides."""
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> OpenDocumentMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    @property
    def slide_count(self) -> int:
        """Number of slides extracted."""
        return len(self.slides)

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for slides in self.slides:
            for img in slides.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for slide in self.slides:
            for table in slide.tables:
                yield TableData(data=table)

    def to_json(self) -> dict:
        return serialize_extraction(self)


##################################
# OpenDocument ODS (Spreadsheet) #
##################################


@dataclass
class OdsUnit(UnitInterface):
    sheet_number: int
    sheet_name: str
    text: str
    images: list[OpenDocumentImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> OdsUnitMetadata:
        return OdsUnitMetadata(
            unit_number=self.sheet_number,
            sheet_number=self.sheet_number,
            sheet_name=self.sheet_name,
        )


@dataclass
class OdsUnitMetadata(UnitMetadataInterface):
    unit_number: int
    sheet_number: int
    sheet_name: str


@dataclass
class OdsSheet(TableInterface):
    """Represents a single sheet in the spreadsheet."""

    name: str = ""
    data: List[List[typing.Any]] = field(default_factory=list)
    text: str = ""
    annotations: List[OpenDocumentAnnotation] = field(default_factory=list)
    images: List[OpenDocumentImage] = field(default_factory=list)

    def get_table(self) -> list[list[typing.Any]]:
        return self.data

    def get_dim(self) -> TableDim:
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class OdsContent(ExtractionInterface):
    """Complete extracted content from an ODS file."""

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    sheets: List[OdsSheet] = field(default_factory=list)

    def iterate_units(self) -> typing.Iterator[OdsUnit]:
        """Iterate over sheets, yielding text per sheet."""
        for sheet_index, sheet in enumerate(self.sheets, start=1):
            yield OdsUnit(
                sheet_number=sheet_index,
                sheet_name=sheet.name,
                images=list(sheet.images),
                tables=[TableData(data=sheet.data)] if sheet.data else [],
                text=(sheet.name + "\n" + sheet.text.strip()).strip(),
            )

    def get_full_text(self) -> str:
        """Get full text of all sheets."""
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> OpenDocumentMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    @property
    def sheet_count(self) -> int:
        """Number of sheets extracted."""
        return len(self.sheets)

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for sheet in self.sheets:
            for img in sheet.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        """Ods is a spreadsheet format. The entire sheet is returned as table object"""
        for sheet in self.sheets:
            yield sheet

    def to_json(self) -> dict:
        return serialize_extraction(self)


####################################
# OpenDocument ODT (Text Document) #
####################################


@dataclass
class OdtUnit(UnitInterface):
    text: str
    unit_number: int
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    kind: str = "body"  # body|annotation
    annotation_creator: str | None = None
    annotation_date: str | None = None
    images: list[ImageInterface] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return list(self.tables)

    def get_metadata(self) -> OdtUnitMetadata:
        return OdtUnitMetadata(
            unit_number=self.unit_number,
            heading_level=self.heading_level,
            heading_path=list(self.heading_path),
            kind=self.kind,
            annotation_creator=self.annotation_creator,
            annotation_date=self.annotation_date,
        )


@dataclass
class OdtUnitMetadata(UnitMetadataInterface):
    unit_number: int
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    kind: str = "body"  # body|annotation
    annotation_creator: str | None = None
    annotation_date: str | None = None


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
class OdtTable(TableInterface):
    """Represents a single table in the document."""

    data: List[List[str]] = field(default_factory=list)

    def get_table(self) -> list[list[typing.Any]]:
        return self.data

    def get_dim(self) -> TableDim:
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class OdtContent(ExtractionInterface):
    """Complete extracted content from an ODT file."""

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    paragraphs: List[OdtParagraph] = field(default_factory=list)
    tables: List[OdtTable] = field(default_factory=list)
    headers: List[OdtHeaderFooter] = field(default_factory=list)
    footers: List[OdtHeaderFooter] = field(default_factory=list)
    images: List[OpenDocumentImage] = field(default_factory=list)
    hyperlinks: List[OdtHyperlink] = field(default_factory=list)
    footnotes: List[OdtNote] = field(default_factory=list)
    endnotes: List[OdtNote] = field(default_factory=list)
    annotations: List[OpenDocumentAnnotation] = field(default_factory=list)
    bookmarks: List[OdtBookmark] = field(default_factory=list)
    styles: List[str] = field(default_factory=list)
    full_text: str = ""

    def iterate_units(self) -> typing.Iterator[OdtUnit]:
        """Iterate over heading-based units.

        Units are built from paragraph runs separated by headings (paragraphs with
        an outline level). Heading text itself becomes part of the unit heading
        path and is not included in the unit body text.
        """
        base_heading_path = [self.metadata.title] if self.metadata.title else []
        units: list[OdtUnit] = []

        if not self.paragraphs:
            heading_path = list(base_heading_path)
            units.append(
                OdtUnit(
                    text=self.full_text,
                    kind="body",
                    unit_number=1,
                    heading_level=1 if heading_path else None,
                    heading_path=heading_path,
                    images=list(self.images),
                    tables=[TableData(data=table.data) for table in self.tables],
                )
            )
            for unit in units:
                yield unit
            return

        heading_stack: list[tuple[int, str]] = []
        current_heading_level: int | None = None
        current_heading_path: list[str] = []
        current_lines: list[str] = []
        current_tables: list[TableData] = []
        unit_index = 1
        any_headings = False

        table_index = 0
        pending_tables: list[TableData] = []
        in_table_block = False

        def flush_current() -> None:
            nonlocal unit_index, current_lines, current_tables
            text = "\n".join(line for line in current_lines if line).strip()
            if not (text or current_tables):
                current_lines = []
                current_tables = []
                return

            unit_heading_path = list(base_heading_path)
            for token in current_heading_path:
                if not unit_heading_path or unit_heading_path[-1] != token:
                    unit_heading_path.append(token)

            units.append(
                OdtUnit(
                    text=text,
                    unit_number=unit_index,
                    heading_level=current_heading_level,
                    heading_path=unit_heading_path,
                    kind="body",
                    tables=list(current_tables),
                )
            )
            unit_index += 1
            current_lines = []
            current_tables = []

        for paragraph in self.paragraphs:
            heading_level = paragraph.outline_level
            if heading_level is not None:
                heading_text = paragraph.text.strip()
                if heading_text:
                    any_headings = True
                    flush_current()

                    while heading_stack and heading_stack[-1][0] >= heading_level:
                        heading_stack.pop()
                    heading_stack.append((heading_level, heading_text))
                    current_heading_level = heading_level
                    current_heading_path = [t for _, t in heading_stack if t]
                    if pending_tables:
                        current_tables.extend(pending_tables)
                        pending_tables = []
                continue

            style = paragraph.style_name or ""
            is_table_paragraph = style.startswith("Table") or "Table_" in style
            if is_table_paragraph:
                if not in_table_block:
                    in_table_block = True
                    if table_index < len(self.tables):
                        table = self.tables[table_index]
                        table_index += 1
                        pending_tables.append(TableData(data=table.data))
                continue
            in_table_block = False

            text = paragraph.text.strip()
            if text:
                current_lines.append(text)

        if pending_tables:
            current_tables.extend(pending_tables)
            pending_tables = []

        flush_current()

        if not any_headings:
            heading_path = list(base_heading_path)
            units = [
                OdtUnit(
                    text=self.full_text,
                    kind="body",
                    unit_number=1,
                    heading_level=1 if heading_path else None,
                    heading_path=heading_path,
                    images=list(self.images),
                    tables=[TableData(data=table.data) for table in self.tables],
                )
            ]
            for image in self.images:
                image.unit_index = 1
            for unit in units:
                yield unit
            return

        # Best-effort mapping of unassigned tables/images to units.
        # (ODT extraction does not currently provide stable positional anchors.)
        if units:
            if table_index < len(self.tables):
                remaining_tables = self.tables[table_index:]
                for table in remaining_tables:
                    table_data = TableData(data=table.data)
                    header_tokens = [
                        str(cell).strip()
                        for cell in (table.data[0] if table.data else [])
                        if str(cell).strip()
                    ]
                    matched_unit: OdtUnit | None = None
                    if header_tokens:
                        for unit in units:
                            if all(token in unit.text for token in header_tokens):
                                matched_unit = unit
                                break
                    (matched_unit or units[-1]).tables.append(table_data)

            for image in self.images:
                matched_unit: OdtUnit | None = None
                for unit in units:
                    if image.caption and image.caption in unit.text:
                        matched_unit = unit
                        break
                    if image.description and image.description in unit.text:
                        matched_unit = unit
                        break
                if matched_unit is None:
                    if len(units) == 1:
                        matched_unit = units[0]
                    else:
                        matched_unit = next(
                            (
                                u
                                for u in reversed(units)
                                if u.heading_level == 1 or u.heading_level is None
                            ),
                            units[-1],
                        )

                image.unit_index = matched_unit.unit_number
                matched_unit.images.append(image)

        for unit in units:
            yield unit

    def get_full_text(self) -> str:
        """Get full text of the document."""
        return self.full_text

    def get_metadata(self) -> OpenDocumentMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        for table in self.tables:
            yield table

    def to_json(self) -> dict:
        return serialize_extraction(self)


#######
# RTF #
#######


@dataclass
class RtfUnitMetadata(UnitMetadataInterface):
    page_number: int


@dataclass
class RtfUnit(UnitInterface):
    page_number: int
    text: str
    images: List[RtfImage] = field(default_factory=list)
    tables: List[RtfTable] = field(default_factory=list)

    def get_text(self) -> str:
        return self.text

    def get_images(self) -> list[ImageInterface]:
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        return [TableData(data=t.data) for t in self.tables]

    def get_metadata(self) -> RtfUnitMetadata:
        return RtfUnitMetadata(
            unit_number=self.page_number, page_number=self.page_number
        )


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
class RtfImage(ImageInterface):
    """Represents an embedded image in an RTF document."""

    image_type: str = ""  # png, jpeg, emf, wmf
    width: int = 0  # in twips (1/1440 inch)
    height: int = 0  # in twips
    data: Optional[bytes] = None  # Binary image data
    image_index: int = 0  # Sequential index of the image (1-based)
    page_number: Optional[int] = None  # Page where image appears (if known)
    caption: str = ""  # Image caption/title if available
    description: str = ""  # Alt text/description if available

    # Content type mapping for RTF image types
    _CONTENT_TYPES: typing.ClassVar[dict[str, str]] = {
        "png": "image/png",
        "jpeg": "image/jpeg",
        "jpg": "image/jpeg",
        "emf": "image/x-emf",
        "wmf": "image/x-wmf",
        "unknown": "application/octet-stream",
    }

    def get_bytes(self) -> io.BytesIO:
        """Returns the bytes of the image as a BytesIO object."""
        if self.data is None:
            return io.BytesIO()
        return io.BytesIO(self.data)

    def get_content_type(self) -> str:
        """Returns the content type of the image as a string."""
        return self._CONTENT_TYPES.get(
            self.image_type.lower(), "application/octet-stream"
        )

    def get_caption(self) -> str:
        """Returns the caption of the image as a string."""
        return self.caption.strip()

    def get_description(self) -> str:
        """Returns the descriptive text of the image as a string."""
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Returns the metadata of the image."""
        # Convert twips to pixels (approximately 1/20 point, 96 dpi)
        # 1 twip = 1/1440 inch, at 96 dpi: pixels = twips * 96 / 1440 = twips / 15
        width_px = self.width // 15 if self.width > 0 else None
        height_px = self.height // 15 if self.height > 0 else None
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.get_content_type(),
            unit_number=self.page_number,
            width=width_px,
            height=height_px,
        )


@dataclass
class RtfTable(TableInterface):
    """Represents a table extracted from an RTF document."""

    data: List[List[str]] = field(default_factory=list)
    table_index: int = 0  # Sequential index of the table (1-based)
    page_number: Optional[int] = None  # Page where table appears (if known)

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table data as a list of rows."""
        return self.data

    def get_dim(self) -> TableDim:
        """Return the table dimensions (rows, columns)."""
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


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
    tables: List[RtfTable] = field(default_factory=list)
    footnotes: List[RtfFootnote] = field(default_factory=list)
    annotations: List[RtfAnnotation] = field(default_factory=list)
    pages: List[str] = field(default_factory=list)  # Text per page (split on \page)
    full_text: str = ""
    raw_text_blocks: List[str] = field(default_factory=list)

    def iterate_units(self) -> typing.Iterator[RtfUnit]:
        """Iterate over pages, yielding text per page.

        RTF documents are split on explicit page breaks (\\page).
        If no page breaks exist, yields the full document as a single unit.
        Images and tables are distributed to units based on their page_number.
        """
        # Group images and tables by page number
        images_by_page: dict[int, List[RtfImage]] = {}
        for img in self.images:
            page = img.page_number or 1
            if page not in images_by_page:
                images_by_page[page] = []
            images_by_page[page].append(img)

        tables_by_page: dict[int, List[RtfTable]] = {}
        for tbl in self.tables:
            page = tbl.page_number or 1
            if page not in tables_by_page:
                tables_by_page[page] = []
            tables_by_page[page].append(tbl)

        if self.pages:
            for page_number, page in enumerate(self.pages, start=1):
                if page.strip():
                    yield RtfUnit(
                        page_number=page_number,
                        text=page,
                        images=images_by_page.get(page_number, []),
                        tables=tables_by_page.get(page_number, []),
                    )
        elif self.full_text:
            yield RtfUnit(
                page_number=1,
                text=self.full_text,
                images=images_by_page.get(1, []),
                tables=tables_by_page.get(1, []),
            )
        else:
            # Fallback: combine all paragraphs
            combined = "\n".join(p.text for p in self.paragraphs if p.text.strip())
            if combined:
                yield RtfUnit(
                    page_number=1,
                    text=combined,
                    images=images_by_page.get(1, []),
                    tables=tables_by_page.get(1, []),
                )

    def get_full_text(self) -> str:
        """Full text of the RTF document as one single block of text."""
        if self.full_text:
            return self.full_text
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> RtfMetadata:
        """Returns the metadata of the extracted file."""
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageInterface, None, None]:
        """Iterate over all images in the document."""
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableInterface, None, None]:
        """Iterate over all tables in the document."""
        for tbl in self.tables:
            yield tbl

    def to_json(self) -> dict:
        return serialize_extraction(self)
