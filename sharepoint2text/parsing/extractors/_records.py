"""Internal parser-native records used while constructing normalized output."""

from __future__ import annotations

import io
import logging
import re
import typing
from abc import abstractmethod
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Protocol

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
class ExtractionRecord(Protocol):
    """Define the common contract implemented by every extraction result."""

    @abstractmethod
    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """
        Returns an iterator over the extracted text i.e., the main text body of a file.
        Additional text areas may be missing if they are not part of the main text body of the file.
        This greatly depends on the underlying data source.
        A PDF returns text per pages, PowerPoint files return slides as units.
        Excel files return sheets.
        Content of footnotes, headers or alike is not part of this iterator's return values.
        The legacy and modern Word documents have no per-page representation in the files, they return only a single unit which is the full text.

        Args:
            ignore_images: If True, units will have empty image lists (images are not
                returned in unit data). This can improve performance when image data
                is not needed. Default is False.
        """
        ...

    @abstractmethod
    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this document.

        Yields:
            Image objects in source order.
        """
        ...

    @abstractmethod
    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this document.

        Yields:
            Table objects in source order.
        """
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
    def get_metadata(self) -> SourceRecord:
        """Return metadata describing this document object.

        Returns:
            The format-specific metadata instance.
        """
        ...


@dataclass
class SourceRecord:
    """Store source identity and decoding metadata shared by extracted files."""

    filename: str | None = None
    file_extension: str | None = None
    file_path: str | None = None
    folder_path: str | None = None
    detected_encoding: str | None = None

    def populate_from_path(self, path: str | Path | None) -> None:
        """Populate source metadata from a filesystem path.

        Args:
            path: Source path used to populate filename and folder fields.

        Existing fields are replaced only when a path is supplied.
        """
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
        """Convert this value to a plain dictionary.

        Returns:
            A dictionary containing the dataclass fields and their values.
        """
        return asdict(self)


@dataclass
class TableRecord(Protocol):
    """Define the tabular-data contract exposed by extraction results."""

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
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
        pass


class ImageRecord(Protocol):
    """Define access to an extracted image and its descriptive metadata."""

    @abstractmethod
    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        pass

    @abstractmethod
    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        pass

    @abstractmethod
    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        pass

    @abstractmethod
    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        pass

    @abstractmethod
    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this document object.

        Returns:
            The format-specific metadata instance.
        """
        pass


@dataclass
class UnitMetadataRecord(Protocol):
    """Mark metadata objects associated with one structural extraction unit."""

    unit_number: int


class UnitRecord(Protocol):
    """Define a structural unit of extracted text, images, tables, and metadata."""

    @abstractmethod
    def get_text(self) -> str:
        """Return the text represented by this document unit.

        Returns:
            Extracted text in reading order.
        """
        ...

    @abstractmethod
    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this document unit.

        Returns:
            A new list containing the unit image objects.
        """
        ...

    @abstractmethod
    def get_tables(self) -> list[TableData]:
        """Return tables associated with this document unit.

        Returns:
            A new list containing the unit table objects.
        """
        ...

    @abstractmethod
    def get_metadata(self) -> UnitMetadataRecord:
        """Return metadata describing this document object.

        Returns:
            The format-specific metadata instance.
        """
        ...


@dataclass
class TableDim:
    """Store the row and column dimensions of a table."""

    rows: int = 0
    columns: int = 0


@dataclass
class TableData(TableRecord):
    """Represent a generic two-dimensional table returned by an extractor."""

    data: list[list[typing.Any]] = field(default_factory=list)

    def __eq__(self, other: object) -> bool:
        """Compare table data against dimensions, wrappers, or raw row lists."""
        if isinstance(other, TableDim):
            return self.get_dim() == other
        if isinstance(other, list):
            return self.data == other
        return super().__eq__(other)

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table as rows of cell values.

        Returns:
            A two-dimensional list whose outer items are rows.
        """
        return self.data

    def get_dim(self) -> TableDim:
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class ImageMetadata:
    """Metadata for an extracted image.

    Provides consistent metadata across all image formats with dict-like
    access for backward compatibility.
    """

    # the number of the unit where this image occurs (1-based for pages/slides)
    # None for formats without pages/slides (e.g. docx, odt, ods, xlsx)
    unit_number: Optional[int] = None
    # A sequential number which shows which nth image this is. The first image has value 1
    image_number: int = 0
    content_type: str = ""
    # Pixel dimensions of the image when available
    width: Optional[int] = None
    height: Optional[int] = None
    # Width-to-height aspect ratio when both dimensions are available
    ratio: Optional[float] = None

    def __post_init__(self) -> None:
        """Derive the width-to-height aspect ratio when dimensions are available."""
        if (
            self.ratio is None
            and self.width is not None
            and self.width > 0
            and self.height is not None
            and self.height > 0
        ):
            self.ratio = self.width / self.height

    def to_dict(self) -> dict:
        """Convert this value to a plain dictionary.

        Returns:
            A dictionary containing the dataclass fields and their values.
        """
        return asdict(self)

    @property
    def unit_index(self) -> Optional[int]:
        """Return the legacy unit_index compatibility alias.

        Returns:
            The corresponding canonical metadata value.
        """
        return self.unit_number

    @unit_index.setter
    def unit_index(self, value: Optional[int]) -> None:
        """Set the legacy unit_index compatibility alias.

        Args:
            value: Replacement value for the compatibility property.

        This updates the corresponding canonical metadata field.
        """
        self.unit_number = value

    @property
    def image_index(self) -> int:
        """Return the legacy image_index compatibility alias.

        Returns:
            The corresponding canonical metadata value.
        """
        return self.image_number

    @image_index.setter
    def image_index(self, value: int) -> None:
        """Set the legacy image_index compatibility alias.

        Args:
            value: Replacement value for the compatibility property.

        This updates the corresponding canonical metadata field.
        """
        self.image_number = value

    # Dict-like access for backward compatibility
    def __getitem__(self, key: str) -> typing.Any:
        """Allow dict-style access for backward compatibility."""
        if hasattr(self, key):
            return getattr(self, key)
        raise KeyError(key)

    def __contains__(self, key: str) -> bool:
        """Allow 'in' operator for backward compatibility."""
        return hasattr(self, key)

    def get(self, key: str, default: typing.Any = None) -> typing.Any:
        """Look up an image metadata field without raising for a missing key.

        Args:
            key: Metadata key to look up.
            default: Value returned when the metadata key is absent.

        Returns:
            The stored value, or the supplied default when the key is absent.
        """
        return getattr(self, key, default)


def _join_unit_text(units: typing.Iterable[UnitRecord]) -> str:
    return ("\n".join(unit.get_text() for unit in units)).strip()


#########
# Email #
#########
@dataclass
class EmailUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a email message."""

    unit_number: int
    body_type: str


@dataclass
class EmailUnitRecord(UnitRecord):
    """Represent one structural text unit from a email message."""

    text: str
    body_type: str = ""  # plain|html|empty

    def get_text(self) -> str:
        """Return the text represented by this email message unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this email message unit.

        Returns:
            A new list containing the unit image objects.
        """
        return []

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this email message unit.

        Returns:
            A new list containing the unit table objects.
        """
        return []

    def get_metadata(self) -> UnitMetadataRecord:
        """Return metadata describing this email message object.

        Returns:
            The format-specific metadata instance.
        """
        return EmailUnitMetadata(unit_number=1, body_type=self.body_type)


@dataclass
class EmailAddress:
    """Represent an email participant with an optional display name and address."""

    name: str = ""
    address: str = ""


@dataclass
class EmailMetadata(SourceRecord):
    """Store metadata extracted from a email message."""

    date: str = ""
    message_id: str = ""


@dataclass
class EmailAttachment:
    """Keep attachment identity, media type, payload, and routing support status together."""

    filename: str
    mime_type: str
    data: io.BytesIO
    is_supported_mime_type: bool = False


@dataclass
class EmailParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a email message."""

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

    def __post_init__(self) -> None:
        self.subject = self.subject.strip()
        self.body_plain = self.body_plain.strip()

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        # ignore_images is a no-op for emails (no images supported)
        """Yield structural text units from this email message.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        if self.body_plain:
            yield EmailUnitRecord(text=self.body_plain, body_type="plain")
            return
        if self.body_html:
            yield EmailUnitRecord(text=self.body_html, body_type="html")
            return
        yield EmailUnitRecord(text="", body_type="empty")

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        # not supported
        """Yield images extracted from this email message.

        Yields:
            Image objects in source order.
        """
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this email message.

        Yields:
            Table objects in source order.
        """
        yield from ()
        return

    def iterate_supported_attachments(
        self,
        *,
        skip_failed: bool = False,
    ) -> typing.Generator[ExtractionRecord, None, None]:
        """Iterate over supported attachments and extract them on demand.

        Args:
            skip_failed: If True, extraction errors for supported attachments are
                logged and skipped. If False (default), supported attachment
                extraction failures raise ``ExtractionFailedError``.
        """
        from sharepoint2text.parsing.exceptions import (
            ExtractionFailedError,
            ExtractionFileEncryptedError,
            ExtractionFileFormatNotSupportedError,
        )
        from sharepoint2text.parsing.mime_types import MIME_TYPE_MAPPING
        from sharepoint2text.parsing.router import _get_extractor

        for attachment in self.attachments:
            if not attachment.is_supported_mime_type:
                logger.debug(
                    "Skipping unsupported attachment: %s (mime=%s)",
                    attachment.filename,
                    attachment.mime_type,
                )
                continue

            try:
                extractor = _get_extractor(attachment.filename)
            except ExtractionFileFormatNotSupportedError:
                file_type = MIME_TYPE_MAPPING.get(attachment.mime_type)
                if not file_type:
                    logger.debug(
                        "Skipping attachment with unknown type: %s (mime=%s)",
                        attachment.filename,
                        attachment.mime_type,
                    )
                    continue
                extractor = _get_extractor(f"attachment.{file_type}")

            attachment.data.seek(0)
            try:
                yield from extractor(attachment.data, attachment.filename)
            except ExtractionFileEncryptedError:
                raise
            except Exception as exc:
                if skip_failed:
                    logger.warning(
                        "Failed to extract attachment: %s (mime=%s) error=%s",
                        attachment.filename,
                        attachment.mime_type,
                        exc,
                    )
                    continue
                raise ExtractionFailedError(
                    "Failed to extract supported attachment: "
                    f"{attachment.filename} (mime={attachment.mime_type})",
                    cause=exc,
                ) from exc
            finally:
                attachment.data.seek(0)

    def get_full_text(self) -> str:
        """Build the default full-text representation of this email message.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> EmailMetadata:
        """Return metadata describing this email message object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata


############
# legacy doc
#############


@dataclass
class DocUnitRecord(UnitRecord):
    """Represent one structural text unit from a legacy Word document."""

    text: str
    unit_number: int = 1
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    images: list[DocImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this legacy Word document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this legacy Word document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this legacy Word document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> DocUnitMeta:
        """Return metadata describing this legacy Word document object.

        Returns:
            The format-specific metadata instance.
        """
        return DocUnitMeta(
            unit_number=self.unit_number,
            location=list(self.location),
            heading_level=self.heading_level,
            heading_path=list(self.heading_path),
        )


@dataclass
class DocUnitMeta(UnitMetadataRecord):
    """Describe the structural position of one unit in a legacy Word document."""

    unit_number: int = 1
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)


@dataclass
class DocMetadata(SourceRecord):
    """Store metadata extracted from a legacy Word document."""

    title: str = ""
    author: str = ""
    subject: str = ""
    keywords: str = ""
    last_saved_by: str = ""
    create_time: str | None = None
    last_saved_time: str | None = None
    num_pages: int = 0
    num_words: int = 0
    num_chars: int = 0


@dataclass
class DocImage(ImageRecord):
    """Represent an image extracted from a legacy Word document."""

    image_index: int
    content_type: str
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None
    caption: str = ""
    description: str = ""
    unit_number: Optional[int] = None

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption.strip()

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this legacy Word document object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.unit_number,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


@dataclass
class DocParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a legacy Word document."""

    main_text: str = ""
    footnotes: str = ""
    headers_footers: str = ""
    annotations: str = ""
    images: List[DocImage] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)
    metadata: DocMetadata = field(default_factory=DocMetadata)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this legacy Word document.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        lines = [line.rstrip() for line in (self.main_text or "").splitlines()]
        if not lines:
            unit_images: list[DocImage] = []
            if not ignore_images:
                unit_images = [
                    DocImage(
                        image_index=image.image_index,
                        content_type=image.content_type,
                        data=image.data,
                        size_bytes=image.size_bytes,
                        width=image.width,
                        height=image.height,
                        caption=image.caption,
                        description=image.description,
                        unit_number=1,
                    )
                    for image in self.images
                ]
            yield DocUnitRecord(text="", unit_number=1, location=[], images=unit_images)
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
            if (
                lowered.startswith("title")
                or lowered.startswith("titel")
                or lowered.endswith(" title")
                or lowered.endswith(" titel")
            ):
                return 0
            if lowered.startswith("sub-section") or lowered.startswith("subsection"):
                return 3
            if lowered.startswith("section"):
                return 2
            if lowered.startswith("chapter") or lowered == "intro":
                return 1
            return None

        units: list[DocUnitRecord] = []
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
                DocUnitRecord(
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
            all_unit_images: list[DocImage] = []
            if not ignore_images:
                all_unit_images = [
                    DocImage(
                        image_index=image.image_index,
                        content_type=image.content_type,
                        data=image.data,
                        size_bytes=image.size_bytes,
                        width=image.width,
                        height=image.height,
                        caption=image.caption,
                        description=image.description,
                        unit_number=1,
                    )
                    for image in self.images
                ]
            yield DocUnitRecord(
                text=self.main_text.strip(),
                unit_number=1,
                location=base_location,
                images=all_unit_images,
                tables=[TableData(data=table) for table in self.tables],
            )
            return

        # Attach unassigned images (no stable anchors in legacy DOC extraction).
        if not ignore_images:
            for image in self.images:
                matched_unit: DocUnitRecord | None = None
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
                        image_index=image.image_index,
                        content_type=image.content_type,
                        data=image.data,
                        size_bytes=image.size_bytes,
                        width=image.width,
                        height=image.height,
                        caption=image.caption,
                        description=image.description,
                        unit_number=matched_unit.unit_number,
                    )
                )

        for unit in units:
            yield unit

    def get_full_text(self) -> str:
        """Build the default full-text representation of this legacy Word document.

        Returns:
            Extracted unit text joined in source order.
        """
        return (
            self.metadata.title + "\n" + _join_unit_text(self.iterate_units())
        ).strip()

    def get_metadata(self) -> SourceRecord:
        """Return metadata describing this legacy Word document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this legacy Word document.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this legacy Word document.

        Yields:
            Table objects in source order.
        """
        for table in self.tables:
            yield TableData(data=table)


##############
# modern docx
###############


@dataclass
class DocxUnitRecord(UnitRecord):
    """Represent one structural text unit from a WordprocessingML document."""

    text: str
    unit_number: int = 1
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    images: list[DocxImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this WordprocessingML document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this WordprocessingML document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this WordprocessingML document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> UnitMetadataRecord:
        """Return metadata describing this WordprocessingML document object.

        Returns:
            The format-specific metadata instance.
        """
        return DocxUnitMetadata(
            unit_number=self.unit_number,
            location=list(self.location),
            heading_level=self.heading_level,
            heading_path=list(self.heading_path),
        )


@dataclass
class DocxUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a WordprocessingML document."""

    unit_number: int
    location: list[str] = field(default_factory=list)
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)


@dataclass
class DocxMetadata(SourceRecord):
    """Store metadata extracted from a WordprocessingML document."""

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
    """Represent one styled run of text from a WordprocessingML paragraph."""

    text: str = ""
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    font_color: Optional[str] = None


@dataclass
class DocxParagraph:
    """Represent a WordprocessingML paragraph and its structural context."""

    text: str = ""
    style: Optional[str] = None
    alignment: Optional[str] = None
    runs: List[DocxRun] = field(default_factory=list)
    has_page_break: bool = False


@dataclass
class DocxHeaderFooter:
    """Represent text extracted from a document header or footer."""

    type: str = ""
    text: str = ""


@dataclass
class DocxImage(ImageRecord):
    """Represent an image extracted from a WordprocessingML document."""

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
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption.strip()

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this WordprocessingML document object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=None,  # DOCX has no page/slide units
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


@dataclass
class DocxHyperlink:
    """Represent visible hyperlink text and its target URL."""

    text: str = ""
    url: str = ""


@dataclass
class DocxNote:
    """Represent a footnote or endnote extracted from a document."""

    id: str = ""
    text: str = ""


@dataclass
class DocxComment:
    """Represent a Word comment and its authoring metadata."""

    id: str = ""
    author: str = ""
    date: str = ""
    text: str = ""


@dataclass
class DocxSection:
    """Describe page and margin settings for one Word document section."""

    page_width_inches: Optional[float] = None
    page_height_inches: Optional[float] = None
    left_margin_inches: Optional[float] = None
    right_margin_inches: Optional[float] = None
    top_margin_inches: Optional[float] = None
    bottom_margin_inches: Optional[float] = None
    orientation: Optional[str] = None


@dataclass
class DocxFormula:
    """Represent an Office Math expression converted to LaTeX."""

    latex: str = ""
    is_display: bool = (
        False  # True for display equations ($$...$$), False for inline ($...$)
    )


@dataclass
class DocxParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a WordprocessingML document."""

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

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this WordprocessingML document.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        heading_re = re.compile(
            r"^(heading|überschrift)\s*(\d+)?\b", flags=re.IGNORECASE
        )
        title_re = re.compile(r"^(title|titel)\b", flags=re.IGNORECASE)

        def heading_level(style: str | None) -> int | None:
            if not style:
                return None
            normalized_style = style.strip()
            if title_re.match(normalized_style):
                return 0

            match = heading_re.match(normalized_style)
            if not match:
                return None
            level_text = match.group(2)
            if level_text is None:
                return 1
            try:
                return int(level_text)
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
        if not ignore_images:
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
        ) -> typing.Iterator[DocxUnitRecord]:
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
                    DocxUnitRecord(
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

        yield DocxUnitRecord(
            text=self.full_text,
            unit_number=1,
            location=[self.metadata.title] if self.metadata.title else [],
            heading_level=None,
            heading_path=[],
            images=[] if ignore_images else list(self.images),
            tables=[TableData(data=table) for table in self.tables],
        )

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this WordprocessingML document.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this WordprocessingML document.

        Yields:
            Table objects in source order.
        """
        for table in self.tables:
            yield TableData(data=table)

    def get_full_text(self) -> str:
        """Build the default full-text representation of this WordprocessingML document.

        Returns:
            Extracted unit text joined in source order.
        """
        return self.full_text

    def get_metadata(self) -> DocxMetadata:
        """Return metadata describing this WordprocessingML document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata


######
# PDF
######


@dataclass
class PdfUnitMetadata(UnitMetadataRecord):
    """PDF unit metadata"""

    unit_number: int

    pass


@dataclass
class PdfUnitRecord(UnitRecord):
    """Represent one structural text unit from a PDF document."""

    page_number: int
    text: str
    images: list[ImageRecord] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this PDF document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this PDF document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this PDF document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> PdfUnitMetadata:
        """Return metadata describing this PDF document object.

        Returns:
            The format-specific metadata instance.
        """
        return PdfUnitMetadata(unit_number=self.page_number)


@dataclass
class PdfImage(ImageRecord):
    """Represent an image extracted from a PDF document."""

    image_index: int = 0
    name: str = ""
    caption: str = ""
    description: str = ""
    width: int = 0
    height: int = 0
    color_space: str = ""
    bits_per_component: int = 8
    filter: str = ""
    data: Optional[io.BytesIO] = None
    format: str = ""
    content_type: str = ""
    unit_number: Optional[int] = None

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption.strip()

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this PDF document object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.get_content_type(),
            unit_number=self.unit_number,
            width=self.width if self.width > 0 else None,
            height=self.height if self.height > 0 else None,
        )


@dataclass
class PdfPage:
    """Keep the text, images, and tables extracted from one PDF page."""

    text: str = ""
    images: List[PdfImage] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)


@dataclass
class PdfMetadata(SourceRecord):
    """Store metadata extracted from a PDF document."""

    total_pages: int = 0


@dataclass
class PdfParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a PDF document."""

    pages: List[PdfPage] = field(default_factory=list)
    metadata: PdfMetadata = field(default_factory=PdfMetadata)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this PDF document.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for page_number, page in enumerate(self.pages, start=1):
            yield PdfUnitRecord(
                page_number=page_number,
                text=page.text,
                images=[] if ignore_images else list(page.images),
                tables=[TableData(data=table) for table in page.tables],
            )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this PDF document.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> PdfMetadata:
        """Return metadata describing this PDF document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this PDF document.

        Yields:
            Image objects in source order.
        """
        for page in self.pages:
            for img in page.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
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


#########
# Plain
#########


@dataclass
class PlainUnitMetadata(UnitMetadataRecord):
    """Plain Unit Metadata"""

    unit_number: int

    pass


@dataclass
class PlainTextUnitRecord(UnitRecord):
    """Represent one structural text unit from a plain-text document."""

    text: str

    def get_text(self) -> str:
        """Return the text represented by this plain-text document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this plain-text document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return []

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this plain-text document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return []

    def get_metadata(self) -> PlainUnitMetadata:
        """Return metadata describing this plain-text document object.

        Returns:
            The format-specific metadata instance.
        """
        return PlainUnitMetadata(unit_number=1)


@dataclass
class PlainTextParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a plain-text document."""

    content: str = ""
    metadata: SourceRecord = field(default_factory=SourceRecord)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        # ignore_images is a no-op for plain text (no images supported)
        """Yield structural text units from this plain-text document.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        yield PlainTextUnitRecord(text=self.content.strip())

    def get_full_text(self) -> str:
        """Build the default full-text representation of this plain-text document.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> SourceRecord:
        """Return metadata describing this plain-text document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this plain-text document.

        Yields:
            Image objects in source order.
        """
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this plain-text document.

        Yields:
            Table objects in source order.
        """
        yield from ()
        return

    def __post_init__(self) -> None:
        self.content = self.content.strip()


############
# CSV / TSV
############


@dataclass
class CsvUnitMetadata(UnitMetadataRecord):
    """CSV/TSV Unit Metadata"""

    unit_number: int

    pass


@dataclass
class CsvUnitRecord(UnitRecord):
    """Represent one structural text unit from a delimited text document."""

    text: str
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this delimited text document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this delimited text document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return []

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this delimited text document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> CsvUnitMetadata:
        """Return metadata describing this delimited text document object.

        Returns:
            The format-specific metadata instance.
        """
        return CsvUnitMetadata(unit_number=1)


@dataclass
class CsvParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a delimited text document."""

    content: str = ""
    table: TableData = field(default_factory=TableData)
    metadata: SourceRecord = field(default_factory=SourceRecord)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this delimited text document.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        tables = [self.table] if self.table.data else []
        yield CsvUnitRecord(text=self.content.strip(), tables=tables)

    def get_full_text(self) -> str:
        """Build the default full-text representation of this delimited text document.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> SourceRecord:
        """Return metadata describing this delimited text document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this delimited text document.

        Yields:
            Image objects in source order.
        """
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this delimited text document.

        Yields:
            Table objects in source order.
        """
        if self.table.data:
            yield self.table


########
# HTML
########


@dataclass
class HtmlUnitMetadata(UnitMetadataRecord):
    """Html Unit Metadata"""

    unit_number: int

    pass


@dataclass
class HtmlUnitRecord(UnitRecord):
    """Represent one structural text unit from a HTML document."""

    text: str

    def get_text(self) -> str:
        """Return the text represented by this HTML document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this HTML document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return []

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this HTML document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return []

    def get_metadata(self) -> HtmlUnitMetadata:
        """Return metadata describing this HTML document object.

        Returns:
            The format-specific metadata instance.
        """
        return HtmlUnitMetadata(unit_number=1)


@dataclass
class HtmlMetadata(SourceRecord):
    """Store metadata extracted from a HTML document."""

    title: str = ""
    language: str = ""
    charset: str = ""
    description: str = ""
    keywords: str = ""
    author: str = ""


@dataclass
class HtmlParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a HTML document."""

    content: str = ""
    tables: List[List[List[str]]] = field(default_factory=list)
    headings: List[Dict[str, str]] = field(
        default_factory=list
    )  # List of {level: "h1", text: "..."}
    links: List[Dict[str, str]] = field(
        default_factory=list
    )  # List of {text: "...", href: "..."}
    metadata: HtmlMetadata = field(default_factory=HtmlMetadata)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        # ignore_images is a no-op for HTML (no images in units)
        """Yield structural text units from this HTML document.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        yield HtmlUnitRecord(text=self.content.strip())

    def get_full_text(self) -> str:
        """Build the default full-text representation of this HTML document.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> HtmlMetadata:
        """Return metadata describing this HTML document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this HTML document.

        Yields:
            Image objects in source order.
        """
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this HTML document.

        Yields:
            Table objects in source order.
        """
        for table in self.tables:
            yield TableData(data=table)


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
class PptUnitMetadata(UnitMetadataRecord):
    """Ppt Unit Metadata"""

    unit_number: int
    title: str = ""

    ...


@dataclass
class PptUnitRecord(UnitRecord):
    """Represents a single legacy PowerPoint slide as an extraction unit."""

    slide_number: int
    text: str
    title: str = ""
    images: list["PptImage"] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this legacy PowerPoint deck unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this legacy PowerPoint deck unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this legacy PowerPoint deck unit.

        Returns:
            A new list containing the unit table objects.
        """
        return []

    def get_metadata(self) -> PptUnitMetadata:
        """Return metadata describing this legacy PowerPoint deck object.

        Returns:
            The format-specific metadata instance.
        """
        return PptUnitMetadata(unit_number=self.slide_number, title=self.title)


@dataclass
class PptImage(ImageRecord):
    """Represents an embedded image in a legacy PPT file."""

    image_index: int = 0
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None
    caption: str = ""
    description: str = ""
    slide_number: int = 0

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption.strip()

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this legacy PowerPoint deck object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.slide_number if self.slide_number > 0 else None,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


@dataclass
class PptMetadata(SourceRecord):
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
        """Return the normalized name of this text-block type.

        Returns:
            A stable string suitable for serialization and diagnostics.
        """
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
        if self.text_type is None:
            return "unknown"
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
        """Combine the primary text fragments for this slide.

        Returns:
            Slide text assembled in extraction order.
        """
        parts = []
        if self.title:
            parts.append(self.title)
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)

    @property
    def unit_text(self) -> str:
        """Build the text exposed when this slide is used as a unit.

        Returns:
            Slide body text with supported annotations included.
        """
        parts = []
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)

    def to_dict(self) -> dict[str, typing.Any]:
        """Convert this value to a plain dictionary.

        Returns:
            A dictionary containing the dataclass fields and their values.
        """
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
class PptParserOutput(ExtractionRecord):
    """Complete extracted content from a PPT file."""

    metadata: PptMetadata = field(default_factory=PptMetadata)
    slides: list[PptSlideContent] = field(default_factory=list)
    master_text: list[str] = field(default_factory=list)  # Text from master slides
    all_text: list[str] = field(default_factory=list)
    streams: list[list[str]] = field(default_factory=list)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this legacy PowerPoint deck.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for slide in self.slides:
            yield PptUnitRecord(
                slide_number=slide.slide_number,
                text=slide.unit_text,
                title=slide.title or "",
                images=[] if ignore_images else list(slide.images),
            )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this legacy PowerPoint deck.

        Returns:
            Extracted unit text joined in source order.
        """
        texts = [slide.text_combined.strip() for slide in self.slides]
        return "\n".join(text for text in texts if text)

    def get_metadata(self) -> PptMetadata:
        """Return metadata describing this legacy PowerPoint deck object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    @property
    def slide_count(self) -> int:
        """Return the number of slides in the presentation.

        Returns:
            Count of extracted slide objects.
        """
        return len(self.slides)

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this legacy PowerPoint deck.

        Yields:
            Image objects in source order.
        """
        for slide in self.slides:
            for img in slide.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this legacy PowerPoint deck.

        Yields:
            Table objects in source order.
        """
        yield from ()
        return


##############
# Modern PPTX
##############


@dataclass
class PptxUnitMetadata(UnitMetadataRecord):
    """Pptx Unit Metadata"""

    unit_number: int
    title: str = ""


@dataclass
class PptxUnitRecord(UnitRecord):
    """Represent one structural text unit from a PresentationML deck."""

    slide_number: int
    text: str
    title: str = ""
    images: list[PptxImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this PresentationML deck unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this PresentationML deck unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this PresentationML deck unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> PptxUnitMetadata:
        """Return metadata describing this PresentationML deck object.

        Returns:
            The format-specific metadata instance.
        """
        return PptxUnitMetadata(unit_number=self.slide_number, title=self.title)


@dataclass
class PptxMetadata(SourceRecord):
    """Store metadata extracted from a PresentationML deck."""

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
class PptxImage(ImageRecord):
    """Represent an image extracted from a PresentationML deck."""

    image_index: int = 0
    filename: str = ""
    content_type: str = ""
    size_bytes: int = 0
    data: Optional[io.BytesIO] = None
    width: Optional[int] = None
    height: Optional[int] = None
    caption: str = ""  # Title/name of the image shape
    description: str = ""  # Alt text / description for accessibility
    slide_number: int = 0

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this PresentationML deck object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.slide_number,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description


@dataclass
class PptxFormula:
    """Represent a slide formula converted from Office Math to LaTeX."""

    latex: str = ""
    is_display: bool = False  # True for display equations, False for inline


@dataclass
class PptxComment:
    """Represent a presentation comment and its authoring information."""

    author: str = ""
    text: str = ""
    date: str = ""


@dataclass
class PptxSlide:
    """Aggregate extracted content and optional annotations for one presentation slide."""

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
        """Return the text represented by this PresentationML deck unit.

        Args:
            include_image_captions: Append available image descriptions when true.

        Returns:
            Extracted text in reading order.
        """
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
class PptxParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a PresentationML deck."""

    metadata: PptxMetadata = field(default_factory=PptxMetadata)
    slides: List[PptxSlide] = field(default_factory=list)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this PresentationML deck.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for slide in self.slides:
            yield PptxUnitRecord(
                slide_number=slide.slide_number,
                title=slide.title,
                images=[] if ignore_images else list(slide.images),
                tables=[TableData(data=table) for table in slide.tables],
                text=slide.get_text().strip(),
            )

    def get_full_text(
        self,
        include_image_captions: bool = False,
    ) -> str:
        """Get full text of all slides.

        Args:
            include_image_captions: Include image captions/alt text in output (default: False)
        """
        return (
            "\n".join(
                slide.get_text(
                    include_image_captions=include_image_captions,
                ).strip()
                for slide in self.slides
            )
        ).strip()

    def get_metadata(self) -> PptxMetadata:
        """Return metadata describing this PresentationML deck object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this PresentationML deck.

        Yields:
            Image objects in source order.
        """
        for slide in self.slides:
            for img in slide.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this PresentationML deck.

        Yields:
            Table objects in source order.
        """
        for slide in self.slides:
            for table in slide.tables:
                yield TableData(data=table)


#############
# Legacy XLS
#############


@dataclass
class XlsUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a legacy Excel workbook."""

    unit_number: int
    sheet_name: str


@dataclass
class XlsUnitRecord(UnitRecord):
    """Represent one structural text unit from a legacy Excel workbook."""

    sheet_number: int
    sheet_name: str
    text: str
    tables: list[TableData] = field(default_factory=list)
    images: list[XlsImage] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this legacy Excel workbook unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this legacy Excel workbook unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this legacy Excel workbook unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> XlsUnitMetadata:
        """Return metadata describing this legacy Excel workbook object.

        Returns:
            The format-specific metadata instance.
        """
        return XlsUnitMetadata(
            unit_number=self.sheet_number, sheet_name=self.sheet_name
        )


@dataclass
class XlsImage(ImageRecord):
    """Represents an embedded image in a legacy XLS file."""

    image_index: int = 0
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None
    caption: str = ""
    description: str = ""

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption.strip()

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this legacy Excel workbook object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=None,  # XLS images are workbook-level, not sheet-level
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


@dataclass
class XlsMetadata(SourceRecord):
    """Store metadata extracted from a legacy Excel workbook."""

    title: str = ""
    author: str = ""
    subject: str = ""
    company: str = ""
    last_saved_by: str = ""
    created: str = ""
    modified: str = ""


@dataclass
class XlsSheet(TableRecord):
    """Represent one worksheet and its tabular content in a legacy Excel workbook."""

    name: str = ""
    data: List[Dict[str, typing.Any]] = field(default_factory=list)
    text: str = ""

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table as rows of cell values.

        Returns:
            A two-dimensional list whose outer items are rows.
        """
        if not self.data:
            return []
        headers = list(self.data[0].keys())
        rows = [headers]
        for row in self.data:
            rows.append([row.get(header, "") for header in headers])
        return rows

    def get_dim(self) -> TableDim:
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
        table = self.get_table()
        rows = len(table)
        columns = max((len(row) for row in table), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class XlsParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a legacy Excel workbook."""

    metadata: XlsMetadata = field(default_factory=XlsMetadata)
    sheets: List[XlsSheet] = field(default_factory=list)
    images: List[XlsImage] = field(default_factory=list)
    full_text: str = ""

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this legacy Excel workbook.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
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
            yield XlsUnitRecord(
                sheet_number=sheet_index,
                sheet_name=sheet.name,
                tables=[TableData(data=normalized_table)] if normalized_table else [],
                images=(
                    (list(self.images) if sheet_index == 1 else [])
                    if not ignore_images
                    else []
                ),
                text=sheet.text.strip(),
            )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this legacy Excel workbook.

        Returns:
            Extracted unit text joined in source order.
        """
        return self.full_text.strip()

    def get_metadata(self) -> XlsMetadata:
        """Return metadata describing this legacy Excel workbook object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this legacy Excel workbook.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this legacy Excel workbook.

        Yields:
            Table objects in source order.
        """
        for sheet in self.sheets:
            yield sheet


##############
# Modern XLSX
##############


@dataclass
class XlsxUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a modern Excel workbook."""

    unit_number: int
    sheet_number: int
    sheet_name: str


@dataclass
class XlsxUnitRecord(UnitRecord):
    """Represent one structural text unit from a modern Excel workbook."""

    sheet_index: int
    sheet_name: str
    text: str
    images: list[XlsxImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this modern Excel workbook unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this modern Excel workbook unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this modern Excel workbook unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> XlsxUnitMetadata:
        """Return metadata describing this modern Excel workbook object.

        Returns:
            The format-specific metadata instance.
        """
        return XlsxUnitMetadata(
            unit_number=self.sheet_index,
            sheet_name=self.sheet_name,
            sheet_number=self.sheet_index,
        )


@dataclass
class XlsxMetadata(SourceRecord):
    """Store metadata extracted from a modern Excel workbook."""

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
class XlsxImage(ImageRecord):
    """Represent an image extracted from a modern Excel workbook."""

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
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this modern Excel workbook object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.sheet_index + 1 if self.sheet_index >= 0 else None,
            width=self.width if self.width > 0 else None,
            height=self.height if self.height > 0 else None,
        )


@dataclass
class XlsxSheet(TableRecord):
    """Represent one worksheet and its tabular content in a modern Excel workbook."""

    name: str = ""
    data: List[List[typing.Any]] = field(default_factory=list)
    text: str = ""
    images: List[XlsxImage] = field(default_factory=list)

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table as rows of cell values.

        Returns:
            A two-dimensional list whose outer items are rows.
        """
        return self.data

    def get_dim(self) -> TableDim:
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class XlsxParserOutput(ExtractionRecord):
    """Aggregate the structured extraction result for a modern Excel workbook."""

    metadata: XlsxMetadata = field(default_factory=XlsxMetadata)
    sheets: List[XlsxSheet] = field(default_factory=list)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this modern Excel workbook.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for sheet_index, sheet in enumerate(self.sheets, start=1):
            yield XlsxUnitRecord(
                sheet_index=sheet_index,
                sheet_name=sheet.name,
                images=[] if ignore_images else list(sheet.images),
                tables=[TableData(data=sheet.data)] if sheet.data else [],
                text=sheet.name + "\n" + sheet.text.strip(),
            )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this modern Excel workbook.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> XlsxMetadata:
        """Return metadata describing this modern Excel workbook object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this modern Excel workbook.

        Yields:
            Image objects in source order.
        """
        for sheet in self.sheets:
            for img in sheet.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this modern Excel workbook.

        Yields:
            Table objects in source order.
        """
        for sheet in self.sheets:
            yield sheet


#############################
# OpenDocument Shared Types #
#############################


@dataclass
class OpenDocumentMetadata(SourceRecord):
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
class OpenDocumentImage(ImageRecord):
    """
    Represents an embedded image in an OpenDocument file.

    Images are stored in the Pictures/ directory within the ODF archive
    and referenced via href attributes in the content.xml.

    Implements ImageRecord for consistent image handling across formats.
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
    unit_number: Optional[int] = None  # Page/slide number (None for ODT/ODS)

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this OpenDocument file object.

        Returns:
            The format-specific metadata instance.
        """
        width_px = _odf_length_to_px(self.width)
        height_px = _odf_length_to_px(self.height)
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.unit_number,
            width=width_px if width_px and width_px > 0 else None,
            height=height_px if height_px and height_px > 0 else None,
        )


###############
# OpenDocument ODG (Drawing) #
###############


@dataclass
class OdgUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a OpenDocument drawing."""

    unit_number: int


@dataclass
class OdgUnitRecord(UnitRecord):
    """Represent one structural text unit from a OpenDocument drawing."""

    text: str
    images: list[OpenDocumentImage] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this OpenDocument drawing unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this OpenDocument drawing unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this OpenDocument drawing unit.

        Returns:
            A new list containing the unit table objects.
        """
        return []

    def get_metadata(self) -> OdgUnitMetadata:
        """Return metadata describing this OpenDocument drawing object.

        Returns:
            The format-specific metadata instance.
        """
        return OdgUnitMetadata(unit_number=1)


@dataclass
class OdgParserOutput(ExtractionRecord):
    """Complete extracted content from an ODG file."""

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    full_text: str = ""
    images: list[OpenDocumentImage] = field(default_factory=list)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this OpenDocument drawing.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        yield OdgUnitRecord(
            text=self.full_text.strip(),
            images=[] if ignore_images else list(self.images),
        )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this OpenDocument drawing.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> OpenDocumentMetadata:
        """Return metadata describing this OpenDocument drawing object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this OpenDocument drawing.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this OpenDocument drawing.

        Yields:
            Table objects in source order.
        """
        yield from ()
        return


###############
# OpenDocument ODF (Formula) #
###############


@dataclass
class OdfUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a OpenDocument formula."""

    unit_number: int


@dataclass
class OdfUnitRecord(UnitRecord):
    """Represent one structural text unit from a OpenDocument formula."""

    text: str

    def get_text(self) -> str:
        """Return the text represented by this OpenDocument formula unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this OpenDocument formula unit.

        Returns:
            A new list containing the unit image objects.
        """
        return []

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this OpenDocument formula unit.

        Returns:
            A new list containing the unit table objects.
        """
        return []

    def get_metadata(self) -> OdfUnitMetadata:
        """Return metadata describing this OpenDocument formula object.

        Returns:
            The format-specific metadata instance.
        """
        return OdfUnitMetadata(unit_number=1)


@dataclass
class OdfParserOutput(ExtractionRecord):
    """Complete extracted content from an ODF file."""

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    full_text: str = ""

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        # ignore_images is a no-op for ODF (no images supported)
        """Yield structural text units from this OpenDocument formula.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        yield OdfUnitRecord(text=self.full_text.strip())

    def get_full_text(self) -> str:
        """Build the default full-text representation of this OpenDocument formula.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> OpenDocumentMetadata:
        """Return metadata describing this OpenDocument formula object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this OpenDocument formula.

        Yields:
            Image objects in source order.
        """
        yield from ()
        return

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this OpenDocument formula.

        Yields:
            Table objects in source order.
        """
        yield from ()
        return


###############
# OpenDocument ODP (Presentation)
###############


@dataclass
class OdpUnitRecord(UnitRecord):
    """Represents a single OpenDocument presentation slide as a unit."""

    slide_number: int
    text: str
    title: str = ""
    location: list[str] = field(default_factory=list)
    images: list[OpenDocumentImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this OpenDocument presentation unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this OpenDocument presentation unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this OpenDocument presentation unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> OdpUnitMetadata:
        """Return metadata describing this OpenDocument presentation object.

        Returns:
            The format-specific metadata instance.
        """
        return OdpUnitMetadata(
            unit_number=self.slide_number,
            title=self.title,
            location=list(self.location),
            slide_number=self.slide_number,
        )


@dataclass
class OdpUnitMetadata(UnitMetadataRecord):
    """Metadata for a single OpenDocument presentation unit."""

    unit_number: int
    title: str = ""
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
        """Combine the primary text fragments for this slide.

        Returns:
            Slide text assembled in extraction order.
        """
        parts = []
        if self.title:
            parts.append(self.title)
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)

    @property
    def unit_text(self) -> str:
        """Build the text exposed when this slide is used as a unit.

        Returns:
            Slide body text with supported annotations included.
        """
        parts = []
        parts.extend(self.body_text)
        parts.extend(self.other_text)
        return "\n".join(parts)


@dataclass
class OdpParserOutput(ExtractionRecord):
    """Complete extracted content from an ODP file."""

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    slides: List[OdpSlide] = field(default_factory=list)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this OpenDocument presentation.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for slide in self.slides:
            yield OdpUnitRecord(
                slide_number=slide.slide_number,
                text=slide.unit_text,
                title=slide.title,
                location=[slide.title] if slide.title else [],
                images=[] if ignore_images else list(slide.images),
                tables=[TableData(data=table) for table in slide.tables],
            )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this OpenDocument presentation.

        Returns:
            Extracted unit text joined in source order.
        """
        texts = [slide.text_combined.strip() for slide in self.slides]
        return "\n".join(text for text in texts if text)

    def get_metadata(self) -> OpenDocumentMetadata:
        """Return metadata describing this OpenDocument presentation object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    @property
    def slide_count(self) -> int:
        """Return the number of slides in the presentation.

        Returns:
            Count of extracted slide objects.
        """
        return len(self.slides)

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this OpenDocument presentation.

        Yields:
            Image objects in source order.
        """
        for slides in self.slides:
            for img in slides.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this OpenDocument presentation.

        Yields:
            Table objects in source order.
        """
        for slide in self.slides:
            for table in slide.tables:
                yield TableData(data=table)


##################################
# OpenDocument ODS (Spreadsheet) #
##################################


@dataclass
class OdsUnitRecord(UnitRecord):
    """Represent one structural text unit from a OpenDocument spreadsheet."""

    sheet_number: int
    sheet_name: str
    text: str
    images: list[OpenDocumentImage] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this OpenDocument spreadsheet unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this OpenDocument spreadsheet unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this OpenDocument spreadsheet unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> OdsUnitMetadata:
        """Return metadata describing this OpenDocument spreadsheet object.

        Returns:
            The format-specific metadata instance.
        """
        return OdsUnitMetadata(
            unit_number=self.sheet_number,
            sheet_number=self.sheet_number,
            sheet_name=self.sheet_name,
        )


@dataclass
class OdsUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a OpenDocument spreadsheet."""

    unit_number: int
    sheet_number: int
    sheet_name: str


@dataclass
class OdsSheet(TableRecord):
    """Represents a single sheet in the spreadsheet."""

    name: str = ""
    data: List[List[typing.Any]] = field(default_factory=list)
    text: str = ""
    annotations: List[OpenDocumentAnnotation] = field(default_factory=list)
    images: List[OpenDocumentImage] = field(default_factory=list)

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table as rows of cell values.

        Returns:
            A two-dimensional list whose outer items are rows.
        """
        return self.data

    def get_dim(self) -> TableDim:
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class OdsParserOutput(ExtractionRecord):
    """Complete extracted content from an ODS file."""

    metadata: OpenDocumentMetadata = field(default_factory=OpenDocumentMetadata)
    sheets: List[OdsSheet] = field(default_factory=list)

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this OpenDocument spreadsheet.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for sheet_index, sheet in enumerate(self.sheets, start=1):
            yield OdsUnitRecord(
                sheet_number=sheet_index,
                sheet_name=sheet.name,
                images=[] if ignore_images else list(sheet.images),
                tables=[TableData(data=sheet.data)] if sheet.data else [],
                text=(sheet.name + "\n" + sheet.text.strip()).strip(),
            )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this OpenDocument spreadsheet.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> OpenDocumentMetadata:
        """Return metadata describing this OpenDocument spreadsheet object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    @property
    def sheet_count(self) -> int:
        """Return the number of worksheets in the workbook.

        Returns:
            Count of extracted sheet objects.
        """
        return len(self.sheets)

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this OpenDocument spreadsheet.

        Yields:
            Image objects in source order.
        """
        for sheet in self.sheets:
            for img in sheet.images:
                yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this OpenDocument spreadsheet.

        Yields:
            Table objects in source order.
        """
        for sheet in self.sheets:
            yield sheet


####################################
# OpenDocument ODT (Text Document) #
####################################


@dataclass
class OdtUnitRecord(UnitRecord):
    """Represent one structural text unit from a OpenDocument text document."""

    text: str
    unit_number: int
    heading_level: int | None = None
    heading_path: list[str] = field(default_factory=list)
    kind: str = "body"  # body|annotation
    annotation_creator: str | None = None
    annotation_date: str | None = None
    images: list[ImageRecord] = field(default_factory=list)
    tables: list[TableData] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this OpenDocument text document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this OpenDocument text document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this OpenDocument text document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return list(self.tables)

    def get_metadata(self) -> OdtUnitMetadata:
        """Return metadata describing this OpenDocument text document object.

        Returns:
            The format-specific metadata instance.
        """
        return OdtUnitMetadata(
            unit_number=self.unit_number,
            heading_level=self.heading_level,
            heading_path=list(self.heading_path),
            kind=self.kind,
            annotation_creator=self.annotation_creator,
            annotation_date=self.annotation_date,
        )


@dataclass
class OdtUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a OpenDocument text document."""

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
class OdtTable(TableRecord):
    """Represents a single table in the document."""

    data: List[List[str]] = field(default_factory=list)

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table as rows of cell values.

        Returns:
            A two-dimensional list whose outer items are rows.
        """
        return self.data

    def get_dim(self) -> TableDim:
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
        rows = len(self.data)
        columns = max((len(row) for row in self.data), default=0)
        return TableDim(rows=rows, columns=columns)


@dataclass
class OdtParserOutput(ExtractionRecord):
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

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Iterate over heading-based units.

        Units are built from paragraph runs separated by headings (paragraphs with
        an outline level). Heading text itself becomes part of the unit heading
        path and is not included in the unit body text.

        Args:
            ignore_images: Omit document images from the yielded units when true.

        Yields:
            ``OdtUnitRecord`` objects in document order, with heading context attached.
        """
        base_heading_path = [self.metadata.title] if self.metadata.title else []
        units: list[OdtUnitRecord] = []
        title_style_prefixes = ("title", "titel")

        if not self.paragraphs:
            heading_path = list(base_heading_path)
            units.append(
                OdtUnitRecord(
                    text=self.full_text,
                    kind="body",
                    unit_number=1,
                    heading_level=1 if heading_path else None,
                    heading_path=heading_path,
                    images=[] if ignore_images else list(self.images),
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
                OdtUnitRecord(
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
            style_name = (paragraph.style_name or "").strip().lower()
            if heading_level is None and any(
                style_name.startswith(prefix) for prefix in title_style_prefixes
            ):
                heading_level = 0
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
                OdtUnitRecord(
                    text=self.full_text,
                    kind="body",
                    unit_number=1,
                    heading_level=1 if heading_path else None,
                    heading_path=heading_path,
                    images=[] if ignore_images else list(self.images),
                    tables=[TableData(data=table.data) for table in self.tables],
                )
            ]
            if not ignore_images:
                for image in self.images:
                    image.unit_number = 1
            for unit in units:
                yield unit
            return

        # Best-effort mapping of unassigned tables/images to units.
        # (ODT extraction does not currently provide stable positional anchors.)
        if units and not ignore_images:
            if table_index < len(self.tables):
                remaining_tables = self.tables[table_index:]
                for table in remaining_tables:
                    table_data = TableData(data=table.data)
                    header_tokens = [
                        str(cell).strip()
                        for cell in (table.data[0] if table.data else [])
                        if str(cell).strip()
                    ]
                    matched_unit: OdtUnitRecord | None = None
                    if header_tokens:
                        for unit in units:
                            if all(token in unit.text for token in header_tokens):
                                matched_unit = unit
                                break
                    (matched_unit or units[-1]).tables.append(table_data)

            for image in self.images:
                matched_image_unit: OdtUnitRecord | None = None
                for unit in units:
                    if image.caption and image.caption in unit.text:
                        matched_image_unit = unit
                        break
                    if image.description and image.description in unit.text:
                        matched_image_unit = unit
                        break
                if matched_image_unit is None:
                    if len(units) == 1:
                        matched_image_unit = units[0]
                    else:
                        matched_image_unit = next(
                            (
                                u
                                for u in reversed(units)
                                if u.heading_level == 1 or u.heading_level is None
                            ),
                            units[-1],
                        )

                image.unit_number = matched_image_unit.unit_number
                matched_image_unit.images.append(image)

        for unit in units:
            yield unit

    def get_full_text(self) -> str:
        """Build the default full-text representation of this OpenDocument text document.

        Returns:
            Extracted unit text joined in source order.
        """
        return self.full_text

    def get_metadata(self) -> OpenDocumentMetadata:
        """Return metadata describing this OpenDocument text document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this OpenDocument text document.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this OpenDocument text document.

        Yields:
            Table objects in source order.
        """
        for table in self.tables:
            yield table


#######
# RTF #
#######


@dataclass
class RtfUnitMetadata(UnitMetadataRecord):
    """Describe the structural position of one unit in a Rich Text Format document."""

    unit_number: int
    page_number: int


@dataclass
class RtfUnitRecord(UnitRecord):
    """Represent one structural text unit from a Rich Text Format document."""

    page_number: int
    text: str
    images: List[RtfImage] = field(default_factory=list)
    tables: List[RtfTable] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this Rich Text Format document unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this Rich Text Format document unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this Rich Text Format document unit.

        Returns:
            A new list containing the unit table objects.
        """
        return [TableData(data=t.data) for t in self.tables]

    def get_metadata(self) -> RtfUnitMetadata:
        """Return metadata describing this Rich Text Format document object.

        Returns:
            The format-specific metadata instance.
        """
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
        """Format this color as a hexadecimal RGB value.

        Returns:
            A six-digit uppercase RGB string prefixed with a hash sign.
        """
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
class RtfMetadata(SourceRecord):
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
class RtfImage(ImageRecord):
    """Represents an embedded image in an RTF document."""

    image_type: str = ""  # png, jpeg, emf, wmf
    width: int = 0  # in twips (1/1440 inch)
    height: int = 0  # in twips
    data: Optional[io.BytesIO] = None  # Binary image data
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
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self._CONTENT_TYPES.get(
            self.image_type.lower(), "application/octet-stream"
        )

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return self.caption.strip()

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return self.description.strip()

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this Rich Text Format document object.

        Returns:
            The format-specific metadata instance.
        """
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
class RtfTable(TableRecord):
    """Represents a table extracted from an RTF document."""

    data: List[List[str]] = field(default_factory=list)
    table_index: int = 0  # Sequential index of the table (1-based)
    page_number: Optional[int] = None  # Page where table appears (if known)

    def get_table(self) -> list[list[typing.Any]]:
        """Return the table as rows of cell values.

        Returns:
            A two-dimensional list whose outer items are rows.
        """
        return self.data

    def get_dim(self) -> TableDim:
        """Calculate the dimensions of the table.

        Returns:
            Row and column counts for the current table data.
        """
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
class RtfParserOutput(ExtractionRecord):
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

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Iterate over pages, yielding text per page.

        RTF documents are split on explicit page breaks (\\page).
        If no page breaks exist, yields the full document as a single unit.
        Images and tables are distributed to units based on their page_number.

        Args:
            ignore_images: Omit page images from the yielded units when true.

        Yields:
            ``RtfUnitRecord`` objects for each non-empty explicit or inferred page.
        """
        # Group images and tables by page number
        images_by_page: dict[int, List[RtfImage]] = {}
        if not ignore_images:
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
            for page_number, page_text in enumerate(self.pages, start=1):
                if page_text.strip():
                    yield RtfUnitRecord(
                        page_number=page_number,
                        text=page_text,
                        images=images_by_page.get(page_number, []),
                        tables=tables_by_page.get(page_number, []),
                    )
        elif self.full_text:
            yield RtfUnitRecord(
                page_number=1,
                text=self.full_text,
                images=images_by_page.get(1, []),
                tables=tables_by_page.get(1, []),
            )
        else:
            # Fallback: combine all paragraphs
            combined = "\n".join(p.text for p in self.paragraphs if p.text.strip())
            if combined:
                yield RtfUnitRecord(
                    page_number=1,
                    text=combined,
                    images=images_by_page.get(1, []),
                    tables=tables_by_page.get(1, []),
                )

    def get_full_text(self) -> str:
        """Build the default full-text representation of this Rich Text Format document.

        Returns:
            Extracted unit text joined in source order.
        """
        if self.full_text:
            return self.full_text
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> RtfMetadata:
        """Return metadata describing this Rich Text Format document object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this Rich Text Format document.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this Rich Text Format document.

        Yields:
            Table objects in source order.
        """
        for tbl in self.tables:
            yield tbl


########
# EPUB #
########


@dataclass
class EpubUnitMetadata(UnitMetadataRecord):
    """EPUB Unit Metadata - represents a chapter/content document."""

    unit_number: int = 1
    href: str = ""  # Path within the EPUB
    title: str = ""  # Chapter/section title if available


@dataclass
class EpubChapter(UnitRecord):
    """A single chapter or content document from an EPUB."""

    chapter_number: int = 1
    href: str = ""  # Path within the EPUB (e.g., "OEBPS/chapter1.xhtml")
    title: str = ""  # Title from spine/manifest or extracted from content
    text: str = ""  # Extracted text content
    images: List["EpubImage"] = field(default_factory=list)
    tables: List[List[List[str]]] = field(default_factory=list)

    def get_text(self) -> str:
        """Return the text represented by this EPUB publication unit.

        Returns:
            Extracted text in reading order.
        """
        return self.text

    def get_images(self) -> list[ImageRecord]:
        """Return images associated with this EPUB publication unit.

        Returns:
            A new list containing the unit image objects.
        """
        return list(self.images)

    def get_tables(self) -> list[TableData]:
        """Return tables associated with this EPUB publication unit.

        Returns:
            A new list containing the unit table objects.
        """
        return [TableData(data=t) for t in self.tables]

    def get_metadata(self) -> EpubUnitMetadata:
        """Return metadata describing this EPUB publication object.

        Returns:
            The format-specific metadata instance.
        """
        return EpubUnitMetadata(
            unit_number=self.chapter_number,
            href=self.href,
            title=self.title,
        )


@dataclass
class EpubImage(ImageRecord):
    """Image embedded in an EPUB file."""

    image_index: int = 0
    href: str = ""  # Path within the EPUB (e.g., "OEBPS/images/cover.jpg")
    content_type: str = ""
    data: Optional[io.BytesIO] = None
    size_bytes: int = 0
    width: Optional[int] = None
    height: Optional[int] = None
    unit_number: Optional[int] = None  # Chapter number where image is referenced

    def get_bytes(self) -> io.BytesIO:
        """Return a readable buffer containing the image payload.

        Returns:
            The image data positioned for reading by the caller.
        """
        if self.data is None:
            return io.BytesIO()
        self.data.seek(0)
        return self.data

    def get_content_type(self) -> str:
        """Return the media type reported for the image.

        Returns:
            A MIME-style content type, or an empty string when unknown.
        """
        return self.content_type.strip()

    def get_caption(self) -> str:
        """Return the human-readable caption associated with the image.

        Returns:
            Caption text, or an empty string when none was extracted.
        """
        return ""

    def get_description(self) -> str:
        """Return accessibility or descriptive text for the image.

        Returns:
            Image description text, or an empty string when unavailable.
        """
        return ""

    def get_metadata(self) -> ImageMetadata:
        """Return metadata describing this EPUB publication object.

        Returns:
            The format-specific metadata instance.
        """
        return ImageMetadata(
            image_number=self.image_index,
            content_type=self.content_type,
            unit_number=self.unit_number,
            width=self.width if self.width is not None and self.width > 0 else None,
            height=self.height if self.height is not None and self.height > 0 else None,
        )


@dataclass
class EpubMetadata(SourceRecord):
    """Metadata from an EPUB file (Dublin Core + EPUB-specific)."""

    # Dublin Core metadata
    title: str = ""
    creator: str = ""  # Author
    language: str = ""
    identifier: str = ""  # ISBN, UUID, or other unique identifier
    publisher: str = ""
    date: str = ""  # Publication date
    description: str = ""
    subject: str = ""  # Keywords/categories
    rights: str = ""  # Copyright info
    contributor: str = ""

    # EPUB-specific
    epub_version: str = ""  # EPUB 2.0, 3.0, etc.


@dataclass
class EpubParserOutput(ExtractionRecord):
    """Complete extracted content from an EPUB file."""

    metadata: EpubMetadata = field(default_factory=EpubMetadata)
    chapters: List[EpubChapter] = field(default_factory=list)
    images: List[EpubImage] = field(default_factory=list)
    toc: List[Dict[str, str]] = field(
        default_factory=list
    )  # Table of contents: [{title, href}, ...]

    def iterate_units(
        self, *, ignore_images: bool = False
    ) -> typing.Iterator[UnitRecord]:
        """Yield structural text units from this EPUB publication.

        Args:
            ignore_images: Exclude image objects from yielded units when true.

        Yields:
            Units in source reading order.
        """
        for chapter in self.chapters:
            if ignore_images:
                # Yield a copy with empty images
                yield EpubChapter(
                    chapter_number=chapter.chapter_number,
                    href=chapter.href,
                    title=chapter.title,
                    text=chapter.text,
                    images=[],
                    tables=chapter.tables,
                )
            else:
                yield chapter

    def iterate_images(self) -> typing.Generator[ImageRecord, None, None]:
        """Yield images extracted from this EPUB publication.

        Yields:
            Image objects in source order.
        """
        for img in self.images:
            yield img

    def iterate_tables(self) -> typing.Generator[TableRecord, None, None]:
        """Yield tables extracted from this EPUB publication.

        Yields:
            Table objects in source order.
        """
        for chapter in self.chapters:
            for table in chapter.tables:
                yield TableData(data=table)

    def get_full_text(self) -> str:
        """Build the default full-text representation of this EPUB publication.

        Returns:
            Extracted unit text joined in source order.
        """
        return _join_unit_text(self.iterate_units())

    def get_metadata(self) -> EpubMetadata:
        """Return metadata describing this EPUB publication object.

        Returns:
            The format-specific metadata instance.
        """
        return self.metadata
