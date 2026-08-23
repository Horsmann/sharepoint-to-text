"""Stable, format-neutral extraction data model."""

from __future__ import annotations

from dataclasses import dataclass, field
from math import isfinite
from typing import Iterator, Literal, TypeVar

JsonScalar = str | int | float | bool | None
JsonValue = JsonScalar | list["JsonValue"] | dict[str, "JsonValue"]
CellValue = str | int | float | bool | None
UnitKind = Literal[
    "document",
    "section",
    "page",
    "slide",
    "sheet",
    "chapter",
    "message",
]
_Record = TypeVar("_Record")


def _normalize_document_records(
    unit_records: list[list[_Record]], document_records: list[_Record]
) -> list[_Record]:
    """Move legacy document records to a unit and return the unit aggregate.

    Args:
        unit_records: Per-unit record collections in document order.
        document_records: Records supplied through the document convenience field.

    Returns:
        Every unit-owned record in document order, preserving object identity.
    """
    aggregate = [record for records in unit_records for record in records]
    if document_records != aggregate:
        owned_record_ids = {id(record) for record in aggregate}
        unit_records[-1].extend(
            record for record in document_records if id(record) not in owned_record_ids
        )
        aggregate = [record for records in unit_records for record in records]
    return aggregate


@dataclass(slots=True)
class SourceMetadata:
    """Identify the input from which content was extracted.

    Attributes:
        filename: Basename of the source, when known.
        extension: File extension, including its leading dot when available.
        path: Source path supplied to the extractor.
        folder: Folder containing the source.
        media_type: MIME media type of the source.
        encoding: Detected character encoding.
    """

    filename: str | None = None
    extension: str | None = None
    path: str | None = None
    folder: str | None = None
    media_type: str | None = None
    encoding: str | None = None


@dataclass(slots=True)
class DocumentMetadata:
    """Store common descriptive metadata and namespaced format properties.

    Attributes:
        title: Human-readable document title.
        author: Primary author or creator.
        subject: Subject or summary text.
        keywords: Ordered descriptive keywords.
        language: Document language identifier.
        created: Source-provided creation timestamp.
        modified: Source-provided modification timestamp.
        properties: Small format-specific JSON values under namespaced keys.
    """

    title: str | None = None
    author: str | None = None
    subject: str | None = None
    keywords: list[str] = field(default_factory=list)
    language: str | None = None
    created: str | None = None
    modified: str | None = None
    properties: dict[str, JsonValue] = field(default_factory=dict)


@dataclass(slots=True)
class ImageAsset:
    """Represent one extracted image and its human-readable description.

    Attributes:
        number: One-based image number in source order.
        data: Immutable image bytes.
        media_type: MIME media type of the image.
        filename: Source-provided image filename.
        caption: Human-readable image caption.
        description: Accessibility or alternative text.
        width: Width in pixels, when known.
        height: Height in pixels, when known.
        ratio: Width-to-height aspect ratio, when both dimensions are known.
        properties: Small format-specific JSON values under namespaced keys.
    """

    number: int
    data: bytes | None = None
    media_type: str | None = None
    filename: str | None = None
    caption: str | None = None
    description: str | None = None
    width: int | None = None
    height: int | None = None
    properties: dict[str, JsonValue] = field(default_factory=dict)
    ratio: float | None = None

    def __post_init__(self) -> None:
        """Reject invalid image values and derive the aspect ratio when possible.

        Raises:
            ValueError: If ``number`` is less than one.
            ValueError: If an explicit ``ratio`` is not positive.
        """
        if self.number < 1:
            raise ValueError("Image numbers must start at 1")
        if self.ratio is not None and (not isfinite(self.ratio) or self.ratio <= 0):
            raise ValueError("Image ratios must be positive")
        if (
            self.ratio is None
            and self.width is not None
            and self.width > 0
            and self.height is not None
            and self.height > 0
        ):
            self.ratio = self.width / self.height


@dataclass(slots=True)
class Table:
    """Represent one rectangular or ragged table in source row order.

    Attributes:
        rows: Rows containing JSON-compatible scalar cell values.
        name: Source-provided table or sheet name.
        caption: Human-readable table caption.
        properties: Small format-specific JSON values under namespaced keys.
    """

    rows: list[list[CellValue]] = field(default_factory=list)
    name: str | None = None
    caption: str | None = None
    properties: dict[str, JsonValue] = field(default_factory=dict)

    @property
    def dimensions(self) -> tuple[int, int]:
        """Return row and maximum column counts.

        Returns:
            A ``(rows, columns)`` tuple that supports ragged tables.

        Example:
            >>> Table(rows=[[1], [2, 3]]).dimensions
            (2, 2)
        """
        return len(self.rows), max((len(row) for row in self.rows), default=0)


@dataclass(slots=True)
class Annotation:
    """Represent supplemental content attached to a document location.

    Attributes:
        kind: Stable lowercase annotation kind.
        text: Human-readable annotation text.
        author: Annotation author or creator.
        target: Link, bookmark, or source target.
        properties: Small format-specific JSON values under namespaced keys.
    """

    kind: str
    text: str = ""
    author: str | None = None
    target: str | None = None
    properties: dict[str, JsonValue] = field(default_factory=dict)


@dataclass(slots=True)
class ContentUnit:
    """Represent an ordered, independently consumable part of a document.

    Attributes:
        number: One-based unit number in source order.
        kind: Format-neutral structural unit kind.
        text: Extracted text in source reading order.
        title: Human-readable unit title.
        heading_path: Hierarchical headings containing this unit.
        images: Images canonically owned by this unit.
        tables: Tables canonically owned by this unit.
        annotations: Supplemental records owned by this unit.
        properties: Small format-specific JSON values under namespaced keys.
    """

    number: int
    kind: UnitKind
    text: str = ""
    title: str | None = None
    heading_path: list[str] = field(default_factory=list)
    images: list[ImageAsset] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    annotations: list[Annotation] = field(default_factory=list)
    properties: dict[str, JsonValue] = field(default_factory=dict)

    def __post_init__(self) -> None:
        """Reject invalid one-based unit numbers.

        Raises:
            ValueError: If ``number`` is less than one.
        """
        if self.number < 1:
            raise ValueError("Unit numbers must start at 1")


@dataclass(slots=True)
class Attachment:
    """Represent an attached file and its optional extracted content.

    Attributes:
        filename: Attachment filename.
        media_type: MIME media type of the attachment.
        data: Immutable attachment bytes.
        extracted_document: Recursively normalized attachment content.
        properties: Small format-specific JSON values under namespaced keys.
    """

    filename: str
    media_type: str | None = None
    data: bytes | None = None
    extracted_document: ExtractedDocument | None = None
    properties: dict[str, JsonValue] = field(default_factory=dict)


@dataclass(slots=True)
class ExtractedDocument:
    """Provide the normalized extraction result returned to consumers.

    Attributes:
        format: Lowercase source format identifier.
        source: Common source identity and decoding metadata.
        metadata: Common descriptive metadata.
        units: Structural content units in source order.
        document_images: Convenience aggregate of all unit-owned images.
        document_tables: Convenience aggregate of all unit-owned tables.
        document_annotations: Convenience aggregate of all unit-owned annotations.
        attachments: Attached files in source order.
        properties: Small format-specific JSON values under namespaced keys.
    """

    format: str
    source: SourceMetadata = field(default_factory=SourceMetadata)
    metadata: DocumentMetadata = field(default_factory=DocumentMetadata)
    units: list[ContentUnit] = field(default_factory=list)
    document_images: list[ImageAsset] = field(default_factory=list)
    document_tables: list[Table] = field(default_factory=list)
    document_annotations: list[Annotation] = field(default_factory=list)
    attachments: list[Attachment] = field(default_factory=list)
    properties: dict[str, JsonValue] = field(default_factory=dict)

    def __post_init__(self) -> None:
        """Create a fallback unit and normalize document convenience fields."""
        if not self.units:
            self.units.append(ContentUnit(number=1, kind="document"))
        self.document_images = _normalize_document_records(
            [unit.images for unit in self.units], self.document_images
        )
        self.document_tables = _normalize_document_records(
            [unit.tables for unit in self.units], self.document_tables
        )
        self.document_annotations = _normalize_document_records(
            [unit.annotations for unit in self.units], self.document_annotations
        )

    @property
    def full_text(self) -> str:
        """Return non-empty unit text joined in source order.

        Returns:
            Newline-delimited text from non-empty units.

        Example:
            >>> document = ExtractedDocument(
            ...     format="txt",
            ...     units=[ContentUnit(number=1, kind="document", text="Hello")],
            ... )
            >>> document.full_text
            'Hello'
        """
        rendered_text = self.properties.get("document.full_text")
        if isinstance(rendered_text, str):
            return rendered_text
        return "\n".join(unit.text for unit in self.units if unit.text).strip()

    def iter_images(self) -> Iterator[ImageAsset]:
        """Yield all document images without copying their records.

        Yields:
            Unit-owned images in canonical document order.
        """
        yield from self.document_images

    def iter_tables(self) -> Iterator[Table]:
        """Yield all document tables without copying their records.

        Yields:
            Unit-owned tables in canonical document order.
        """
        yield from self.document_tables
