"""Internal adapters from extractor records to the normalized model."""

from __future__ import annotations

import io
import mimetypes
import re
from dataclasses import fields, is_dataclass
from datetime import date, datetime
from typing import TypeVar, cast

from sharepoint2text.parsing.extractors._legacy_types import (
    ExtractionInterface,
    ImageInterface,
    TableInterface,
    UnitInterface,
)
from sharepoint2text.parsing.models.core import (
    Annotation,
    Attachment,
    CellValue,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    JsonValue,
    SourceMetadata,
    Table,
    UnitKind,
)

_FORMAT_BY_CONTENT_TYPE = {
    "CsvContent": "csv",
    "DocContent": "doc",
    "DocxContent": "docx",
    "EmailContent": "email",
    "EpubContent": "epub",
    "HtmlContent": "html",
    "OdfContent": "odf",
    "OdgContent": "odg",
    "OdpContent": "odp",
    "OdsContent": "ods",
    "OdtContent": "odt",
    "PdfContent": "pdf",
    "PlainTextContent": "txt",
    "PptContent": "ppt",
    "PptxContent": "pptx",
    "RtfContent": "rtf",
    "XlsContent": "xls",
    "XlsxContent": "xlsx",
}
_SOURCE_FIELDS = {
    "filename",
    "file_extension",
    "file_path",
    "folder_path",
    "detected_encoding",
}
_COMMON_METADATA_FIELDS = _SOURCE_FIELDS | {
    "title",
    "author",
    "creator",
    "initial_creator",
    "subject",
    "description",
    "keywords",
    "language",
    "created",
    "create_time",
    "creation_date",
    "modified",
    "date",
    "last_saved_time",
}
_COMMON_IMAGE_FIELDS = {
    "data",
    "image_index",
    "content_type",
    "filename",
    "name",
    "caption",
    "description",
    "width",
    "height",
}
_UNIT_KINDS_BY_FORMAT: dict[str, UnitKind] = {
    "email": "message",
    "eml": "message",
    "mbox": "message",
    "msg": "message",
    "epub": "chapter",
    "pdf": "page",
    "rtf": "page",
    "ppt": "slide",
    "pptx": "slide",
    "odp": "slide",
    "xls": "sheet",
    "xlsx": "sheet",
    "ods": "sheet",
}
_UNSUPPORTED = object()
_Item = TypeVar("_Item")


def _nonempty_text(value: object) -> str | None:
    """Return stripped text or ``None`` for absent and empty values."""
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    return stripped or None


def _first_text(value: object, *names: str) -> str | None:
    """Return the first non-empty text attribute from a legacy object."""
    for name in names:
        result = _nonempty_text(getattr(value, name, None))
        if result is not None:
            return result
    return None


def _legacy_format(value: ExtractionInterface, metadata: object) -> str:
    """Derive a lowercase source format from metadata or the content type."""
    extension = _first_text(metadata, "file_extension")
    if extension:
        return extension.lower().lstrip(".")
    return _FORMAT_BY_CONTENT_TYPE.get(
        type(value).__name__, type(value).__name__.lower()
    )


def _source_metadata(metadata: object, source_format: str) -> SourceMetadata:
    """Map shared legacy source fields to stable source metadata."""
    filename = _first_text(metadata, "filename")
    extension = _first_text(metadata, "file_extension")
    media_name = filename or (f"source.{source_format}" if source_format else None)
    media_type = mimetypes.guess_type(media_name or "")[0]
    return SourceMetadata(
        filename=filename,
        extension=extension,
        path=_first_text(metadata, "file_path"),
        folder=_first_text(metadata, "folder_path"),
        media_type=media_type,
        encoding=_first_text(metadata, "detected_encoding"),
    )


def _json_value(value: object) -> JsonValue | object:
    """Convert a small legacy value to JSON data or mark it unsupported."""
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if is_dataclass(value) and not isinstance(value, type):
        converted_dataclass = {
            item.name: _json_value(getattr(value, item.name)) for item in fields(value)
        }
        if all(item is not _UNSUPPORTED for item in converted_dataclass.values()):
            return cast(dict[str, JsonValue], converted_dataclass)
    if isinstance(value, (list, tuple)):
        converted = [_json_value(item) for item in value]
        return (
            converted
            if all(item is not _UNSUPPORTED for item in converted)
            else _UNSUPPORTED
        )
    if isinstance(value, dict) and all(isinstance(key, str) for key in value):
        converted_dict = {key: _json_value(item) for key, item in value.items()}
        if all(item is not _UNSUPPORTED for item in converted_dict.values()):
            return cast(dict[str, JsonValue], converted_dict)
    return _UNSUPPORTED


def _namespaced_fields(
    value: object, namespace: str, excluded: set[str]
) -> dict[str, JsonValue]:
    """Preserve small format-specific dataclass fields under stable keys."""
    if not is_dataclass(value) or isinstance(value, type):
        return {}
    properties: dict[str, JsonValue] = {}
    for item in fields(value):
        if item.name in excluded:
            continue
        converted = _json_value(getattr(value, item.name))
        if converted is not _UNSUPPORTED and converted not in (None, "", [], {}):
            properties[f"{namespace}.{item.name}"] = cast(JsonValue, converted)
    return properties


def _keywords(value: object) -> list[str]:
    """Normalize legacy keyword strings and sequences."""
    if isinstance(value, str):
        return [item.strip() for item in re.split(r"[;,]", value) if item.strip()]
    if isinstance(value, (list, tuple)):
        return [text for item in value if (text := _nonempty_text(item))]
    return []


def _document_metadata(
    legacy: ExtractionInterface, metadata: object, source_format: str
) -> DocumentMetadata:
    """Map descriptive legacy metadata and preserve uncommon scalar facts."""
    title = _first_text(metadata, "title")
    if title is None:
        title = _first_text(legacy, "subject")
    author = _first_text(metadata, "author", "creator", "initial_creator")
    if author is None and hasattr(legacy, "from_email"):
        sender = getattr(legacy, "from_email")
        author = _first_text(sender, "address", "name")
    return DocumentMetadata(
        title=title,
        author=author,
        subject=_first_text(metadata, "subject", "description"),
        keywords=_keywords(getattr(metadata, "keywords", None)),
        language=_first_text(metadata, "language"),
        created=_first_text(
            metadata, "created", "create_time", "creation_date", "date"
        ),
        modified=_first_text(metadata, "modified", "last_saved_time"),
        properties=_namespaced_fields(metadata, source_format, _COMMON_METADATA_FIELDS),
    )


def _stream_bytes(value: object) -> bytes | None:
    """Read immutable bytes without changing a legacy stream cursor."""
    if value is None:
        return None
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if not isinstance(value, io.BytesIO):
        return None
    position = value.tell()
    try:
        value.seek(0)
        return value.read()
    finally:
        value.seek(position)


def _positive_dimension(value: object) -> int | None:
    """Return a positive integer dimension or ``None``."""
    return (
        value
        if isinstance(value, int) and not isinstance(value, bool) and value > 0
        else None
    )


def _image_asset(
    image: ImageInterface, source_format: str, fallback: int
) -> ImageAsset:
    """Normalize one legacy image without leaking mutable stream state."""
    metadata = image.get_metadata()
    legacy_number = getattr(metadata, "image_number", None)
    number = (
        legacy_number
        if isinstance(legacy_number, int) and legacy_number > 0
        else fallback
    )
    return ImageAsset(
        number=number,
        data=_stream_bytes(getattr(image, "data", None)),
        media_type=_nonempty_text(image.get_content_type()),
        filename=_first_text(image, "filename", "name", "href"),
        caption=_nonempty_text(image.get_caption()),
        description=_nonempty_text(image.get_description()),
        width=_positive_dimension(getattr(metadata, "width", None)),
        height=_positive_dimension(getattr(metadata, "height", None)),
        properties=_namespaced_fields(image, source_format, _COMMON_IMAGE_FIELDS),
    )


def _cell_value(value: object) -> CellValue:
    """Convert a legacy table value to a stable scalar cell value."""
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(value)


def _table(table: TableInterface, source_format: str) -> Table:
    """Normalize one legacy table while preserving ragged row order."""
    rows = [[_cell_value(cell) for cell in row] for row in table.get_table()]
    excluded = {"data", "name", "caption", "text", "images", "annotations"}
    return Table(
        rows=rows,
        name=_first_text(table, "name"),
        caption=_first_text(table, "caption"),
        properties=_namespaced_fields(table, source_format, excluded),
    )


def _annotation(record: object, kind: str, source_format: str) -> Annotation:
    """Normalize one legacy supplemental record."""
    if isinstance(record, str):
        return Annotation(kind=kind, text=record)
    text = _first_text(record, "text", "latex", "field_result", "name") or ""
    author = _first_text(record, "author", "creator")
    target = _first_text(record, "url", "href")
    excluded = {
        "text",
        "author",
        "creator",
        "url",
        "href",
        "latex",
        "is_display",
    }
    properties = _namespaced_fields(record, source_format, excluded)
    if kind == "formula":
        properties["math.latex"] = text
        properties["math.display"] = bool(getattr(record, "is_display", False))
    return Annotation(
        kind=kind, text=text, author=author, target=target, properties=properties
    )


def _record_annotations(owner: object, source_format: str) -> list[Annotation]:
    """Collect supported annotation families from a legacy owner."""
    mappings = (
        ("comments", "comment"),
        ("annotations", "comment"),
        ("footnotes", "footnote"),
        ("endnotes", "endnote"),
        ("formulas", "formula"),
        ("hyperlinks", "hyperlink"),
        ("links", "hyperlink"),
        ("bookmarks", "bookmark"),
        ("fields", "field"),
        ("notes", "note"),
        ("headers", "header"),
        ("footers", "footer"),
        ("headers_footers", "header_footer"),
    )
    result: list[Annotation] = []
    for field_name, kind in mappings:
        records = getattr(owner, field_name, ())
        if isinstance(records, str) and records.strip():
            result.append(_annotation(records, kind, source_format))
        elif isinstance(records, (list, tuple)):
            result.extend(
                _annotation(record, kind, source_format) for record in records
            )
    return result


def _unit_kind(source_format: str, unit: UnitInterface, metadata: object) -> UnitKind:
    """Map a legacy structural unit to a format-neutral kind."""
    mapped = _UNIT_KINDS_BY_FORMAT.get(source_format)
    if mapped is not None:
        return mapped
    if hasattr(unit, "body_type"):
        return "message"
    if hasattr(unit, "slide_number"):
        return "slide"
    if hasattr(unit, "sheet_name"):
        return "sheet"
    if hasattr(unit, "page_number"):
        return "page"
    if type(unit).__name__ == "EpubChapter":
        return "chapter"
    heading_path = getattr(metadata, "heading_path", None)
    return "section" if heading_path else "document"


def _unit_annotations(
    legacy: ExtractionInterface, index: int, source_format: str
) -> list[Annotation]:
    """Collect annotations from the corresponding slide or sheet record."""
    for collection_name in ("slides", "sheets", "chapters"):
        collection = getattr(legacy, collection_name, None)
        if isinstance(collection, list) and index < len(collection):
            return _record_annotations(collection[index], source_format)
    return []


def _content_unit(
    legacy: ExtractionInterface, unit: UnitInterface, index: int, source_format: str
) -> ContentUnit:
    """Normalize one legacy content unit and its canonical assets."""
    metadata = unit.get_metadata()
    legacy_number = getattr(metadata, "unit_number", None)
    number = (
        legacy_number
        if isinstance(legacy_number, int) and legacy_number > 0
        else index + 1
    )
    heading_path = getattr(metadata, "heading_path", [])
    safe_heading_path = (
        [str(item) for item in heading_path] if isinstance(heading_path, list) else []
    )
    excluded = {"unit_number", "title", "sheet_name", "heading_path"}
    return ContentUnit(
        number=number,
        kind=_unit_kind(source_format, unit, metadata),
        text=unit.get_text(),
        title=_first_text(metadata, "title", "sheet_name"),
        heading_path=safe_heading_path,
        images=[
            _image_asset(image, source_format, position)
            for position, image in enumerate(unit.get_images(), start=1)
        ],
        tables=[_table(table, source_format) for table in unit.get_tables()],
        annotations=_unit_annotations(legacy, index, source_format),
        properties=_namespaced_fields(metadata, source_format, excluded),
    )


def _unowned(all_items: list[_Item], owned_items: list[_Item]) -> list[_Item]:
    """Subtract a multiset of unit-owned objects without changing order."""
    remaining_owned = list(owned_items)
    result: list[_Item] = []
    for item in all_items:
        try:
            index = remaining_owned.index(item)
        except ValueError:
            result.append(item)
        else:
            remaining_owned.pop(index)
    expected_count = max(0, len(all_items) - len(owned_items))
    if len(result) <= expected_count:
        return result
    return result[-expected_count:] if expected_count else []


def _attachments(legacy: ExtractionInterface, source_format: str) -> list[Attachment]:
    """Normalize email attachments without extracting them eagerly."""
    records = getattr(legacy, "attachments", ())
    if not isinstance(records, (list, tuple)):
        return []
    return [
        Attachment(
            filename=str(getattr(record, "filename", "")),
            media_type=_first_text(record, "mime_type"),
            data=_stream_bytes(getattr(record, "data", None)),
            properties=_namespaced_fields(
                record, source_format, {"filename", "mime_type", "data"}
            ),
        )
        for record in records
    ]


def _document_properties(
    legacy: ExtractionInterface, source_format: str
) -> dict[str, JsonValue]:
    """Preserve selected stable content-level facts that are not parser trees."""
    properties: dict[str, JsonValue] = {}
    for field_name in (
        "in_reply_to",
        "reply_to",
        "to_emails",
        "to_cc",
        "to_bcc",
        "toc",
    ):
        value = getattr(legacy, field_name, None)
        if is_dataclass(value) and not isinstance(value, type):
            continue
        converted = _json_value(value)
        if converted is not _UNSUPPORTED and converted not in (None, "", [], {}):
            properties[f"{source_format}.{field_name}"] = cast(JsonValue, converted)
    return properties


def _normalized_units(
    legacy: ExtractionInterface, source_format: str
) -> list[ContentUnit]:
    """Normalize all structural units in source order."""
    return [
        _content_unit(legacy, unit, index, source_format)
        for index, unit in enumerate(legacy.iterate_units())
    ]


def _normalized_images(
    legacy: ExtractionInterface, source_format: str
) -> list[ImageAsset]:
    """Normalize the legacy document-wide image iterator."""
    return [
        _image_asset(image, source_format, index)
        for index, image in enumerate(legacy.iterate_images(), start=1)
    ]


def _normalized_tables(legacy: ExtractionInterface, source_format: str) -> list[Table]:
    """Normalize the legacy document-wide table iterator."""
    return [_table(table, source_format) for table in legacy.iterate_tables()]


def _build_document(
    legacy: ExtractionInterface, metadata: object, source_format: str
) -> ExtractedDocument:
    """Build a normalized document after validating the legacy boundary."""
    units = _normalized_units(legacy, source_format)
    all_images = _normalized_images(legacy, source_format)
    all_tables = _normalized_tables(legacy, source_format)
    owned_images = [item for unit in units for item in unit.images]
    owned_tables = [item for unit in units for item in unit.tables]
    return ExtractedDocument(
        format=source_format,
        source=_source_metadata(metadata, source_format),
        metadata=_document_metadata(legacy, metadata, source_format),
        units=units,
        document_images=_unowned(all_images, owned_images),
        document_tables=_unowned(all_tables, owned_tables),
        document_annotations=_record_annotations(legacy, source_format),
        attachments=_attachments(legacy, source_format),
        properties=_document_properties(legacy, source_format),
    )


def _normalize_extraction(legacy_result: ExtractionInterface) -> ExtractedDocument:
    """Convert an internal extractor result to the normalized data model.

    Args:
        legacy_result: Format-specific internal extraction result.

    Returns:
        Normalized document with one canonical owner for each asset.

    Raises:
        TypeError: If the value does not implement the internal extraction protocol.

    Example:
        >>> from sharepoint2text.parsing.extractors._legacy_types import PlainTextContent
        >>> _normalize_extraction(PlainTextContent(content="Hello")).full_text
        'Hello'
    """
    required = ("get_metadata", "iterate_units", "iterate_images", "iterate_tables")
    if not all(callable(getattr(legacy_result, name, None)) for name in required):
        raise TypeError("legacy_result must implement the internal extraction protocol")
    metadata = legacy_result.get_metadata()
    source_format = _legacy_format(legacy_result, metadata)
    return _build_document(legacy_result, metadata, source_format)
