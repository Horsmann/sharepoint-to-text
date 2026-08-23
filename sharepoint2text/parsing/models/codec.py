"""Explicit codec for the version-2 extraction wire schema."""

from __future__ import annotations

import base64
import binascii
import json
import math
from dataclasses import dataclass
from typing import Literal, cast

from sharepoint2text.parsing.models.core import (
    Annotation,
    Attachment,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    JsonValue,
    SourceMetadata,
    Table,
    UnitKind,
)

BinaryMode = Literal["omit", "base64"]
SCHEMA_NAME = "sharepoint2text.extraction"
SCHEMA_VERSION = 2
DEFAULT_MAX_BINARY_BYTES: int | None = None
_UNIT_KINDS = {
    "document",
    "section",
    "page",
    "slide",
    "sheet",
    "chapter",
    "message",
}


@dataclass
class _DecodeState:
    """Track cumulative binary allocation and an optional decoding ceiling."""

    maximum_bytes: int | None
    decoded_bytes: int = 0


def _without_none(values: dict[str, JsonValue]) -> dict[str, JsonValue]:
    """Return a field mapping with absent optional values removed."""
    return {key: value for key, value in values.items() if value is not None}


def _sorted_json(value: JsonValue) -> JsonValue:
    """Return JSON data with deterministic recursive dictionary ordering."""
    if isinstance(value, dict):
        return {key: _sorted_json(value[key]) for key in sorted(value)}
    if isinstance(value, list):
        return [_sorted_json(item) for item in value]
    return value


def _validate_json(value: JsonValue, path: str) -> JsonValue:
    """Validate and deterministically order one JSON-compatible value."""
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(f"{path} must not contain non-finite numbers")
        return value
    if isinstance(value, list):
        return [_validate_json(item, f"{path}[]") for item in value]
    if isinstance(value, dict):
        return {
            key: _validate_json(item, f"{path}.{key}")
            for key, item in sorted(value.items())
        }
    raise ValueError(f"{path} contains a non-JSON value")


def _validate_properties(
    values: dict[str, JsonValue], path: str
) -> dict[str, JsonValue]:
    """Validate namespaced property keys and their JSON values."""
    invalid = [key for key in values if "." not in key or key.startswith(".")]
    if invalid:
        raise ValueError(f"{path} keys must be namespaced: {invalid[0]!r}")
    return cast(dict[str, JsonValue], _validate_json(values, path))


def _encode_binary(data: bytes | None, binary: BinaryMode) -> str | None:
    """Encode bytes only when the caller selects base64 mode."""
    if data is None or binary == "omit":
        return None
    return base64.b64encode(data).decode("ascii")


def _source_to_dict(source: SourceMetadata) -> dict[str, JsonValue]:
    """Encode source identity using stable wire names."""
    return _without_none(
        {
            "filename": source.filename,
            "extension": source.extension,
            "path": source.path,
            "folder": source.folder,
            "media_type": source.media_type,
            "encoding": source.encoding,
            "size_bytes": source.size_bytes,
        }
    )


def _metadata_to_dict(metadata: DocumentMetadata) -> dict[str, JsonValue]:
    """Encode common document metadata and format properties."""
    values = _without_none(
        {
            "title": metadata.title,
            "author": metadata.author,
            "subject": metadata.subject,
            "language": metadata.language,
            "created": metadata.created,
            "modified": metadata.modified,
        }
    )
    values["keywords"] = list(metadata.keywords)
    values["properties"] = _validate_properties(
        metadata.properties, "metadata.properties"
    )
    return values


def _image_to_dict(image: ImageAsset, binary: BinaryMode) -> dict[str, JsonValue]:
    """Encode one image asset."""
    values = _without_none(
        {
            "number": image.number,
            "media_type": image.media_type,
            "filename": image.filename,
            "caption": image.caption,
            "description": image.description,
            "width": image.width,
            "height": image.height,
            "ratio": image.ratio,
        }
    )
    encoded = _encode_binary(image.data, binary)
    if encoded is not None:
        values["data"] = encoded
    values["properties"] = _validate_properties(image.properties, "image.properties")
    return values


def _table_to_dict(table: Table) -> dict[str, JsonValue]:
    """Encode one table in source row order."""
    values = _without_none({"name": table.name, "caption": table.caption})
    values["rows"] = _validate_json(cast(JsonValue, table.rows), "table.rows")
    values["properties"] = _validate_properties(table.properties, "table.properties")
    return values


def _annotation_to_dict(annotation: Annotation) -> dict[str, JsonValue]:
    """Encode one annotation."""
    values = _without_none(
        {
            "kind": annotation.kind,
            "text": annotation.text,
            "author": annotation.author,
            "target": annotation.target,
        }
    )
    values["properties"] = _validate_properties(
        annotation.properties, "annotation.properties"
    )
    return values


def _unit_to_dict(unit: ContentUnit, binary: BinaryMode) -> dict[str, JsonValue]:
    """Encode one structural content unit."""
    values = _without_none(
        {
            "number": unit.number,
            "kind": unit.kind,
            "text": unit.text,
            "title": unit.title,
        }
    )
    values["heading_path"] = list(unit.heading_path)
    values["images"] = [_image_to_dict(image, binary) for image in unit.images]
    values["tables"] = [_table_to_dict(table) for table in unit.tables]
    values["annotations"] = [
        _annotation_to_dict(annotation) for annotation in unit.annotations
    ]
    values["properties"] = _validate_properties(unit.properties, "unit.properties")
    return values


def _attachment_to_dict(
    attachment: Attachment, binary: BinaryMode
) -> dict[str, JsonValue]:
    """Encode one attachment and optional recursive extraction."""
    values = _without_none(
        {"filename": attachment.filename, "media_type": attachment.media_type}
    )
    encoded = _encode_binary(attachment.data, binary)
    if encoded is not None:
        values["data"] = encoded
    if attachment.extracted_document is not None:
        values["extracted_document"] = _document_body_to_dict(
            attachment.extracted_document, binary
        )
    values["properties"] = _validate_properties(
        attachment.properties, "attachment.properties"
    )
    return values


def _document_body_to_dict(
    document: ExtractedDocument, binary: BinaryMode
) -> dict[str, JsonValue]:
    """Encode the document body shared by envelopes and attachments."""
    return {
        "format": document.format,
        "source": _source_to_dict(document.source),
        "metadata": _metadata_to_dict(document.metadata),
        "units": [_unit_to_dict(unit, binary) for unit in document.units],
        "document_images": [
            _image_to_dict(image, binary) for image in document.document_images
        ],
        "document_tables": [
            _table_to_dict(table) for table in document.document_tables
        ],
        "document_annotations": [
            _annotation_to_dict(item) for item in document.document_annotations
        ],
        "attachments": [
            _attachment_to_dict(item, binary) for item in document.attachments
        ],
        "properties": _validate_properties(document.properties, "document.properties"),
    }


def document_to_dict(
    document: ExtractedDocument, *, binary: BinaryMode = "omit"
) -> dict[str, JsonValue]:
    """Encode a document using the explicit version-2 wire schema.

    Args:
        document: Normalized document to encode.
        binary: Omit payloads by default or encode them as base64 strings.

    Returns:
        Versioned, JSON-compatible extraction envelope.

    Raises:
        ValueError: If ``binary`` is unsupported or properties are not valid JSON.

    Example:
        >>> payload = document_to_dict(ExtractedDocument(format="txt"))
        >>> payload["version"]
        2
    """
    if binary not in ("omit", "base64"):
        raise ValueError(f"Unsupported binary mode: {binary}")
    return {
        "schema": SCHEMA_NAME,
        "version": SCHEMA_VERSION,
        "document": _document_body_to_dict(document, binary),
    }


def _as_mapping(value: JsonValue, path: str) -> dict[str, JsonValue]:
    """Require a JSON object at a wire-schema path."""
    if not isinstance(value, dict):
        raise ValueError(f"{path} must be an object")
    return value


def _as_list(value: JsonValue, path: str) -> list[JsonValue]:
    """Require a JSON array at a wire-schema path."""
    if not isinstance(value, list):
        raise ValueError(f"{path} must be an array")
    return value


def _optional_string(value: JsonValue, path: str) -> str | None:
    """Decode an optional string field."""
    if value is None or isinstance(value, str):
        return value
    raise ValueError(f"{path} must be a string or null")


def _required_string(value: JsonValue, path: str) -> str:
    """Decode a required string field."""
    result = _optional_string(value, path)
    if result is None:
        raise ValueError(f"{path} is required")
    return result


def _optional_int(value: JsonValue, path: str) -> int | None:
    """Decode an optional integer field while rejecting booleans."""
    if value is None:
        return None
    if isinstance(value, int) and not isinstance(value, bool):
        return value
    raise ValueError(f"{path} must be an integer or null")


def _optional_float(value: JsonValue, path: str) -> float | None:
    """Decode an optional finite number while rejecting booleans."""
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        result = float(value)
        if math.isfinite(result):
            return result
    raise ValueError(f"{path} must be a finite number or null")


def _required_int(value: JsonValue, path: str) -> int:
    """Decode a required integer field."""
    result = _optional_int(value, path)
    if result is None:
        raise ValueError(f"{path} is required")
    return result


def _properties(data: dict[str, JsonValue], path: str) -> dict[str, JsonValue]:
    """Decode and validate a namespaced properties mapping."""
    value = data.get("properties", {})
    mapping = _as_mapping(value, f"{path}.properties")
    return _validate_properties(mapping, f"{path}.properties")


def _string_list(value: JsonValue, path: str) -> list[str]:
    """Decode a list containing only strings."""
    items = _as_list(value, path)
    return [_required_string(item, f"{path}[]") for item in items]


def _decode_binary(value: JsonValue, path: str, state: _DecodeState) -> bytes | None:
    """Decode a base64 field subject to an optional cumulative allocation limit."""
    if value is None:
        return None
    encoded = _required_string(value, path)
    if state.maximum_bytes is not None:
        remaining_bytes = state.maximum_bytes - state.decoded_bytes
        maximum_encoded = 4 * math.ceil(remaining_bytes / 3)
        if len(encoded) > maximum_encoded:
            raise ValueError(f"Decoded binary data exceeds {state.maximum_bytes} bytes")
    try:
        decoded = base64.b64decode(encoded, validate=True)
    except (binascii.Error, ValueError) as error:
        raise ValueError(f"{path} is not valid base64") from error
    state.decoded_bytes += len(decoded)
    if state.maximum_bytes is not None and state.decoded_bytes > state.maximum_bytes:
        raise ValueError(f"Decoded binary data exceeds {state.maximum_bytes} bytes")
    return decoded


def _source_from_dict(data: dict[str, JsonValue]) -> SourceMetadata:
    """Decode common source identity fields."""
    return SourceMetadata(
        filename=_optional_string(data.get("filename"), "source.filename"),
        extension=_optional_string(data.get("extension"), "source.extension"),
        path=_optional_string(data.get("path"), "source.path"),
        folder=_optional_string(data.get("folder"), "source.folder"),
        media_type=_optional_string(data.get("media_type"), "source.media_type"),
        encoding=_optional_string(data.get("encoding"), "source.encoding"),
        size_bytes=_optional_int(data.get("size_bytes"), "source.size_bytes"),
    )


def _metadata_from_dict(data: dict[str, JsonValue]) -> DocumentMetadata:
    """Decode common descriptive metadata fields."""
    return DocumentMetadata(
        title=_optional_string(data.get("title"), "metadata.title"),
        author=_optional_string(data.get("author"), "metadata.author"),
        subject=_optional_string(data.get("subject"), "metadata.subject"),
        keywords=_string_list(data.get("keywords", []), "metadata.keywords"),
        language=_optional_string(data.get("language"), "metadata.language"),
        created=_optional_string(data.get("created"), "metadata.created"),
        modified=_optional_string(data.get("modified"), "metadata.modified"),
        properties=_properties(data, "metadata"),
    )


def _image_from_dict(
    data: dict[str, JsonValue], path: str, state: _DecodeState
) -> ImageAsset:
    """Decode one image asset."""
    return ImageAsset(
        number=_required_int(data.get("number"), f"{path}.number"),
        data=_decode_binary(data.get("data"), f"{path}.data", state),
        media_type=_optional_string(data.get("media_type"), f"{path}.media_type"),
        filename=_optional_string(data.get("filename"), f"{path}.filename"),
        caption=_optional_string(data.get("caption"), f"{path}.caption"),
        description=_optional_string(data.get("description"), f"{path}.description"),
        width=_optional_int(data.get("width"), f"{path}.width"),
        height=_optional_int(data.get("height"), f"{path}.height"),
        ratio=_optional_float(data.get("ratio"), f"{path}.ratio"),
        properties=_properties(data, path),
    )


def _cell(value: JsonValue, path: str) -> str | int | float | bool | None:
    """Decode one supported scalar table cell."""
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float) and math.isfinite(value):
        return value
    raise ValueError(f"{path} must be a finite JSON scalar")


def _table_from_dict(data: dict[str, JsonValue], path: str) -> Table:
    """Decode one table and validate every cell."""
    rows_data = _as_list(data.get("rows", []), f"{path}.rows")
    rows = [
        [_cell(cell, f"{path}.rows[]") for cell in _as_list(row, f"{path}.rows[]")]
        for row in rows_data
    ]
    return Table(
        rows=rows,
        name=_optional_string(data.get("name"), f"{path}.name"),
        caption=_optional_string(data.get("caption"), f"{path}.caption"),
        properties=_properties(data, path),
    )


def _annotation_from_dict(data: dict[str, JsonValue], path: str) -> Annotation:
    """Decode one supplemental annotation."""
    return Annotation(
        kind=_required_string(data.get("kind"), f"{path}.kind"),
        text=_required_string(data.get("text", ""), f"{path}.text"),
        author=_optional_string(data.get("author"), f"{path}.author"),
        target=_optional_string(data.get("target"), f"{path}.target"),
        properties=_properties(data, path),
    )


def _objects(value: JsonValue, path: str) -> list[tuple[dict[str, JsonValue], str]]:
    """Decode an object array while retaining indexed error paths."""
    return [
        (_as_mapping(item, f"{path}[{index}]"), f"{path}[{index}]")
        for index, item in enumerate(_as_list(value, path))
    ]


def _unit_from_dict(
    data: dict[str, JsonValue], path: str, state: _DecodeState
) -> ContentUnit:
    """Decode one structural content unit."""
    kind_value = _required_string(data.get("kind"), f"{path}.kind")
    if kind_value not in _UNIT_KINDS:
        raise ValueError(f"{path}.kind is unsupported: {kind_value}")
    return ContentUnit(
        number=_required_int(data.get("number"), f"{path}.number"),
        kind=cast(UnitKind, kind_value),
        text=_required_string(data.get("text", ""), f"{path}.text"),
        title=_optional_string(data.get("title"), f"{path}.title"),
        heading_path=_string_list(data.get("heading_path", []), f"{path}.heading_path"),
        images=[
            _image_from_dict(item, item_path, state)
            for item, item_path in _objects(data.get("images", []), f"{path}.images")
        ],
        tables=[
            _table_from_dict(item, item_path)
            for item, item_path in _objects(data.get("tables", []), f"{path}.tables")
        ],
        annotations=[
            _annotation_from_dict(item, item_path)
            for item, item_path in _objects(
                data.get("annotations", []), f"{path}.annotations"
            )
        ],
        properties=_properties(data, path),
    )


def _attachment_from_dict(
    data: dict[str, JsonValue], path: str, state: _DecodeState
) -> Attachment:
    """Decode one attachment and its optional recursive document."""
    nested_data = data.get("extracted_document")
    nested = None
    if nested_data is not None:
        nested = _document_body_from_dict(
            _as_mapping(nested_data, f"{path}.extracted_document"), state
        )
    return Attachment(
        filename=_required_string(data.get("filename"), f"{path}.filename"),
        media_type=_optional_string(data.get("media_type"), f"{path}.media_type"),
        data=_decode_binary(data.get("data"), f"{path}.data", state),
        extracted_document=nested,
        properties=_properties(data, path),
    )


def _decode_units(data: dict[str, JsonValue], state: _DecodeState) -> list[ContentUnit]:
    """Decode all structural units in a document body."""
    values = _objects(data.get("units", []), "document.units")
    return [_unit_from_dict(item, path, state) for item, path in values]


def _decode_document_images(
    data: dict[str, JsonValue], state: _DecodeState
) -> list[ImageAsset]:
    """Decode the document-level image convenience aggregate."""
    path = "document.document_images"
    values = _objects(data.get("document_images", []), path)
    return [_image_from_dict(item, item_path, state) for item, item_path in values]


def _decode_document_tables(data: dict[str, JsonValue]) -> list[Table]:
    """Decode the document-level table convenience aggregate."""
    path = "document.document_tables"
    values = _objects(data.get("document_tables", []), path)
    return [_table_from_dict(item, item_path) for item, item_path in values]


def _decode_document_annotations(data: dict[str, JsonValue]) -> list[Annotation]:
    """Decode the document-level annotation convenience aggregate."""
    path = "document.document_annotations"
    values = _objects(data.get("document_annotations", []), path)
    return [_annotation_from_dict(item, item_path) for item, item_path in values]


def _decode_attachments(
    data: dict[str, JsonValue], state: _DecodeState
) -> list[Attachment]:
    """Decode all document attachments."""
    values = _objects(data.get("attachments", []), "document.attachments")
    return [_attachment_from_dict(item, path, state) for item, path in values]


def _document_body_from_dict(
    data: dict[str, JsonValue], state: _DecodeState
) -> ExtractedDocument:
    """Decode a normalized document body."""
    source = _as_mapping(data.get("source", {}), "document.source")
    metadata = _as_mapping(data.get("metadata", {}), "document.metadata")
    return ExtractedDocument(
        format=_required_string(data.get("format"), "document.format"),
        source=_source_from_dict(source),
        metadata=_metadata_from_dict(metadata),
        units=_decode_units(data, state),
        document_images=_decode_document_images(data, state),
        document_tables=_decode_document_tables(data),
        document_annotations=_decode_document_annotations(data),
        attachments=_decode_attachments(data, state),
        properties=_properties(data, "document"),
    )


def document_from_dict(
    data: dict[str, JsonValue],
    *,
    max_binary_bytes: int | None = DEFAULT_MAX_BINARY_BYTES,
) -> ExtractedDocument:
    """Decode and validate a version-2 extraction document.

    Args:
        data: Versioned extraction envelope.
        max_binary_bytes: Optional maximum cumulative decoded payload size.
            ``None`` decodes the complete payload without a size ceiling.

    Returns:
        Validated normalized document.

    Raises:
        ValueError: If the schema is invalid, unsupported, or exceeds an
            explicitly configured binary limit.

    Example:
        >>> original = ExtractedDocument(format="txt")
        >>> document_from_dict(document_to_dict(original)) == original
        True
    """
    if max_binary_bytes is not None and max_binary_bytes < 0:
        raise ValueError("max_binary_bytes must not be negative")
    if data.get("schema") != SCHEMA_NAME:
        raise ValueError(f"Unsupported extraction schema: {data.get('schema')!r}")
    if data.get("version") != SCHEMA_VERSION:
        raise ValueError(
            f"Unsupported extraction schema version: {data.get('version')!r}"
        )
    body = _as_mapping(data.get("document"), "document")
    return _document_body_from_dict(body, _DecodeState(maximum_bytes=max_binary_bytes))


def document_to_json(
    document: ExtractedDocument, *, binary: BinaryMode = "omit"
) -> str:
    """Encode a normalized document as deterministic compact JSON.

    Args:
        document: Normalized document to encode.
        binary: Omit payloads by default or encode them as base64 strings.

    Returns:
        Compact JSON with recursively sorted object keys.
    """
    payload = _sorted_json(document_to_dict(document, binary=binary))
    return json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def document_from_json(
    value: str,
    *,
    max_binary_bytes: int | None = DEFAULT_MAX_BINARY_BYTES,
) -> ExtractedDocument:
    """Decode a normalized document from a JSON string.

    Args:
        value: JSON extraction envelope.
        max_binary_bytes: Optional maximum cumulative decoded payload size.
            ``None`` decodes the complete payload without a size ceiling.

    Returns:
        Validated normalized document.

    Raises:
        ValueError: If JSON or the extraction schema is invalid, or an
            explicitly configured binary limit is exceeded.
    """
    try:
        parsed: object = json.loads(value)
    except json.JSONDecodeError as error:
        raise ValueError("Invalid extraction JSON") from error
    if not isinstance(parsed, dict):
        raise ValueError("Extraction JSON must contain an object")
    return document_from_dict(
        cast(dict[str, JsonValue], parsed), max_binary_bytes=max_binary_bytes
    )
