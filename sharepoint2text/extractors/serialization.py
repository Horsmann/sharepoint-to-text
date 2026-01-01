import base64
import io
import typing
from dataclasses import fields, is_dataclass

# Type marker key used for serialization/deserialization
_TYPE_KEY = "_type"

# Registry mapping type names to classes (populated lazily)
_TYPE_REGISTRY: dict[str, type] = {}


def _bytes_to_base64(data: bytes | bytearray) -> str:
    return base64.b64encode(bytes(data)).decode("utf-8")


def _base64_to_bytes(data: str) -> bytes:
    return base64.b64decode(data.encode("utf-8"))


def _bytesio_to_base64(buffer: io.BytesIO) -> str:
    position = buffer.tell()
    buffer.seek(0)
    encoded = base64.b64encode(buffer.read()).decode("utf-8")
    buffer.seek(position)
    return encoded


def _base64_to_bytesio(data: str) -> io.BytesIO:
    return io.BytesIO(base64.b64decode(data.encode("utf-8")))


def _serialize_for_json(value: typing.Any) -> typing.Any:
    if isinstance(value, io.BytesIO):
        return {"_bytesio": _bytesio_to_base64(value)}
    if isinstance(value, (bytes, bytearray)):
        return {"_bytes": _bytes_to_base64(value)}
    if is_dataclass(value) and not isinstance(value, type):
        result = {
            _TYPE_KEY: type(value).__name__,
        }
        for item in fields(value):
            result[item.name] = _serialize_for_json(getattr(value, item.name))
        return result
    if isinstance(value, dict):
        return {str(key): _serialize_for_json(val) for key, val in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_serialize_for_json(item) for item in value]
    return value


def serialize_extraction(value: typing.Any) -> dict:
    serialized = _serialize_for_json(value)
    if isinstance(serialized, dict):
        return serialized
    return {"value": serialized}


def _get_type_registry() -> dict[str, type]:
    """Lazily populate and return the type registry."""
    if _TYPE_REGISTRY:
        return _TYPE_REGISTRY

    # Import all dataclass types from data_types module
    from sharepoint2text.extractors import data_types

    # Get all classes from data_types that are dataclasses
    for name in dir(data_types):
        obj = getattr(data_types, name)
        if isinstance(obj, type) and is_dataclass(obj):
            _TYPE_REGISTRY[name] = obj

    return _TYPE_REGISTRY


def _get_field_types(cls: type) -> dict[str, typing.Any]:
    """Get the type hints for all fields of a dataclass."""
    hints = typing.get_type_hints(cls)
    return hints


def _unwrap_optional(tp: typing.Any) -> tuple[typing.Any, bool]:
    """Unwrap Optional[X] to (X, True) or return (tp, False) if not Optional."""
    origin = typing.get_origin(tp)
    if origin is typing.Union:
        args = typing.get_args(tp)
        # Optional[X] is Union[X, None]
        non_none_args = [a for a in args if a is not type(None)]
        if len(non_none_args) == 1 and len(args) == 2:
            return non_none_args[0], True
    return tp, False


def _deserialize_value(value: typing.Any, expected_type: typing.Any) -> typing.Any:
    """Deserialize a value according to its expected type."""
    if value is None:
        return None

    # Handle Optional types
    inner_type, is_optional = _unwrap_optional(expected_type)
    if is_optional:
        expected_type = inner_type

    # Handle special encoded types
    if isinstance(value, dict):
        if "_bytesio" in value:
            return _base64_to_bytesio(value["_bytesio"])
        if "_bytes" in value:
            return _base64_to_bytes(value["_bytes"])
        if _TYPE_KEY in value:
            return _deserialize_dataclass(value)

    # Handle generic types (List, Dict, etc.)
    origin = typing.get_origin(expected_type)

    if origin is list:
        item_type = typing.get_args(expected_type)
        item_type = item_type[0] if item_type else typing.Any
        if isinstance(value, list):
            return [_deserialize_value(item, item_type) for item in value]
        return value

    if origin is dict:
        args = typing.get_args(expected_type)
        value_type = args[1] if len(args) > 1 else typing.Any
        if isinstance(value, dict):
            return {k: _deserialize_value(v, value_type) for k, v in value.items()}
        return value

    # Handle bytes type explicitly
    if expected_type is bytes or expected_type is bytearray:
        if isinstance(value, dict) and "_bytes" in value:
            return _base64_to_bytes(value["_bytes"])
        if isinstance(value, str):
            # Assume it's base64 encoded
            return _base64_to_bytes(value)
        return value

    # Handle BytesIO type
    if expected_type is io.BytesIO:
        if isinstance(value, dict) and "_bytesio" in value:
            return _base64_to_bytesio(value["_bytesio"])
        if isinstance(value, str):
            return _base64_to_bytesio(value)
        return value

    # Handle dataclass types
    registry = _get_type_registry()
    if isinstance(expected_type, type) and expected_type.__name__ in registry:
        if isinstance(value, dict):
            return _deserialize_dataclass(value, expected_type)
        return value

    # Return primitive values as-is
    return value


def _deserialize_dataclass(
    data: dict, expected_class: typing.Optional[type] = None
) -> typing.Any:
    """Deserialize a dictionary to a dataclass instance."""
    registry = _get_type_registry()

    # Determine the target class
    type_name = data.get(_TYPE_KEY)
    if type_name and type_name in registry:
        cls = registry[type_name]
    elif expected_class is not None:
        cls = expected_class
    else:
        # Can't determine the class, return dict as-is
        return data

    # Get field type hints
    field_types = _get_field_types(cls)
    field_names = {f.name for f in fields(cls)}

    # Build kwargs for constructor
    kwargs = {}
    for field_name in field_names:
        if field_name in data:
            field_type = field_types.get(field_name, typing.Any)
            kwargs[field_name] = _deserialize_value(data[field_name], field_type)

    return cls(**kwargs)


def deserialize_extraction(data: dict) -> typing.Any:
    """
    Deserialize a JSON dictionary back to an ExtractionInterface instance.

    This is the inverse of serialize_extraction(). It reconstructs the original
    dataclass hierarchy from the serialized JSON representation.

    Args:
        data: A dictionary produced by serialize_extraction() or to_json()

    Returns:
        An instance of the appropriate ExtractionInterface subclass

    Raises:
        ValueError: If the data doesn't contain valid type information
        KeyError: If the type name is not recognized

    Example:
        >>> content = read_file("document.docx")
        >>> json_data = content.to_json()
        >>> restored = deserialize_extraction(json_data)
        >>> assert restored.get_full_text() == content.get_full_text()
    """
    if not isinstance(data, dict):
        raise ValueError("Input must be a dictionary")

    if _TYPE_KEY not in data:
        raise ValueError(
            f"Input dictionary must contain '{_TYPE_KEY}' key for deserialization"
        )

    return _deserialize_dataclass(data)
