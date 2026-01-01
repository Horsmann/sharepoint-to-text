import base64
import io
import typing
from dataclasses import fields, is_dataclass


def _bytes_to_base64(data: bytes | bytearray) -> str:
    return base64.b64encode(bytes(data)).decode("utf-8")


def _bytesio_to_base64(buffer: io.BytesIO) -> str:
    position = buffer.tell()
    buffer.seek(0)
    encoded = base64.b64encode(buffer.read()).decode("utf-8")
    buffer.seek(position)
    return encoded


def _serialize_for_json(value: typing.Any) -> typing.Any:
    if isinstance(value, io.BytesIO):
        return _bytesio_to_base64(value)
    if isinstance(value, (bytes, bytearray)):
        return _bytes_to_base64(value)
    if is_dataclass(value):
        return {
            item.name: _serialize_for_json(getattr(value, item.name))
            for item in fields(value)
        }
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
