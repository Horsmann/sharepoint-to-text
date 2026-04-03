#!/usr/bin/env python3
"""
Extract textual content from a modern Apple Pages `.pages` archive.

This intentionally reads only proprietary `Index/*.iwa` members and ignores preview
images. The extractor combines two generic strategies:
1. Scan IWA payload bytes for visible UTF-8 runs to recover rich document text.
2. Decode simple protobuf-like records used by Pages tables to recover headers and
   scalar cell values from `Index/Tables/*.iwa`.
"""

from __future__ import annotations

import argparse
import io
import math
import re
import string
import struct
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from pathlib import Path

PRINTABLE_ASCII = set(string.printable) - {"\x0b", "\x0c"}


@dataclass(frozen=True)
class Candidate:
    source: str
    kind: str
    text: str


@dataclass(frozen=True)
class Table:
    headers: list[str]
    rows: list[list[str]]
    raw_values: list[str]
    metadata: dict[str, object]


@dataclass(frozen=True)
class DecodedMessage:
    object_id: int
    message_type: int
    body: bytes
    source: str


@dataclass(frozen=True)
class TableOwner:
    owner_id: int
    tile_object_id: int | None
    string_table_object_id: int | None
    title: str | None
    row_count: int | None
    column_count: int | None


@dataclass(frozen=True)
class DocumentFlow:
    segments: list[str]
    placeholder_count: int
    source_object_id: int


def read_varint(data: bytes, pos: int) -> tuple[int, int]:
    value = 0
    shift = 0
    while pos < len(data):
        byte = data[pos]
        pos += 1
        value |= (byte & 0x7F) << shift
        if byte < 0x80:
            return value, pos
        shift += 7
    raise ValueError("unterminated varint")


def parse_proto_fields(data: bytes) -> list[tuple[int, int, object]]:
    fields: list[tuple[int, int, object]] = []
    pos = 0
    while pos < len(data):
        try:
            tag, pos = read_varint(data, pos)
        except ValueError:
            break

        field_number = tag >> 3
        wire_type = tag & 0x07
        if field_number == 0:
            break

        if wire_type == 0:
            value, pos = read_varint(data, pos)
            fields.append((field_number, wire_type, value))
        elif wire_type == 1:
            if pos + 8 > len(data):
                break
            fields.append((field_number, wire_type, data[pos : pos + 8]))
            pos += 8
        elif wire_type == 2:
            try:
                size, pos = read_varint(data, pos)
            except ValueError:
                break
            blob = data[pos : pos + size]
            if len(blob) != size:
                break
            pos += size
            fields.append((field_number, wire_type, blob))
        elif wire_type == 5:
            if pos + 4 > len(data):
                break
            fields.append((field_number, wire_type, data[pos : pos + 4]))
            pos += 4
        else:
            break

    return fields


def snappy_raw_decompress(data: bytes) -> bytes:
    _, pos = read_varint(data, 0)
    out = bytearray()
    while pos < len(data):
        tag = data[pos]
        pos += 1
        wire_type = tag & 0x03

        if wire_type == 0:
            literal_length = tag >> 2
            if literal_length < 60:
                literal_length += 1
            else:
                extra = literal_length - 59
                literal_length = int.from_bytes(data[pos : pos + extra], "little") + 1
                pos += extra
            out.extend(data[pos : pos + literal_length])
            pos += literal_length
        elif wire_type == 1:
            length = 4 + ((tag >> 2) & 0x07)
            offset = ((tag & 0xE0) << 3) | data[pos]
            pos += 1
            for _ in range(length):
                out.append(out[-offset])
        elif wire_type == 2:
            length = 1 + (tag >> 2)
            offset = int.from_bytes(data[pos : pos + 2], "little")
            pos += 2
            for _ in range(length):
                out.append(out[-offset])
        else:
            length = 1 + (tag >> 2)
            offset = int.from_bytes(data[pos : pos + 4], "little")
            pos += 4
            for _ in range(length):
                out.append(out[-offset])

    return bytes(out)


def decode_iwa_messages(blob: bytes, source: str) -> list[DecodedMessage]:
    messages: list[DecodedMessage] = []
    pos = 0

    while pos + 4 <= len(blob):
        compressed_length = int.from_bytes(blob[pos + 1 : pos + 4], "little")
        body_start = pos + 4
        body_end = body_start + compressed_length
        if body_end > len(blob):
            break

        decompressed = snappy_raw_decompress(blob[body_start:body_end])
        chunk_pos = 0
        while chunk_pos < len(decompressed):
            try:
                archive_length, archive_pos = read_varint(decompressed, chunk_pos)
            except ValueError:
                break
            archive_end = archive_pos + archive_length
            archive_info = decompressed[archive_pos:archive_end]
            archive_fields = parse_proto_fields(archive_info)

            object_id = next(
                int(value)
                for field, wire, value in archive_fields
                if field == 1 and wire == 0
            )
            message_infos = [
                parse_proto_fields(value)
                for field, wire, value in archive_fields
                if field == 2 and wire == 2
            ]

            body_pos = archive_end
            for message_info in message_infos:
                message_type = next(
                    int(value)
                    for field, wire, value in message_info
                    if field == 1 and wire == 0
                )
                body_length = next(
                    int(value)
                    for field, wire, value in message_info
                    if field == 3 and wire == 0
                )
                body = decompressed[body_pos : body_pos + body_length]
                body_pos += body_length
                messages.append(
                    DecodedMessage(
                        object_id=object_id,
                        message_type=message_type,
                        body=body,
                        source=source,
                    )
                )

            chunk_pos = body_pos
        pos = body_end

    return messages


def normalize_text(text: str) -> str:
    text = text.replace("\ufffc", " ")
    text = " ".join(text.split())
    return text.strip(" \t\r\n\"',*")


def looks_like_text(text: str, min_chars: int = 1) -> bool:
    if len(text) < min_chars:
        return False

    letters = sum(ch.isalpha() for ch in text)
    digits = sum(ch.isdigit() for ch in text)
    if letters + digits == 0:
        return False

    visible = sum(ch in PRINTABLE_ASCII or ord(ch) > 127 for ch in text)
    return visible / max(len(text), 1) >= 0.95


def iter_iwa_payloads(blob: bytes) -> list[bytes]:
    payloads: list[bytes] = []
    pos = 0

    while pos + 8 <= len(blob):
        chunk_len = int.from_bytes(blob[pos + 1 : pos + 4], "little")
        chunk_end = pos + 4 + chunk_len
        if chunk_len < 4 or chunk_end > len(blob):
            break

        payload_start = pos + 8
        payloads.append(blob[payload_start:chunk_end])
        pos = chunk_end

    return payloads or [blob]


def consume_utf8_char(data: bytes, pos: int) -> tuple[str, int] | None:
    first = data[pos]

    if 0x20 <= first <= 0x7E:
        return chr(first), pos + 1

    if first in (0x09, 0x0A, 0x0D):
        return " ", pos + 1

    if 0xC2 <= first <= 0xDF and pos + 1 < len(data):
        size = 2
    elif 0xE0 <= first <= 0xEF and pos + 2 < len(data):
        size = 3
    elif 0xF0 <= first <= 0xF4 and pos + 3 < len(data):
        size = 4
    else:
        return None

    chunk = data[pos : pos + size]
    try:
        char = chunk.decode("utf-8")
    except UnicodeDecodeError:
        return None

    if not char.isprintable():
        return None

    return char, pos + size


def iter_visible_strings(data: bytes, min_chars: int = 4) -> list[str]:
    found: list[str] = []
    pos = 0

    while pos < len(data):
        item = consume_utf8_char(data, pos)
        if item is None:
            pos += 1
            continue

        chars: list[str] = []
        next_pos = pos
        while next_pos < len(data):
            item = consume_utf8_char(data, next_pos)
            if item is None:
                break
            char, next_pos = item
            chars.append(char)

        text = normalize_text("".join(chars))
        if looks_like_text(text, min_chars=min_chars):
            found.append(text)

        pos = next_pos if next_pos > pos else pos + 1

    return found


def parse_top_level_message(
    payload: bytes,
) -> tuple[int | None, list[tuple[int, int, object]]]:
    try:
        object_id, pos = read_varint(payload, 0)
    except ValueError:
        return None, []

    fields: list[tuple[int, int, object]] = []
    while pos < len(payload):
        try:
            tag, pos = read_varint(payload, pos)
        except ValueError:
            break

        field_number = tag >> 3
        wire_type = tag & 0x07
        if field_number == 0:
            break

        if wire_type == 0:
            value, pos = read_varint(payload, pos)
            fields.append((field_number, wire_type, value))
        elif wire_type == 1:
            if pos + 8 > len(payload):
                break
            fields.append((field_number, wire_type, payload[pos : pos + 8]))
            pos += 8
        elif wire_type == 2:
            try:
                size, pos = read_varint(payload, pos)
            except ValueError:
                break
            blob = payload[pos : pos + size]
            if len(blob) != size:
                break
            pos += size
            fields.append((field_number, wire_type, blob))
        elif wire_type == 5:
            if pos + 4 > len(payload):
                break
            fields.append((field_number, wire_type, payload[pos : pos + 4]))
            pos += 4
        else:
            break

    return object_id, fields


def extract_table_record_text(blob: bytes) -> list[str]:
    texts: list[str] = []
    pos = 0
    while pos < len(blob):
        try:
            tag, pos = read_varint(blob, pos)
        except ValueError:
            break
        field_number = tag >> 3
        wire_type = tag & 0x07
        if field_number == 0:
            break

        if wire_type == 0:
            _, pos = read_varint(blob, pos)
        elif wire_type == 2:
            size, pos = read_varint(blob, pos)
            nested = blob[pos : pos + size]
            pos += size
            text = normalize_text(nested.decode("utf-8", errors="ignore"))
            if looks_like_text(text):
                texts.append(text)
        elif wire_type == 1:
            pos += 8
        elif wire_type == 5:
            pos += 4
        else:
            break

    return texts


def extract_table_scalar(fields: list[tuple[int, int, object]]) -> int | None:
    scalars = [
        (field, value)
        for field, wire, value in fields
        if wire == 0 and isinstance(value, int)
    ]
    if not scalars:
        return None

    # Pages table cell records consistently expose the displayed scalar as field 1.
    # Other small varints are flags and formatting markers.
    for field, value in scalars:
        if field == 1:
            return value

    return None


def extract_table_width(fields: list[tuple[int, int, object]]) -> int | None:
    numeric_fields = [
        (field, value)
        for field, wire, value in fields
        if wire == 0 and isinstance(value, int)
    ]
    for field, value in numeric_fields:
        if field == 2 and value > 0:
            return value
    return None


def normalize_header_tokens(tokens: list[str]) -> list[str]:
    normalized: list[str] = []
    prev_ascii: str | None = None

    for token in tokens:
        if len(token) == 1 and token in string.ascii_uppercase:
            normalized.append(token)
            prev_ascii = token
            continue

        if len(token) == 1 and token.isalpha():
            # Prefer ASCII header sequences when available and ignore stray non-ASCII
            # glyphs that appear next to encoded text boundaries.
            continue

        ascii_letters = [char for char in token if char in string.ascii_uppercase]
        if not ascii_letters:
            continue

        if prev_ascii is not None:
            target = ord(prev_ascii) + 1
            chosen = min(
                ascii_letters,
                key=lambda char: (abs(ord(char) - target), -ord(char)),
            )
        else:
            chosen = ascii_letters[-1]

        normalized.append(chosen)
        prev_ascii = chosen

    return list(OrderedDict.fromkeys(normalized))


def extract_table_headers(payload: bytes) -> list[str]:
    tokens = [
        text
        for text in iter_visible_strings(payload, min_chars=1)
        if text.isalpha() and len(text) <= 8
    ]

    filtered = [token for token in tokens if not (len(token) == 1 and token.islower())]
    return normalize_header_tokens(filtered)


def decode_table_string_table(messages: list[DecodedMessage]) -> dict[int, str]:
    strings: dict[int, str] = {}
    for message in messages:
        if message.message_type != 6005:
            continue

        body_fields = parse_proto_fields(message.body)
        list_type = next(
            (
                int(value)
                for field, wire, value in body_fields
                if field == 1 and wire == 0
            ),
            None,
        )
        if list_type != 1:
            continue

        for field, wire, value in body_fields:
            if field != 3 or wire != 2 or not isinstance(value, bytes):
                continue
            entry_fields = parse_proto_fields(value)
            key = next(
                (
                    int(entry_value)
                    for entry_field, entry_wire, entry_value in entry_fields
                    if entry_field == 1 and entry_wire == 0
                ),
                None,
            )
            text_blob = next(
                (
                    entry_value
                    for entry_field, entry_wire, entry_value in entry_fields
                    if entry_field == 3 and entry_wire == 2
                ),
                None,
            )
            if key is None or not isinstance(text_blob, bytes):
                continue
            strings[key] = text_blob.decode("utf-8", errors="replace")

    return strings


def is_root_string_table_message(message: DecodedMessage) -> bool:
    if message.message_type != 6005:
        return False

    body_fields = parse_proto_fields(message.body)
    list_type = next(
        (int(value) for field, wire, value in body_fields if field == 1 and wire == 0),
        None,
    )
    entry_count = sum(
        1
        for field, wire, value in body_fields
        if field == 3 and wire == 2 and isinstance(value, bytes)
    )
    return list_type == 1 and entry_count > 0


def is_root_tile_message(message: DecodedMessage) -> bool:
    if message.message_type != 6002:
        return False

    row_count = sum(
        1
        for field, wire, value in parse_proto_fields(message.body)
        if field == 5 and wire == 2
    )
    return row_count > 0


def pair_table_roots(
    messages: list[DecodedMessage],
) -> list[tuple[DecodedMessage, DecodedMessage | None]]:
    tile_roots = sorted(
        [message for message in messages if is_root_tile_message(message)],
        key=lambda message: message.object_id,
    )
    string_roots = sorted(
        [message for message in messages if is_root_string_table_message(message)],
        key=lambda message: message.object_id,
    )

    if not tile_roots:
        return []

    pairs: list[tuple[DecodedMessage, DecodedMessage | None]] = []
    remaining_strings = string_roots.copy()
    for tile in tile_roots:
        paired: DecodedMessage | None = None
        if remaining_strings:
            paired = min(
                remaining_strings,
                key=lambda message: (
                    abs(message.object_id - tile.object_id),
                    message.object_id,
                ),
            )
            remaining_strings.remove(paired)
        pairs.append((tile, paired))

    return pairs


def collect_varint_hits(blob: bytes, targets: set[int]) -> list[int]:
    hits: list[int] = []
    try:
        fields = parse_proto_fields(blob)
    except Exception:
        return hits

    for field, wire, value in fields:
        if wire == 0 and isinstance(value, int) and value in targets:
            hits.append(value)
        elif wire == 2 and isinstance(value, bytes):
            hits.extend(collect_varint_hits(value, targets))
    return hits


def decode_table_owners(
    messages: list[DecodedMessage], tile_ids: set[int], string_table_ids: set[int]
) -> dict[int, TableOwner]:
    owners: dict[int, TableOwner] = {}

    for message in messages:
        if message.message_type != 6001:
            continue

        refs = collect_varint_hits(message.body, tile_ids | string_table_ids)
        tile_object_id = next((ref for ref in refs if ref in tile_ids), None)
        string_table_object_id = next(
            (ref for ref in refs if ref in string_table_ids), None
        )
        if tile_object_id is None:
            continue

        fields = parse_proto_fields(message.body)
        title_blob = next(
            (
                value
                for field, wire, value in fields
                if field == 8 and wire == 2 and isinstance(value, bytes)
            ),
            None,
        )
        row_count = next(
            (int(value) for field, wire, value in fields if field == 6 and wire == 0),
            None,
        )
        column_count = next(
            (int(value) for field, wire, value in fields if field == 7 and wire == 0),
            None,
        )
        title = title_blob.decode("utf-8", errors="replace") if title_blob else None

        owners[message.object_id] = TableOwner(
            owner_id=message.object_id,
            tile_object_id=tile_object_id,
            string_table_object_id=string_table_object_id,
            title=title,
            row_count=row_count,
            column_count=column_count,
        )

    return owners


def decode_decimal128_integer(raw: bytes) -> str:
    if len(raw) != 16:
        return raw.hex()

    # Pages stores small integers in a decimal128 form whose low 32 bits carry the
    # integer value and whose high bytes are constant in these table cells.
    if raw[4:14] == b"\x00" * 10 and raw[14:] == b"\x40\x30":
        return str(int.from_bytes(raw[:4], "little"))

    return str(int.from_bytes(raw[:4], "little"))


def decode_v5_cell(cell: bytes, string_table: dict[int, str]) -> str:
    if len(cell) < 12 or cell[0] != 5:
        return ""

    cell_type = cell[1]
    bitmask = int.from_bytes(cell[8:12], "little")
    pos = 12
    decoded_fields: dict[int, bytes] = {}

    for mask, size in (
        (0x000001, 16),
        (0x000002, 8),
        (0x000004, 8),
        (0x000008, 4),
        (0x000010, 4),
        (0x000020, 4),
        (0x000040, 4),
        (0x000080, 4),
        (0x000100, 4),
        (0x000200, 4),
        (0x000400, 4),
        (0x000800, 4),
        (0x001000, 4),
        (0x002000, 4),
        (0x004000, 4),
        (0x008000, 4),
        (0x010000, 4),
        (0x020000, 4),
        (0x040000, 4),
        (0x080000, 4),
        (0x100000, 4),
    ):
        if bitmask & mask:
            decoded_fields[mask] = cell[pos : pos + size]
            pos += size

    if cell_type == 3 and 0x000008 in decoded_fields:
        string_id = int.from_bytes(decoded_fields[0x000008], "little")
        return string_table.get(string_id, f"<str:{string_id}>")

    if cell_type == 2:
        if 0x000001 in decoded_fields:
            return decode_decimal128_integer(decoded_fields[0x000001])
        if 0x000002 in decoded_fields:
            number = struct.unpack("<d", decoded_fields[0x000002])[0]
            return str(int(number) if number.is_integer() else number)

    return ""


def decode_tile_tables(
    messages: list[DecodedMessage], string_table: dict[int, str]
) -> list[Table]:
    row_map: dict[int, list[str]] = {}
    width = 0

    for message in messages:
        if message.message_type != 6002:
            continue

        for field, wire, value in parse_proto_fields(message.body):
            if field != 5 or wire != 2 or not isinstance(value, bytes):
                continue

            row_fields = parse_proto_fields(value)
            row_index = next(
                (int(v) for f, w, v in row_fields if f == 1 and w == 0),
                None,
            )
            if row_index is None:
                continue

            storage = next(
                (v for f, w, v in row_fields if f == 6 and w == 2),
                None,
            ) or next((v for f, w, v in row_fields if f == 3 and w == 2), None)
            offsets_blob = next(
                (v for f, w, v in row_fields if f == 7 and w == 2),
                None,
            ) or next((v for f, w, v in row_fields if f == 4 and w == 2), None)

            if not isinstance(storage, bytes) or not isinstance(offsets_blob, bytes):
                continue

            offsets = [
                int.from_bytes(offsets_blob[index : index + 2], "little")
                for index in range(0, len(offsets_blob), 2)
            ]
            last_used = max(
                (index for index, offset in enumerate(offsets) if offset != 0xFFFF),
                default=-1,
            )
            if last_used < 0:
                continue

            width = max(width, last_used + 1)
            row_cells = [""] * (last_used + 1)

            for column in range(last_used + 1):
                start = offsets[column]
                if start == 0xFFFF:
                    continue
                next_offsets = [
                    offset
                    for offset in offsets[column + 1 : last_used + 1]
                    if offset != 0xFFFF
                ]
                end = next_offsets[0] if next_offsets else len(storage)
                row_cells[column] = decode_v5_cell(storage[start:end], string_table)

            row_map[row_index] = row_cells

    if not row_map:
        return []

    ordered_rows = [row_map[index] for index in sorted(row_map)]
    padded_rows = [row + [""] * (width - len(row)) for row in ordered_rows]

    header_row = padded_rows[0]
    body_rows = padded_rows[1:]
    raw_values = [cell for row in body_rows for cell in row if cell != ""]
    metadata = {
        "width": width,
        "row_count": len(body_rows),
        "decoder": "full_tile_map_v5",
        "string_table": string_table,
        "row_indices": sorted(row_map),
        "object_ids": sorted({message.object_id for message in messages}),
        "sources": sorted({message.source for message in messages}),
    }
    return [
        Table(
            headers=header_row, rows=body_rows, raw_values=raw_values, metadata=metadata
        )
    ]


def extract_tables_from_pages(pages_path: Path) -> list[Table]:
    table_messages: list[DecodedMessage] = []
    calc_messages: list[DecodedMessage] = []
    with zipfile.ZipFile(pages_path) as archive:
        for name in archive.namelist():
            if not name.startswith("Index/") or not name.endswith(".iwa"):
                continue

            try:
                messages = decode_iwa_messages(archive.read(name), source=name)
            except Exception:
                continue
            if name.startswith("Index/Tables/"):
                table_messages.extend(messages)
            elif name.startswith("Index/CalculationEngine"):
                calc_messages.extend(messages)

    tables: list[Table] = []
    root_pairs = pair_table_roots(table_messages)
    tile_ids = {tile_root.object_id for tile_root, _ in root_pairs}
    string_table_ids = {
        string_root.object_id
        for _, string_root in root_pairs
        if string_root is not None
    }
    owners = decode_table_owners(calc_messages, tile_ids, string_table_ids)
    tile_to_owner = {
        owner.tile_object_id: owner
        for owner in owners.values()
        if owner.tile_object_id is not None
    }

    for tile_root, string_root in root_pairs:
        paired_messages = [tile_root]
        string_table: dict[int, str] = {}
        if string_root is not None:
            paired_messages.append(string_root)
            string_table = decode_table_string_table([string_root])

        decoded = decode_tile_tables([tile_root], string_table)
        if not decoded:
            continue

        table = decoded[0]
        metadata = dict(table.metadata)
        owner = tile_to_owner.get(tile_root.object_id)
        metadata["tile_object_id"] = tile_root.object_id
        metadata["tile_source"] = tile_root.source
        if string_root is not None:
            metadata["string_table_object_id"] = string_root.object_id
            metadata["string_table_source"] = string_root.source
        if owner is not None:
            metadata["owner_object_id"] = owner.owner_id
            metadata["table_title"] = owner.title
            metadata["declared_row_count"] = owner.row_count
            metadata["declared_column_count"] = owner.column_count
        tables.append(
            Table(
                headers=table.headers,
                rows=table.rows,
                raw_values=table.raw_values,
                metadata=metadata,
            )
        )

    return tables


def extract_row_count_hint(
    candidates: list[Candidate], width: int, value_count: int
) -> int | None:
    if width <= 0 or value_count <= 0:
        return None

    minimum_rows = math.ceil(value_count / width)
    hints = sorted(
        {
            int(candidate.text)
            for candidate in candidates
            if candidate.kind == "table_text" and candidate.text.isdigit()
        }
    )
    for hint in hints:
        if minimum_rows <= hint <= minimum_rows + 2:
            return hint
    return None


def score_candidate(candidate: Candidate) -> tuple[int, int, int]:
    text = candidate.text
    words = text.split()
    metadata_penalty = any(
        marker in text
        for marker in (
            "Application/Blank",
            "DocumentMetadata",
            "CalculationEngine",
            "gregorian",
            "HelveticaNeue",
            "paragraphStyle",
            "Stylesheet",
        )
    )
    return (
        20 if candidate.kind == "document" and len(words) >= 4 else 0,
        10 if candidate.kind == "table_header" else 0,
        len(words) * 3 + len(text) - (30 if metadata_penalty else 0),
    )


def is_plain_document_line(candidate: Candidate) -> bool:
    if candidate.kind != "document":
        return False
    if candidate.source != "Index/Document.iwa":
        return False

    text = candidate.text
    words = text.split()
    alpha = sum(ch.isalpha() for ch in text)
    bad_markers = (
        "/",
        "\\",
        "_",
        ".jpg",
        ".pdf",
        "Stylesheet",
        "paragraphStyle",
        "HeaderStorageBucket",
        "CalculationEngine",
        "DocumentMetadata",
        "Europe/Berlin",
    )
    if any(marker in text for marker in bad_markers):
        return False
    if len(words) < 3:
        return False
    if alpha < 12:
        return False

    punctuation = sum(ch in ".,;:!?" for ch in text)
    uppercase_words = sum(word[:1].isupper() for word in words if word)
    digits = sum(ch.isdigit() for ch in text)

    # Prefer prose-like text and reject formatter fragments or locale snippets.
    return punctuation > 0 or (
        len(words) >= 5 and digits == 0 and uppercase_words < len(words)
    )


def clean_document_line(text: str) -> str:
    # Pages sometimes leaks a small numeric run directly into the prose payload.
    # Strip only a short digit prefix when it is glued to a normal sentence start.
    return re.sub(r"^\d{1,3}(?=[A-ZÄÖÜ])", "", text)


def clean_document_segment(text: str) -> str:
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n")
    cleaned = cleaned.replace("\ufffc", "")
    cleaned = cleaned.strip("\n")
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    lines = [clean_document_line(line.strip()) for line in cleaned.split("\n")]
    lines = [line for line in lines if line]
    return "\n".join(lines).strip()


def extract_document_flow(pages_path: Path) -> DocumentFlow | None:
    best: DocumentFlow | None = None

    with zipfile.ZipFile(pages_path) as archive:
        raw = archive.read("Index/Document.iwa")
        messages = decode_iwa_messages(raw, source="Index/Document.iwa")

    for message in messages:
        for field, wire, value in parse_proto_fields(message.body):
            if field != 3 or wire != 2 or not isinstance(value, bytes):
                continue
            try:
                text = value.decode("utf-8")
            except UnicodeDecodeError:
                continue

            placeholder_count = text.count("\ufffc")
            if placeholder_count <= 0:
                continue

            segments = [clean_document_segment(part) for part in text.split("\ufffc")]
            alpha_count = sum(ch.isalpha() for ch in text)
            if alpha_count < 12:
                continue

            candidate = DocumentFlow(
                segments=segments,
                placeholder_count=placeholder_count,
                source_object_id=message.object_id,
            )
            if best is None:
                best = candidate
                continue

            best_score = (
                best.placeholder_count,
                sum(len(part) for part in best.segments),
            )
            candidate_score = (
                candidate.placeholder_count,
                sum(len(part) for part in candidate.segments),
            )
            if candidate_score > best_score:
                best = candidate

    return best


def extract_candidates(pages_path: Path) -> list[Candidate]:
    candidates: list[Candidate] = []

    with zipfile.ZipFile(pages_path) as archive:
        for name in archive.namelist():
            if not name.startswith("Index/") or not name.endswith(".iwa"):
                continue

            raw = archive.read(name)
            for payload in iter_iwa_payloads(raw):
                if name.startswith("Index/Tables/"):
                    object_id, fields = parse_top_level_message(payload)
                    if fields:
                        width = extract_table_width(fields)
                        if width is not None and name.endswith("DataList.iwa"):
                            candidates.append(
                                Candidate(name, "table_width", str(width))
                            )

                        if name.endswith("DataList.iwa"):
                            for header in extract_table_headers(payload):
                                candidates.append(
                                    Candidate(name, "table_header", header)
                                )

                        scalar = extract_table_scalar(fields)
                        if (
                            scalar is not None
                            and "DataList-" in name
                            and name.endswith("-2.iwa")
                        ):
                            candidates.append(
                                Candidate(name, "table_scalar", str(scalar))
                            )

                    if "HeaderStorageBucket" in name:
                        for text in iter_visible_strings(payload, min_chars=1):
                            if text.isdigit():
                                candidates.append(Candidate(name, "table_text", text))
                    continue

                for text in iter_visible_strings(payload, min_chars=4):
                    candidates.append(Candidate(name, "document", text))

    return candidates


def dedupe_preserve_best(candidates: list[Candidate]) -> list[Candidate]:
    best: OrderedDict[tuple[str, str], Candidate] = OrderedDict()
    for candidate in candidates:
        key = (candidate.kind, candidate.text)
        current = best.get(key)
        if current is None or score_candidate(candidate) > score_candidate(current):
            best[key] = candidate

    return list(best.values())


def render_output(
    candidates: list[Candidate],
    tables: list[Table],
    show_source: bool,
    document_flow: DocumentFlow | None = None,
) -> str:
    buffer = io.StringIO()

    def write(line: str) -> None:
        buffer.write(line)
        buffer.write("\n")

    def format_row(cells: list[str]) -> str:
        return " | ".join(cells)

    def write_table(table: Table) -> None:
        write(format_row(table.headers))
        for row in table.rows:
            write(format_row(row))

    if document_flow is not None:
        table_index = 0
        wrote_any = False

        for segment_index, segment in enumerate(document_flow.segments):
            if segment:
                if wrote_any:
                    write("")
                if show_source:
                    write(
                        f"[Index/Document.iwa#{document_flow.source_object_id}] {segment}"
                    )
                else:
                    for line in segment.split("\n"):
                        write(line)
                wrote_any = True

            if segment_index < document_flow.placeholder_count and table_index < len(
                tables
            ):
                if wrote_any:
                    write("")
                write_table(tables[table_index])
                table_index += 1
                wrote_any = True

        while table_index < len(tables):
            if wrote_any:
                write("")
            write_table(tables[table_index])
            table_index += 1
            wrote_any = True

        return buffer.getvalue()

    document_lines = [c for c in candidates if is_plain_document_line(c)]
    if not document_lines:
        document_lines = [
            c
            for c in candidates
            if c.kind == "document" and c.source == "Index/Document.iwa"
        ][:1]

    if document_lines:
        for item in document_lines:
            cleaned = clean_document_line(item.text)
            if show_source:
                write(f"[{item.source}] {cleaned}")
            else:
                write(cleaned)

    for index, table in enumerate(tables, start=1):
        if (document_lines and index == 1) or index > 1:
            write("")
        write_table(table)

    return buffer.getvalue()


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extract text from Apple Pages IWA content"
    )
    parser.add_argument("pages_file", type=Path, help="Path to the .pages file")
    parser.add_argument(
        "--show-source", action="store_true", help="Show the source IWA member"
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Path to the output text file. Defaults to a .txt file next to this script.",
    )
    args = parser.parse_args()

    ranked = dedupe_preserve_best(extract_candidates(args.pages_file))
    if not ranked:
        raise SystemExit("No content extracted from Index/*.iwa")

    output_path = args.output or Path(__file__).resolve().with_name(
        f"{args.pages_file.stem}_extracted.txt"
    )
    tables = extract_tables_from_pages(args.pages_file)
    document_flow = extract_document_flow(args.pages_file)
    output_text = render_output(
        ranked,
        tables,
        show_source=args.show_source,
        document_flow=document_flow,
    )
    output_path.write_text(output_text, encoding="utf-8")
    print(output_path)
    return 0


def build_tables(candidates: list[Candidate]) -> list[Table]:
    headers = [
        c.text
        for c in candidates
        if c.kind == "table_header" and len(c.text) <= 16 and looks_like_text(c.text)
    ]
    if not headers:
        return []

    scalar_values = [
        int(c.text) for c in candidates if c.kind == "table_scalar" and c.text.isdigit()
    ]
    extra_numeric_text = [
        int(c.text)
        for c in candidates
        if c.kind in {"table_text", "document"} and c.text.isdigit()
    ]

    ordered_unique_headers = list(OrderedDict.fromkeys(headers))
    ordered_unique_values = list(
        OrderedDict.fromkeys(scalar_values + extra_numeric_text)
    )
    numeric_values = sorted(
        value for value in ordered_unique_values if isinstance(value, int)
    )

    width = len(ordered_unique_headers)
    if width <= 0:
        return []

    row_count_hint = extract_row_count_hint(candidates, width, len(numeric_values))
    rows, reconstruction = infer_table_rows(
        ordered_unique_headers,
        numeric_values,
        row_count_hint=row_count_hint,
    )
    raw_values = [str(value) for value in numeric_values]
    metadata = {
        "width": width,
        "row_count_hint": row_count_hint,
        "raw_values": numeric_values,
        "reconstruction": reconstruction,
        "hypothesis": (
            "Pages appears to store table cells sparsely: table headers and dimensions live "
            "in Index/Tables/DataList.iwa, populated cell values in DataList-*-2.iwa, and "
            "empty cells are omitted from the scalar stream. When numeric values form an "
            "almost contiguous ascending sequence, missing integers are treated as blank cells."
        ),
    }
    return [
        Table(
            headers=ordered_unique_headers,
            rows=rows,
            raw_values=raw_values,
            metadata=metadata,
        )
    ]


def infer_table_rows(
    headers: list[str],
    values: list[int],
    row_count_hint: int | None = None,
) -> tuple[list[list[str]], dict[str, object]]:
    if not headers or not values:
        return [], {"mode": "empty"}

    width = len(headers)
    if width <= 0:
        return []

    # Common Pages numeric tables often serialize cell values as a mostly complete
    # integer sequence. Prefer a contiguous range when one can be inferred.
    value_set = set(values)
    if values and min(value_set) == 1:
        contiguous = []
        current = 1
        while current in value_set:
            contiguous.append(current)
            current += 1
        if len(contiguous) >= max(width, len(values) - 2):
            values = contiguous

    target_rows = max(math.ceil(len(values) / width), row_count_hint or 0)
    if target_rows <= 0:
        target_rows = math.ceil(len(values) / width)

    expanded_values: list[str] = [str(value) for value in values]
    reconstruction: dict[str, object] = {
        "mode": "dense_row_major",
        "target_rows": target_rows,
        "width": width,
    }

    if values == sorted(values) and len(values) >= 3:
        minimum = values[0]
        maximum = values[-1]
        missing_count = (maximum - minimum + 1) - len(values)
        if minimum == 1 and 0 < missing_count <= width:
            present = set(values)
            expanded_values = [
                str(number) if number in present else ""
                for number in range(minimum, maximum + 1)
            ]
            reconstruction = {
                "mode": "sparse_sequence_with_blanks",
                "target_rows": target_rows,
                "width": width,
                "sequence_start": minimum,
                "sequence_end": maximum,
                "missing_numbers": [
                    number
                    for number in range(minimum, maximum + 1)
                    if number not in present
                ],
            }

    rows: list[list[str]] = []
    position = 0
    for _ in range(target_rows):
        row: list[str] = []
        for _ in range(width):
            if position < len(expanded_values):
                row.append(expanded_values[position])
                position += 1
            else:
                row.append("")
        rows.append(row)
    return rows, reconstruction


if __name__ == "__main__":
    raise SystemExit(main())
