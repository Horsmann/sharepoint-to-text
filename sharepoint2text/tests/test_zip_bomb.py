import io
import zipfile

import pytest

from sharepoint2text.exceptions import ExtractionZipBombError
from sharepoint2text.extractors.util.zip_bomb import ZipBombLimits, validate_zip_bytesio


def _make_zip_bytesio(files: dict[str, bytes]) -> io.BytesIO:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, data in files.items():
            zf.writestr(name, data)
    buffer.seek(0)
    return buffer


def test_zip_bomb_detection_can_use_low_thresholds__compression_ratio() -> None:
    buffer = _make_zip_bytesio({"a.txt": b"A" * 10_000})

    with pytest.raises(ExtractionZipBombError):
        validate_zip_bytesio(
            buffer,
            limits=ZipBombLimits(
                max_entry_compression_ratio=10.0,
                max_total_compression_ratio=10.0,
            ),
            source="test",
        )

    validate_zip_bytesio(
        buffer,
        limits=ZipBombLimits(
            max_entry_compression_ratio=10_000.0,
            max_total_compression_ratio=10_000.0,
        ),
        source="test",
    )


def test_zip_bomb_detection_can_use_low_thresholds__entry_count() -> None:
    buffer = _make_zip_bytesio(
        {
            "a.txt": b"a",
            "b.txt": b"b",
            "c.txt": b"c",
        }
    )

    with pytest.raises(ExtractionZipBombError):
        validate_zip_bytesio(
            buffer,
            limits=ZipBombLimits(max_entries=2),
            source="test",
        )
