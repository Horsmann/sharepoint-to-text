import io
import zipfile

import pytest

from sharepoint2text.parsing.exceptions import ExtractionZipBombError
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    ZipBombLimits,
    get_zip_bomb_limits,
    open_zipfile,
    reset_zip_bomb_limits,
    set_zip_bomb_limits,
    validate_zip_bytesio,
)


@pytest.fixture(autouse=True)
def _restore_default_limits() -> object:
    """Ensure each test starts and ends with the default process-wide limits."""
    reset_zip_bomb_limits()
    yield
    reset_zip_bomb_limits()


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


def test_set_zip_bomb_limits_applies_at_call_time() -> None:
    """Setting central limits must affect calls that pass no explicit limits."""
    buffer = _make_zip_bytesio({"a.txt": b"A" * 10_000})

    # Tighten the central limits: this container should now be rejected even
    # though no explicit ``limits`` argument is passed.
    set_zip_bomb_limits(
        ZipBombLimits(
            max_entry_compression_ratio=10.0,
            max_total_compression_ratio=10.0,
        )
    )
    with pytest.raises(ExtractionZipBombError):
        validate_zip_bytesio(buffer, source="central")

    # Relax the central limits: the same container should now pass.
    set_zip_bomb_limits(
        ZipBombLimits(
            max_entry_compression_ratio=10_000.0,
            max_total_compression_ratio=10_000.0,
        )
    )
    validate_zip_bytesio(buffer, source="central")


def test_set_zip_bomb_limits_affects_open_zipfile() -> None:
    """open_zipfile (used by OOXML/ODF contexts) honors the central limits."""
    buffer = _make_zip_bytesio({"slide.xml": b"A" * 10_000})

    set_zip_bomb_limits(ZipBombLimits(max_entry_compression_ratio=2.0))
    with pytest.raises(ExtractionZipBombError):
        open_zipfile(buffer, source="_PptxContext")


def test_explicit_limits_override_central_limits() -> None:
    """A per-call ``limits`` argument takes precedence over central limits."""
    buffer = _make_zip_bytesio({"a.txt": b"A" * 10_000})

    set_zip_bomb_limits(ZipBombLimits(max_entry_compression_ratio=2.0))
    # Explicit generous limits must win over the tight central limits.
    validate_zip_bytesio(
        buffer,
        limits=ZipBombLimits(
            max_entry_compression_ratio=10_000.0,
            max_total_compression_ratio=10_000.0,
        ),
        source="explicit",
    )


def test_reset_zip_bomb_limits_restores_defaults() -> None:
    """reset_zip_bomb_limits returns the active limits to library defaults."""
    set_zip_bomb_limits(ZipBombLimits(max_entry_compression_ratio=1.0))
    reset_zip_bomb_limits()
    assert get_zip_bomb_limits() == ZipBombLimits()


def test_set_zip_bomb_limits_rejects_wrong_type() -> None:
    """set_zip_bomb_limits validates its argument type."""
    with pytest.raises(TypeError):
        set_zip_bomb_limits(object())  # type: ignore[arg-type]


def test_limit_error_message_points_to_switch() -> None:
    """Limit-exceeded errors must guide the reader to set_zip_bomb_limits."""
    buffer = _make_zip_bytesio({"a.txt": b"A" * 10_000})

    with pytest.raises(ExtractionZipBombError) as excinfo:
        validate_zip_bytesio(
            buffer,
            limits=ZipBombLimits(max_entry_compression_ratio=2.0),
            source="test",
        )

    message = str(excinfo.value)
    assert "set_zip_bomb_limits" in message
    assert "[test]" in message
