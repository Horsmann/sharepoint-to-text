import io
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from threading import Barrier
from typing import Any

import pytest

import sharepoint2text
from sharepoint2text import ZipBombLimits, read_bytes, read_file, read_many
from sharepoint2text.parsing.exceptions import ExtractionZipBombError
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    validate_zip_bytesio,
)


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


def test_explicit_limits_are_honored_by_low_level_helpers() -> None:
    """Low-level helper calls should continue to honor explicit limits."""
    buffer = _make_zip_bytesio({"a.txt": b"A" * 10_000})

    validate_zip_bytesio(
        buffer,
        limits=ZipBombLimits(
            max_entry_compression_ratio=10_000.0,
            max_total_compression_ratio=10_000.0,
        ),
        source="explicit",
    )


def test_read_bytes_limits_apply_only_to_one_call() -> None:
    """Relaxed limits should not affect a subsequent in-memory extraction."""
    data = _make_zip_bytesio({"a.txt": b"A" * 10_000}).getvalue()
    relaxed_limits = ZipBombLimits(
        max_entry_compression_ratio=10_000.0,
        max_total_compression_ratio=10_000.0,
    )

    documents = list(read_bytes(data, extension="zip", zip_bomb_limits=relaxed_limits))
    assert documents

    with pytest.raises(ExtractionZipBombError):
        list(read_bytes(data, extension="zip"))


def test_read_file_limits_apply_only_to_one_call(tmp_path: Path) -> None:
    """Relaxed limits should not affect a subsequent file extraction."""
    archive_path = tmp_path / "trusted.zip"
    archive_path.write_bytes(_make_zip_bytesio({"a.txt": b"A" * 10_000}).getvalue())
    relaxed_limits = ZipBombLimits(
        max_entry_compression_ratio=10_000.0,
        max_total_compression_ratio=10_000.0,
    )

    assert list(read_file(archive_path, zip_bomb_limits=relaxed_limits))

    with pytest.raises(ExtractionZipBombError):
        list(read_file(archive_path))


def test_limits_reset_while_generator_is_suspended() -> None:
    """A generator paused after one result must not retain relaxed limits."""
    archive = _make_zip_bytesio(
        {
            "first.txt": b"A" * 10_000,
            "second.txt": b"second",
        }
    )
    relaxed_limits = ZipBombLimits(
        max_entry_compression_ratio=10_000.0,
        max_total_compression_ratio=10_000.0,
    )
    documents = read_bytes(
        archive.getvalue(),
        extension="zip",
        zip_bomb_limits=relaxed_limits,
    )

    next(documents)
    try:
        with pytest.raises(ExtractionZipBombError):
            validate_zip_bytesio(archive, source="suspended")
    finally:
        documents.close()


def test_limits_reset_after_extraction_failure() -> None:
    """A failed extraction must restore defaults before propagating its error."""
    moderate_archive = _make_zip_bytesio({"a.txt": b"A" * 1_000})
    strict_limits = ZipBombLimits(
        max_entry_compression_ratio=2.0,
        max_total_compression_ratio=2.0,
    )

    with pytest.raises(ExtractionZipBombError):
        list(
            read_bytes(
                moderate_archive.getvalue(),
                extension="zip",
                zip_bomb_limits=strict_limits,
            )
        )

    validate_zip_bytesio(moderate_archive, source="after-failure")


def test_concurrent_calls_keep_limits_isolated(monkeypatch: pytest.MonkeyPatch) -> None:
    """Concurrent extraction calls must resolve their own ZIP-bomb limits."""
    import sharepoint2text.parsing.extractors.util.zip_bomb as zip_bomb_module

    data = _make_zip_bytesio({"a.txt": b"A" * 10_000}).getvalue()
    relaxed_limits = ZipBombLimits(
        max_entry_compression_ratio=10_000.0,
        max_total_compression_ratio=10_000.0,
    )
    barrier = Barrier(2)
    original_validate_zipfile = zip_bomb_module.validate_zipfile

    def synchronized_validate_zipfile(*args: Any, **kwargs: Any) -> None:
        barrier.wait(timeout=5)
        original_validate_zipfile(*args, **kwargs)

    monkeypatch.setattr(
        zip_bomb_module,
        "validate_zipfile",
        synchronized_validate_zipfile,
    )

    def extract(limits: ZipBombLimits | None) -> bool:
        try:
            list(read_bytes(data, extension="zip", zip_bomb_limits=limits))
        except ExtractionZipBombError:
            return False
        return True

    with ThreadPoolExecutor(max_workers=2) as executor:
        relaxed_result = executor.submit(extract, relaxed_limits)
        default_result = executor.submit(extract, None)

    assert relaxed_result.result() is True
    assert default_result.result() is False


def test_read_many_forwards_limits_to_each_file(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Batch extraction should forward one limits object to every file call."""
    (tmp_path / "first.txt").write_text("first")
    (tmp_path / "second.txt").write_text("second")
    limits = ZipBombLimits(max_entries=123)
    received_limits: list[ZipBombLimits | None] = []
    original_read_file = read_many.__globals__["read_file"]

    def recording_read_file(path: Any, **kwargs: Any) -> Any:
        received_limits.append(kwargs.get("zip_bomb_limits"))
        return original_read_file(path, **kwargs)

    monkeypatch.setattr("sharepoint2text._api.read_file", recording_read_file)

    documents = list(read_many(tmp_path, suffixes=[".txt"], zip_bomb_limits=limits))

    assert len(documents) == 2
    assert received_limits == [limits, limits]


def test_process_wide_limit_helpers_are_not_public() -> None:
    """The removed process-wide configuration helpers must not be exported."""
    removed_names = {
        "get_zip_bomb_limits",
        "reset_zip_bomb_limits",
        "set_zip_bomb_limits",
    }

    assert removed_names.isdisjoint(sharepoint2text.__all__)
    assert all(not hasattr(sharepoint2text, name) for name in removed_names)


def test_read_bytes_rejects_invalid_limits_type() -> None:
    """Public extraction calls should validate the per-call limits type."""
    with pytest.raises(TypeError):
        list(
            read_bytes(
                b"hello",
                extension="txt",
                zip_bomb_limits=object(),  # type: ignore[arg-type]
            )
        )


def test_limit_error_message_points_to_per_call_argument() -> None:
    """Limit errors should guide callers to the isolated public argument."""
    buffer = _make_zip_bytesio({"a.txt": b"A" * 10_000})

    with pytest.raises(ExtractionZipBombError) as excinfo:
        validate_zip_bytesio(
            buffer,
            limits=ZipBombLimits(max_entry_compression_ratio=2.0),
            source="test",
        )

    message = str(excinfo.value)
    assert "zip_bomb_limits" in message
    assert "[test]" in message
