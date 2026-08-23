"""Regression tests for per-call pypdf decompression limit isolation."""

from __future__ import annotations

import sys
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path
from threading import Event
from typing import BinaryIO, Generator

import pypdf.filters
import pytest

import sharepoint2text._api as extraction_api
from sharepoint2text import (
    ExtractedDocument,
    ExtractionFailedError,
    read_bytes,
    read_file,
)


def _current_pypdf_limits() -> dict[str, object]:
    """Return all pypdf limit globals supported by the installed version."""
    return {
        attribute: getattr(pypdf.filters, attribute)
        for attribute in extraction_api._PYPDF_LIMIT_ATTRIBUTES
        if hasattr(pypdf.filters, attribute)
    }


def _failing_extractor(
    file_like: BinaryIO, path: str | None
) -> Generator[object, None, None]:
    """Fail while asserting that the relaxed limit is active."""
    del file_like, path
    assert set(_current_pypdf_limits().values()) == {sys.maxsize}
    raise ValueError("simulated PDF failure")
    yield  # pragma: no cover


def _normalize_document(
    record: object, *, include_image_data: bool = True
) -> ExtractedDocument:
    """Convert a fake extractor record into a public document."""
    del record, include_image_data
    return ExtractedDocument(format="pdf")


@dataclass
class _ConcurrentLimitProbe:
    """Coordinate two fake PDF extractors and record their active limits."""

    relaxed_started: Event = field(default_factory=Event)
    release_relaxed: Event = field(default_factory=Event)
    strict_started: Event = field(default_factory=Event)
    observed_limits: dict[bytes, dict[str, object]] = field(default_factory=dict)

    def extract(
        self, file_like: BinaryIO, path: str | None
    ) -> Generator[object, None, None]:
        """Record active limits and pause the relaxed extraction."""
        del path
        marker = file_like.read()
        self.observed_limits[marker] = _current_pypdf_limits()
        if marker == b"relaxed":
            self.relaxed_started.set()
            assert self.release_relaxed.wait(timeout=2)
        else:
            self.strict_started.set()
        yield object()


def _run_concurrent_extractions(probe: _ConcurrentLimitProbe) -> None:
    """Run a relaxed PDF call concurrently with a strict PDF call."""
    with ThreadPoolExecutor(max_workers=2) as executor:
        relaxed = executor.submit(
            list, read_bytes(b"relaxed", extension="pdf", max_file_size=0)
        )
        assert probe.relaxed_started.wait(timeout=1)
        strict = executor.submit(
            list, read_bytes(b"strict", extension="pdf", max_file_size=10)
        )
        try:
            assert not probe.strict_started.wait(timeout=0.1)
        finally:
            probe.release_relaxed.set()
        relaxed.result(timeout=2)
        strict.result(timeout=2)


def test_non_pdf_extraction_does_not_modify_pypdf_limits() -> None:
    """Verify non-PDF calls cannot relax process-wide pypdf settings."""
    original_limits = _current_pypdf_limits()

    assert list(read_bytes(b"plain text", extension="txt", max_file_size=0))

    assert _current_pypdf_limits() == original_limits


def test_pdf_limits_are_restored_before_a_result_is_yielded() -> None:
    """Verify a suspended public generator exposes no relaxed pypdf limits."""
    original_limits = _current_pypdf_limits()
    pdf_path = Path(__file__).parent / "resources" / "pdf" / "sample.pdf"
    documents = read_file(pdf_path, max_file_size=0)

    try:
        assert next(documents).format == "pdf"
        assert _current_pypdf_limits() == original_limits
    finally:
        documents.close()

    assert _current_pypdf_limits() == original_limits


def test_pdf_limits_are_restored_after_extraction_failure(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify extractor exceptions cannot leak relaxed pypdf limits."""
    original_limits = _current_pypdf_limits()

    monkeypatch.setattr(
        extraction_api, "_get_extractor", lambda *args, **kwargs: _failing_extractor
    )

    with pytest.raises(ExtractionFailedError, match="Failed to extract"):
        list(read_bytes(b"%PDF", extension="pdf", max_file_size=0))

    assert _current_pypdf_limits() == original_limits


def test_concurrent_pdf_calls_cannot_observe_relaxed_limits(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify strict PDF parsing waits for a concurrent relaxed call."""
    original_limits = _current_pypdf_limits()
    probe = _ConcurrentLimitProbe()

    monkeypatch.setattr(
        extraction_api, "_get_extractor", lambda *args, **kwargs: probe.extract
    )
    monkeypatch.setattr(extraction_api, "_normalize_record", _normalize_document)

    _run_concurrent_extractions(probe)

    assert set(probe.observed_limits[b"relaxed"].values()) == {sys.maxsize}
    assert probe.observed_limits[b"strict"] == original_limits
    assert _current_pypdf_limits() == original_limits
