"""Integration tests for the normalized public extraction API."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, BinaryIO, Callable, Generator, cast

import pytest

import sharepoint2text
from sharepoint2text import (
    ExtractedDocument,
    ExtractionFileFormatNotSupportedError,
    read_bytes,
    read_file,
)

_SAMPLES = [
    "plain_text/plain.txt",
    "plain_text/plain.csv",
    "html/sample.html",
    "html/sample.mhtml",
    "pdf/sample.pdf",
    "epub/sample.epub",
    "modern_ms/sample.docm",
    "modern_ms/mwe.xlsx",
    "modern_ms/slide_titles.pptx",
    "legacy_ms/headings.doc",
    "legacy_ms/mwe.xls",
    "legacy_ms/slide_headlines.ppt",
    "legacy_ms/2025.144.un.rtf",
    "open_office/sample_document.odt",
    "open_office/sample_spreadsheet.ods",
    "open_office/sample_presentation.odp",
    "open_office/drawing.odg",
    "open_office/formular.odf",
    "mails/basic_email.eml",
    "mails/basic_email.msg",
    "mails/basic_email.mbox",
]


@pytest.mark.parametrize("relative_path", _SAMPLES)
def test_read_file_returns_only_normalized_documents(relative_path: str) -> None:
    """Verify representative formats share the same public result type."""
    path = Path(__file__).parent / "resources" / relative_path

    results = list(read_file(path))

    assert results
    assert all(isinstance(result, ExtractedDocument) for result in results)
    assert all(result.source.path for result in results)


def test_read_bytes_returns_a_normalized_document() -> None:
    """Verify in-memory extraction uses the same public result type."""
    result = next(read_bytes(b"hello", extension="txt"))

    assert isinstance(result, ExtractedDocument)
    assert result.format == "txt"
    assert result.full_text == "hello"


def test_read_bytes_skips_image_extraction_when_requested() -> None:
    """Avoid returning image records for an in-memory image-bearing document."""
    path = (
        Path(__file__).parent / "resources" / "modern_ms" / "document_with_image.docx"
    )

    result = next(read_bytes(path.read_bytes(), extension="docx", ignore_images=True))

    assert list(result.iter_images()) == []


def test_read_file_validates_missing_path_eagerly(tmp_path: Path) -> None:
    """Raise path validation errors when the public API is called."""
    with pytest.raises(FileNotFoundError):
        read_file(tmp_path / "missing.txt")


def test_read_file_validates_routing_eagerly(tmp_path: Path) -> None:
    """Reject unsupported filesystem sources when the public API is called."""
    unsupported_path = tmp_path / "document.unsupported"
    unsupported_path.write_bytes(b"content")

    with pytest.raises(ExtractionFileFormatNotSupportedError):
        read_file(unsupported_path)


def test_read_bytes_validates_configuration_eagerly() -> None:
    """Raise byte-input validation errors when the public API is called."""
    with pytest.raises(TypeError, match="data must be bytes or io.BytesIO"):
        read_bytes(cast(Any, "not bytes"), extension="txt")

    with pytest.raises(ValueError, match="Either mime_type or extension"):
        read_bytes(b"content")

    with pytest.raises(ExtractionFileFormatNotSupportedError):
        read_bytes(b"content", extension="unsupported")


def test_read_file_defers_opening_until_iteration(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Keep file I/O lazy after eager public validation succeeds."""
    path = Path(__file__).parent / "resources" / "plain_text" / "plain.txt"
    documents = read_file(path)

    def reject_open(*args: Any, **kwargs: Any) -> Any:
        """Fail when lazy extraction attempts to open the source file."""
        del args, kwargs
        raise RuntimeError("file opened")

    monkeypatch.setattr("builtins.open", reject_open)

    with pytest.raises(RuntimeError, match="file opened"):
        next(documents)


def test_read_bytes_defers_extraction_until_iteration(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Keep in-memory parsing lazy after eager public validation succeeds."""
    extraction_started = False

    def recording_extractor(
        file_like: BinaryIO, path: str | None
    ) -> Generator[Any, None, None]:
        """Record when the returned extractor is first advanced."""
        nonlocal extraction_started
        del file_like, path
        extraction_started = True
        yield from ()

    def get_recording_extractor(
        *args: Any, **kwargs: Any
    ) -> Callable[[BinaryIO, str | None], Generator[Any, None, None]]:
        """Return the recording extractor without advancing it."""
        del args, kwargs
        return recording_extractor

    monkeypatch.setattr(
        "sharepoint2text._api._get_extractor",
        get_recording_extractor,
    )

    documents = read_bytes(b"content", extension="txt")
    assert not extraction_started

    assert list(documents) == []
    assert extraction_started


def test_read_file_logs_one_debug_completion_without_info_noise(
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Aggregate multi-document extraction into one detailed completion event."""
    path = Path(__file__).parent / "resources" / "mails" / "basic_email.mbox"

    with caplog.at_level(logging.DEBUG, logger="sharepoint2text"):
        results = list(read_file(path))

    api_records = [
        record for record in caplog.records if record.name == "sharepoint2text._api"
    ]
    completion_messages = [
        record.getMessage()
        for record in api_records
        if record.getMessage().startswith("Extracted file:")
    ]

    assert len(results) == 2
    assert not [record for record in api_records if record.levelno == logging.INFO]
    assert len(completion_messages) == 1
    assert completion_messages[0].endswith("(2 documents)")


def test_package_exports_only_the_normalized_api() -> None:
    """Verify the package has one explicit normalized public surface."""
    expected = {
        "Annotation",
        "Attachment",
        "BatchFileResult",
        "ContentUnit",
        "DocumentMetadata",
        "ExtractedDocument",
        "ImageAsset",
        "SourceMetadata",
        "Table",
        "document_from_dict",
        "document_from_json",
        "document_to_dict",
        "document_to_json",
        "read_bytes",
        "read_file",
        "read_many",
        "render_markdown",
    }

    assert expected <= set(sharepoint2text.__all__)
