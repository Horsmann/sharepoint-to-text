"""Integration tests for the normalized public extraction API."""

from __future__ import annotations

import logging
from pathlib import Path

import pytest

import sharepoint2text
from sharepoint2text import ExtractedDocument, read_bytes, read_file

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
