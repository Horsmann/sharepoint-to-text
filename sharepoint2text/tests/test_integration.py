"""Integration tests for the normalized public extraction API."""

from __future__ import annotations

import importlib
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


@pytest.mark.parametrize(
    "removed_name",
    ["DocxContent", "ExtractionInterface", "normalize_extraction", "read_docx"],
)
def test_version_one_names_are_absent_from_package(removed_name: str) -> None:
    """Verify version-one records, adapters, and readers are not exported."""
    assert not hasattr(sharepoint2text, removed_name)


@pytest.mark.parametrize(
    "removed_module",
    [
        "sharepoint2text.parsing.extractors.data_types",
        "sharepoint2text.parsing.extractors.serialization",
        "sharepoint2text.parsing.models.legacy",
    ],
)
def test_version_one_module_paths_are_removed(removed_module: str) -> None:
    """Verify the prior version-one import paths no longer resolve."""
    with pytest.raises(ModuleNotFoundError):
        importlib.import_module(removed_module)
