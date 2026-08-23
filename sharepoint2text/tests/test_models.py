"""Tests for the normalized extraction model and versioned codec."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
from typing import cast

import pytest

from sharepoint2text import read_file
from sharepoint2text.parsing.models import (
    Annotation,
    Attachment,
    ContentUnit,
    DocumentMetadata,
    ExtractedDocument,
    ImageAsset,
    JsonValue,
    SourceMetadata,
    Table,
    document_from_dict,
    document_from_json,
    document_to_dict,
    document_to_json,
    render_markdown,
)

_REAL_FORMAT_SAMPLES = [
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


@pytest.mark.parametrize("relative_path", _REAL_FORMAT_SAMPLES)
def test_real_supported_format_round_trips_v2(relative_path: str) -> None:
    """Verify representative real files round-trip through the public codec."""
    resource = Path(__file__).parent / "resources" / relative_path
    results = list(read_file(resource))

    assert results
    for document in results:
        payload = document_to_dict(document, binary="base64")
        assert document_from_dict(payload) == document


def test_core_invariants_and_canonical_iteration() -> None:
    """Verify one-based numbering and canonical asset order."""
    first = ImageAsset(number=1, data=b"first")
    second = ImageAsset(number=2, data=b"second")
    unit_table = Table(rows=[["unit"]])
    document_table = Table(rows=[["document"]])
    document = ExtractedDocument(
        format="pdf",
        units=[
            ContentUnit(
                number=1,
                kind="page",
                text="first",
                images=[first],
                tables=[unit_table],
            ),
            ContentUnit(
                number=2, kind="page", text="", images=[second], tables=[document_table]
            ),
        ],
    )

    assert document.full_text == "first"
    assert list(document.iter_images()) == [first, second]
    assert list(document.iter_tables()) == [unit_table, document_table]
    assert Table(rows=[[1], [2, 3]]).dimensions == (2, 2)
    with pytest.raises(ValueError, match="Unit numbers"):
        ContentUnit(number=0, kind="page")
    with pytest.raises(ValueError, match="Image numbers"):
        ImageAsset(number=0)


def test_document_iterators_yield_unit_records_in_order() -> None:
    """Iterate over unit-owned records in document order."""
    first_image = ImageAsset(number=1, filename="first.png")
    second_image = ImageAsset(number=2, filename="second.png")
    first_table = Table(rows=[["first"]])
    second_table = Table(rows=[["second"]])
    first_annotation = Annotation(kind="note", text="first")
    second_annotation = Annotation(kind="note", text="second")
    document = ExtractedDocument(
        format="pdf",
        units=[
            ContentUnit(
                number=1,
                kind="page",
                images=[first_image],
                tables=[first_table],
                annotations=[first_annotation],
            ),
            ContentUnit(
                number=2,
                kind="page",
                images=[second_image],
                tables=[second_table],
                annotations=[second_annotation],
            ),
        ],
    )

    assert list(document.iter_images()) == [first_image, second_image]
    assert list(document.iter_tables()) == [first_table, second_table]
    assert list(document.iter_annotations()) == [first_annotation, second_annotation]
    assert list(document.iter_images())[0] is document.units[0].images[0]
    assert list(document.iter_tables())[1] is document.units[1].tables[0]
    assert list(document.iter_annotations())[1] is document.units[1].annotations[0]


def test_document_without_units_creates_fallback_unit() -> None:
    """Create a fallback document unit when no units are provided."""
    document = ExtractedDocument(format="txt")

    assert len(document.units) == 1
    assert document.units[0].number == 1
    assert document.units[0].kind == "document"
    assert document.units[0].images == []
    assert document.units[0].tables == []
    assert document.units[0].annotations == []


def test_source_serialization_omits_size_bytes() -> None:
    """Keep source size outside the normalized extraction contract."""
    document = ExtractedDocument(
        format="txt",
        source=SourceMetadata(filename="report.txt"),
    )

    payload = document_to_dict(document)

    assert payload["document"]["source"] == {"filename": "report.txt"}


def test_image_asset_derives_ratio_from_available_dimensions() -> None:
    """Derive image ratio only when both positive dimensions are available."""
    assert ImageAsset(number=1, width=600, height=300).ratio == 2.0
    assert ImageAsset(number=1, width=600).ratio is None
    assert ImageAsset(number=1, width=600, height=0).ratio is None
    with pytest.raises(ValueError, match="ratios"):
        ImageAsset(number=1, ratio=0.0)


def _rich_document() -> ExtractedDocument:
    """Build a document exercising every recursive codec record."""
    nested = ExtractedDocument(
        format="txt", units=[ContentUnit(1, "document", "nested")]
    )
    return ExtractedDocument(
        format="pdf",
        source=SourceMetadata(filename="report.pdf", media_type="application/pdf"),
        metadata=DocumentMetadata(
            title="Report", keywords=["one", "two"], properties={"pdf.pages": 1}
        ),
        units=[
            ContentUnit(
                number=1,
                kind="page",
                text="Body",
                images=[
                    ImageAsset(
                        number=1,
                        data=b"image",
                        media_type="image/png",
                        width=600,
                        height=300,
                    )
                ],
                tables=[Table(rows=[["name", "value"], ["a", 1]])],
                annotations=[Annotation(kind="comment", text="Review")],
                properties={"pdf.label": "i"},
            )
        ],
        attachments=[
            Attachment(filename="nested.txt", data=b"file", extracted_document=nested)
        ],
    )


def test_v2_codec_round_trips_binary_and_nested_documents() -> None:
    """Verify explicit base64 mode round-trips the complete model graph."""
    document = _rich_document()

    payload = document_to_dict(document, binary="base64")
    restored = document_from_dict(payload)

    assert payload["schema"] == "sharepoint2text.extraction"
    assert payload["version"] == 2
    assert restored == document
    assert document_from_json(document_to_json(document, binary="base64")) == document
    assert list(restored.iter_images())[0] is restored.units[0].images[0]
    assert list(restored.iter_tables())[0] is restored.units[0].tables[0]
    assert list(restored.iter_annotations())[0] is restored.units[0].annotations[0]


def test_v2_codec_has_no_default_binary_decode_limit() -> None:
    """Decode binary payloads without an implicit cumulative size ceiling."""

    class ReportedLargeBase64(str):
        """Simulate encoded data larger than the former limit without allocating it."""

        def __len__(self) -> int:
            """Report a size beyond the former 100 MiB decoded-data ceiling."""
            return 140 * 1024 * 1024

    document = _rich_document()
    payload = document_to_dict(document, binary="base64")
    body = cast(dict[str, JsonValue], payload["document"])
    units = cast(list[JsonValue], body["units"])
    unit = cast(dict[str, JsonValue], units[0])
    images = cast(list[JsonValue], unit["images"])
    image = cast(dict[str, JsonValue], images[0])
    image["data"] = ReportedLargeBase64(cast(str, image["data"]))

    assert document_from_dict(payload) == document
    assert document_from_dict(payload, max_binary_bytes=None) == document


def test_v2_codec_omits_binary_by_default() -> None:
    """Verify the default schema contains no implicit binary payloads."""
    payload = document_to_dict(_rich_document())
    body = cast(dict[str, JsonValue], payload["document"])
    units = cast(list[JsonValue], body["units"])
    attachments = cast(list[JsonValue], body["attachments"])
    unit = cast(dict[str, JsonValue], units[0])
    attachment = cast(dict[str, JsonValue], attachments[0])
    images = cast(list[JsonValue], unit["images"])
    image = cast(dict[str, JsonValue], images[0])
    assert "data" not in image
    assert image["width"] == 600
    assert image["height"] == 300
    assert image["ratio"] == 2.0
    assert "data" not in attachment


def test_v2_json_is_deterministic_for_property_insertion_order() -> None:
    """Verify namespaced properties serialize in deterministic key order."""
    left = ExtractedDocument(format="txt", properties={"z.value": 1, "a.value": 2})
    right = ExtractedDocument(format="txt", properties={"a.value": 2, "z.value": 1})

    assert document_to_json(left) == document_to_json(right)


def test_v2_codec_requires_namespaced_property_keys() -> None:
    """Verify extension properties cannot become an unscoped dumping ground."""
    document = ExtractedDocument(format="txt", properties={"ambiguous": True})

    with pytest.raises(ValueError, match="namespaced"):
        document_to_dict(document)


def _unsupported_version(payload: dict[str, JsonValue]) -> None:
    """Replace the supported wire-schema version."""
    payload["version"] = 3


def _unsupported_kind(payload: dict[str, JsonValue]) -> None:
    """Replace the unit discriminator with an unsupported value."""
    body = cast(dict[str, JsonValue], payload["document"])
    body["units"] = [{"number": 1, "kind": "frame"}]


@pytest.mark.parametrize(
    "mutation, message",
    [(_unsupported_version, "version"), (_unsupported_kind, "kind")],
)
def test_v2_codec_rejects_unsupported_schema_values(
    mutation: Callable[[dict[str, JsonValue]], None], message: str
) -> None:
    """Verify unsupported schema versions and discriminators fail clearly."""
    payload = document_to_dict(ExtractedDocument(format="txt"))
    mutation(payload)

    with pytest.raises(ValueError, match=message):
        document_from_dict(payload)


def test_v2_codec_rejects_invalid_or_excessive_binary() -> None:
    """Verify untrusted base64 is validated before excessive allocation."""
    payload = document_to_dict(_rich_document(), binary="base64")
    body = cast(dict[str, JsonValue], payload["document"])
    units = cast(list[JsonValue], body["units"])
    unit = cast(dict[str, JsonValue], units[0])
    images = cast(list[JsonValue], unit["images"])
    image = cast(dict[str, JsonValue], images[0])
    image["data"] = "not base64!"
    with pytest.raises(ValueError, match="valid base64"):
        document_from_dict(payload)

    limited = document_to_dict(_rich_document(), binary="base64")
    with pytest.raises(ValueError, match="exceeds 4 bytes"):
        document_from_dict(limited, max_binary_bytes=4)


def test_render_markdown_uses_normalized_units_and_tables() -> None:
    """Verify the renderer consumes only normalized public concepts."""
    document = ExtractedDocument(
        format="xlsx",
        units=[
            ContentUnit(
                number=1,
                kind="sheet",
                title="People",
                text="Names",
                tables=[Table(rows=[["Name"], ["Ada"]])],
            )
        ],
    )

    rendered = render_markdown(document)

    assert rendered.startswith("Names\n\n## Tables")
    assert "| Name |" in rendered
    assert "| Ada  |" in rendered
