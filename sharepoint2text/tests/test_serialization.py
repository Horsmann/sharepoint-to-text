import io
import json
import logging
import unittest

import pytest

from sharepoint2text import read_file
from sharepoint2text.parsing.extractors.data_types import (
    DocContent,
    DocxContent,
    EmailContent,
    ExtractionInterface,
    HtmlContent,
    ImageMetadata,
    OdpContent,
    OdsContent,
    OdtContent,
    PdfContent,
    PlainTextContent,
    PptContent,
    PptxContent,
    RtfContent,
    XlsContent,
    XlsxContent,
)
from sharepoint2text.parsing.extractors.serialization import deserialize_extraction

logger = logging.getLogger(__name__)

tc = unittest.TestCase()
tc.maxDiff = None


def _read_file_to_file_like(path: str) -> io.BytesIO:
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)
        return file_like


def test_serialize_for_json() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    obj = next(read_file(path))
    tc.assertIsInstance(obj, PptxContent)

    payload = obj.to_json()
    tc.assertIsInstance(payload, dict)

    try:
        json.dumps(payload)
    except Exception as e:
        tc.fail("Unexpected exception: {}".format(e))


def test_image_metadata_json_serializable() -> None:
    meta = ImageMetadata(
        unit_number=3,
        image_number=7,
        content_type="image/png",
        width=120,
        height=80,
    )

    try:
        json.dumps(meta)
    except Exception as e:
        tc.fail(f"JSON dump failed on object {e}")

    tc.assertEqual(
        {
            "unit_number": 3,
            "image_number": 7,
            "content_type": "image/png",
            "width": 120,
            "height": 80,
        },
        meta.to_dict(),
    )


def test_deserialize_pptx() -> None:
    """Test round-trip serialization/deserialization for PPTX."""
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    original = next(read_file(path))
    tc.assertIsInstance(original, PptxContent)

    # Serialize and deserialize
    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    # Verify type and content
    tc.assertIsInstance(restored, PptxContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())
    tc.assertEqual(len(restored.slides), len(original.slides))
    tc.assertEqual(restored.metadata.title, original.metadata.title)


def test_deserialize_docx() -> None:
    """Test round-trip serialization/deserialization for DOCX."""
    path = (
        "sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx"
    )
    original = next(read_file(path))
    tc.assertIsInstance(original, DocxContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, DocxContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())
    tc.assertEqual(len(restored.paragraphs), len(original.paragraphs))


def test_deserialize_xlsx() -> None:
    """Test round-trip serialization/deserialization for XLSX."""
    path = "sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx"
    original = next(read_file(path))
    tc.assertIsInstance(original, XlsxContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, XlsxContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())
    tc.assertEqual(len(restored.sheets), len(original.sheets))


def test_deserialize_pdf() -> None:
    """Test round-trip serialization/deserialization for PDF."""
    path = "sharepoint2text/tests/resources/pdf/sample.pdf"
    original = next(read_file(path))
    tc.assertIsInstance(original, PdfContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, PdfContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())
    tc.assertEqual(len(restored.pages), len(original.pages))


def test_deserialize_html() -> None:
    """Test round-trip serialization/deserialization for HTML."""
    path = "sharepoint2text/tests/resources/html/sample.html"
    original = next(read_file(path))
    tc.assertIsInstance(original, HtmlContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, HtmlContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_plain_text() -> None:
    """Test round-trip serialization/deserialization for plain text."""
    path = "sharepoint2text/tests/resources/plain_text/plain.txt"
    original = next(read_file(path))
    tc.assertIsInstance(original, PlainTextContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, PlainTextContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_odt() -> None:
    """Test round-trip serialization/deserialization for ODT."""
    path = "sharepoint2text/tests/resources/open_office/sample_document.odt"
    original = next(read_file(path))
    tc.assertIsInstance(original, OdtContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, OdtContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_odp() -> None:
    """Test round-trip serialization/deserialization for ODP."""
    path = "sharepoint2text/tests/resources/open_office/sample_presentation.odp"
    original = next(read_file(path))
    tc.assertIsInstance(original, OdpContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, OdpContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_ods() -> None:
    """Test round-trip serialization/deserialization for ODS."""
    path = "sharepoint2text/tests/resources/open_office/sample_spreadsheet.ods"
    original = next(read_file(path))
    tc.assertIsInstance(original, OdsContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, OdsContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_legacy_doc() -> None:
    """Test round-trip serialization/deserialization for legacy DOC."""
    path = "sharepoint2text/tests/resources/legacy_ms/Speech_Prime_Minister_of_The_Netherlands_EN.doc"
    original = next(read_file(path))
    tc.assertIsInstance(original, DocContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, DocContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_legacy_ppt() -> None:
    """Test round-trip serialization/deserialization for legacy PPT."""
    path = "sharepoint2text/tests/resources/legacy_ms/eurouni2.ppt"
    original = next(read_file(path))
    tc.assertIsInstance(original, PptContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, PptContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_legacy_xls() -> None:
    """Test round-trip serialization/deserialization for legacy XLS."""
    path = "sharepoint2text/tests/resources/legacy_ms/mwe.xls"
    original = next(read_file(path))
    tc.assertIsInstance(original, XlsContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, XlsContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_rtf() -> None:
    """Test round-trip serialization/deserialization for RTF."""
    path = "sharepoint2text/tests/resources/legacy_ms/2025.144.un.rtf"
    original = next(read_file(path))
    tc.assertIsInstance(original, RtfContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, RtfContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_email() -> None:
    """Test round-trip serialization/deserialization for email."""
    path = "sharepoint2text/tests/resources/mails/basic_email.eml"
    original = next(read_file(path))
    tc.assertIsInstance(original, EmailContent)

    json_data = original.to_json()
    restored = deserialize_extraction(json_data)

    tc.assertIsInstance(restored, EmailContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())
    tc.assertEqual(restored.subject, original.subject)
    tc.assertEqual(restored.from_email.address, original.from_email.address)


def test_deserialize_with_images() -> None:
    """Test that images are properly serialized/deserialized with binary data."""
    path = "sharepoint2text/tests/resources/pdf/multi_image.pdf"
    original = next(read_file(path))
    tc.assertIsInstance(original, PdfContent)

    # Verify there are images
    original_images = list(original.iterate_images())
    tc.assertGreater(len(original_images), 0)

    json_data = original.to_json()

    # Verify JSON is serializable
    json_str = json.dumps(json_data)
    parsed_json = json.loads(json_str)

    restored = deserialize_extraction(parsed_json)
    tc.assertIsInstance(restored, PdfContent)

    restored_images = list(restored.iterate_images())
    tc.assertEqual(len(restored_images), len(original_images))

    # Verify image data is preserved
    for orig_img, rest_img in zip(original_images, restored_images):
        orig_bytes = orig_img.get_bytes().read()
        rest_bytes = rest_img.get_bytes().read()
        tc.assertEqual(orig_bytes, rest_bytes)


def test_deserialize_from_json_method() -> None:
    """Test using the from_json class method on ExtractionInterface."""
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    original = next(read_file(path))

    json_data = original.to_json()
    restored = ExtractionInterface.from_json(json_data)

    tc.assertIsInstance(restored, PptxContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())


def test_deserialize_invalid_input() -> None:
    """Test error handling for invalid input."""
    with pytest.raises(ValueError, match="Input must be a dictionary"):
        deserialize_extraction("not a dict")

    with pytest.raises(ValueError, match="must contain '_type' key"):
        deserialize_extraction({})

    with pytest.raises(ValueError, match="must contain '_type' key"):
        deserialize_extraction({"some_field": "value"})


def test_json_roundtrip_full() -> None:
    """Test full JSON string round-trip (serialize -> json.dumps -> json.loads -> deserialize)."""
    path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )
    original = next(read_file(path))
    tc.assertIsInstance(original, DocxContent)

    # Full round-trip through JSON string
    json_data = original.to_json()
    json_str = json.dumps(json_data)
    parsed_data = json.loads(json_str)
    restored = deserialize_extraction(parsed_data)

    tc.assertIsInstance(restored, DocxContent)
    tc.assertEqual(restored.get_full_text(), original.get_full_text())
    tc.assertEqual(len(restored.tables), len(original.tables))
    tc.assertEqual(len(restored.comments), len(original.comments))
