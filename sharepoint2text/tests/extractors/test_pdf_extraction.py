import logging
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.pdf.pdf_extractor import read_pdf
from sharepoint2text.parsing.models import ExtractedDocument
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def test_pdf__1() -> None:
    path = "sharepoint2text/tests/resources/pdf/sample.pdf"
    pdf: ExtractedDocument = next(read_pdf(file_like=read_file_to_file_like(path=path)))

    tc.assertEqual(2, pdf.metadata.properties["pdf.total_pages"])
    tc.assertEqual(2, len(pdf.units))

    # Text page 1
    expected = (
        "This is a test sentence" + "\n"
        "This is a table" + "\n"
        "C1 C2" + "\n"
        "R1 V1" + "\n"
        "R2 V2"
    )
    page_1_text = pdf.units[0].text
    tc.assertEqual(
        expected.strip().replace("\n", " "), page_1_text.strip().replace("\n", " ")
    )

    # Text page 2
    expected = "This is page 2" "\n" "An image of the Google landing page"
    page_2_text = pdf.units[1].text
    tc.assertEqual(
        expected.strip().replace("\n", " "), page_2_text.strip().replace("\n", " ")
    )

    # Image data
    tc.assertEqual(0, len(pdf.units[0].images))
    tc.assertEqual(1, len(pdf.units[1].images))

    # test iterator
    tc.assertEqual(2, len(list(pdf.units)))
    tc.assertEqual(1, len(list(pdf.iter_images())))
    image = list(pdf.iter_images())[0]
    tc.assertEqual(1, image.number)
    tc.assertEqual("image/png", image.media_type)
    tc.assertEqual(910, image.width)
    tc.assertEqual(344, image.height)
    tc.assertEqual(2, image.properties["pdf.unit_number"])

    # test full text
    tc.assertEqual("This is a test sentence", pdf.full_text[:23])

    # units
    units = list(pdf.units)
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("page", units[0].kind)

    # tables
    tc.assertEqual(0, len(list(pdf.iter_tables())))


def test_pdf__2() -> None:
    path = "sharepoint2text/tests/resources/pdf/multi_image.pdf"
    pdf: ExtractedDocument = next(
        read_pdf(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(1, len(pdf.units))
    tc.assertEqual(2, len(pdf.units[0].images))

    images = pdf.units[0].images
    img_1 = images[0]
    tc.assertEqual(1, img_1.number)
    tc.assertEqual("image/png", img_1.media_type)
    tc.assertEqual(1030, img_1.width)
    tc.assertEqual(454, img_1.height)
    tc.assertEqual(1, img_1.properties["pdf.unit_number"])
    tc.assertEqual("The OpenDocument table", img_1.caption)

    img_2 = images[1]
    tc.assertEqual(2, img_2.number)
    tc.assertEqual("image/png", img_2.media_type)
    tc.assertEqual(1172, img_2.width)
    tc.assertEqual(430, img_2.height)
    tc.assertEqual(1, img_2.properties["pdf.unit_number"])
    tc.assertEqual("The modern office table", img_2.caption)

    tc.assertEqual(1, pdf.metadata.properties["pdf.total_pages"])
    tc.assertEqual("multi_image.pdf", pdf.source.filename)
    tc.assertEqual(".pdf", pdf.source.extension)

    tc.assertEqual(0, len(list(pdf.iter_tables())))

    # units
    units = list(pdf.units)
    tc.assertEqual(2, len(units[0].images))
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("page", units[0].kind)


def test_pdf__3() -> None:
    path = (
        "sharepoint2text/tests/resources/pdf/vendor-creation-form-english-version.pdf"
    )
    pdf: ExtractedDocument = next(
        read_pdf(file_like=read_file_to_file_like(path=path), path=path)
    )

    full_text = pdf.full_text
    tc.assertTrue(len(full_text) > 0)
    tc.assertIn("Supplier Registration Form", full_text)


def test_password_protected__pdf() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/pdf-password-protected-pw123.pdf"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_pdf(file_like=read_file_to_file_like(path=path), path=path))
