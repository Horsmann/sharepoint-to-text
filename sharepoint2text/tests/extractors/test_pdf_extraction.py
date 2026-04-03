import logging
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    ImageMetadata,
    PdfContent,
    PdfUnitMetadata,
)
from sharepoint2text.parsing.extractors.pdf.pdf_extractor import read_pdf
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def test_pdf__1() -> None:
    path = "sharepoint2text/tests/resources/pdf/sample.pdf"
    pdf: PdfContent = next(read_pdf(file_like=read_file_to_file_like(path=path)))

    tc.assertEqual(2, pdf.metadata.total_pages)
    tc.assertEqual(2, len(pdf.pages))

    # Text page 1
    expected = (
        "This is a test sentence" + "\n"
        "This is a table" + "\n"
        "C1 C2" + "\n"
        "R1 V1" + "\n"
        "R2 V2"
    )
    page_1_text = pdf.pages[0].text
    tc.assertEqual(
        expected.strip().replace("\n", " "), page_1_text.strip().replace("\n", " ")
    )

    # Text page 2
    expected = "This is page 2" "\n" "An image of the Google landing page"
    page_2_text = pdf.pages[1].text
    tc.assertEqual(
        expected.strip().replace("\n", " "), page_2_text.strip().replace("\n", " ")
    )

    # Image data
    tc.assertEqual(0, len(pdf.pages[0].images))
    tc.assertEqual(1, len(pdf.pages[1].images))

    # test iterator
    tc.assertEqual(2, len(list(pdf.iterate_units())))
    tc.assertEqual(1, len(list(pdf.iterate_images())))
    tc.assertEqual(
        ImageMetadata(
            unit_number=2,
            image_number=1,
            content_type="image/png",
            width=910,
            height=344,
        ),
        list(pdf.iterate_images())[0].get_metadata(),
    )

    # test full text
    tc.assertEqual("This is a test sentence", pdf.get_full_text()[:23])

    # units
    units = list(pdf.iterate_units())
    tc.assertEqual(PdfUnitMetadata(unit_number=1), units[0].get_metadata())

    # tables
    tc.assertEqual(0, len(list(pdf.iterate_tables())))


def test_pdf__2() -> None:
    path = "sharepoint2text/tests/resources/pdf/multi_image.pdf"
    pdf: PdfContent = next(
        read_pdf(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(1, len(pdf.pages))
    tc.assertEqual(2, len(pdf.pages[0].images))

    images = pdf.pages[0].images
    img_1 = images[0]
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/png",
            width=1030,
            height=454,
        ),
        img_1.get_metadata(),
    )
    tc.assertEqual("The OpenDocument table", img_1.get_caption())

    img_2 = images[1]
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=2,
            content_type="image/png",
            width=1172,
            height=430,
        ),
        img_2.get_metadata(),
    )
    tc.assertEqual("The modern office table", img_2.get_caption())

    metadata = pdf.get_metadata()
    tc.assertEqual(1, metadata.total_pages)
    tc.assertEqual("multi_image.pdf", metadata.filename)
    tc.assertEqual(".pdf", metadata.file_extension)

    tc.assertEqual(0, len(list(pdf.iterate_tables())))

    # units
    units = list(pdf.iterate_units())
    tc.assertEqual(2, len(units[0].get_images()))
    tc.assertEqual(PdfUnitMetadata(unit_number=1), units[0].get_metadata())


def test_pdf__3() -> None:
    path = (
        "sharepoint2text/tests/resources/pdf/vendor-creation-form-english-version.pdf"
    )
    pdf: PdfContent = next(
        read_pdf(file_like=read_file_to_file_like(path=path), path=path)
    )

    full_text = pdf.get_full_text()
    tc.assertTrue(len(full_text) > 0)
    tc.assertIn("Supplier Registration Form", full_text)


def test_password_protected__pdf() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/pdf-password-protected-pw123.pdf"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_pdf(file_like=read_file_to_file_like(path=path), path=path))
