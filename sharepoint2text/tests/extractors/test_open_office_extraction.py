import logging
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    ImageMetadata,
    OdfContent,
    OdgContent,
    OdpContent,
    OdpUnitMetadata,
    OdsContent,
    OdtContent,
    OdtHeaderFooter,
    OdtNote,
    OdtTable,
    OdtUnitMetadata,
    OpenDocumentAnnotation,
    TableDim,
)
from sharepoint2text.parsing.extractors.open_office.odf_extractor import read_odf
from sharepoint2text.parsing.extractors.open_office.odg_extractor import read_odg
from sharepoint2text.parsing.extractors.open_office.odp_extractor import read_odp
from sharepoint2text.parsing.extractors.open_office.ods_extractor import read_ods
from sharepoint2text.parsing.extractors.open_office.odt_extractor import read_odt
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


######################
# Password protected #
######################


def test_password_protected__odt() -> None:
    path = "sharepoint2text/tests/resources/open_office/password_protected/odt-password-protected-pw123.odt"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_odt(file_like=read_file_to_file_like(path=path), path=path))


def test_password_protected__ods() -> None:
    path = "sharepoint2text/tests/resources/open_office/password_protected/ods-password-protected-pw123.ods"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_ods(file_like=read_file_to_file_like(path=path), path=path))


def test_password_protected__odp() -> None:
    path = "sharepoint2text/tests/resources/open_office/password_protected/odp-password-protected-pw123.odp"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_odp(file_like=read_file_to_file_like(path=path), path=path))


###############
# Open Office #
###############


def test_read_open_office__document() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_document.odt"
    odt: OdtContent = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(".odt", odt.get_metadata().to_dict().get("file_extension"))
    tc.assertEqual("sample_document.odt", odt.get_metadata().to_dict().get("filename"))

    # comments
    tc.assertListEqual(
        [
            OpenDocumentAnnotation(
                creator="User",
                date="2025-12-28T12:00:00",
                text="This is a comment by User on the sample text.",
            )
        ],
        odt.annotations,
    )

    # footer/headers
    tc.assertListEqual(
        [OdtHeaderFooter(type="header", text="Document Header - My ODT Document")],
        odt.headers,
    )
    tc.assertListEqual(
        [OdtHeaderFooter(type="footer", text="Footer - Page 1 | Confidential")],
        odt.footers,
    )

    # endnote
    tc.assertListEqual(
        [
            OdtNote(
                id="en1",
                note_class="endnote",
                text="This is an endnote that appears at the end of the document.",
            )
        ],
        odt.endnotes,
    )

    # images
    tc.assertEqual(0, len(list(odt.images)))

    # tables
    tc.assertEqual(1, len(odt.tables))
    tc.assertEqual(
        OdtTable(data=[["Header 1", "Header 2"], ["Cell A", "Cell B"]]),
        odt.tables[0],
    )
    tc.assertListEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        list(odt.iterate_tables())[0].get_table(),
    )
    tc.assertEqual(TableDim(rows=2, columns=2), list(odt.iterate_tables())[0].get_dim())

    # full text with defaults
    tc.assertEqual(
        "Hello World Document\n"
        "Hello World! This is a sample ODT document created with Python.\n"
        "This paragraph contains an endnote reference for demonstration purposes.\n"
        "Header 1\n"
        "Header 2\n"
        "Cell A\n"
        "Cell B\n"
        "End of document.",
        odt.get_full_text(),
    )

    tc.assertEqual(
        "Hello World Document\n"
        "Hello World! This is a sample ODT document created with Python.\n"
        "This paragraph contains an endnote reference for demonstration purposes.\n"
        "Header 1\n"
        "Header 2\n"
        "Cell A\n"
        "Cell B\n"
        "End of document.",
        odt.get_full_text(),
    )

    tc.assertEqual(0, len(list(odt.iterate_images())))
    tc.assertEqual(
        OdtTable(data=[["Header 1", "Header 2"], ["Cell A", "Cell B"]]),
        list(odt.iterate_tables())[0],
    )
    tc.assertEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        list(odt.iterate_tables())[0].get_table(),
    )
    tc.assertEqual(TableDim(rows=2, columns=2), list(odt.iterate_tables())[0].get_dim())

    #########
    # Units #
    #########
    units = list(odt.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertListEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        units[0].get_tables()[0].get_table(),
    )
    tc.assertTrue(isinstance(units[0].get_metadata(), OdtUnitMetadata))
    tc.assertEqual(
        OdtUnitMetadata(
            unit_number=1,
            heading_level=1,
            heading_path=["Hello World Document"],
            kind="body",
            annotation_creator=None,
            annotation_date=None,
        ),
        units[0].get_metadata(),
    )


def test_read_open_office__document_aoo() -> None:
    # the dialects are not fully compatible
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_document.odt"
    odt: OdtContent = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Apache OO document\nA\nB\n1\n2", odt.get_full_text())
    tc.assertListEqual(
        [["A", "B"], ["1", "2"]], list(odt.iterate_tables())[0].get_table()
    )
    tc.assertEqual(1, len(list(odt.iterate_images())))


def test_read_open_office__presentation_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_presentation.odp"
    odp: OdpContent = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Test text", odp.get_full_text())


def test_read_open_office__spreadsheet_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_spreadsheet.ods"
    ods: OdsContent = next(
        read_ods(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Sheet1\nHey!\nSheet2\nSheet3", ods.get_full_text())


def test_read_open_office__drawing_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_drawing.odg"
    odg: OdgContent = next(
        read_odg(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("A text shape\nOne more...", odg.get_full_text())


def test_read_open_office__formular_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_formular.odf"
    odf: OdfContent = next(
        read_odf(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("4 + 5 * 8 div 7", odf.get_full_text())


def test_read_open_office__presentation_with_notes() -> None:
    path = "sharepoint2text/tests/resources/open_office/slide_with_notes.odp"
    odp: OdpContent = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertListEqual(
        ["This is an example text in the notes section"], odp.slides[0].notes
    )


def test_read_open_office__presentation_with_table() -> None:
    path = "sharepoint2text/tests/resources/open_office/odp_with_table.odp"
    odp: OdpContent = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("A slide with table", odp.get_full_text())

    #########
    # Units #
    #########
    units = list(odp.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertEqual(
        [["A", "B"], ["1", "2"]],
        list(odp.iterate_units())[0].get_tables()[0].get_table(),
    )

    tc.assertTrue(isinstance(units[0].get_metadata(), OdpUnitMetadata))
    tc.assertEqual(
        OdpUnitMetadata(unit_number=1, location=[], slide_number=1),
        units[0].get_metadata(),
    )


def test_read_open_office__drawing_odg() -> None:
    path = "sharepoint2text/tests/resources/open_office/drawing.odg"
    odg: OdgContent = next(
        read_odg(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Hello there!", odg.get_full_text())
    tc.assertEqual(1, len(list(odg.iterate_images())))
    tc.assertEqual(1, len(list(odg.iterate_units())))


def test_read_open_office__formula_odf() -> None:
    path = "sharepoint2text/tests/resources/open_office/formular.odf"
    odf: OdfContent = next(
        read_odf(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("4/7", odf.get_full_text())


def test_read_open_office__heading_units() -> None:
    path = "sharepoint2text/tests/resources/open_office/headings.odt"
    odt: OdtContent = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(1, len(list(odt.iterate_tables())))
    tc.assertEqual(1, len(list(odt.iterate_images())))

    # unit extraction
    units = list(odt.iterate_units())
    tc.assertEqual(5, len(units))

    # 1
    tc.assertListEqual(["Intro"], units[0].get_metadata().heading_path)
    tc.assertEqual("This is the intro text.", units[0].get_text())
    tc.assertTrue(isinstance(units[0].get_metadata(), OdtUnitMetadata))
    tc.assertEqual(
        OdtUnitMetadata(
            unit_number=1,
            heading_level=1,
            heading_path=["Intro"],
            kind="body",
            annotation_creator=None,
            annotation_date=None,
        ),
        units[0].get_metadata(),
    )

    # 2
    tc.assertListEqual(["Chapter 1"], units[1].get_metadata().heading_path)
    tc.assertEqual("Welcome to chapter 1", units[1].get_text())

    # 3
    tc.assertListEqual(
        ["Chapter 1", "Subsection in Chapter 1"], units[2].get_metadata().heading_path
    )
    tc.assertEqual("This is a subsection in chapter 1", units[2].get_text())
    tc.assertListEqual(
        [["A", "B", "C", "D"], ["1", "2", "3", "4"]],
        units[3].get_tables()[0].get_table(),
    )

    # 4
    tc.assertListEqual(["Chapter 2"], units[3].get_metadata().heading_path)
    tc.assertEqual("Welcome to chapter 2", units[3].get_text())
    tc.assertEqual(1, len(list(units[3].get_images())))
    tc.assertEqual(62421, len(list(units[3].get_images())[0].get_bytes().getvalue()))
    tc.assertEqual(
        ImageMetadata(
            unit_number=4,
            image_number=1,
            content_type="image/png",
            width=412,
            height=195,
        ),
        list(units[3].get_images())[0].get_metadata(),
    )

    # 5
    tc.assertListEqual(
        ["Chapter 2", "Subsection in Chapter 2"], units[4].get_metadata().heading_path
    )
    tc.assertEqual("This is a subsection in chapter 2", units[4].get_text())


def test_read_open_office__presentation() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_presentation.odp"
    odp: OdpContent = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    # File metadata
    tc.assertEqual(".odp", odp.get_metadata().to_dict().get("file_extension"))
    tc.assertEqual(
        "sample_presentation.odp", odp.get_metadata().to_dict().get("filename")
    )

    # Document metadata
    tc.assertEqual("ODFPY/1.4.1", odp.metadata.generator)

    # Slides
    tc.assertEqual(3, len(odp.slides))
    tc.assertEqual(3, odp.slide_count)

    # Slide 1
    tc.assertEqual(1, odp.slides[0].slide_number)
    tc.assertEqual("Slide1", odp.slides[0].name)
    tc.assertEqual("Hello World Presentation", odp.slides[0].title)
    tc.assertIn("Created with Python and odfpy", odp.slides[0].body_text)
    tc.assertIn("Sample Presentation - Header", odp.slides[0].other_text)
    tc.assertIn("Confidential | Page 1 | 2025", odp.slides[0].other_text)
    tc.assertEqual(
        ["Speaker notes for Slide 1: Welcome the audience and introduce the topic."],
        odp.slides[0].notes,
    )
    # No images in this sample
    tc.assertEqual(0, len(odp.slides[0].images))

    # Slide 2
    tc.assertEqual(2, odp.slides[1].slide_number)
    tc.assertEqual("Slide2", odp.slides[1].name)
    tc.assertEqual("Content Slide", odp.slides[1].title)
    # Body text contains annotation marker that gets extracted separately
    tc.assertTrue(any("ODP features" in text for text in odp.slides[1].body_text))
    # Table on slide 2
    tc.assertEqual(1, len(odp.slides[1].tables))
    tc.assertEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]], odp.slides[1].tables[0]
    )
    tc.assertEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        list(odp.iterate_tables())[0].get_table(),
    )
    # Annotation on slide 2
    tc.assertEqual(1, len(odp.slides[1].annotations))
    tc.assertEqual(
        OpenDocumentAnnotation(
            creator="User",
            date="2025-12-28T12:00:00",
            text="This is a comment by User on the presentation content.",
        ),
        odp.slides[1].annotations[0],
    )
    tc.assertEqual(
        [
            "Speaker notes for Slide 2: Explain the table data and highlight key features."
        ],
        odp.slides[1].notes,
    )

    # Slide 3
    tc.assertEqual(3, odp.slides[2].slide_number)
    tc.assertEqual("Slide3", odp.slides[2].name)
    tc.assertEqual("Thank You!", odp.slides[2].title)
    tc.assertIn("Questions? Contact: user@example.com", odp.slides[2].body_text)
    tc.assertEqual(
        ["Speaker notes for Slide 3: Thank the audience and open for Q&A."],
        odp.slides[2].notes,
    )

    # Iterator yields 3 items (one per slide)
    tc.assertEqual(3, len(list(odp.iterate_units())))

    # Full text (default - no annotations, no notes)
    full_text = odp.get_full_text()
    tc.assertIn("Hello World Presentation", full_text)
    tc.assertIn("Content Slide", full_text)
    tc.assertIn("Thank You!", full_text)

    tc.assertEqual(0, len(list(odp.iterate_images())))


def test_read_open_office__spreadsheet() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_spreadsheet.ods"
    ods: OdsContent = next(
        read_ods(file_like=read_file_to_file_like(path=path), path=path)
    )

    # File metadata
    tc.assertEqual(".ods", ods.get_metadata().to_dict().get("file_extension"))
    tc.assertEqual(
        "sample_spreadsheet.ods", ods.get_metadata().to_dict().get("filename")
    )

    # Document metadata
    tc.assertEqual("ODFPY/1.4.1", ods.metadata.generator)

    # Sheets
    tc.assertEqual(2, len(ods.sheets))
    tc.assertEqual(2, ods.sheet_count)

    # Sheet 1: Sales Data
    tc.assertEqual("Sales Data", ods.sheets[0].name)
    # Verify data rows exist
    tc.assertEqual(8, len(ods.sheets[0].data))
    # Verify header row content
    tc.assertIn("Product", ods.sheets[0].text)
    tc.assertIn("Q1", ods.sheets[0].text)
    tc.assertIn("Q2", ods.sheets[0].text)
    tc.assertIn("Q3", ods.sheets[0].text)
    tc.assertIn("Q4", ods.sheets[0].text)
    tc.assertIn("Total", ods.sheets[0].text)
    # Verify product data
    tc.assertIn("Widget A", ods.sheets[0].text)
    tc.assertIn("Widget B", ods.sheets[0].text)
    tc.assertIn("Widget C", ods.sheets[0].text)
    tc.assertIn("Widget D", ods.sheets[0].text)
    # Verify numeric values (from office:value attribute)
    tc.assertIn("1500", ods.sheets[0].text)
    tc.assertIn("2200", ods.sheets[0].text)
    # Annotations on Sales Data sheet - should have 2 annotations
    tc.assertEqual(2, len(ods.sheets[0].annotations))
    # First annotation: on Widget A cell
    tc.assertEqual(
        OpenDocumentAnnotation(
            creator="User",
            date="2025-12-28T12:00:00",
            text="This is our best-selling product line.",
        ),
        ods.sheets[0].annotations[0],
    )
    # Second annotation: on the notes row
    tc.assertEqual(
        OpenDocumentAnnotation(
            creator="User",
            date="2025-12-28T14:30:00",
            text="Remember to update these figures after the quarterly review meeting.",
        ),
        ods.sheets[0].annotations[1],
    )
    # No images in this sample
    tc.assertEqual(0, len(ods.sheets[0].images))

    # Sheet 2: Summary
    tc.assertEqual("Summary", ods.sheets[1].name)
    tc.assertIn("Metric", ods.sheets[1].text)
    tc.assertIn("Value", ods.sheets[1].text)
    tc.assertIn("Total Revenue", ods.sheets[1].text)
    tc.assertIn("Average per Product", ods.sheets[1].text)
    # Summary sheet has 1 annotation
    tc.assertEqual(1, len(ods.sheets[1].annotations))
    tc.assertEqual(
        OpenDocumentAnnotation(
            creator="User",
            date="2025-12-28T15:00:00",
            text="These formulas reference the Sales Data sheet. Update source data to refresh.",
        ),
        ods.sheets[1].annotations[0],
    )

    # Iterator yields 2 items (one per sheet)
    tc.assertEqual(2, len(list(ods.iterate_units())))
    tc.assertEqual(0, len(list(ods.iterate_images())))
    tc.assertEqual(2, len(list(ods.iterate_tables())))

    # check length of full text with length of all sheets
    total_length_iteration = sum(len(unit.get_text()) for unit in ods.iterate_units())
    # one line break is added
    length_total = len(ods.get_full_text()) - 1
    tc.assertEqual(total_length_iteration, length_total)

    # Full text contains data from both sheets
    full_text = ods.get_full_text()
    tc.assertEqual(
        "Sales Data\n" "Product\tQ1\tQ2\tQ3\tQ4\tTotal\nWidget",
        full_text[:44].strip(),
    )

    #########
    # Units #
    #########
    tc.assertEqual(2, len(list(ods.iterate_units())))


def test_read_open_office__spreadsheet_2() -> None:
    """Verifies the treatment of empty rows and columns in a sheet

    We want that the list of rows is easily processable with Pandas or Polars to create
    dataframes. This requires that None/Nulls are not accidentally pruned. The rows must have
    the same number of columns for this to work
    """
    path = "sharepoint2text/tests/resources/modern_ms/empty_row_columns.ods"
    ods: OdsContent = next(read_ods(file_like=read_file_to_file_like(path=path)))

    tc.assertEqual(3, len(ods.sheets))
    expected_rows = [
        [None, "Name", None, "Age"],
        [None, "A", None, 25],
        [None, None, None, None],
        [None, "B", None, 28],
    ]
    tc.assertListEqual(
        expected_rows,
        ods.sheets[0].data,
    )
    tc.assertListEqual(expected_rows, ods.sheets[0].get_table())
    tc.assertEqual(0, len(list(ods.iterate_images())))
    tc.assertEqual(3, len(list(ods.iterate_tables())))
    tc.assertEqual(TableDim(rows=4, columns=4), list(ods.iterate_tables())[0].get_dim())

    #########
    # Units #
    #########
    tc.assertEqual(3, len(list(ods.iterate_units())))
    units = list(ods.iterate_units())
    tc.assertEqual("Sheet1\nName\tAge\nA\t25\nB\t28", units[0].get_text())
    tc.assertEqual("Sheet1", units[0].get_metadata().sheet_name)


def test_open_office__document_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odt"
    odt: OdtContent = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(2, len(odt.images))
    tc.assertEqual(2, len(list(odt.iterate_images())))
    tc.assertEqual(0, len(list(odt.iterate_tables())))
    tc.assertEqual(
        "Illustration 1: Screenshot from the Open Office download website",
        odt.images[0].get_caption(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=1,
            content_type="image/png",
            width=643,
            height=92,
        ),
        odt.images[0].get_metadata(),
    )
    tc.assertEqual(90038, len(odt.images[0].get_bytes().getvalue()))
    tc.assertEqual(
        "Illustration 2: Another Image from the download website",
        odt.images[1].get_caption(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=2,
            content_type="image/png",
            width=643,
            height=70,
        ),
        odt.images[1].get_metadata(),
    )
    tc.assertEqual(82881, len(odt.images[1].get_bytes().getvalue()))


def test_open_office__document_image_interface__no_images() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odt"
    odt: OdtContent = next(
        read_odt(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )

    tc.assertEqual(0, len(odt.images))


def test_open_office__presentation_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odp"
    odp: OdpContent = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )
    tc.assertEqual(1, len(odp.slides[0].images))
    tc.assertEqual(1, len(list(odp.iterate_images())))
    tc.assertEqual(35712, len(list(odp.iterate_images())[0].get_bytes().getvalue()))
    tc.assertEqual(0, len(list(odp.iterate_tables())))
    tc.assertEqual(
        "",
        odp.slides[0].images[0].get_caption(),
    )
    tc.assertEqual(
        "Screenshot test image\nA test image from the Internet",
        odp.slides[0].images[0].get_description(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/png",
            width=924,
            height=163,
        ),
        list(odp.iterate_images())[0].get_metadata(),
    )

    #########
    # Units #
    #########
    tc.assertEqual(1, len(list(odp.iterate_units())))
    units = list(odp.iterate_units())
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/png",
            width=924,
            height=163,
        ),
        units[0].get_images()[0].get_metadata(),
    )
    tc.assertEqual(35712, len(units[0].get_images()[0].get_bytes().getvalue()))


def test_open_office__presentation_image_interface__no_image_flag() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odp"
    odp: OdpContent = next(
        read_odp(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )
    tc.assertEqual(0, len(odp.slides[0].images))


def test_open_office__spreadsheet_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.ods"
    ods: OdsContent = next(
        read_ods(file_like=read_file_to_file_like(path=path), path=path)
    )
    tc.assertEqual(3, len(ods.sheets))
    tc.assertEqual(1, len(ods.sheets[0].images))
    tc.assertEqual(1, len(list(ods.iterate_images())))
    tc.assertEqual(3, len(list(ods.iterate_tables())))

    tc.assertEqual(
        "",
        ods.sheets[0].images[0].get_caption(),
    )
    tc.assertEqual(
        "A description title\nThe description text of the image",
        ods.sheets[0].images[0].get_description(),
    )


def test_open_office__spreadsheet_image_interface__no_images() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.ods"
    ods: OdsContent = next(
        read_ods(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )
    tc.assertEqual(0, len(ods.sheets[0].images))


def test_read_odt__unit_structure() -> None:
    path = "sharepoint2text/tests/resources/open_office/word_structure.odt"
    doc: OdtContent = next(read_odt(file_like=read_file_to_file_like(path=path)))

    units = list(doc.iterate_units())
    tc.assertEqual(5, len(units))

    unit1 = units[0]
    tc.assertListEqual(["A title"], unit1.get_metadata().heading_path)
    tc.assertEqual("blabla", unit1.get_text())

    unit2 = units[1]
    tc.assertListEqual(["A title", "Chapter 1"], unit2.get_metadata().heading_path)
    tc.assertEqual("A chapter", unit2.get_text())

    unit3 = units[2]
    tc.assertListEqual(
        ["A title", "Chapter 1", "Section 1.1"],
        unit3.get_metadata().heading_path,
    )
    tc.assertEqual("A section", unit3.get_text())

    unit4 = units[3]
    tc.assertListEqual(
        ["A title", "Chapter 1", "Section 1.1", "Sub-Section 1.1.1"],
        unit4.get_metadata().heading_path,
    )
    tc.assertEqual("A sub-section", unit4.get_text())

    unit5 = units[4]
    tc.assertListEqual(["A title", "Chapter 2"], unit5.get_metadata().heading_path)
    tc.assertEqual("The text of chapter 2", unit5.get_text())
