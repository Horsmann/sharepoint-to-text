import logging
import typing
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.open_office.odf_extractor import read_odf
from sharepoint2text.parsing.extractors.open_office.odg_extractor import read_odg
from sharepoint2text.parsing.extractors.open_office.odp_extractor import read_odp
from sharepoint2text.parsing.extractors.open_office.ods_extractor import read_ods
from sharepoint2text.parsing.extractors.open_office.odt_extractor import read_odt
from sharepoint2text.parsing.models import (
    Annotation,
    ContentUnit,
    ExtractedDocument,
    ImageAsset,
)
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def _annotations(owner: ExtractedDocument | ContentUnit, kind: str) -> list[Annotation]:
    records = (
        owner.document_annotations
        if isinstance(owner, ExtractedDocument)
        else owner.annotations
    )
    return [annotation for annotation in records if annotation.kind == kind]


def _assert_image(
    image: ImageAsset,
    *,
    number: int,
    unit_number: int,
    media_type: str,
    width: int,
    height: int,
    format_name: str,
) -> None:
    tc.assertEqual(number, image.number)
    tc.assertEqual(unit_number, image.properties[f"{format_name}.unit_number"])
    tc.assertEqual(media_type, image.media_type)
    tc.assertEqual(width, image.width)
    tc.assertEqual(height, image.height)


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
    odt: ExtractedDocument = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(".odt", odt.source.extension)
    tc.assertEqual("sample_document.odt", odt.source.filename)

    # comments
    comments = _annotations(odt, "comment")
    tc.assertEqual(1, len(comments))
    tc.assertEqual("User", comments[0].author)
    tc.assertEqual("2025-12-28T12:00:00", comments[0].properties["odt.date"])
    tc.assertEqual("This is a comment by User on the sample text.", comments[0].text)

    # footer/headers
    tc.assertListEqual(
        ["Document Header - My ODT Document"],
        [annotation.text for annotation in _annotations(odt, "header")],
    )
    tc.assertListEqual(
        ["Footer - Page 1 | Confidential"],
        [annotation.text for annotation in _annotations(odt, "footer")],
    )

    # endnote
    endnotes = _annotations(odt, "endnote")
    tc.assertEqual(1, len(endnotes))
    tc.assertEqual("en1", endnotes[0].properties["odt.id"])
    tc.assertEqual(
        "This is an endnote that appears at the end of the document.",
        endnotes[0].text,
    )

    # images
    tc.assertEqual(0, len(list(odt.iter_images())))

    # tables
    tc.assertEqual(1, len(list(odt.iter_tables())))
    tc.assertListEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        list(odt.iter_tables())[0].rows,
    )
    tc.assertEqual((2, 2), list(odt.iter_tables())[0].dimensions)

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
        odt.full_text,
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
        odt.full_text,
    )

    tc.assertEqual(0, len(list(odt.iter_images())))
    tc.assertEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        list(odt.iter_tables())[0].rows,
    )
    tc.assertEqual((2, 2), list(odt.iter_tables())[0].dimensions)

    #########
    # Units #
    #########
    units = list(odt.units)
    tc.assertEqual(1, len(units))
    tc.assertListEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        units[0].tables[0].rows,
    )
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("section", units[0].kind)
    tc.assertListEqual(["Hello World Document"], units[0].heading_path)


def test_read_open_office__document_aoo() -> None:
    # the dialects are not fully compatible
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_document.odt"
    odt: ExtractedDocument = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Apache OO document\nA\nB\n1\n2", odt.full_text)
    tc.assertListEqual([["A", "B"], ["1", "2"]], list(odt.iter_tables())[0].rows)
    tc.assertEqual(1, len(list(odt.iter_images())))


def test_read_open_office__presentation_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_presentation.odp"
    odp: ExtractedDocument = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Test text", odp.full_text)


def test_read_open_office__spreadsheet_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_spreadsheet.ods"
    ods: ExtractedDocument = next(
        read_ods(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Sheet1\nHey!\nSheet2\nSheet3", ods.full_text)


def test_read_open_office__drawing_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_drawing.odg"
    odg: ExtractedDocument = next(
        read_odg(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("A text shape\nOne more...", odg.full_text)


def test_read_open_office__formular_aoo() -> None:
    path = "sharepoint2text/tests/resources/open_office/apache_oo/aoo_formular.odf"
    odf: ExtractedDocument = next(
        read_odf(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("4 + 5 * 8 div 7", odf.full_text)


def test_read_open_office__presentation_with_notes() -> None:
    path = "sharepoint2text/tests/resources/open_office/slide_with_notes.odp"
    odp: ExtractedDocument = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertListEqual(
        ["This is an example text in the notes section"],
        [annotation.text for annotation in _annotations(odp.units[0], "note")],
    )


def test_read_open_office__presentation_with_table() -> None:
    path = "sharepoint2text/tests/resources/open_office/odp_with_table.odp"
    odp: ExtractedDocument = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("A slide with table", odp.full_text)

    #########
    # Units #
    #########
    units = list(odp.units)
    tc.assertEqual(1, len(units))
    tc.assertEqual(
        [["A", "B"], ["1", "2"]],
        list(odp.units)[0].tables[0].rows,
    )

    tc.assertEqual(1, units[0].number)
    tc.assertEqual("slide", units[0].kind)
    tc.assertEqual("A slide with table", units[0].title)
    tc.assertEqual(1, units[0].properties["odp.slide_number"])
    tc.assertEqual(["A slide with table"], units[0].properties["odp.location"])


def test_read_open_office__drawing_odg() -> None:
    path = "sharepoint2text/tests/resources/open_office/drawing.odg"
    odg: ExtractedDocument = next(
        read_odg(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Hello there!", odg.full_text)
    tc.assertEqual(1, len(list(odg.iter_images())))
    tc.assertEqual(1, len(list(odg.units)))


def test_read_open_office__formula_odf() -> None:
    path = "sharepoint2text/tests/resources/open_office/formular.odf"
    odf: ExtractedDocument = next(
        read_odf(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("4/7", odf.full_text)


def test_read_open_office__heading_units() -> None:
    path = "sharepoint2text/tests/resources/open_office/headings.odt"
    odt: ExtractedDocument = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(1, len(list(odt.iter_tables())))
    tc.assertEqual(1, len(list(odt.iter_images())))

    # unit extraction
    units = list(odt.units)
    tc.assertEqual(5, len(units))

    # 1
    tc.assertListEqual(["Intro"], units[0].heading_path)
    tc.assertEqual("This is the intro text.", units[0].text)
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("section", units[0].kind)
    tc.assertEqual(1, units[0].properties["odt.outline_level"])

    # 2
    tc.assertListEqual(["Chapter 1"], units[1].heading_path)
    tc.assertEqual("Welcome to chapter 1", units[1].text)

    # 3
    tc.assertListEqual(["Chapter 1", "Subsection in Chapter 1"], units[2].heading_path)
    tc.assertEqual("This is a subsection in chapter 1", units[2].text)
    tc.assertListEqual(
        [["A", "B", "C", "D"], ["1", "2", "3", "4"]],
        units[3].tables[0].rows,
    )

    # 4
    tc.assertListEqual(["Chapter 2"], units[3].heading_path)
    tc.assertEqual("Welcome to chapter 2", units[3].text)
    tc.assertEqual(1, len(list(units[3].images)))
    image = units[3].images[0]
    tc.assertEqual(62421, len(image.data or b""))
    _assert_image(
        image,
        number=1,
        unit_number=4,
        media_type="image/png",
        width=412,
        height=195,
        format_name="odt",
    )

    # 5
    tc.assertListEqual(["Chapter 2", "Subsection in Chapter 2"], units[4].heading_path)
    tc.assertEqual("This is a subsection in chapter 2", units[4].text)


def test_read_open_office__presentation() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_presentation.odp"
    odp: ExtractedDocument = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )

    # File metadata
    tc.assertEqual(".odp", odp.source.extension)
    tc.assertEqual("sample_presentation.odp", odp.source.filename)

    # Document metadata
    tc.assertEqual("ODFPY/1.4.1", odp.metadata.properties["odf.generator"])

    # Slides
    tc.assertEqual(3, len(odp.units))
    tc.assertEqual(3, len(odp.units))

    # Slide 1
    tc.assertEqual(1, odp.units[0].number)
    tc.assertEqual("Slide1", odp.units[0].properties["odp.name"])
    tc.assertEqual("Hello World Presentation", odp.units[0].title)
    body_text = typing.cast(list[str], odp.units[0].properties["odp.body_text"])
    other_text = typing.cast(list[str], odp.units[0].properties["odp.other_text"])
    tc.assertIn("Created with Python and odfpy", body_text)
    tc.assertIn("Sample Presentation - Header", other_text)
    tc.assertIn("Confidential | Page 1 | 2025", other_text)
    tc.assertEqual(
        ["Speaker notes for Slide 1: Welcome the audience and introduce the topic."],
        [annotation.text for annotation in _annotations(odp.units[0], "note")],
    )
    # No images in this sample
    tc.assertEqual(0, len(odp.units[0].images))

    # Slide 2
    tc.assertEqual(2, odp.units[1].number)
    tc.assertEqual("Slide2", odp.units[1].properties["odp.name"])
    tc.assertEqual("Content Slide", odp.units[1].title)
    # Body text contains annotation marker that gets extracted separately
    body_text = typing.cast(list[str], odp.units[1].properties["odp.body_text"])
    tc.assertTrue(any("ODP features" in text for text in body_text))
    # Table on slide 2
    tc.assertEqual(1, len(odp.units[1].tables))
    tc.assertEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        odp.units[1].tables[0].rows,
    )
    tc.assertEqual(
        [["Header 1", "Header 2"], ["Cell A", "Cell B"]],
        list(odp.iter_tables())[0].rows,
    )
    # Annotation on slide 2
    comments = _annotations(odp.units[1], "comment")
    tc.assertEqual(1, len(comments))
    tc.assertEqual("User", comments[0].author)
    tc.assertEqual("2025-12-28T12:00:00", comments[0].properties["odp.date"])
    tc.assertEqual(
        "This is a comment by User on the presentation content.", comments[0].text
    )
    tc.assertEqual(
        [
            "Speaker notes for Slide 2: Explain the table data and highlight key features."
        ],
        [annotation.text for annotation in _annotations(odp.units[1], "note")],
    )

    # Slide 3
    tc.assertEqual(3, odp.units[2].number)
    tc.assertEqual("Slide3", odp.units[2].properties["odp.name"])
    tc.assertEqual("Thank You!", odp.units[2].title)
    body_text = typing.cast(list[str], odp.units[2].properties["odp.body_text"])
    tc.assertIn("Questions? Contact: user@example.com", body_text)
    tc.assertEqual(
        ["Speaker notes for Slide 3: Thank the audience and open for Q&A."],
        [annotation.text for annotation in _annotations(odp.units[2], "note")],
    )

    # Iterator yields 3 items (one per slide)
    tc.assertEqual(3, len(list(odp.units)))

    # Full text (default - no annotations, no notes)
    full_text = odp.full_text
    tc.assertIn("Hello World Presentation", full_text)
    tc.assertIn("Content Slide", full_text)
    tc.assertIn("Thank You!", full_text)

    tc.assertEqual(0, len(list(odp.iter_images())))


def test_read_open_office__spreadsheet() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_spreadsheet.ods"
    ods: ExtractedDocument = next(
        read_ods(file_like=read_file_to_file_like(path=path), path=path)
    )

    # File metadata
    tc.assertEqual(".ods", ods.source.extension)
    tc.assertEqual("sample_spreadsheet.ods", ods.source.filename)

    # Document metadata
    tc.assertEqual("ODFPY/1.4.1", ods.metadata.properties["odf.generator"])

    # Sheets
    tc.assertEqual(2, len(ods.units))
    tc.assertEqual(2, len(ods.units))

    # Sheet 1: Sales Data
    tc.assertEqual("Sales Data", ods.units[0].title)
    # Verify data rows exist
    tc.assertEqual(8, len(ods.units[0].tables[0].rows))
    # Verify header row content
    tc.assertIn("Product", ods.units[0].text)
    tc.assertIn("Q1", ods.units[0].text)
    tc.assertIn("Q2", ods.units[0].text)
    tc.assertIn("Q3", ods.units[0].text)
    tc.assertIn("Q4", ods.units[0].text)
    tc.assertIn("Total", ods.units[0].text)
    # Verify product data
    tc.assertIn("Widget A", ods.units[0].text)
    tc.assertIn("Widget B", ods.units[0].text)
    tc.assertIn("Widget C", ods.units[0].text)
    tc.assertIn("Widget D", ods.units[0].text)
    # Verify numeric values (from office:value attribute)
    tc.assertIn("1500", ods.units[0].text)
    tc.assertIn("2200", ods.units[0].text)
    # Annotations on Sales Data sheet - should have 2 annotations
    tc.assertEqual(2, len(ods.units[0].annotations))
    # First annotation: on Widget A cell
    tc.assertEqual("User", ods.units[0].annotations[0].author)
    tc.assertEqual(
        "2025-12-28T12:00:00",
        ods.units[0].annotations[0].properties["ods.date"],
    )
    tc.assertEqual(
        "This is our best-selling product line.", ods.units[0].annotations[0].text
    )
    # Second annotation: on the notes row
    tc.assertEqual("User", ods.units[0].annotations[1].author)
    tc.assertEqual(
        "2025-12-28T14:30:00",
        ods.units[0].annotations[1].properties["ods.date"],
    )
    tc.assertEqual(
        "Remember to update these figures after the quarterly review meeting.",
        ods.units[0].annotations[1].text,
    )
    # No images in this sample
    tc.assertEqual(0, len(ods.units[0].images))

    # Sheet 2: Summary
    tc.assertEqual("Summary", ods.units[1].title)
    tc.assertIn("Metric", ods.units[1].text)
    tc.assertIn("Value", ods.units[1].text)
    tc.assertIn("Total Revenue", ods.units[1].text)
    tc.assertIn("Average per Product", ods.units[1].text)
    # Summary sheet has 1 annotation
    tc.assertEqual(1, len(ods.units[1].annotations))
    tc.assertEqual("User", ods.units[1].annotations[0].author)
    tc.assertEqual(
        "2025-12-28T15:00:00",
        ods.units[1].annotations[0].properties["ods.date"],
    )
    tc.assertEqual(
        "These formulas reference the Sales Data sheet. Update source data to refresh.",
        ods.units[1].annotations[0].text,
    )

    # Iterator yields 2 items (one per sheet)
    tc.assertEqual(2, len(list(ods.units)))
    tc.assertEqual(0, len(list(ods.iter_images())))
    tc.assertEqual(2, len(list(ods.iter_tables())))

    # check length of full text with length of all sheets
    total_length_iteration = sum(len(unit.text) for unit in ods.units)
    # one line break is added
    length_total = len(ods.full_text) - 1
    tc.assertEqual(total_length_iteration, length_total)

    # Full text contains data from both sheets
    full_text = ods.full_text
    tc.assertEqual(
        "Sales Data\n" "Product\tQ1\tQ2\tQ3\tQ4\tTotal\nWidget",
        full_text[:44].strip(),
    )

    #########
    # Units #
    #########
    tc.assertEqual(2, len(list(ods.units)))


def test_read_open_office__spreadsheet_2() -> None:
    """Verifies the treatment of empty rows and columns in a sheet

    We want that the list of rows is easily processable with Pandas or Polars to create
    dataframes. This requires that None/Nulls are not accidentally pruned. The rows must have
    the same number of columns for this to work
    """
    path = "sharepoint2text/tests/resources/modern_ms/empty_row_columns.ods"
    ods: ExtractedDocument = next(read_ods(file_like=read_file_to_file_like(path=path)))

    tc.assertEqual(3, len(ods.units))
    expected_rows = [
        [None, "Name", None, "Age"],
        [None, "A", None, 25],
        [None, None, None, None],
        [None, "B", None, 28],
    ]
    tc.assertListEqual(
        expected_rows,
        ods.units[0].tables[0].rows,
    )
    tc.assertListEqual(expected_rows, ods.units[0].tables[0].rows)
    tc.assertEqual(0, len(list(ods.iter_images())))
    tc.assertEqual(3, len(list(ods.iter_tables())))
    tc.assertEqual((4, 4), list(ods.iter_tables())[0].dimensions)

    #########
    # Units #
    #########
    tc.assertEqual(3, len(list(ods.units)))
    units = list(ods.units)
    tc.assertEqual("Sheet1\nName\tAge\nA\t25\nB\t28", units[0].text)
    tc.assertEqual("Sheet1", units[0].title)


def test_open_office__document_image_interface() -> None:
    """Verify ODT image content through the canonical image model."""
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odt"
    odt: ExtractedDocument = next(
        read_odt(file_like=read_file_to_file_like(path=path), path=path)
    )

    images = list(odt.iter_images())
    tc.assertEqual(2, len(images))
    tc.assertEqual(2, len(images))
    tc.assertEqual(0, len(list(odt.iter_tables())))
    tc.assertEqual(
        "Illustration 1: Screenshot from the Open Office download website",
        images[0].caption,
    )
    _assert_image(
        images[0],
        number=1,
        unit_number=1,
        media_type="image/png",
        width=643,
        height=92,
        format_name="odt",
    )
    tc.assertEqual(90038, len(images[0].data or b""))
    tc.assertEqual(
        "Illustration 2: Another Image from the download website",
        images[1].caption,
    )
    _assert_image(
        images[1],
        number=2,
        unit_number=1,
        media_type="image/png",
        width=643,
        height=70,
        format_name="odt",
    )
    tc.assertEqual(82881, len(images[1].data or b""))


def test_open_office__document_image_interface__no_images() -> None:
    """Verify that excluding images leaves canonical ODT image collections empty."""
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odt"
    odt: ExtractedDocument = next(
        read_odt(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )

    tc.assertEqual(0, len(list(odt.iter_images())))


def test_open_office__presentation_image_interface() -> None:
    """Verify ODP image content through the canonical image model."""
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odp"
    odp: ExtractedDocument = next(
        read_odp(file_like=read_file_to_file_like(path=path), path=path)
    )
    tc.assertEqual(1, len(odp.units[0].images))
    tc.assertEqual(1, len(list(odp.iter_images())))
    tc.assertEqual(35712, len((list(odp.iter_images())[0].data or b"")))
    tc.assertEqual(0, len(list(odp.iter_tables())))
    tc.assertEqual(
        None,
        odp.units[0].images[0].caption,
    )
    tc.assertEqual(
        "Screenshot test image\nA test image from the Internet",
        odp.units[0].images[0].description,
    )
    _assert_image(
        list(odp.iter_images())[0],
        number=1,
        unit_number=1,
        media_type="image/png",
        width=924,
        height=163,
        format_name="odp",
    )

    #########
    # Units #
    #########
    tc.assertEqual(1, len(list(odp.units)))
    units = list(odp.units)
    _assert_image(
        units[0].images[0],
        number=1,
        unit_number=1,
        media_type="image/png",
        width=924,
        height=163,
        format_name="odp",
    )
    tc.assertEqual(35712, len(units[0].images[0].data or b""))


def test_open_office__presentation_image_interface__no_image_flag() -> None:
    """Verify that excluding images leaves canonical ODP image collections empty."""
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odp"
    odp: ExtractedDocument = next(
        read_odp(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )
    tc.assertEqual(0, len(odp.units[0].images))


def test_open_office__spreadsheet_image_interface() -> None:
    """Verify ODS image content through the canonical image model."""
    path = "sharepoint2text/tests/resources/open_office/image_extraction.ods"
    ods: ExtractedDocument = next(
        read_ods(file_like=read_file_to_file_like(path=path), path=path)
    )
    tc.assertEqual(3, len(ods.units))
    tc.assertEqual(1, len(ods.units[0].images))
    tc.assertEqual(1, len(list(ods.iter_images())))
    tc.assertEqual(3, len(list(ods.iter_tables())))

    tc.assertEqual(
        None,
        ods.units[0].images[0].caption,
    )
    tc.assertEqual(
        "A description title\nThe description text of the image",
        ods.units[0].images[0].description,
    )


def test_open_office__spreadsheet_image_interface__no_images() -> None:
    """Verify that excluding images leaves canonical ODS image collections empty."""
    path = "sharepoint2text/tests/resources/open_office/image_extraction.ods"
    ods: ExtractedDocument = next(
        read_ods(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )
    tc.assertEqual(0, len(ods.units[0].images))


def test_read_odt__unit_structure() -> None:
    path = "sharepoint2text/tests/resources/open_office/word_structure.odt"
    doc: ExtractedDocument = next(read_odt(file_like=read_file_to_file_like(path=path)))

    units = list(doc.units)
    tc.assertEqual(5, len(units))

    unit1 = units[0]
    tc.assertListEqual(["A title"], unit1.heading_path)
    tc.assertEqual("blabla", unit1.text)

    unit2 = units[1]
    tc.assertListEqual(["A title", "Chapter 1"], unit2.heading_path)
    tc.assertEqual("A chapter", unit2.text)

    unit3 = units[2]
    tc.assertListEqual(
        ["A title", "Chapter 1", "Section 1.1"],
        unit3.heading_path,
    )
    tc.assertEqual("A section", unit3.text)

    unit4 = units[3]
    tc.assertListEqual(
        ["A title", "Chapter 1", "Section 1.1", "Sub-Section 1.1.1"],
        unit4.heading_path,
    )
    tc.assertEqual("A sub-section", unit4.text)

    unit5 = units[4]
    tc.assertListEqual(["A title", "Chapter 2"], unit5.heading_path)
    tc.assertEqual("The text of chapter 2", unit5.text)


def test_read_odt_units() -> None:
    path = "sharepoint2text/tests/resources/open_office/slide_headlines.odp"
    odt: ExtractedDocument = next(read_odp(read_file_to_file_like(path=path)))

    units = list(odt.units)
    tc.assertEqual(2, len(units))
    tc.assertEqual("My Slide Title", units[0].title)
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("", units[0].text)
    tc.assertEqual("Another Slide", units[1].title)
    tc.assertEqual("Good day!", units[1].text)
    tc.assertEqual(2, units[1].number)
