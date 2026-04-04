import logging
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    DocxComment,
    DocxContent,
    DocxFormula,
    DocxNote,
    DocxUnit,
    ImageMetadata,
    PptxComment,
    PptxContent,
    PptxUnitMetadata,
    TableData,
    TableDim,
    XlsxContent,
    XlsxSheet,
    XlsxUnitMetadata,
)
from sharepoint2text.parsing.extractors.ms_modern.docx_extractor import read_docx
from sharepoint2text.parsing.extractors.ms_modern.pptx_extractor import read_pptx
from sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor import read_xlsx
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


####################
# Modern Microsoft #
####################
def test_read_xlsx_1() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx"
    xlsx: XlsxContent = next(read_xlsx(file_like=read_file_to_file_like(path=path)))

    tc.assertEqual("2006-09-16T00:00:00", xlsx.metadata.created)
    tc.assertEqual("2015-05-06T11:46:24", xlsx.metadata.modified)

    tc.assertEqual(3, len(xlsx.sheets))
    tc.assertEqual(3, len(list(xlsx.iterate_tables())))
    tc.assertListEqual(
        sorted(["Sheet1", "Sheet2", "Sheet3"]), sorted([s.name for s in xlsx.sheets])
    )
    tc.assertListEqual(
        ["AREA", "CODE", "COUNTRY NAME"], list(xlsx.iterate_tables())[0].get_table()[0]
    )
    tc.assertEqual(TableDim(52, 3), list(xlsx.iterate_tables())[0].get_dim())

    # check raw data and table interface
    # check that the first row in the first sheet is the headline
    tc.assertListEqual(["AREA", "CODE", "COUNTRY NAME"], xlsx.sheets[0].data[0])
    tc.assertListEqual(["AREA", "CODE", "COUNTRY NAME"], xlsx.sheets[0].get_table()[0])
    tc.assertListEqual(
        ["European Union (EU)", "EU-28 ", "European Union (28 countries)"],
        xlsx.sheets[0].get_table()[1],
    )

    tc.assertEqual(3, len(list(xlsx.iterate_units())))

    tc.assertEqual("Sheet1\nAREA     CODE", xlsx.get_full_text()[:20])

    tc.assertDictEqual(
        {
            "filename": None,
            "file_extension": None,
            "file_path": None,
            "folder_path": None,
            "detected_encoding": None,
            "title": "",
            "description": "",
            "creator": "",
            "last_modified_by": "",
            "created": "2006-09-16T00:00:00",
            "modified": "2015-05-06T11:46:24",
            "keywords": "",
            "language": "",
            "revision": None,
        },
        xlsx.get_metadata().to_dict(),
    )
    tc.assertEqual(0, len(list(xlsx.iterate_images())))


def test_read_xlsx_2() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/mwe.xlsx"
    xlsx: XlsxContent = next(read_xlsx(file_like=read_file_to_file_like(path=path)))
    tc.assertEqual(
        "Blatt 1\nTabelle 1 Unnamed: 1\n     ColA       ColB\n        1          2",
        xlsx.get_full_text(),
    )
    tc.assertListEqual([["ColA", "ColB"], [1, 2]], xlsx.sheets[0].data)
    tc.assertEqual(0, len(list(xlsx.iterate_images())))


def test_read_xlsx_3() -> None:
    """Verifies the treatment of empty rows and columns in a sheet

    We want that the list of rows is easily processable with Pandas or Polars to create
    dataframes. This requires that None/Nulls are not accidentally pruned. The rows must have
    the same number of columns for this to work
    """
    path = "sharepoint2text/tests/resources/modern_ms/empty_row_columns.xlsx"

    xlsx: XlsxContent = next(read_xlsx(file_like=read_file_to_file_like(path=path)))
    tc.assertListEqual(
        [
            [None, "Name", None, "Age"],
            [None, "A", None, 25],
            [None, None, None, None],
            [None, "B", None, 28],
        ],
        xlsx.sheets[0].data,
    )
    tc.assertEqual(0, len(list(xlsx.iterate_images())))
    tc.assertEqual(TableDim(4, 4), list(xlsx.iterate_tables())[0].get_dim())

    #########
    # Units #
    #########
    units = list(xlsx.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertListEqual(
        [
            [None, "Name", None, "Age"],
            [None, "A", None, 25],
            [None, None, None, None],
            [None, "B", None, 28],
        ],
        units[0].get_tables()[0].get_table(),
    )
    tc.assertEqual(0, len(units[0].get_images()))
    tc.assertEqual(
        XlsxUnitMetadata(unit_number=1, sheet_number=1, sheet_name="Blatt 1"),
        units[0].get_metadata(),
    )


def test_read_xlsx_4__image_extraction() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/image_in_excel.xlsx"

    xlsx: XlsxContent = next(read_xlsx(file_like=read_file_to_file_like(path=path)))
    tc.assertEqual("Image Sheet", xlsx.sheets[0].name)
    tc.assertEqual(1, len(xlsx.sheets[0].images))

    image = xlsx.sheets[0].images[0]
    tc.assertEqual(7280, len(image.get_bytes().getvalue()))
    tc.assertEqual("Image 1", image.get_caption())
    tc.assertEqual("Picture", image.get_description())
    tc.assertEqual(600, image.width)
    tc.assertEqual(300, image.height)

    tc.assertEqual(1, len(list(xlsx.iterate_images())))
    img_meta = list(xlsx.iterate_images())[0].get_metadata()
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/png",
            width=600,
            height=300,
        ),
        img_meta,
    )
    tc.assertEqual(1, img_meta.unit_number)
    tc.assertEqual(600, img_meta.width)
    tc.assertEqual(300, img_meta.height)

    #########
    # Units #
    #########
    units = list(xlsx.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertEqual("image/png", list(units[0].get_images())[0].get_content_type())
    tc.assertEqual(7280, len(list(units[0].get_images())[0].get_bytes().getvalue()))
    tc.assertEqual("Image 1", list(units[0].get_images())[0].get_caption())
    tc.assertEqual("Picture", list(units[0].get_images())[0].get_description())


def test_read_pptx_1() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/eu-visibility_rules_00704232-AF9F-1A18-BD782C469454ADAD_68401.pptx"
    pptx: PptxContent = next(read_pptx(read_file_to_file_like(path=path)))

    # metadata
    tc.assertEqual("IVAN Anda-Otilia", pptx.metadata.author)
    tc.assertEqual("MAGLI Mia (JUST)", pptx.metadata.last_modified_by)
    tc.assertEqual("2011-10-28T10:25:18", pptx.metadata.created)
    tc.assertEqual("2020-07-12T09:25:35", pptx.metadata.modified)

    tc.assertEqual(3, len(pptx.slides))
    tc.assertEqual(5, len(list(pptx.iterate_images())))
    tc.assertEqual(0, len(list(pptx.iterate_tables())))
    tc.assertEqual(
        ImageMetadata(
            unit_number=2,
            image_number=1,
            content_type="image/png",
            width=130,
            height=111,
        ),
        list(pptx.iterate_images())[0].get_metadata(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=2,
            image_number=2,
            content_type="image/jpeg",
            width=264,
            height=255,
        ),
        list(pptx.iterate_images())[1].get_metadata(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=2,
            image_number=3,
            content_type="image/jpeg",
            width=279,
            height=186,
        ),
        list(pptx.iterate_images())[2].get_metadata(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=2,
            image_number=4,
            content_type="image/jpeg",
            width=305,
            height=250,
        ),
        list(pptx.iterate_images())[3].get_metadata(),
    )
    tc.assertEqual(
        ImageMetadata(
            unit_number=2,
            image_number=5,
            content_type="image/jpeg",
            width=286,
            height=191,
        ),
        list(pptx.iterate_images())[4].get_metadata(),
    )
    ##########
    # SLIDES #
    ##########
    # slide 1
    tc.assertEqual("EU-funding visibility - art. 22 GA", pptx.slides[0].title)
    expected = [
        'To be applied on all materials and communication activities:\n\nThe correct EU emblem: https://europa.eu/european-union/about-eu/symbols/flag_en; \nThe reference to the correct funding programme (to be put next to the EU emblem): “This [e.g. project, report, publication, conference, website, etc.] was funded by the European Union’s Justice Programme (2014-2020) or by the Rights, Equality and Citizenship Programme (REC 2014-2020)“;\n The following disclaimer: "The content of this [insert appropriate description, e.g. report, publication, conference, etc.] represents the views of the author only and is his/her sole responsibility. The European Commission does not accept any responsibility for use that may be made of the information it contains”.'
    ]
    tc.assertListEqual(expected, pptx.slides[0].content_placeholders)

    tc.assertListEqual(["1"], pptx.slides[0].other_textboxes)
    tc.assertEqual(1, pptx.slides[0].slide_number)

    # images
    tc.assertEqual(0, len(pptx.slides[0].images))

    # slide 2
    tc.assertEqual("EU-funding visibility", pptx.slides[1].title)

    expected = [
        "! Please choose the correct name of the funding Programme, either JUSTICE or REC, depending under which Programme your call for proposals was published and your grant was awarded:\n\nSupported by the Rights, Equality \x0band Citizenship Programme \nof the European Union (2014-2020) \x0b\n     This project is funded by the Justice \n      Programme of the European Union \n      (2014-2020)"
    ]
    tc.assertListEqual(expected, pptx.slides[1].content_placeholders)

    # Order reflects visual position on slide (top to bottom, left to right)
    expected = ["This is the wrong EU emblem", "Don’t use this emblem!", "2"]
    tc.assertListEqual(expected, pptx.slides[1].other_textboxes)

    # images (sorted by position on slide)
    tc.assertEqual(5, len(pptx.slides[1].images))
    # test presence of metadata for first image (now image.png due to position sort)
    tc.assertEqual(1, pptx.slides[1].images[0].image_index)
    tc.assertEqual("image.png", pptx.slides[1].images[0].filename)
    tc.assertEqual("image/png", pptx.slides[1].images[0].content_type)
    tc.assertEqual(12538, pptx.slides[1].images[0].size_bytes)
    tc.assertIsNotNone(pptx.slides[1].images[0].data)

    # full text
    expected = (
        "EU-funding visibility - art. 22 GA"
        + "\n"
        + "To be applied on all materials and communica"
    )
    tc.assertEqual(expected, pptx.get_full_text()[:79])

    #########
    # Units #
    #########
    tc.assertEqual(3, len(list(pptx.iterate_units())))
    tc.assertEqual(len(pptx.slides), len(list(pptx.iterate_units())))
    units = list(pptx.iterate_units())

    tc.assertEqual(0, len(units[0].get_images()))
    tc.assertEqual("EU-funding visibility", units[0].get_text()[:21])
    tc.assertEqual(PptxUnitMetadata(unit_number=1), units[0].get_metadata())

    tc.assertEqual(5, len(units[1].get_images()))
    tc.assertEqual("This is the wrong EU ", units[1].get_text()[:21])

    tc.assertEqual(0, len(units[2].get_images()))
    tc.assertEqual("EU-funding visibility", units[2].get_text()[:21])


def test_read_pptx_2() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    pptx: PptxContent = next(read_pptx(read_file_to_file_like(path=path)))

    # Test default get_full_text() - formulas included (no comments or image captions)
    # Note: "A beach" is a regular textbox, not an image caption
    base_text = pptx.get_full_text()
    tc.assertEqual(
        "The slide title\nThe first text line\n\n\n\n\nThe last text line\nA beach\n$$f(x)=\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}e^{-\\frac{(x-\\mu)^{2}}{2\\sigma^{2}}}$$",
        base_text,
    )

    # images
    tc.assertEqual(1, len(list(pptx.iterate_images())))
    tc.assertEqual(1, len(pptx.slides[0].images))
    tc.assertEqual(1, pptx.slides[0].images[0].image_index)
    tc.assertEqual("image/jpeg", pptx.slides[0].images[0].content_type)
    tc.assertEqual(1, pptx.slides[0].images[0].slide_number)
    tc.assertEqual(1535390, pptx.slides[0].images[0].size_bytes)
    # description is the alt text for accessibility (from descr attribute)
    tc.assertEqual(
        "Sandiger Weg zwischen zwei Hügeln, die ans Meer führen",
        pptx.slides[0].images[0].description,
    )
    # caption is the shape name/title (from name attribute)
    # Note: in this file, name and descr have the same value
    tc.assertEqual(
        "Sandiger Weg zwischen zwei Hügeln, die ans Meer führen",
        pptx.slides[0].images[0].caption,
    )

    # image interface - get_description() returns the caption (title/name)
    tc.assertEqual(
        "Sandiger Weg zwischen zwei Hügeln, die ans Meer führen",
        pptx.slides[0].images[0].get_description(),
    )
    tc.assertEqual(1535390, len(pptx.slides[0].images[0].get_bytes().getvalue()))
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/jpeg",
            width=1647,
            height=1098,
        ),
        pptx.slides[0].images[0].get_metadata(),
    )

    # comments go separately - they are not part of the full text body
    tc.assertListEqual(
        [PptxComment(author="0", text="Not second?", date="2025-12-28T11:15:49.694")],
        pptx.slides[0].comments,
    )
    tc.assertNotIn("Not second?", pptx.get_full_text())


def test_read_pptx_3() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_table.pptx"
    pptx: PptxContent = next(read_pptx(read_file_to_file_like(path=path)))

    tc.assertEqual(1, len(list(pptx.iterate_tables())))
    table_1 = list(pptx.iterate_tables())[0]
    tc.assertListEqual(
        [
            ["", "2020", "2021", "2022"],
            ["A", "1", "2", "3"],
            ["B", "4", "5", "6"],
            ["C", "7", "8", "9"],
            ["D", "10", "11", "12"],
        ],
        table_1.get_table(),
    )
    tc.assertEqual(TableDim(rows=5, columns=4), table_1.get_dim())
    tc.assertEqual(
        "2020\t2021\t2022\nA\t1\t2\t3\nB\t4\t5\t6\nC\t7\t8\t9\nD\t10\t11\t12",
        pptx.get_full_text(),
    )

    #########
    # Units #
    #########
    units = list(pptx.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertListEqual(
        [
            ["", "2020", "2021", "2022"],
            ["A", "1", "2", "3"],
            ["B", "4", "5", "6"],
            ["C", "7", "8", "9"],
            ["D", "10", "11", "12"],
        ],
        units[0].get_tables()[0].get_table(),
    )


def test_read_pptx__image_flag():
    path = "sharepoint2text/tests/resources/modern_ms/pptx_images.pptx"
    pptx: PptxContent = next(
        read_pptx(read_file_to_file_like(path=path), ignore_images=False)
    )
    tc.assertEqual(1, len(pptx.slides[0].images))
    tc.assertEqual("PPTX text", pptx.get_full_text())

    pptx: PptxContent = next(
        read_pptx(read_file_to_file_like(path=path), ignore_images=True)
    )
    tc.assertEqual(0, len(pptx.slides[0].images))
    tc.assertEqual("PPTX text", pptx.get_full_text())


def test_read_docx_1() -> None:
    # An actual document from the web - this is likely created on a Windows client
    path = (
        "sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx"
    )
    docx: DocxContent = next(read_docx(read_file_to_file_like(path=path)))

    # text is long. Verify only beginning
    tc.assertEqual("Welcome to the Government", docx.full_text[:25].strip())

    tc.assertEqual(230, len(docx.paragraphs))

    tc.assertEqual(17, docx.metadata.revision)
    # Raw XML format uses 'Z' for UTC timezone
    tc.assertEqual("2023-01-20T16:07:00Z", docx.metadata.modified)
    tc.assertEqual("2022-04-19T14:03:00Z", docx.metadata.created)

    # test iterator
    tc.assertEqual(1, len(list(docx.iterate_units())))
    tc.assertEqual(1, len(docx.images))
    tc.assertEqual(1, len(list(docx.iterate_images())))
    tc.assertEqual(7, len(list(docx.iterate_tables())))
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=1,
            content_type="image/png",
            width=1823,
            height=1052,
        ),
        list(docx.iterate_images())[0].get_metadata(),
    )

    # test full text
    tc.assertEqual("Welcome to the Government", docx.get_full_text()[:25].strip())


def test_read_docx_2() -> None:
    # A converted docx from OSX pages - may not populate like a true MS client .docx
    # dedicated test for comment, table and footnote extraction
    path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )

    docx: DocxContent = next(read_docx(read_file_to_file_like(path=path)))
    # Formula with properly converted multiplication sign
    tc.assertEqual(
        "Hello World!\nAn image of space\nIncome\ntax\n119\n19\nAnother sentence after the table.\n$$\\frac{3}{4}\\times4=\\sqrt{9}$$",
        docx.full_text,
    )
    tc.assertEqual(docx.full_text, docx.get_full_text())
    tc.assertNotIn("Nice!", docx.get_full_text())
    tc.assertListEqual(
        [DocxComment(id="0", author="User", date="2025-12-28T09:16:57Z", text="Nice!")],
        docx.comments,
    )
    tc.assertListEqual(
        [
            # I am not sure where this is coming from
            DocxNote(id="-2", text=""),
            DocxNote(id="1", text="A simple footnote"),
        ],
        docx.footnotes,
    )
    tc.assertListEqual([[["Income", "tax"], ["119", "19"]]], docx.tables)

    # formulas (with converted multiplication sign)
    tc.assertListEqual(
        [DocxFormula(latex="\\frac{3}{4}\\times4=\\sqrt{9}", is_display=True)],
        docx.formulas,
    )

    # section object
    tc.assertEqual(1, len(docx.sections))
    tc.assertAlmostEqual(8.268, docx.sections[0].page_width_inches, places=1)
    tc.assertAlmostEqual(11.693, docx.sections[0].page_height_inches, places=1)
    tc.assertAlmostEqual(0.7875, docx.sections[0].left_margin_inches, places=1)
    tc.assertAlmostEqual(0.7875, docx.sections[0].right_margin_inches, places=1)
    tc.assertAlmostEqual(0.7875, docx.sections[0].top_margin_inches, places=1)
    tc.assertAlmostEqual(0.7875, docx.sections[0].bottom_margin_inches, places=1)
    tc.assertIsNone(docx.sections[0].orientation)

    # images
    tc.assertEqual(1, len(docx.images))
    tc.assertEqual(1, len(list(docx.iterate_images())))
    tc.assertEqual(1, len(list(docx.iterate_tables())))
    tc.assertEqual(
        TableData(data=[["Income", "tax"], ["119", "19"]]),
        list(docx.iterate_tables())[0],
    )
    tc.assertEqual(1, docx.images[0].image_index)
    tc.assertEqual("image1.png", docx.images[0].filename)
    tc.assertEqual("image/png", docx.images[0].content_type)
    # description (alt text) is from pic:cNvPr[@descr]
    tc.assertEqual("Space", docx.images[0].description)
    # caption is from the text box content (wps:txbx)
    tc.assertEqual("An image of space", docx.images[0].caption)

    # ImageInterface methods
    tc.assertEqual("image/png", docx.images[0].get_content_type())
    tc.assertEqual("Space", docx.images[0].get_description())
    tc.assertEqual("An image of space", docx.images[0].get_caption())
    # get_bytes returns BytesIO with image data
    image_bytes = docx.images[0].get_bytes()
    tc.assertEqual(828786, len(image_bytes.getvalue()))
    tc.assertEqual(docx.images[0].size_bytes, len(image_bytes.getvalue()))
    # get_metadata returns ImageMetadata
    img_meta = docx.images[0].get_metadata()
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=1,
            content_type="image/png",
            width=930,
            height=506,
        ),
        img_meta,
    )


def test_read_docx__image_flag() -> None:
    # A converted docx from OSX pages - may not populate like a true MS client .docx
    # dedicated test for comment, table and footnote extraction
    path = "sharepoint2text/tests/resources/modern_ms/document_with_image.docx"
    docx: DocxContent = next(
        read_docx(read_file_to_file_like(path=path), ignore_images=False)
    )
    tc.assertEqual(1, len(docx.images))
    tc.assertEqual("Docx with image", docx.get_full_text())

    docx: DocxContent = next(
        read_docx(read_file_to_file_like(path=path), ignore_images=True)
    )
    tc.assertEqual(0, len(docx.images))
    tc.assertEqual("Docx with image", docx.get_full_text())


def test_read_docx__image_extraction_1() -> None:
    # Test for caption extraction from following paragraph with caption style
    path = "sharepoint2text/tests/resources/modern_ms/vorlage-abschlussarbeit.docx"
    docx: DocxContent = next(read_docx(read_file_to_file_like(path=path)))

    tc.assertEqual(1, len(docx.images))
    tc.assertEqual(1, len(list(docx.iterate_images())))
    tc.assertEqual(0, len(list(docx.iterate_tables())))
    # image interface - caption from following paragraph with "HA-Bildunterschrift" style
    expected_caption = (
        "Abb. 1: Eine aus dem Internet heruntergeladene Bilddatei mit einer "
        "Bildunterschrift. Die Abbildungen und Tabellen bitte nicht als "
        "textumflossene Objekte, sondern so wie dies Bild als Absatz in den "
        "Text einbinden. Dieser Untertext hat die Formatvorlage "
        "\u201eHA-Bildunterschrift\u201c."
    )
    tc.assertEqual(expected_caption, docx.images[0].get_caption())
    # description is the alt text (URL in this case)
    tc.assertEqual(
        "http://omgunmen.de/wp-content/uploads/2011/03/but-on-math-it-is.png",
        docx.images[0].get_description(),
    )


def test_read_docx__image_extraction_2() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/thesis-template.docx"
    docx: DocxContent = next(read_docx(read_file_to_file_like(path=path)))

    tc.assertEqual(2, len(docx.images))
    tc.assertEqual(2, len(list(docx.iterate_images())))
    tc.assertEqual(4, len(list(docx.iterate_tables())))
    tc.assertEqual("Illustration 1: [Figure title]", docx.images[1].get_caption())
    tc.assertEqual(
        """Ein Bild, das Zeichnung "Marketing" enthält.""",
        docx.images[1].get_description(),
    )

    # units
    tc.assertEqual(17, len(list(docx.iterate_units())))
    units = list(docx.iterate_units())
    tc.assertListEqual(["II. List of figures"], units[0].get_metadata().location)
    tc.assertListEqual(["III. List of tables"], units[1].get_metadata().location)
    tc.assertListEqual(["IV. List of formulas"], units[2].get_metadata().location)
    tc.assertListEqual(["V. List of abbreviations"], units[3].get_metadata().location)
    tc.assertListEqual(["VI. List of symbols"], units[4].get_metadata().location)
    tc.assertListEqual(["Title 1 Chapter"], units[5].get_metadata().location)
    tc.assertListEqual(["Title 2 Chapter"], units[6].get_metadata().location)
    tc.assertListEqual(
        ["Title 2 Chapter", "2.1 Title Subchapter"], units[7].get_metadata().location
    )
    # unit has an image
    tc.assertListEqual(
        ["Title 2 Chapter", "2.1 Title Subchapter", "2.1.1 Title Subchapter"],
        units[8].get_metadata().location,
    )
    tc.assertEqual(54423, len(units[8].get_images()[0].get_bytes().getvalue()))

    # unit has an table
    tc.assertListEqual(
        ["Title 2 Chapter", "2.1 Title Subchapter", "2.1.2 Title Subchapter"],
        units[9].get_metadata().location,
    )
    tc.assertEqual(TableDim(rows=3, columns=4), units[9].get_tables()[0].get_dim())

    tc.assertListEqual(
        ["Title 2 Chapter", "2.2 Title Subchapter"],
        units[10].get_metadata().location,
    )
    tc.assertListEqual(["Title 3 Chapter"], units[11].get_metadata().location)
    tc.assertListEqual(["Title 4 Chapter"], units[12].get_metadata().location)
    tc.assertListEqual(["VII. Appendix"], units[13].get_metadata().location)
    tc.assertListEqual(["VIII. Bibliography"], units[14].get_metadata().location)
    tc.assertListEqual(["VIII. Bibliography"], units[15].get_metadata().location)
    tc.assertListEqual(["IX. Affidavit"], units[16].get_metadata().location)


def test_read_docx__units() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/headings.docx"
    docx: DocxContent = next(read_docx(file_like=read_file_to_file_like(path=path)))

    units = list(docx.iterate_units())
    tc.assertEqual(8, len(units))

    tc.assertTrue(hasattr(units[0], "get_images"))
    tc.assertTrue(hasattr(units[0], "get_tables"))

    # first unit
    unit_meta: DocxUnit = units[0].get_metadata()
    tc.assertEqual(["Sample Document"], unit_meta.location)
    tc.assertEqual(
        "This document was created using accessibility techniques for headings, lists, image alternate text, tables, and columns. It should be completely accessible using assistive technologies such as screen readers.",
        units[0].get_text(),
    )
    tc.assertEqual(0, len(units[0].get_images()))
    tc.assertEqual(0, len(units[0].get_tables()))

    # second unit
    tc.assertEqual(["Sample Document", "Headings"], units[1].get_metadata().location)
    tc.assertEqual(
        'There are eight section headings in this document. At the beginning, "Sample Document" is a level 1 heading. The main section headings, such as "Headings" and "Lists" are level 2 headings. The Tables section contains two sub-headings, "Simple Table" and "Complex Table," which are both level 3 headings.',
        units[1].get_text(),
    )
    tc.assertEqual(0, len(units[1].get_images()))
    tc.assertEqual(0, len(units[1].get_tables()))

    # third unit
    tc.assertEqual(["Sample Document", "Lists"], units[2].get_metadata().location)
    tc.assertEqual(
        (
            "The following outline of the sections of this document is an ordered "
            '(numbered) list with six items. The fifth item, "Tables," contains a nested '
            "unordered (bulleted) list with two items.\n"
            "Headings\n"
            "Lists\n"
            "Links\n"
            "Images\n"
            "Tables\n"
            "Simple Tables\n"
            "Complex Tables\n"
            "Columns"
        ),
        units[2].get_text(),
    )
    tc.assertEqual(0, len(units[2].get_images()))
    tc.assertEqual(0, len(units[2].get_tables()))

    # Images section
    tc.assertEqual(["Sample Document", "Images"], units[4].get_metadata().location)
    tc.assertEqual(2, len(units[4].get_images()))
    tc.assertSetEqual(
        {"image1.gif", "image2.png"}, {img.filename for img in units[4].get_images()}
    )
    tc.assertEqual(5437, len(units[4].get_images()[0].get_bytes().getvalue()))
    tc.assertEqual(7570, len(units[4].get_images()[1].get_bytes().getvalue()))
    tc.assertEqual(0, len(units[4].get_tables()))

    # Tables section
    tc.assertEqual(1, len(units[5].get_tables()))
    tc.assertEqual(docx.tables[0], units[5].get_tables()[0].get_table())
    tc.assertEqual(1, len(units[6].get_tables()))
    tc.assertEqual(docx.tables[1], units[6].get_tables()[0].get_table())


def test_read_docx__unit_structure() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/word_structure.docx"
    docx: DocxContent = next(read_docx(file_like=read_file_to_file_like(path=path)))

    units = list(docx.iterate_units())
    tc.assertEqual(5, len(units))

    unit1 = units[0]
    tc.assertEqual(["The document title"], unit1.get_metadata().heading_path)
    tc.assertEqual("blabla", unit1.get_text())

    unit2 = units[1]
    tc.assertEqual(
        ["The document title", "Chapter 1"], unit2.get_metadata().heading_path
    )
    tc.assertEqual("This is chapter 1", unit2.get_text())

    unit3 = units[2]
    tc.assertEqual(
        ["The document title", "Chapter 1", "Section 1.1"],
        unit3.get_metadata().heading_path,
    )
    tc.assertEqual("A subsection", unit3.get_text())

    unit4 = units[3]
    tc.assertEqual(
        ["The document title", "Chapter 2"], unit4.get_metadata().heading_path
    )
    tc.assertEqual("This is chapter 2", unit4.get_text())

    unit5 = units[4]
    tc.assertEqual(
        ["The document title", "Chapter 3"], unit5.get_metadata().heading_path
    )
    tc.assertEqual("This is chapter 3", unit5.get_text())


def test_read_macro_enabled_docm() -> None:
    """Test .docm (macro-enabled Word) extraction - same structure as .docx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.docm"
    result: DocxContent = next(
        read_docx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as DocxContent (same as .docx)
    tc.assertIsInstance(result, DocxContent)
    tc.assertTrue(len(result.get_full_text()) > 0)


def test_read_macro_enabled_xlsm() -> None:
    """Test .xlsm (macro-enabled Excel) extraction - same structure as .xlsx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.xlsm"
    result: XlsxContent = next(
        read_xlsx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as XlsxContent (same as .xlsx)
    tc.assertIsInstance(result, XlsxContent)
    tc.assertTrue(len(result.sheets) > 0)


def test_read_xlsb() -> None:
    """Test .xlsm (macro-enabled Excel) extraction - same structure as .xlsx."""
    path = "sharepoint2text/tests/resources/modern_ms/excel.xlsb"
    result: XlsxContent = next(
        read_xlsx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as XlsxContent (same as .xlsx)
    tc.assertEqual(
        """Sheet2
A
A1
A2
A3
B
Atable
Btable
Ctable
Ytable
Zparam
Y
X
W
XWtable
Utable
Stable
S
.
H
R
P
O
B1
B2
B3
POtable
I1
I2
I3
I4
Itable
I
Ttable
Dtable
+
-
Etable
Q
RQtable
H1
H2
H3
PPtable
PP""",
        result.get_full_text(),
    )

    tc.assertEqual(1, len(list(result.iterate_tables())))
    sheet: XlsxSheet = list(result.iterate_tables())[0]
    tc.assertListEqual(
        [
            ["A"],
            ["A1"],
            ["A2"],
            ["A3"],
            ["B"],
            ["Atable"],
            ["Btable"],
            ["Ctable"],
            ["Ytable"],
            ["Zparam"],
            ["Y"],
            ["X"],
            ["W"],
            ["XWtable"],
            ["Utable"],
            ["Stable"],
            ["S"],
            ["."],
            ["H"],
            ["R"],
            ["P"],
            ["O"],
            ["B1"],
            ["B2"],
            ["B3"],
            ["POtable"],
            ["I1"],
            ["I2"],
            ["I3"],
            ["I4"],
            ["Itable"],
            ["I"],
            ["Ttable"],
            ["Dtable"],
            ["+"],
            ["-"],
            ["Etable"],
            ["Q"],
            ["RQtable"],
            ["H1"],
            ["H2"],
            ["H3"],
            ["PPtable"],
            ["PP"],
        ],
        sheet.data,
    )


def test_read_xlsx__image_flag() -> None:
    """Test .xlsm (macro-enabled Excel) extraction - same structure as .xlsx."""
    path = "sharepoint2text/tests/resources/modern_ms/excel_images.xlsx"
    result: XlsxContent = next(
        read_xlsx(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=False
        ),
    )
    tc.assertEqual(1, len(result.sheets[0].images))

    result: XlsxContent = next(
        read_xlsx(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        ),
    )
    tc.assertEqual(0, len(result.sheets[0].images))


def test_read_macro_enabled_pptm() -> None:
    """Test .pptm (macro-enabled PowerPoint) extraction - same structure as .pptx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.pptm"
    result: PptxContent = next(
        read_pptx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as PptxContent (same as .pptx)
    tc.assertIsInstance(result, PptxContent)
    tc.assertTrue(len(result.slides) > 0)


def test_markdown_export():
    """Test markdown export functionality."""

    path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )

    docx: DocxContent = next(read_docx(read_file_to_file_like(path=path)))

    tc.assertEqual(
        "Hello World!\nAn image of space\nIncome\ntax\n119\n19\n"
        "Another sentence after the table.\n$$\\frac{3}{4}\\times4=\\sqrt{9}$$\n\n"
        "## Tables\n\n| Income | tax |\n|--------|-----|\n| 119    | 19  |",
        docx.get_full_markdown(),
    )


def test_password_protected__docx() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/docx-password-protected-pw123.docx"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_docx(file_like=read_file_to_file_like(path=path), path=path))


def test_password_protected__xlsx() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/xslx-password-protected-pw123.xlsx"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_xlsx(file_like=read_file_to_file_like(path=path), path=path))


def test_password_protected__pptx() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/pptx-password-protected-pw123.pptx"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_pptx(file_like=read_file_to_file_like(path=path), path=path))
