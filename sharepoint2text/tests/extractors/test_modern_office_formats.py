import logging
import typing
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.ms_modern.docx_extractor import read_docx
from sharepoint2text.parsing.extractors.ms_modern.pptx_extractor import read_pptx
from sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor import read_xlsx
from sharepoint2text.parsing.models import (
    ExtractedDocument,
    ImageAsset,
    render_markdown,
)
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None

_PptxImageMetadata = tuple[int, int, str, int | None, int | None]


def _assert_image(
    image: ImageAsset,
    *,
    number: int,
    unit_number: int,
    media_type: str,
    width: int | None,
    height: int | None,
    format_name: str,
) -> None:
    tc.assertEqual(number, image.number)
    tc.assertEqual(unit_number, image.properties[f"{format_name}.slide_number"])
    tc.assertEqual(media_type, image.media_type)
    tc.assertEqual(width, image.width)
    tc.assertEqual(height, image.height)


def _assert_pptx_image(image: ImageAsset, expected: _PptxImageMetadata) -> None:
    """Assert image number, slide number, media type, width, and height."""
    number, slide_number, media_type, width, height = expected
    _assert_image(
        image,
        number=number,
        unit_number=slide_number,
        media_type=media_type,
        width=width,
        height=height,
        format_name="pptx",
    )


####################
# Modern Microsoft #
####################
def test_read_xlsx_1() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx"
    xlsx: ExtractedDocument = next(
        read_xlsx(file_like=read_file_to_file_like(path=path))
    )

    tc.assertEqual("2006-09-16T00:00:00", xlsx.metadata.created)
    tc.assertEqual("2015-05-06T11:46:24", xlsx.metadata.modified)

    tc.assertEqual(3, len(xlsx.units))
    tc.assertEqual(3, len(list(xlsx.iter_tables())))
    tc.assertListEqual(
        sorted(["Sheet1", "Sheet2", "Sheet3"]),
        sorted(sheet.title or "" for sheet in xlsx.units),
    )
    tc.assertListEqual(
        ["AREA", "CODE", "COUNTRY NAME"], list(xlsx.iter_tables())[0].rows[0]
    )
    tc.assertEqual((52, 3), list(xlsx.iter_tables())[0].dimensions)

    # check raw data and table interface
    # check that the first row in the first sheet is the headline
    tc.assertListEqual(
        ["AREA", "CODE", "COUNTRY NAME"], xlsx.units[0].tables[0].rows[0]
    )
    tc.assertListEqual(
        ["AREA", "CODE", "COUNTRY NAME"], xlsx.units[0].tables[0].rows[0]
    )
    tc.assertListEqual(
        ["European Union (EU)", "EU-28 ", "European Union (28 countries)"],
        xlsx.units[0].tables[0].rows[1],
    )

    tc.assertEqual(3, len(list(xlsx.units)))

    tc.assertEqual("Sheet1\nAREA     CODE", xlsx.full_text[:20])

    tc.assertIsNone(xlsx.source.filename)
    tc.assertIsNone(xlsx.source.extension)
    tc.assertIsNone(xlsx.metadata.title)
    tc.assertIsNone(xlsx.metadata.subject)
    tc.assertIsNone(xlsx.metadata.author)
    tc.assertEqual("2006-09-16T00:00:00", xlsx.metadata.created)
    tc.assertEqual("2015-05-06T11:46:24", xlsx.metadata.modified)
    tc.assertListEqual([], xlsx.metadata.keywords)
    tc.assertIsNone(xlsx.metadata.language)
    tc.assertNotIn("xlsx.revision", xlsx.metadata.properties)
    tc.assertEqual(0, len(list(xlsx.iter_images())))


def test_read_xlsx_2() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/mwe.xlsx"
    xlsx: ExtractedDocument = next(
        read_xlsx(file_like=read_file_to_file_like(path=path))
    )
    tc.assertEqual(
        "Blatt 1\nTabelle 1 Unnamed: 1\n     ColA       ColB\n        1          2",
        xlsx.full_text,
    )
    tc.assertListEqual([["ColA", "ColB"], [1, 2]], xlsx.units[0].tables[0].rows)
    tc.assertEqual(0, len(list(xlsx.iter_images())))


def test_read_xlsx_3() -> None:
    """Verifies the treatment of empty rows and columns in a sheet

    We want that the list of rows is easily processable with Pandas or Polars to create
    dataframes. This requires that None/Nulls are not accidentally pruned. The rows must have
    the same number of columns for this to work
    """
    path = "sharepoint2text/tests/resources/modern_ms/empty_row_columns.xlsx"

    xlsx: ExtractedDocument = next(
        read_xlsx(file_like=read_file_to_file_like(path=path))
    )
    tc.assertListEqual(
        [
            [None, "Name", None, "Age"],
            [None, "A", None, 25],
            [None, None, None, None],
            [None, "B", None, 28],
        ],
        xlsx.units[0].tables[0].rows,
    )
    tc.assertEqual(0, len(list(xlsx.iter_images())))
    tc.assertEqual((4, 4), list(xlsx.iter_tables())[0].dimensions)

    #########
    # Units #
    #########
    units = list(xlsx.units)
    tc.assertEqual(1, len(units))
    tc.assertListEqual(
        [
            [None, "Name", None, "Age"],
            [None, "A", None, 25],
            [None, None, None, None],
            [None, "B", None, 28],
        ],
        units[0].tables[0].rows,
    )
    tc.assertEqual(0, len(units[0].images))
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("sheet", units[0].kind)
    tc.assertEqual("Blatt 1", units[0].title)


def test_read_xlsx_4__image_extraction() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/image_in_excel.xlsx"

    xlsx: ExtractedDocument = next(
        read_xlsx(file_like=read_file_to_file_like(path=path))
    )
    tc.assertEqual("Image Sheet", xlsx.units[0].title)
    tc.assertEqual(1, len(xlsx.units[0].images))

    image = xlsx.units[0].images[0]
    tc.assertEqual(7280, len(image.data or b""))
    tc.assertEqual("Image 1", image.caption)
    tc.assertEqual("Picture", image.description)
    tc.assertEqual(600, image.width)
    tc.assertEqual(300, image.height)

    tc.assertEqual(1, len(list(xlsx.iter_images())))
    extracted_image = list(xlsx.iter_images())[0]
    tc.assertEqual(1, extracted_image.number)
    tc.assertEqual(1, extracted_image.properties["xlsx.sheet_index"])
    tc.assertEqual("image/png", extracted_image.media_type)
    tc.assertEqual(600, extracted_image.width)
    tc.assertEqual(300, extracted_image.height)
    tc.assertAlmostEqual(600 / 300, extracted_image.ratio or 0.0)

    #########
    # Units #
    #########
    units = list(xlsx.units)
    tc.assertEqual(1, len(units))
    tc.assertEqual("image/png", units[0].images[0].media_type)
    tc.assertEqual(7280, len(units[0].images[0].data or b""))
    tc.assertEqual("Image 1", list(units[0].images)[0].caption)
    tc.assertEqual("Picture", list(units[0].images)[0].description)


def test_read_pptx_1() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/eu-visibility_rules_00704232-AF9F-1A18-BD782C469454ADAD_68401.pptx"
    pptx: ExtractedDocument = next(read_pptx(read_file_to_file_like(path=path)))

    # metadata
    tc.assertEqual("IVAN Anda-Otilia", pptx.metadata.author)
    tc.assertEqual(
        "MAGLI Mia (JUST)", pptx.metadata.properties["pptx.last_modified_by"]
    )
    tc.assertEqual("2011-10-28T10:25:18", pptx.metadata.created)
    tc.assertEqual("2020-07-12T09:25:35", pptx.metadata.modified)

    tc.assertEqual(3, len(pptx.units))
    tc.assertEqual(5, len(list(pptx.iter_images())))
    tc.assertEqual(0, len(list(pptx.iter_tables())))
    images = list(pptx.iter_images())
    _assert_image(
        images[0],
        number=1,
        unit_number=2,
        media_type="image/png",
        width=130,
        height=111,
        format_name="pptx",
    )
    _assert_image(
        images[1],
        number=2,
        unit_number=2,
        media_type="image/jpeg",
        width=264,
        height=255,
        format_name="pptx",
    )
    _assert_image(
        images[2],
        number=3,
        unit_number=2,
        media_type="image/jpeg",
        width=279,
        height=186,
        format_name="pptx",
    )
    _assert_image(
        images[3],
        number=4,
        unit_number=2,
        media_type="image/jpeg",
        width=305,
        height=250,
        format_name="pptx",
    )
    _assert_image(
        images[4],
        number=5,
        unit_number=2,
        media_type="image/jpeg",
        width=286,
        height=191,
        format_name="pptx",
    )
    ##########
    # SLIDES #
    ##########
    # slide 1
    tc.assertEqual("EU-funding visibility - art. 22 GA", pptx.units[0].title)
    expected = [
        'To be applied on all materials and communication activities:\n\nThe correct EU emblem: https://europa.eu/european-union/about-eu/symbols/flag_en; \nThe reference to the correct funding programme (to be put next to the EU emblem): “This [e.g. project, report, publication, conference, website, etc.] was funded by the European Union’s Justice Programme (2014-2020) or by the Rights, Equality and Citizenship Programme (REC 2014-2020)“;\n The following disclaimer: "The content of this [insert appropriate description, e.g. report, publication, conference, etc.] represents the views of the author only and is his/her sole responsibility. The European Commission does not accept any responsibility for use that may be made of the information it contains”.'
    ]
    tc.assertListEqual(expected, pptx.units[0].properties["pptx.content_placeholders"])

    tc.assertListEqual(["1"], pptx.units[0].properties["pptx.other_textboxes"])
    tc.assertEqual(1, pptx.units[0].number)

    # images
    tc.assertEqual(0, len(pptx.units[0].images))

    # slide 2
    tc.assertEqual("EU-funding visibility", pptx.units[1].title)

    expected = [
        "! Please choose the correct name of the funding Programme, either JUSTICE or REC, depending under which Programme your call for proposals was published and your grant was awarded:\n\nSupported by the Rights, Equality \x0band Citizenship Programme \nof the European Union (2014-2020) \x0b\n     This project is funded by the Justice \n      Programme of the European Union \n      (2014-2020)"
    ]
    tc.assertListEqual(expected, pptx.units[1].properties["pptx.content_placeholders"])

    # Order reflects visual position on slide (top to bottom, left to right)
    expected = ["This is the wrong EU emblem", "Don’t use this emblem!", "2"]
    tc.assertListEqual(expected, pptx.units[1].properties["pptx.other_textboxes"])

    # images (sorted by position on slide)
    tc.assertEqual(5, len(pptx.units[1].images))
    # test presence of metadata for first image (now image.png due to position sort)
    tc.assertEqual(1, pptx.units[1].images[0].number)
    tc.assertEqual("image.png", pptx.units[1].images[0].filename)
    tc.assertEqual("image/png", pptx.units[1].images[0].media_type)
    tc.assertEqual(12538, pptx.units[1].images[0].properties["pptx.size_bytes"])
    tc.assertIsNotNone(pptx.units[1].images[0].data)

    # full text
    expected = (
        "EU-funding visibility - art. 22 GA"
        + "\n"
        + "To be applied on all materials and communica"
    )
    tc.assertEqual(expected, pptx.full_text[:79])

    #########
    # Units #
    #########
    tc.assertEqual(3, len(list(pptx.units)))
    tc.assertEqual(len(pptx.units), len(list(pptx.units)))
    units = list(pptx.units)

    tc.assertEqual(0, len(units[0].images))
    tc.assertEqual("EU-funding visibility", units[0].text[:21])
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("slide", units[0].kind)
    tc.assertEqual("EU-funding visibility - art. 22 GA", units[0].title)

    tc.assertEqual(5, len(units[1].images))
    tc.assertEqual("This is the wrong EU ", units[1].text[:21])

    tc.assertEqual(0, len(units[2].images))
    tc.assertEqual("EU-funding visibility", units[2].text[:21])


def test_read_pptx_2() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    pptx: ExtractedDocument = next(read_pptx(read_file_to_file_like(path=path)))

    # Test canonical full text - formulas included (no comments or image captions)
    # Note: "A beach" is a regular textbox, not an image caption
    base_text = pptx.full_text
    tc.assertEqual(
        "The slide title\nThe first text line\n\n\n\n\nThe last text line\nA beach\n$$f(x)=\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}e^{-\\frac{(x-\\mu)^{2}}{2\\sigma^{2}}}$$",
        base_text,
    )

    # images
    tc.assertEqual(1, len(list(pptx.iter_images())))
    tc.assertEqual(1, len(pptx.units[0].images))
    tc.assertEqual(1, pptx.units[0].images[0].number)
    tc.assertEqual("image/jpeg", pptx.units[0].images[0].media_type)
    tc.assertEqual(1, pptx.units[0].images[0].properties["pptx.slide_number"])
    tc.assertEqual(1535390, pptx.units[0].images[0].properties["pptx.size_bytes"])
    # description is the alt text for accessibility (from descr attribute)
    tc.assertEqual(
        "Sandiger Weg zwischen zwei Hügeln, die ans Meer führen",
        pptx.units[0].images[0].description,
    )
    # caption is the shape name/title (from name attribute)
    # Note: in this file, name and descr have the same value
    tc.assertEqual(
        "Sandiger Weg zwischen zwei Hügeln, die ans Meer führen",
        pptx.units[0].images[0].caption,
    )

    # Canonical image description contains the caption (title/name)
    tc.assertEqual(
        "Sandiger Weg zwischen zwei Hügeln, die ans Meer führen",
        pptx.units[0].images[0].description,
    )
    tc.assertEqual(1535390, len(pptx.units[0].images[0].data or b""))
    _assert_image(
        pptx.units[0].images[0],
        number=1,
        unit_number=1,
        media_type="image/jpeg",
        width=1647,
        height=1098,
        format_name="pptx",
    )

    # comments go separately - they are not part of the full text body
    comments = [
        annotation
        for annotation in pptx.units[0].annotations
        if annotation.kind == "comment"
    ]
    tc.assertEqual(1, len(comments))
    tc.assertEqual("0", comments[0].author)
    tc.assertEqual("Not second?", comments[0].text)
    tc.assertEqual("2025-12-28T11:15:49.694", comments[0].properties["pptx.date"])
    tc.assertNotIn("Not second?", pptx.full_text)


def test_read_pptx_3() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_table.pptx"
    pptx: ExtractedDocument = next(read_pptx(read_file_to_file_like(path=path)))

    tc.assertEqual(1, len(pptx.document_tables))
    tc.assertEqual(1, len(list(pptx.iter_tables())))
    table_1 = list(pptx.iter_tables())[0]
    tc.assertListEqual(
        [
            ["", "2020", "2021", "2022"],
            ["A", "1", "2", "3"],
            ["B", "4", "5", "6"],
            ["C", "7", "8", "9"],
            ["D", "10", "11", "12"],
        ],
        table_1.rows,
    )
    tc.assertListEqual(pptx.document_tables[0].rows, table_1.rows)
    tc.assertEqual((5, 4), table_1.dimensions)
    tc.assertEqual(
        "2020\t2021\t2022\nA\t1\t2\t3\nB\t4\t5\t6\nC\t7\t8\t9\nD\t10\t11\t12",
        pptx.full_text,
    )

    #########
    # Units #
    #########
    units = list(pptx.units)
    tc.assertEqual(1, len(units))
    tc.assertListEqual(
        [
            ["", "2020", "2021", "2022"],
            ["A", "1", "2", "3"],
            ["B", "4", "5", "6"],
            ["C", "7", "8", "9"],
            ["D", "10", "11", "12"],
        ],
        units[0].tables[0].rows,
    )


def test_read_pptx_4_units() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/slide_titles.pptx"
    pptx: ExtractedDocument = next(read_pptx(read_file_to_file_like(path=path)))

    units = list(pptx.units)
    tc.assertEqual(2, len(units))
    tc.assertEqual("Title Slide 1", units[0].title)
    tc.assertEqual("Title Slide 2", units[1].title)


def test_read_pptx_5_full() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/Scope Ratings_Italy_update_October_2022.pptx"
    pptx: ExtractedDocument = next(read_pptx(read_file_to_file_like(path=path)))

    # full content
    tc.assertTrue(
        pptx.full_text.startswith("Italy’s credit outlook after the election")
    )
    tc.assertTrue(pptx.full_text.endswith("e +49 30 27891-0.\n15"))

    tc.assertEqual(17, len(pptx.document_images))

    # units
    tc.assertEqual(15, len(pptx.units))
    expected_text_boundaries = [
        ("Italy’s credit outlo", "ngs.com\nOctober 2022"),
        ("Italy’s credit outlo", "s over the long term"),
        ("Italy’s credit outlo", "&P, Moody’s, Fitch\n3"),
        ("Italy’s credit outlo", "irs, Scope Ratings\n4"),
        ("Italy’s credit outlo", "ond, Scope Ratings\n5"),
        ("Italy’s credit outlo", "ond, Scope Ratings\n6"),
        ("Italy’s credit outlo", "nce, Scope Ratings\n7"),
        ("Italy’s credit outlo", "ion, Scope Ratings\n8"),
        ("Italy’s credit outlo", "ond, Scope Ratings\n9"),
        ("Italy’s credit outlo", "nd, Scope Ratings\n10"),
        ("Italy’s credit outlo", " growth potential\n11"),
        ("Annex: Documentation", " housing methodology"),
        ("Headquarters\x0bBERLIN\x0b", "ww.scopeexplorer.com"),
        ("About Scope Group\nSc", "tive consistency.\n14"),
        ("Disclaimer\n“Scope Gr", "e +49 30 27891-0.\n15"),
    ]
    for unit, (expected_start, expected_end) in zip(
        pptx.units, expected_text_boundaries, strict=True
    ):
        tc.assertEqual(expected_start, unit.text[:20])
        tc.assertEqual(expected_end, unit.text[-20:])

    # 0
    tc.assertEqual(2, len(pptx.units[0].images))
    _assert_pptx_image(pptx.units[0].images[0], (1, 1, "image/jpeg", 302, 149))
    _assert_pptx_image(pptx.units[0].images[1], (2, 1, "image/png", 1104, 605))
    tc.assertIsNone(pptx.units[0].title)
    tc.assertListEqual(
        [
            "Italy’s credit outlook after the elections",
            "Alvise Lennkh-Yunus, CFA\nExecutive Director, Sovereign and Public Sector\x0ba.lennkh@scoperatings.com\n\nGiulia Branz, CFA\nSenior Analyst, Sovereign and Public Sector\x0bg.branz@scoperatings.com\n\nAlessandra Poli\nAssociate Analyst, Sovereign and Public Sector\x0ba.poli@scoperatings.com",
            "October 2022",
        ],
        pptx.units[0].properties["pptx.other_textboxes"],
    )

    # 1
    tc.assertEqual(0, len(pptx.units[1].images))
    # 2
    tc.assertEqual(1, len(pptx.units[2].images))
    _assert_pptx_image(pptx.units[2].images[0], (1, 3, "image/x-emf", None, None))
    # 3
    tc.assertEqual(2, len(pptx.units[3].images))
    _assert_pptx_image(pptx.units[3].images[0], (1, 4, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[3].images[1], (2, 4, "image/x-emf", None, None))
    # 4
    tc.assertEqual(2, len(pptx.units[4].images))
    _assert_pptx_image(pptx.units[4].images[0], (1, 5, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[4].images[1], (2, 5, "image/x-emf", None, None))
    # 5
    tc.assertEqual(2, len(pptx.units[5].images))
    _assert_pptx_image(pptx.units[5].images[0], (1, 6, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[5].images[1], (2, 6, "image/x-emf", None, None))
    # 6
    tc.assertEqual(2, len(pptx.units[6].images))
    _assert_pptx_image(pptx.units[6].images[0], (1, 7, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[6].images[1], (2, 7, "image/x-emf", None, None))
    # 7
    tc.assertEqual(2, len(pptx.units[7].images))
    _assert_pptx_image(pptx.units[7].images[0], (1, 8, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[7].images[1], (2, 8, "image/x-emf", None, None))
    # 8
    tc.assertEqual(2, len(pptx.units[8].images))
    _assert_pptx_image(pptx.units[8].images[0], (1, 9, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[8].images[1], (2, 9, "image/x-emf", None, None))
    # 9
    tc.assertEqual(2, len(pptx.units[9].images))
    _assert_pptx_image(pptx.units[9].images[0], (1, 10, "image/x-emf", None, None))
    _assert_pptx_image(pptx.units[9].images[1], (2, 10, "image/x-emf", None, None))
    # 10
    tc.assertEqual(0, len(pptx.units[10].images))
    # 11
    tc.assertEqual(0, len(pptx.units[11].images))
    tc.assertTrue(
        pptx.units[11].text.startswith("Annex: Documentation\nAdditional document")
    )
    # 12
    tc.assertEqual(0, len(pptx.units[12].images))
    # 13
    tc.assertEqual(0, len(pptx.units[13].images))
    # 14
    tc.assertEqual(0, len(pptx.units[14].images))


def test_read_pptx__image_flag() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/pptx_images.pptx"
    pptx: ExtractedDocument = next(
        read_pptx(read_file_to_file_like(path=path), ignore_images=False)
    )
    tc.assertEqual(1, len(pptx.units[0].images))
    tc.assertEqual("PPTX text", pptx.full_text)

    pptx: ExtractedDocument = next(
        read_pptx(read_file_to_file_like(path=path), ignore_images=True)
    )
    tc.assertEqual(0, len(pptx.units[0].images))
    tc.assertEqual("PPTX text", pptx.full_text)


def test_read_docx_1() -> None:
    # An actual document from the web - this is likely created on a Windows client
    path = (
        "sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx"
    )
    docx: ExtractedDocument = next(read_docx(read_file_to_file_like(path=path)))

    # text is long. Verify only beginning
    tc.assertEqual("Welcome to the Government", docx.full_text[:25].strip())

    tc.assertEqual(230, docx.properties["docx.paragraph_count"])

    tc.assertEqual(17, docx.metadata.properties["docx.revision"])
    # Raw XML format uses 'Z' for UTC timezone
    tc.assertEqual("2023-01-20T16:07:00Z", docx.metadata.modified)
    tc.assertEqual("2022-04-19T14:03:00Z", docx.metadata.created)

    # test iterator
    tc.assertEqual(1, len(list(docx.units)))

    tc.assertEqual(1, len(list(docx.iter_images())))
    tc.assertEqual(1, len(docx.document_images))

    # images
    tc.assertEqual(1, len(list(docx.iter_images())))
    tc.assertEqual(docx.document_images[0], docx.units[0].images[0])

    tc.assertEqual(7, len(list(docx.iter_tables())))
    image = list(docx.iter_images())[0]
    tc.assertEqual(1, image.number)
    tc.assertEqual("image/png", image.media_type)
    tc.assertEqual(1823, image.width)
    tc.assertEqual(1052, image.height)

    # test full text
    tc.assertEqual("Welcome to the Government", docx.full_text[:25].strip())


def test_read_docx_2() -> None:
    # A converted docx from OSX pages - may not populate like a true MS client .docx
    # dedicated test for comment, table and footnote extraction
    path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )

    docx: ExtractedDocument = next(read_docx(read_file_to_file_like(path=path)))
    # Formula with properly converted multiplication sign
    tc.assertEqual(
        "Hello World!\nAn image of space\nIncome\ntax\n119\n19\nAnother sentence after the table.\n$$\\frac{3}{4}\\times4=\\sqrt{9}$$",
        docx.full_text,
    )
    tc.assertEqual(docx.full_text, docx.units[0].text)
    tc.assertNotIn("Nice!", docx.full_text)
    comments = [
        annotation
        for annotation in docx.document_annotations
        if annotation.kind == "comment"
    ]
    tc.assertEqual(1, len(comments))
    tc.assertEqual("0", comments[0].properties["docx.id"])
    tc.assertEqual("User", comments[0].author)
    tc.assertEqual("2025-12-28T09:16:57Z", comments[0].properties["docx.date"])
    tc.assertEqual("Nice!", comments[0].text)

    footnotes = [
        annotation
        for annotation in docx.document_annotations
        if annotation.kind == "footnote"
    ]
    tc.assertListEqual(["-2", "1"], [note.properties["docx.id"] for note in footnotes])
    tc.assertListEqual(["", "A simple footnote"], [note.text for note in footnotes])

    headers = [
        annotation
        for annotation in docx.document_annotations
        if annotation.kind == "header"
    ]
    footers = [
        annotation
        for annotation in docx.document_annotations
        if annotation.kind == "footer"
    ]
    tc.assertListEqual(["My header"], [header.text for header in headers])
    tc.assertListEqual(
        ["default"], [header.properties["docx.type"] for header in headers]
    )
    tc.assertListEqual(["My footer"], [footer.text for footer in footers])
    tc.assertListEqual(
        ["default"], [footer.properties["docx.type"] for footer in footers]
    )

    tc.assertListEqual(
        [["Income", "tax"], ["119", "19"]], list(docx.iter_tables())[0].rows
    )

    # formulas (with converted multiplication sign)
    formulas = [
        annotation
        for annotation in docx.document_annotations
        if annotation.kind == "formula"
    ]
    tc.assertEqual(1, len(formulas))
    tc.assertEqual("\\frac{3}{4}\\times4=\\sqrt{9}", formulas[0].text)
    tc.assertIs(formulas[0].properties["docx.is_display"], True)

    # section object
    sections = typing.cast(
        list[dict[str, float | str | None]], docx.properties["docx.sections"]
    )
    tc.assertEqual(1, len(sections))
    tc.assertAlmostEqual(8.268, sections[0]["page_width_inches"], places=1)
    tc.assertAlmostEqual(11.693, sections[0]["page_height_inches"], places=1)
    tc.assertAlmostEqual(0.7875, sections[0]["left_margin_inches"], places=1)
    tc.assertAlmostEqual(0.7875, sections[0]["right_margin_inches"], places=1)
    tc.assertAlmostEqual(0.7875, sections[0]["top_margin_inches"], places=1)
    tc.assertAlmostEqual(0.7875, sections[0]["bottom_margin_inches"], places=1)
    tc.assertIsNone(sections[0]["orientation"])

    # images
    images = list(docx.iter_images())
    tc.assertEqual(1, len(images))
    tc.assertEqual(1, len(list(docx.iter_images())))
    tc.assertEqual(1, len(list(docx.iter_tables())))
    tc.assertListEqual(
        [["Income", "tax"], ["119", "19"]], list(docx.iter_tables())[0].rows
    )
    tc.assertEqual(1, images[0].number)
    tc.assertEqual("image1.png", images[0].filename)
    tc.assertEqual("image/png", images[0].media_type)
    # description (alt text) is from pic:cNvPr[@descr]
    tc.assertEqual("Space", images[0].description)
    # caption is from the text box content (wps:txbx)
    tc.assertEqual("An image of space", images[0].caption)

    # Canonical image fields
    tc.assertEqual("image/png", images[0].media_type)
    tc.assertEqual("Space", images[0].description)
    tc.assertEqual("An image of space", images[0].caption)
    tc.assertEqual(828786, len(images[0].data or b""))
    tc.assertEqual(828786, images[0].properties["docx.size_bytes"])
    tc.assertEqual(930, images[0].width)
    tc.assertEqual(506, images[0].height)


def test_read_docx__image_flag() -> None:
    # A converted docx from OSX pages - may not populate like a true MS client .docx
    # dedicated test for comment, table and footnote extraction
    path = "sharepoint2text/tests/resources/modern_ms/document_with_image.docx"
    docx: ExtractedDocument = next(
        read_docx(read_file_to_file_like(path=path), ignore_images=False)
    )
    tc.assertEqual(1, len(list(docx.iter_images())))
    tc.assertEqual("Docx with image", docx.full_text)

    docx: ExtractedDocument = next(
        read_docx(read_file_to_file_like(path=path), ignore_images=True)
    )
    tc.assertEqual(0, len(list(docx.iter_images())))
    tc.assertEqual("Docx with image", docx.full_text)


def test_read_docx__image_extraction_1() -> None:
    # Test for caption extraction from following paragraph with caption style
    path = "sharepoint2text/tests/resources/modern_ms/vorlage-abschlussarbeit.docx"
    docx: ExtractedDocument = next(read_docx(read_file_to_file_like(path=path)))

    images = list(docx.iter_images())
    tc.assertEqual(1, len(images))
    tc.assertEqual(1, len(list(docx.iter_images())))
    tc.assertEqual(0, len(list(docx.iter_tables())))
    # image interface - caption from following paragraph with "HA-Bildunterschrift" style
    expected_caption = (
        "Abb. 1: Eine aus dem Internet heruntergeladene Bilddatei mit einer "
        "Bildunterschrift. Die Abbildungen und Tabellen bitte nicht als "
        "textumflossene Objekte, sondern so wie dies Bild als Absatz in den "
        "Text einbinden. Dieser Untertext hat die Formatvorlage "
        "\u201eHA-Bildunterschrift\u201c."
    )
    tc.assertEqual(expected_caption, images[0].caption)
    # description is the alt text (URL in this case)
    tc.assertEqual(
        "http://omgunmen.de/wp-content/uploads/2011/03/but-on-math-it-is.png",
        images[0].description,
    )


def test_read_docx__image_extraction_2() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/thesis-template.docx"
    docx: ExtractedDocument = next(read_docx(read_file_to_file_like(path=path)))

    images = list(docx.iter_images())
    tc.assertEqual(2, len(images))
    tc.assertEqual(2, len(list(docx.iter_images())))
    tc.assertEqual(4, len(list(docx.iter_tables())))
    figure_image = next(image for image in images if image.number == 2)
    tc.assertEqual("Illustration 1: [Figure title]", figure_image.caption)
    tc.assertEqual(
        """Ein Bild, das Zeichnung "Marketing" enthält.""",
        (figure_image.description or "").strip(),
    )

    # units
    tc.assertEqual(17, len(list(docx.units)))
    units = list(docx.units)
    tc.assertListEqual(["II. List of figures"], units[0].heading_path)
    tc.assertListEqual(["III. List of tables"], units[1].heading_path)
    tc.assertListEqual(["IV. List of formulas"], units[2].heading_path)
    tc.assertListEqual(["V. List of abbreviations"], units[3].heading_path)
    tc.assertListEqual(["VI. List of symbols"], units[4].heading_path)
    tc.assertListEqual(["Title 1 Chapter"], units[5].heading_path)
    tc.assertListEqual(["Title 2 Chapter"], units[6].heading_path)
    tc.assertListEqual(
        ["Title 2 Chapter", "2.1 Title Subchapter"], units[7].heading_path
    )
    # unit has an image
    tc.assertListEqual(
        ["Title 2 Chapter", "2.1 Title Subchapter", "2.1.1 Title Subchapter"],
        units[8].heading_path,
    )
    tc.assertEqual(54423, len(units[8].images[0].data or b""))

    # unit has an table
    tc.assertListEqual(
        ["Title 2 Chapter", "2.1 Title Subchapter", "2.1.2 Title Subchapter"],
        units[9].heading_path,
    )
    tc.assertEqual((3, 4), units[9].tables[0].dimensions)

    tc.assertListEqual(
        ["Title 2 Chapter", "2.2 Title Subchapter"],
        units[10].heading_path,
    )
    tc.assertListEqual(["Title 3 Chapter"], units[11].heading_path)
    tc.assertListEqual(["Title 4 Chapter"], units[12].heading_path)
    tc.assertListEqual(["VII. Appendix"], units[13].heading_path)
    tc.assertListEqual(["VIII. Bibliography"], units[14].heading_path)
    tc.assertListEqual(["VIII. Bibliography"], units[15].heading_path)
    tc.assertListEqual(["IX. Affidavit"], units[16].heading_path)


def test_read_docx__units() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/headings.docx"
    docx: ExtractedDocument = next(
        read_docx(file_like=read_file_to_file_like(path=path))
    )

    units = list(docx.units)
    tc.assertEqual(8, len(units))

    tc.assertIsInstance(units[0].images, list)
    tc.assertIsInstance(units[0].tables, list)

    # first unit
    tc.assertEqual(["Sample Document"], units[0].heading_path)
    tc.assertEqual(
        "This document was created using accessibility techniques for headings, lists, image alternate text, tables, and columns. It should be completely accessible using assistive technologies such as screen readers.",
        units[0].text,
    )
    tc.assertEqual(0, len(units[0].images))
    tc.assertEqual(0, len(units[0].tables))

    # second unit
    tc.assertEqual(["Sample Document", "Headings"], units[1].heading_path)
    tc.assertEqual(
        'There are eight section headings in this document. At the beginning, "Sample Document" is a level 1 heading. The main section headings, such as "Headings" and "Lists" are level 2 headings. The Tables section contains two sub-headings, "Simple Table" and "Complex Table," which are both level 3 headings.',
        units[1].text,
    )
    tc.assertEqual(0, len(units[1].images))
    tc.assertEqual(0, len(units[1].tables))

    # third unit
    tc.assertEqual(["Sample Document", "Lists"], units[2].heading_path)
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
        units[2].text,
    )
    tc.assertEqual(0, len(units[2].images))
    tc.assertEqual(0, len(units[2].tables))

    # Images section
    tc.assertEqual(["Sample Document", "Images"], units[4].heading_path)
    tc.assertEqual(2, len(units[4].images))
    tc.assertSetEqual(
        {"image1.gif", "image2.png"}, {img.filename for img in units[4].images}
    )
    tc.assertEqual(5437, len(units[4].images[0].data or b""))
    tc.assertEqual(7570, len(units[4].images[1].data or b""))
    tc.assertEqual(0, len(units[4].tables))

    # Tables section
    tc.assertEqual(1, len(units[5].tables))
    tc.assertEqual(list(docx.iter_tables())[0].rows, units[5].tables[0].rows)
    tc.assertEqual(1, len(units[6].tables))
    tc.assertEqual(list(docx.iter_tables())[1].rows, units[6].tables[0].rows)


def test_read_docx__unit_structure() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/word_structure.docx"
    docx: ExtractedDocument = next(
        read_docx(file_like=read_file_to_file_like(path=path))
    )

    units = list(docx.units)
    tc.assertEqual(5, len(units))

    unit1 = units[0]
    tc.assertEqual(["The document title"], unit1.heading_path)
    tc.assertEqual("blabla", unit1.text)

    unit2 = units[1]
    tc.assertEqual(["The document title", "Chapter 1"], unit2.heading_path)
    tc.assertEqual("This is chapter 1", unit2.text)

    unit3 = units[2]
    tc.assertEqual(
        ["The document title", "Chapter 1", "Section 1.1"],
        unit3.heading_path,
    )
    tc.assertEqual("A subsection", unit3.text)

    unit4 = units[3]
    tc.assertEqual(["The document title", "Chapter 2"], unit4.heading_path)
    tc.assertEqual("This is chapter 2", unit4.text)

    unit5 = units[4]
    tc.assertEqual(["The document title", "Chapter 3"], unit5.heading_path)
    tc.assertEqual("This is chapter 3", unit5.text)


def test_read_macro_enabled_docm() -> None:
    """Test .docm (macro-enabled Word) extraction - same structure as .docx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.docm"
    result: ExtractedDocument = next(
        read_docx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as ExtractedDocument (same as .docx)
    tc.assertIsInstance(result, ExtractedDocument)
    tc.assertTrue(len(result.full_text) > 0)


def test_read_macro_enabled_xlsm() -> None:
    """Test .xlsm (macro-enabled Excel) extraction - same structure as .xlsx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.xlsm"
    result: ExtractedDocument = next(
        read_xlsx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as ExtractedDocument (same as .xlsx)
    tc.assertIsInstance(result, ExtractedDocument)
    tc.assertTrue(len(result.units) > 0)


def test_read_xlsb() -> None:
    """Test that XLSB extraction preserves worksheets, rows, and cell positions."""
    path = "sharepoint2text/tests/resources/modern_ms/excel.xlsb"
    result: ExtractedDocument = next(
        read_xlsx(file_like=read_file_to_file_like(path=path), path=path)
    )
    tc.assertListEqual(
        ["Sheet1", "Sheet2", "Sheet3"], [sheet.title for sheet in result.units]
    )
    tc.assertEqual(3, len(list(result.iter_tables())))

    sheet = result.units[0]
    tc.assertEqual((11, 52), sheet.tables[0].dimensions)
    tc.assertEqual("Atable", sheet.tables[0].rows[0][0])
    tc.assertEqual("Btable", sheet.tables[0].rows[0][2])
    tc.assertEqual("Zparam", sheet.tables[0].rows[0][18])
    tc.assertEqual(1.01, sheet.tables[0].rows[1][18])
    tc.assertEqual(1.0, sheet.tables[0].rows[2][2])
    tc.assertListEqual([], result.units[1].tables[0].rows)
    tc.assertListEqual([], result.units[2].tables[0].rows)

    full_text = result.full_text
    tc.assertTrue(full_text.startswith("Sheet1\nAtable"))
    tc.assertIn("Sheet2", full_text)
    tc.assertIn("Sheet3", full_text)


def test_read_xlsx__image_flag() -> None:
    """Test .xlsm (macro-enabled Excel) extraction - same structure as .xlsx."""
    path = "sharepoint2text/tests/resources/modern_ms/excel_images.xlsx"
    result: ExtractedDocument = next(
        read_xlsx(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=False
        ),
    )
    tc.assertEqual(1, len(result.units[0].images))

    result: ExtractedDocument = next(
        read_xlsx(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        ),
    )
    tc.assertEqual(0, len(result.units[0].images))


def test_read_macro_enabled_pptm() -> None:
    """Test .pptm (macro-enabled PowerPoint) extraction - same structure as .pptx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.pptm"
    result: ExtractedDocument = next(
        read_pptx(file_like=read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as ExtractedDocument (same as .pptx)
    tc.assertIsInstance(result, ExtractedDocument)
    tc.assertTrue(len(result.units) > 0)


def test_markdown_export() -> None:
    """Retain the exact DOCX Markdown content produced on the main branch."""
    path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )
    document = next(read_docx(read_file_to_file_like(path=path)))

    tc.assertEqual(
        "Hello World!\nAn image of space\nIncome\ntax\n119\n19\n"
        "Another sentence after the table.\n$$\\frac{3}{4}\\times4=\\sqrt{9}$$\n\n"
        "## Tables\n\n| Income | tax |\n|--------|-----|\n| 119    | 19  |",
        render_markdown(document),
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


def test_extract_annotations__docx() -> None:
    """Test that extract_annotations flag includes comments in text and ensures consistency."""
    path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )

    # Without extract_annotations - comments should NOT be in full_text
    docx: ExtractedDocument = next(
        read_docx(read_file_to_file_like(path=path), extract_annotations=False)
    )
    tc.assertNotIn("Nice!", docx.full_text)
    tc.assertNotIn("[Comment:", docx.full_text)

    # With extract_annotations - comments should be in full_text
    docx_with_annotations: ExtractedDocument = next(
        read_docx(read_file_to_file_like(path=path), extract_annotations=True)
    )
    tc.assertIn("[Comment:", docx_with_annotations.full_text)
    tc.assertIn("Nice!", docx_with_annotations.full_text)

    # Verify full_text equals concatenated unit text (key consistency requirement)
    concatenated_unit_text = "\n".join(
        unit.text for unit in docx_with_annotations.units if unit.text
    ).strip()
    tc.assertEqual(docx_with_annotations.full_text, concatenated_unit_text)


def test_extract_annotations__pptx() -> None:
    """Test that extract_annotations flag includes comments in text and ensures consistency."""
    path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"

    # Without extract_annotations - comments should NOT be in full_text
    pptx: ExtractedDocument = next(
        read_pptx(read_file_to_file_like(path=path), extract_annotations=False)
    )
    tc.assertNotIn("Not second?", pptx.full_text)
    tc.assertNotIn("[Comment:", pptx.full_text)

    # With extract_annotations - comments should be in full_text
    pptx_with_annotations: ExtractedDocument = next(
        read_pptx(read_file_to_file_like(path=path), extract_annotations=True)
    )
    tc.assertIn("[Comment:", pptx_with_annotations.full_text)
    tc.assertIn("Not second?", pptx_with_annotations.full_text)

    # Verify full_text equals concatenated unit text (key consistency requirement)
    concatenated_unit_text = "\n".join(
        unit.text for unit in pptx_with_annotations.units if unit.text
    ).strip()
    tc.assertEqual(pptx_with_annotations.full_text, concatenated_unit_text)


def test_extract_annotations__api_integration() -> None:
    """Test that extract_annotations flag works via the read_file API."""
    import sharepoint2text

    docx_path = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )
    pptx_path = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"

    # DOCX via read_file API
    docx: ExtractedDocument = next(
        sharepoint2text.read_file(docx_path, extract_annotations=True)
    )
    tc.assertIn("[Comment:", docx.full_text)
    tc.assertIn("Nice!", docx.full_text)
    # Verify consistency
    concatenated = "\n".join(unit.text for unit in docx.units if unit.text).strip()
    tc.assertEqual(docx.full_text, concatenated)

    # PPTX via read_file API
    pptx: ExtractedDocument = next(
        sharepoint2text.read_file(pptx_path, extract_annotations=True)
    )
    tc.assertIn("[Comment:", pptx.full_text)
    tc.assertIn("Not second?", pptx.full_text)
    # Verify consistency
    concatenated = "\n".join(unit.text for unit in pptx.units if unit.text).strip()
    tc.assertEqual(pptx.full_text, concatenated)
