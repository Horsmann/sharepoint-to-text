import io
import io as std_io
import logging
import typing
import zipfile
from unittest import TestCase

from sharepoint2text.parsing.exceptions import ExtractionFileEncryptedError
from sharepoint2text.parsing.extractors.archive_extractor import read_archive
from sharepoint2text.parsing.extractors.data_types import (
    DocContent,
    DocImage,
    DocxComment,
    DocxContent,
    DocxFormula,
    DocxNote,
    EmailContent,
    EmailUnitMetadata,
    EpubContent,
    EpubUnitMetadata,
    FileMetadataInterface,
    HtmlContent,
    HtmlUnitMetadata,
    ImageInterface,
    ImageMetadata,
    OdpContent,
    OdpUnitMetadata,
    OdsContent,
    OdtContent,
    OdtHeaderFooter,
    OdtNote,
    OdtTable,
    OdtUnitMetadata,
    OpenDocumentAnnotation,
    PdfContent,
    PdfUnitMetadata,
    PlainTextContent,
    PlainUnitMetadata,
    PptContent,
    PptImage,
    PptxComment,
    PptxContent,
    PptxUnitMetadata,
    RtfContent,
    RtfUnitMetadata,
    TableData,
    TableDim,
    XlsContent,
    XlsUnitMetadata,
    XlsxContent,
    XlsxUnitMetadata,
)
from sharepoint2text.parsing.extractors.epub_extractor import read_epub
from sharepoint2text.parsing.extractors.html_extractor import read_html
from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
    read_eml_format_mail,
)
from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
    read_mbox_format_mail,
)
from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
    read_msg_format_mail,
)
from sharepoint2text.parsing.extractors.mhtml_extractor import read_mhtml
from sharepoint2text.parsing.extractors.ms_legacy.doc_extractor import read_doc
from sharepoint2text.parsing.extractors.ms_legacy.ppt_extractor import read_ppt
from sharepoint2text.parsing.extractors.ms_legacy.rtf_extractor import read_rtf
from sharepoint2text.parsing.extractors.ms_legacy.xls_extractor import read_xls
from sharepoint2text.parsing.extractors.ms_modern.docx_extractor import read_docx
from sharepoint2text.parsing.extractors.ms_modern.pptx_extractor import read_pptx
from sharepoint2text.parsing.extractors.ms_modern.xlsx_extractor import read_xlsx
from sharepoint2text.parsing.extractors.open_office.odp_extractor import read_odp
from sharepoint2text.parsing.extractors.open_office.ods_extractor import read_ods
from sharepoint2text.parsing.extractors.open_office.odt_extractor import read_odt
from sharepoint2text.parsing.extractors.pdf.pdf_extractor import read_pdf
from sharepoint2text.parsing.extractors.plain_extractor import read_plain_text

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def _read_file_to_file_like(path: str) -> io.BytesIO:
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)
        return file_like


#############
# Interface #
#############


def test_file_metadata_extraction() -> None:
    meta = FileMetadataInterface()
    meta.populate_from_path("my/dummy/path.txt")

    tc.assertEqual("path.txt", meta.filename)
    tc.assertEqual(".txt", meta.file_extension)
    tc.assertEqual("my/dummy/path.txt", meta.file_path)
    tc.assertEqual("my/dummy", meta.folder_path)

    tc.assertDictEqual(
        {
            "filename": "path.txt",
            "file_extension": ".txt",
            "file_path": "my/dummy/path.txt",
            "folder_path": "my/dummy",
            "detected_encoding": None,
        },
        meta.to_dict(),
    )


#########
# Plain #
#########


def test_read_text() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.txt"
    plain: PlainTextContent = next(
        read_plain_text(file_like=_read_file_to_file_like(path), path=path)
    )

    tc.assertEqual("Hello World", plain.content)
    tc.assertEqual("Hello World", plain.get_full_text())
    tc.assertEqual(1, len(list(plain.iterate_units())))
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))

    units = list(plain.iterate_units())
    tc.assertTrue(isinstance(units[0].get_metadata(), PlainUnitMetadata))
    tc.assertEqual(PlainUnitMetadata(unit_number=1), units[0].get_metadata())

    meta = plain.get_metadata()
    tc.assertEqual("ascii", meta.detected_encoding)
    tc.assertEqual("plain.txt", meta.filename)
    tc.assertEqual(".txt", meta.file_extension)


def test_read_plain_csv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.csv"
    plain: PlainTextContent = next(
        read_plain_text(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual('Text; Date\n"Hello World"; "2025-12-25"', plain.content)

    tc.assertEqual(
        'Text; Date\n"Hello World"; "2025-12-25"',
        "\n".join(unit.get_text() for unit in plain.iterate_units()),
    )
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))


def test_read_plain_tsv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.tsv"
    plain = next(
        read_plain_text(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Text\tDate\nHello World\t2025-12-25", plain.content)
    tc.assertEqual("Text\tDate\nHello World\t2025-12-25", plain.get_full_text())
    tc.assertEqual(
        "Text\tDate\nHello World\t2025-12-25",
        "\n".join(unit.get_text() for unit in plain.iterate_units()),
    )
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))


def test_read_plain_markdown() -> None:
    path = "sharepoint2text/tests/resources/plain_text/document.md"
    plain = next(
        read_plain_text(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("# Markdown file\n\nThis is a text", plain.content)
    tc.assertEqual("# Markdown file\n\nThis is a text", plain.get_full_text())
    tc.assertEqual(
        "# Markdown file\n\nThis is a text",
        "\n".join(unit.get_text() for unit in plain.iterate_units()),
    )
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))


####################
# Modern Microsoft #
####################
def test_read_xlsx_1() -> None:
    path = "sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx"
    xlsx: XlsxContent = next(read_xlsx(file_like=_read_file_to_file_like(path=path)))

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
    xlsx: XlsxContent = next(read_xlsx(file_like=_read_file_to_file_like(path=path)))
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

    xlsx: XlsxContent = next(read_xlsx(file_like=_read_file_to_file_like(path=path)))
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

    xlsx: XlsxContent = next(read_xlsx(file_like=_read_file_to_file_like(path=path)))
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
            unit_number=None,
            image_number=1,
            content_type="image/png",
            width=600,
            height=300,
        ),
        img_meta,
    )
    tc.assertIsNone(img_meta.unit_number)
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
    pptx: PptxContent = next(read_pptx(_read_file_to_file_like(path=path)))

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
    tc.assertIsNotNone(pptx.slides[1].images[0].blob)

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
    pptx: PptxContent = next(read_pptx(_read_file_to_file_like(path=path)))

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
    pptx: PptxContent = next(read_pptx(_read_file_to_file_like(path=path)))

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


def test_read_docx_1() -> None:
    # An actual document from the web - this is likely created on a Windows client
    path = (
        "sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx"
    )
    docx: DocxContent = next(read_docx(_read_file_to_file_like(path=path)))

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

    docx: DocxContent = next(read_docx(_read_file_to_file_like(path=path)))
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


def test_read_docx__image_extraction_1() -> None:
    # Test for caption extraction from following paragraph with caption style
    path = "sharepoint2text/tests/resources/modern_ms/vorlage-abschlussarbeit.docx"
    docx: DocxContent = next(read_docx(_read_file_to_file_like(path=path)))

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
    docx: DocxContent = next(read_docx(_read_file_to_file_like(path=path)))

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
    docx: DocxContent = next(read_docx(file_like=_read_file_to_file_like(path=path)))

    units = list(docx.iterate_units())
    tc.assertEqual(8, len(units))

    tc.assertTrue(hasattr(units[0], "get_images"))
    tc.assertTrue(hasattr(units[0], "get_tables"))

    # first unit
    tc.assertEqual(["Sample Document"], units[0].get_metadata().location)
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


####################
# Legacy Microsoft #
####################


def test_read_xls_1() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/pb_2011_1_gen_web.xls"

    xls: XlsContent = next(read_xls(file_like=_read_file_to_file_like(path=path)))

    tc.assertEqual(13, len(xls.sheets))
    tc.assertEqual("2007-09-19T14:21:02", xls.metadata.created)
    tc.assertEqual("2011-06-01T13:54:08", xls.metadata.modified)
    tc.assertEqual("European Commission", xls.metadata.company)

    # iterator
    tc.assertEqual(0, len(list(xls.iterate_images())))
    tc.assertEqual(13, len(list(xls.iterate_tables())))

    xls_it = xls.iterate_units()
    # test first page
    s1 = next(xls_it).get_text()
    expected = (
        "EUROPEAN UNION\n"
        "                             European Commission\n"
        "  Directorate-General for Mobility and Transport\n"
    )
    tc.assertEqual(expected, s1[:113])

    # test second page
    s2 = next(xls_it).get_text()
    tc.assertIn(
        "The content of this pocketbook is based on a range of sources including Eurostat",
        s2,
    )

    # all text
    tc.assertIsNotNone(xls.get_full_text())

    #########
    # Units #
    #########
    units = list(xls.iterate_units())
    tc.assertEqual(
        XlsUnitMetadata(unit_number=1, sheet_name="Title"), units[0].get_metadata()
    )


def test_read_xls_2() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/mwe.xls"
    xls: XlsContent = next(read_xls(file_like=_read_file_to_file_like(path=path)))
    tc.assertEqual(
        "colA  colB\n   1     2",
        xls.get_full_text(),
    )

    tc.assertEqual(0, len(list(xls.iterate_images())))
    tc.assertEqual(1, len(list(xls.iterate_tables())))
    tc.assertEqual(TableDim(rows=2, columns=2), list(xls.iterate_tables())[0].get_dim())

    #########
    # Units #
    #########
    tc.assertEqual(1, len(list(xls.iterate_units())))
    units = list(xls.iterate_units())
    tc.assertListEqual(
        [["colA", "colB"], ["1", "2"]], units[0].get_tables()[0].get_table()
    )
    tc.assertEqual(TableDim(rows=2, columns=2), units[0].get_tables()[0])


def test_read_xls_3_images() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/xls_with_images.xls"
    xls: XlsContent = next(read_xls(file_like=_read_file_to_file_like(path=path)))

    tc.assertEqual(1, len(xls.images))
    tc.assertEqual(1, len(list(xls.iterate_images())))
    tc.assertEqual(1, xls.images[0].image_index)
    tc.assertEqual(183928, xls.images[0].size_bytes)
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=1,
            content_type="image/jpeg",
            width=800,
            height=450,
        ),
        xls.images[0].get_metadata(),
    )

    #########
    # Units #
    #########
    tc.assertEqual(3, len(list(xls.iterate_units())))
    units = list(xls.iterate_units())
    tc.assertEqual(183928, len(units[0].get_images()[0].get_bytes().getvalue()))
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=1,
            content_type="image/jpeg",
            width=800,
            height=450,
        ),
        units[0].get_images()[0].get_metadata(),
    )


def test_read_ppt() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/eurouni2.ppt"
    ppt: PptContent = next(read_ppt(_read_file_to_file_like(path=path)))

    tc.assertEqual(48, ppt.slide_count)
    tc.assertEqual(48, len(ppt.slides))
    # test first slide
    slide_1 = ppt.slides[0]
    tc.assertEqual("European Union", slide_1.title)
    tc.assertEqual(1, slide_1.slide_number)
    tc.assertListEqual(["Institutions and functions"], slide_1.body_text)
    tc.assertListEqual([], slide_1.other_text)
    tc.assertListEqual([], slide_1.notes)

    # test iterator
    tc.assertEqual(48, len(list(ppt.iterate_units())))
    tc.assertEqual(6, len(list(ppt.iterate_images())))
    tc.assertEqual(0, len(list(ppt.iterate_tables())))

    # test full text
    tc.assertEqual("European Union", ppt.get_full_text()[:14])


def test_read_ppt__presentation_with_notes() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/slide_with_notes.ppt"
    ppt: PptContent = next(
        read_ppt(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertListEqual(
        ["This is an example text in the notes section"], ppt.slides[0].notes
    )


def test_read_ppt__image_extraction() -> None:
    """Test image extraction from legacy PPT files."""
    path = "sharepoint2text/tests/resources/legacy_ms/ppt_with_images.ppt"
    ppt: PptContent = next(read_ppt(_read_file_to_file_like(path=path)))

    tc.assertEqual("", ppt.get_full_text())

    # Basic structure
    tc.assertEqual(2, ppt.slide_count)
    tc.assertEqual(2, len(ppt.slides))

    # Image extraction
    images = list(ppt.iterate_images())
    tc.assertEqual(2, len(images))

    # First image (PNG)
    img1: PptImage | ImageInterface = images[0]
    tc.assertEqual("image/png", img1.get_content_type())
    tc.assertEqual(1, img1.image_index)
    tc.assertEqual(1, img1.slide_number)
    tc.assertEqual(1718, img1.width)
    tc.assertEqual(348, img1.height)
    tc.assertEqual(83623, img1.size_bytes)

    # Verify PNG data starts with correct signature
    img1_bytes = img1.get_bytes()
    tc.assertEqual(b"\x89PNG\r\n\x1a\n", img1_bytes.read(8))

    # Second image (JPEG)
    img2: PptImage | ImageInterface = images[1]
    tc.assertEqual("image/jpeg", img2.get_content_type())
    tc.assertEqual(2, img2.image_index)
    tc.assertEqual(2, img2.slide_number)
    tc.assertEqual(800, img2.width)
    tc.assertEqual(450, img2.height)
    tc.assertEqual(183928, img2.size_bytes)

    # Verify JPEG data starts with correct signature
    img2_bytes = img2.get_bytes()
    tc.assertEqual(b"\xff\xd8\xff", img2_bytes.read(3))

    # Check ImageMetadata
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/png",
            width=1718,
            height=348,
        ),
        img1.get_metadata(),
    )

    #########
    # Units #
    #########
    tc.assertEqual(2, len(list(ppt.iterate_units())))
    tc.assertEqual("", list(ppt.iterate_units())[0].get_text())


def test_read_doc() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/Speech_Prime_Minister_of_The_Netherlands_EN.doc"
    doc: DocContent = next(read_doc(file_like=_read_file_to_file_like(path=path)))

    # Text content
    expected = """
    Welcome by the Prime Minister of the Kingdom of the Netherlands, Dr Jan Peter Balkenende, at the Inaugural Session of the International Criminal Court, The Hague, 11 March 2003 \n\n(Check against delivery)\n\nYour Royal Highnesses, Secretary-General, Your Excellencies, ladies and gentlemen,\n\nA very warm welcome to The Hague, the heart of Dutch democracy. The Netherlands is proud to be your host. \n\nAnd a special welcome to today’s eighteen most important people, who will shortly be sworn in as the first judges at the International Criminal Court. My sincere congratulations on your election.\n\nFour hundred and twenty years ago, the great legal thinker Hugo Grotius was born in Delft, less than ten kilometres from this spot. He was active in Dutch and European politics. \n\nFate did not smile on him. He fell victim to internal political conflicts, and was imprisoned in Loevestein castle. But he escaped by hiding in a chest of books. Dutch schoolchildren still love that story.\n\nGrotius fled to France, where he wrote the book that was to make him famous and which was translated into many languages: On the Law of War and Peace. \n\nIn it Grotius sets out his ideal: a system of international law, with clear agreements and procedures for countries to comply with. He believed that a system of this kind was necessary for international justice and stability.\n\nToday, ladies and gentlemen, nearly four centuries later, we move a step closer to that ideal. The International Criminal Court adds a crucial new element to the international legal system. \n\nIt makes it possible to prosecute the most serious crimes (genocide, crimes against humanity and war crimes) if they are not prosecuted at national level.\n\nSo today, the eleventh of March two thousand and three, is a historic day. Today the international community shows that it is still committed to justice, despite the many bloody conflicts and treaty violations we have seen since the Second World War.\n\nSuspicion and pessimism often dominate international politics. But today we are showing the world that there are also grounds for joy, optimism and hope.\n\nOf course, there is still a long way to go. We know that some countries are reluctant to sign up. The International Criminal Court is like a young swan. It needs time to grow bigger and stronger, then it can spread its wings and everyone will see it fly. Our work is not yet done. But with all of our help the ICC will succeed.\n\nMany people have been looking forward to this day. Many people have worked hard to bring it about. In particular, President Arthur Robinson of Trinidad and Tobago, who put the ICC onto the United Nations’ agenda in the late nineteen-eighties. And the UN Secretary-General, Kofi Annan, who did so much to speed up its establishment.\n\nI would also mention the judges and staff of other international courts, especially the Yugoslavia and Rwanda tribunals. Their experience has been and will be most valuable to the ICC.\n\nAnd finally I would mention the non-governmental organisations that have given their backing. Without your enthusiasm and support, it would all have taken far longer.\n\nThe Netherlands, and The Hague in particular, is honoured to be the ICC’s host. Since the first international peace conference was held here, over a century ago, The Hague has developed into the judicial capital of the world. We are proud of that.\n\nBut today, all of us can be proud.\n\nHugo Grotius’s last words were: “I have attempted much but achieved nothing”. \n\nToday we can say we have achieved something Grotius could only dream of: an international criminal court as part of an international legal order. And that takes us a big step closer to international justice.\n\nIt now gives me great pleasure to give the floor to the President of the Assembly of States Parties, His Royal Highness Prince Zeid Ra’ad Zeid al-Hussein.\n\nThank you.
    """
    tc.assertEqual(expected.strip(), doc.main_text)

    # Metadata
    tc.assertEqual(
        "Short dinner speech by the Prime Minister of the Kingdom of the Netherlands, Dr Jan Peter Balkenende",
        doc.metadata.title,
    )
    tc.assertEqual("Toby Screech", doc.metadata.author)
    tc.assertEqual("", doc.metadata.keywords)
    tc.assertEqual(580, doc.metadata.num_words)
    tc.assertEqual("2003-03-13T09:03:00", doc.metadata.create_time)
    tc.assertEqual("2003-03-13T09:03:00", doc.metadata.last_saved_time)

    # test iterator
    tc.assertEqual(1, len(list(doc.iterate_units())))
    tc.assertEqual(0, len(list(doc.iterate_images())))
    tc.assertEqual(0, len(list(doc.iterate_tables())))

    # test full text
    tc.assertEqual(
        "Short dinner speech by the Prime Minister of the Kingdom of the Netherlands, Dr Jan Peter Balkenende"
        + "\n"
        + "Welcome by the Prime Minister of the Kingdom",
        doc.get_full_text()[:145],
    )


def test_read_doc__image_extraction_1() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/legacy_doc_image.doc"

    doc: DocContent = next(read_doc(file_like=_read_file_to_file_like(path=path)))

    images: list[DocImage | ImageInterface] = list(doc.iterate_images())
    tc.assertEqual(1, len(images))
    tc.assertEqual(0, len(list(doc.iterate_tables())))

    tc.assertEqual("image/bmp", images[0].get_content_type())
    tc.assertEqual("Illustration 1: A GitHub screenshot", images[0].get_caption())
    tc.assertEqual(1304, images[0].width)
    tc.assertEqual(660, images[0].height)
    tc.assertEqual(
        ImageMetadata(
            unit_number=None,
            image_number=1,
            content_type="image/bmp",
            width=1304,
            height=660,
        ),
        images[0].get_metadata(),
    )


def test_read_doc__image_extraction_2() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/legacy_doc_multi_image.doc"
    doc: DocContent = next(read_doc(file_like=_read_file_to_file_like(path=path)))
    images: list[DocImage | ImageInterface] = list(doc.iterate_images())
    tc.assertEqual(2, len(images))
    tc.assertEqual(0, len(list(doc.iterate_tables())))

    # image 1
    tc.assertEqual("image/bmp", images[0].get_content_type())
    tc.assertEqual("Drawing 1: Second image", images[0].get_caption())
    tc.assertEqual(1038, images[0].width)
    tc.assertEqual(144, images[0].height)

    # image 2
    tc.assertEqual("image/bmp", images[1].get_content_type())
    tc.assertEqual("", images[1].get_caption())
    tc.assertEqual(1716, images[1].width)
    tc.assertEqual(336, images[1].height)


def test_read_doc__heading_units() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/headings.doc"
    doc: DocContent = next(
        read_doc(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(1, len(list(doc.iterate_tables())))
    tc.assertEqual(1, len(list(doc.iterate_images())))

    # unit extraction
    units = list(doc.iterate_units())
    tc.assertEqual(5, len(units))

    # 1
    tc.assertListEqual(["Intro"], units[0].get_metadata().location)
    tc.assertEqual("This is the intro text.", units[0].get_text())

    # 2
    tc.assertListEqual(["Chapter 1"], units[1].get_metadata().location)
    tc.assertEqual("Welcome to chapter 1", units[1].get_text())

    # 3
    tc.assertListEqual(
        ["Chapter 1", "Subsection in Chapter 1"], units[2].get_metadata().location
    )
    tc.assertEqual("This is a subsection in chapter 1", units[2].get_text())
    tc.assertListEqual(
        [["A", "B", "C", "D"], ["1", "2", "3", "4"]],
        units[3].get_tables()[0].get_table(),
    )

    # 4
    tc.assertListEqual(["Chapter 2"], units[3].get_metadata().location)
    tc.assertEqual("Welcome to chapter 2", units[3].get_text())
    tc.assertEqual(1, len(list(units[3].get_images())))
    tc.assertEqual(62421, len(list(units[3].get_images())[0].get_bytes().getvalue()))
    tc.assertEqual(
        ImageMetadata(
            unit_number=4,
            image_number=1,
            content_type="image/png",
            width=948,
            height=400,
        ),
        list(units[3].get_images())[0].get_metadata(),
    )

    # 5
    tc.assertListEqual(
        ["Chapter 2", "Subsection in Chapter 2"], units[4].get_metadata().location
    )
    tc.assertEqual("This is a subsection in chapter 2", units[4].get_text())


def test_read_rtf() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/2025.144.un.rtf"
    rtf_gen: typing.Generator[RtfContent] = read_rtf(
        file_like=_read_file_to_file_like(path=path)
    )

    rtfs = list(rtf_gen)
    tc.assertEqual(1, len(rtfs))

    rtf = rtfs[0]
    full_text = rtf.get_full_text()
    tc.assertEqual("c1\nSouth Australia", full_text[:18])
    tc.assertEqual("\non 18 December 2025\nNo 144 of 2025", full_text[-35:])

    tc.assertEqual(0, len(list(rtf.iterate_images())))
    tc.assertEqual(0, len(list(rtf.iterate_tables())))

    units = list(rtf.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertEqual("c1\n\nSouth Australia", units[0].get_text()[:19])
    tc.assertEqual(
        RtfUnitMetadata(unit_number=1, page_number=1), units[0].get_metadata()
    )


def test_read_rtf_tables_1() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/CULT-OJ-2024-10-03-1_DE.rtf"
    rtf_gen: typing.Generator[RtfContent] = read_rtf(
        file_like=_read_file_to_file_like(path=path), path=path
    )

    rtfs = list(rtf_gen)
    tc.assertEqual(1, len(rtfs))
    tc.assertEqual("Europäisches Parlament\n2024-2029", rtfs[0].get_full_text()[:32])
    tables = list(rtfs[0].iterate_tables())
    tc.assertEqual(2, len(tables))

    tc.assertEqual(TableDim(rows=2, columns=2), tables[0].get_dim())
    tc.assertEqual(TableDim(rows=4, columns=4), tables[1].get_dim())

    tc.assertListEqual(
        [["Europäisches Parlament\n2024-2029", ""], ["", ""]],
        list(rtfs[0].iterate_units())[0].get_tables()[0].get_table(),
    )
    tc.assertListEqual(
        [
            ["Verfasserin der Stellungnahme:", "", "", ""],
            ["", "Nela Riehl (Verts/ALE)", "", ""],
            ["Federführend:", "", "", ""],
            [
                "",
                "BUDG",
                "Victor Negrescu (S&D)\nNiclas Herbst (PPE)",
                "DT\xa0–\xa0PE763.050v01-00",
            ],
        ],
        list(rtfs[0].iterate_units())[0].get_tables()[1].get_table(),
    )


def test_read_rtf_tables_2() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/02_dept_transport.rtf"
    rtf_gen: typing.Generator[RtfContent] = read_rtf(
        file_like=_read_file_to_file_like(path=path)
    )

    rtfs = list(rtf_gen)
    tc.assertEqual(1, len(rtfs))

    tables = list(rtfs[0].iterate_tables())
    tc.assertEqual(23, len(tables))


#################
# Email formats #
#################
def test_email__eml_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.eml"
    mail_gen: typing.Generator[EmailContent, None, None] = read_eml_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    tc.assertEqual(1, len(mails))

    mail = mails[0]

    # from
    tc.assertEqual("Mikel Lindsaar", mail.from_email.name)
    tc.assertEqual("test@lindsaar.net", mail.from_email.address)
    # to
    tc.assertEqual(1, len(mail.to_emails))
    tc.assertEqual("Mikel Lindsaar", mail.to_emails[0].name)
    tc.assertEqual("raasdnil@gmail.com", mail.to_emails[0].address)

    # to-cc
    tc.assertEqual(2, len(mail.to_cc))
    tc.assertEqual("Jane Doe", mail.to_cc[0].name)
    tc.assertEqual("jane.doe@example.test", mail.to_cc[0].address)
    tc.assertEqual("Bob Smith", mail.to_cc[1].name)
    tc.assertEqual("bob.smith@example.test", mail.to_cc[1].address)

    # to-bcc
    tc.assertEqual(2, len(mail.to_bcc))
    tc.assertEqual("Hidden Tester", mail.to_bcc[0].name)
    tc.assertEqual("hidden.tester@example.test", mail.to_bcc[0].address)
    tc.assertEqual("Silent Observer", mail.to_bcc[1].name)
    tc.assertEqual("silent.observer@example.test", mail.to_bcc[1].address)

    # body
    tc.assertEqual("Plain email.\n\nHope it works well!\n\nMikel", mail.body_plain)

    # subject
    tc.assertEqual("Testing 123", mail.subject)

    # interface methods
    tc.assertEqual("Plain email.\n\nHope it works well!\n\nMikel", mail.get_full_text())
    tc.assertEqual(
        "Plain email.\n\nHope it works well!\n\nMikel",
        list(mail.iterate_units())[0].get_text(),
    )
    tc.assertEqual(0, len(list(mail.iterate_images())))
    tc.assertEqual(0, len(list(mail.iterate_tables())))

    # metadata
    mail_meta = mail.get_metadata()
    tc.assertEqual("basic_email.eml", mail_meta.filename)
    tc.assertEqual(".eml", mail_meta.file_extension)
    tc.assertEqual("2008-11-22T04:04:59+00:00", mail_meta.date)
    tc.assertEqual(
        "<6B7EC235-5B17-4CA8-B2B8-39290DEB43A3@test.lindsaar.net>", mail_meta.message_id
    )

    #########
    # Units #
    #########
    units = list(mail.iterate_units())
    tc.assertTrue(isinstance(units[0].get_metadata(), EmailUnitMetadata))
    tc.assertEqual(
        EmailUnitMetadata(unit_number=1, body_type="plain"), units[0].get_metadata()
    )


def test_email__msg_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.msg"
    mail_gen: typing.Generator[EmailContent, None, None] = read_msg_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    tc.assertEqual(1, len(mails))

    mail = mails[0]

    # from
    tc.assertEqual("Brian Zhou", mail.from_email.name)
    tc.assertEqual("brizhou@gmail.com", mail.from_email.address)
    # to
    tc.assertEqual(1, len(mail.to_emails))
    tc.assertEqual("", mail.to_emails[0].name)
    tc.assertEqual("brianzhou@me.com", mail.to_emails[0].address)

    # cc
    tc.assertEqual(1, len(mail.to_cc))
    tc.assertEqual("Brian Zhou", mail.to_cc[0].name)
    tc.assertEqual("brizhou@gmail.com", mail.to_cc[0].address)

    # bcc
    tc.assertEqual(0, len(mail.to_bcc))
    tc.assertListEqual([], mail.to_bcc)

    # subject
    tc.assertEqual("Test for TIF files", mail.subject)
    # body
    tc.assertEqual("This is a test email to experiment with", mail.body_plain[:39])

    # metadata
    mail_meta = mail.get_metadata()
    tc.assertEqual("basic_email.msg", mail_meta.filename)
    tc.assertEqual(".msg", mail_meta.file_extension)
    tc.assertEqual("2013-11-18T10:26:24+02:00", mail_meta.date)
    tc.assertEqual(
        "<CADtJ4eNjQSkGcBtVteCiTF+YFG89+AcHxK3QZ=-Mt48xygkvdQ@mail.gmail.com>",
        mail_meta.message_id,
    )

    tc.assertEqual(0, len(list(mail.iterate_images())))
    tc.assertEqual(0, len(list(mail.iterate_tables())))


def test_email__msg_format_with_attachment() -> None:
    path = "sharepoint2text/tests/resources/mails/msg_with_attachment.msg"
    mail_gen: typing.Generator[EmailContent, None, None] = read_msg_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    tc.assertEqual(1, len(mails))

    mail = mails[0]

    # from
    tc.assertIsNotNone(mail.from_email.name)
    tc.assertIsNotNone(mail.from_email.address)
    # to
    tc.assertEqual(1, len(mail.to_emails))
    tc.assertIsNotNone(mail.to_emails[0].name)
    tc.assertEqual("", mail.to_emails[0].address)

    # cc
    tc.assertEqual(0, len(mail.to_cc))
    tc.assertListEqual([], mail.to_cc)

    # bcc
    tc.assertEqual(0, len(mail.to_bcc))
    tc.assertListEqual([], mail.to_bcc)

    # subject
    tc.assertEqual("Test .msg with attachment", mail.subject)
    # body
    tc.assertEqual("", mail.body_plain)
    tc.assertEqual("<html><head>", mail.body_html[:12])

    # metadata
    mail_meta = mail.get_metadata()
    tc.assertEqual("msg_with_attachment.msg", mail_meta.filename)
    tc.assertEqual(".msg", mail_meta.file_extension)
    tc.assertEqual("2025-12-31T12:32:42+00:00", mail_meta.date)
    tc.assertEqual(
        "<VE1PR10MB3790E964D9B988D177790593FABDA@VE1PR10MB3790.EURPRD10.PROD.OUTLOOK.COM>",
        mail_meta.message_id,
    )

    tc.assertEqual(2, len(mail.attachments))
    attachments_by_name = {att.filename: att for att in mail.attachments}
    tc.assertIn("sample.pdf", attachments_by_name)
    tc.assertIn("pptx_formula_image.pptx", attachments_by_name)

    pdf_attachment = attachments_by_name["sample.pdf"]
    tc.assertEqual("application/pdf", pdf_attachment.mime_type)
    tc.assertIsInstance(pdf_attachment.data, io.BytesIO)
    tc.assertEqual(0, pdf_attachment.data.tell())
    tc.assertEqual(249095, len(pdf_attachment.data.getvalue()))
    tc.assertTrue(pdf_attachment.is_supported_mime_type)

    attachments = list(mail.iterate_supported_attachments())
    tc.assertEqual(2, len(attachments))
    tc.assertIsInstance(attachments[0], PdfContent)
    tc.assertIsInstance(attachments[1], PptxContent)
    tc.assertEqual(
        "This is a test sentence\n"
        "This is a table\n"
        "C1 C2\n"
        "R1 V1\n"
        "R2 V2\n"
        "This is page 2\n"
        "An image of the Google landing page",
        attachments[0].get_full_text(),
    )
    tc.assertEqual(1, len(list(attachments[0].iterate_images())))
    tc.assertEqual(
        "The slide title\nThe first text line\n\n\n\n\nThe last text line\nA beach\n$$f(x)=\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}e^{-\\frac{(x-\\mu)^{2}}{2\\sigma^{2}}}$$",
        attachments[1].get_full_text(),
    )
    tc.assertEqual(1, len(list(attachments[1].iterate_images())))

    pptx_attachment = attachments_by_name["pptx_formula_image.pptx"]
    tc.assertEqual(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        pptx_attachment.mime_type,
    )
    tc.assertIsInstance(pptx_attachment.data, io.BytesIO)
    tc.assertEqual(0, pptx_attachment.data.tell())
    tc.assertEqual(1566612, len(pptx_attachment.data.getvalue()))
    tc.assertTrue(pptx_attachment.is_supported_mime_type)

    tc.assertEqual(0, len(list(mail.iterate_images())))
    tc.assertEqual(0, len(list(mail.iterate_tables())))


def test_email__eml_format_with_attachment() -> None:
    path = "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
    mail_gen: typing.Generator[EmailContent, None, None] = read_eml_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    tc.assertEqual(1, len(mails))

    mail = mails[0]

    # from
    tc.assertIsNotNone(mail.from_email.name)
    tc.assertIsNotNone(mail.from_email.address)
    # to
    tc.assertEqual(1, len(mail.to_emails))
    tc.assertIsNotNone(mail.to_emails[0].name)
    tc.assertIsNotNone(mail.to_emails[0].address)

    # cc
    tc.assertEqual(0, len(mail.to_cc))
    tc.assertListEqual([], mail.to_cc)

    # bcc
    tc.assertEqual(0, len(mail.to_bcc))
    tc.assertListEqual([], mail.to_bcc)

    # subject
    tc.assertEqual("Test .msg with attachment", mail.subject)
    # body
    tc.assertEqual("<html><head>", mail.body_html[:12])

    # metadata
    mail_meta = mail.get_metadata()
    tc.assertEqual("msg_with_attachment.eml", mail_meta.filename)
    tc.assertEqual(".eml", mail_meta.file_extension)
    tc.assertEqual("2025-12-31T12:32:42+00:00", mail_meta.date)
    tc.assertEqual(
        "<VE1PR10MB3790E964D9B988D177790593FABDA@VE1PR10MB3790.EURPRD10.PROD.OUTLOOK.COM>",
        mail_meta.message_id,
    )

    tc.assertEqual(2, len(mail.attachments))
    attachments_by_name = {att.filename: att for att in mail.attachments}
    tc.assertIn("sample.pdf", attachments_by_name)
    tc.assertIn("pptx_formula_image.pptx", attachments_by_name)

    pdf_attachment = attachments_by_name["sample.pdf"]
    tc.assertEqual("application/pdf", pdf_attachment.mime_type)
    tc.assertIsInstance(pdf_attachment.data, io.BytesIO)
    tc.assertEqual(0, pdf_attachment.data.tell())
    tc.assertEqual(249095, len(pdf_attachment.data.getvalue()))
    tc.assertTrue(pdf_attachment.is_supported_mime_type)

    attachments = list(mail.iterate_supported_attachments())
    tc.assertEqual(2, len(attachments))
    tc.assertIsInstance(attachments[0], PdfContent)
    tc.assertIsInstance(attachments[1], PptxContent)

    pptx_attachment = attachments_by_name["pptx_formula_image.pptx"]
    tc.assertEqual(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        pptx_attachment.mime_type,
    )
    tc.assertIsInstance(pptx_attachment.data, io.BytesIO)
    tc.assertEqual(0, pptx_attachment.data.tell())
    tc.assertEqual(1566612, len(pptx_attachment.data.getvalue()))
    tc.assertTrue(pptx_attachment.is_supported_mime_type)

    tc.assertEqual(0, len(list(mail.iterate_images())))
    tc.assertEqual(0, len(list(mail.iterate_tables())))


def test_password_protected__doc() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/doc-password-protected-pw123.doc"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_doc(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__odt() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/odt-password-protected-pw123.odt"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_odt(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__pdf() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/pdf-password-protected-pw123.pdf"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_pdf(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__ods() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/ods-password-protected-pw123.ods"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_ods(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__xls() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/xls-password-protected-pw123.xls"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_xls(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__odp() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/odp-password-protected-pw123.odp"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_odp(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__docx() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/docx-password-protected-pw123.docx"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_docx(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__xlsx() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/xslx-password-protected-pw123.xlsx"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_xlsx(file_like=_read_file_to_file_like(path=path), path=path))


def test_password_protected__pptx() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/pptx-password-protected-pw123.pptx"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_pptx(file_like=_read_file_to_file_like(path=path), path=path))


def test_email__mbox_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.mbox"

    mail_gen: typing.Generator[EmailContent, None, None] = read_mbox_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    # number of mails
    tc.assertEqual(2, len(mails))

    # 1st mail
    # subject
    tc.assertEqual("Test Email 1", mails[0].subject)
    # body
    tc.assertEqual("This is the body", mails[0].body_plain[:16])
    # sender
    tc.assertEqual("John Doe", mails[0].from_email.name)
    tc.assertEqual("john@example.com", mails[0].from_email.address)

    # receiver
    tc.assertEqual(1, len(mails[0].to_emails))
    tc.assertEqual("Jane Smith", mails[0].to_emails[0].name)
    tc.assertEqual("jane@example.com", mails[0].to_emails[0].address)

    # cc
    tc.assertEqual(0, len(mails[0].to_cc))

    # bcc
    tc.assertEqual(0, len(mails[0].to_bcc))

    # metadata
    mail_meta = mails[0].get_metadata()
    tc.assertEqual("basic_email.mbox", mail_meta.filename)
    tc.assertEqual(".mbox", mail_meta.file_extension)
    tc.assertEqual("2025-12-27T10:00:00+00:00", mail_meta.date)
    tc.assertEqual("<msg001@example.com>", mail_meta.message_id)

    tc.assertEqual(0, len(list(mails[0].iterate_images())))
    tc.assertEqual(0, len(list(mails[0].iterate_tables())))


###############
# Open Office #
###############


def test_read_open_office__document() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_document.odt"
    odt: OdtContent = next(
        read_odt(file_like=_read_file_to_file_like(path=path), path=path)
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


def test_read_open_office__presentation_with_notes() -> None:
    path = "sharepoint2text/tests/resources/open_office/slide_with_notes.odp"
    odp: OdpContent = next(
        read_odp(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertListEqual(
        ["This is an example text in the notes section"], odp.slides[0].notes
    )


def test_read_open_office__presentation_with_table() -> None:
    path = "sharepoint2text/tests/resources/open_office/odp_with_table.odp"
    odp: OdpContent = next(
        read_odp(file_like=_read_file_to_file_like(path=path), path=path)
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


def test_read_open_office__heading_units() -> None:
    path = "sharepoint2text/tests/resources/open_office/headings.odt"
    odt: OdtContent = next(
        read_odt(file_like=_read_file_to_file_like(path=path), path=path)
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
        read_odp(file_like=_read_file_to_file_like(path=path), path=path)
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
        read_ods(file_like=_read_file_to_file_like(path=path), path=path)
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
    ods: OdsContent = next(read_ods(file_like=_read_file_to_file_like(path=path)))

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
        read_odt(file_like=_read_file_to_file_like(path=path), path=path)
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


def test_open_office__presentation_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odp"
    odp: OdpContent = next(
        read_odp(file_like=_read_file_to_file_like(path=path), path=path)
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


def test_open_office__spreadsheet_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.ods"
    ods: OdsContent = next(
        read_ods(file_like=_read_file_to_file_like(path=path), path=path)
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


#########
# Other #
#########


def test_read_pdf_1() -> None:
    path = "sharepoint2text/tests/resources/pdf/sample.pdf"
    pdf: PdfContent = next(read_pdf(file_like=_read_file_to_file_like(path=path)))

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
    tables = list(pdf.iterate_tables())
    tc.assertEqual(1, len(tables))
    tc.assertListEqual(
        [["C1", "C2"], ["R1", "V1"], ["R2", "V2"]], tables[0].get_table()
    )
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
    tc.assertListEqual(
        [["C1", "C2"], ["R1", "V1"], ["R2", "V2"]], units[0].get_tables()[0].get_table()
    )
    tc.assertEqual(PdfUnitMetadata(unit_number=1), units[0].get_metadata())


def test_read_pdf_2() -> None:
    path = "sharepoint2text/tests/resources/pdf/multi_image.pdf"
    pdf: PdfContent = next(
        read_pdf(file_like=_read_file_to_file_like(path=path), path=path)
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


def test_read_pdf_3() -> None:
    path = "sharepoint2text/tests/resources/pdf/large_table_1.pdf"
    pdf: PdfContent = next(
        read_pdf(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(1, len(list(pdf.iterate_tables())))
    first_table = list(pdf.iterate_tables())[0].get_table()
    # begin of table
    tc.assertEqual(
        ["€ million", "2018", "2019", "2020", "2021", "2022"], first_table[0]
    )
    tc.assertEqual(["Bayer Group financial KPIs", "", "", "", "", ""], first_table[1])
    tc.assertEqual(
        ["Sales", "36,742", "43,545", "41,400", "44,081", "50,739"], first_table[2]
    )
    tc.assertEqual(
        ["EBITDA", "9,695", "9,529", "(2,910)", "6,409", "13,515"], first_table[3]
    )
    # end of table
    tc.assertEqual(
        ["Water use (million m3)", "42", "59", "57", "55", "53"], first_table[-9]
    )
    tc.assertEqual(
        [
            "2021 figures restated; figures for 2018 - 2020 as last reported",
            "",
            "",
            "",
            "",
            "",
        ],
        first_table[-8],
    )
    tc.assertEqual(
        [
            "1 For definition see A 2.3 “Alternative Performance Measures Used by the Bayer Group.”",
            "",
            "",
            "",
            "",
            "",
        ],
        first_table[-7],
    )
    tc.assertEqual(
        ["2 For more in formation see A 1.2.1.", "", "", "", "", ""], first_table[-6]
    )
    tc.assertEqual(["3 Economically or medically", "", "", "", "", ""], first_table[-5])
    tc.assertEqual(
        [
            "4 The in crease in R&D expenses in 2020 was mainly due to special charges in connection with impairment charges at Crop Science.",
            "",
            "",
            "",
            "",
            "",
        ],
        first_table[-4],
    )
    tc.assertEqual(
        ["5 R&D expenses before special items", "", "", "", "", ""], first_table[-3]
    )
    tc.assertEqual(
        ["6 Employees calculated as full-time equivalents (FTEs)", "", "", "", "", ""],
        first_table[-2],
    )
    tc.assertEqual(
        [
            "7 Quotient of total energy consumption and external sales",
            "",
            "",
            "",
            "",
            "",
        ],
        first_table[-1],
    )
    tc.assertTrue(all(len(row) == 6 for row in first_table))
    tc.assertEqual(
        TableDim(rows=52, columns=6), list(pdf.iterate_tables())[0].get_dim()
    )


def test_read_pdf_4() -> None:
    path = "sharepoint2text/tests/resources/pdf/multi_table.pdf"
    pdf: PdfContent = next(
        read_pdf(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tables = list(pdf.iterate_tables())
    tc.assertEqual(2, len(tables))

    # tabel 1
    tc.assertEqual(TableDim(rows=6, columns=3), list(pdf.iterate_tables())[0].get_dim())
    tc.assertListEqual(
        [
            ["€ million", "2022", "2021"],
            ["Financial statement audit services", "26", "22"],
            ["Other assurance services", "6", "3"],
            ["Tax advisory services", "0", "7"],
            ["Other services", "2", "4"],
            ["", "34", "36"],
        ],
        list(pdf.iterate_tables())[0].get_table(),
    )

    # tabel 2
    tc.assertEqual(TableDim(rows=4, columns=3), list(pdf.iterate_tables())[1].get_dim())
    tc.assertListEqual(
        [
            ["€ million", "2022", "2021"],
            ["Wages and salaries", "37,529", "34,644"],
            [
                "Social security, post-employment and other employee benefit costs",
                "9,473",
                "9,033",
            ],
            ["", "47,002", "43,677"],
        ],
        list(pdf.iterate_tables())[1].get_table(),
    )


def test_read_pdf_5() -> None:
    path = "sharepoint2text/tests/resources/pdf/two_tables_horizontal.pdf"
    pdf: PdfContent = next(
        read_pdf(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(2, len(list(pdf.iterate_tables())))
    first_table = list(pdf.iterate_tables())[0].get_table()
    # beginning of table
    tc.assertListEqual(["Assets", "31/12/2024", "31/12/2023"], first_table[0])
    tc.assertListEqual(["in €m", "", ""], first_table[1])
    tc.assertListEqual(["Property, plant and equipment", "3.4", "3.9"], first_table[2])
    tc.assertListEqual(["Investments in associates", "0.5", "34.4"], first_table[3])
    # end of table
    tc.assertListEqual(["Current assets", "168.6", "152.8"], first_table[-2])
    tc.assertListEqual(["Overall total", "6,723.6", "8,882.3"], first_table[-1])
    tc.assertEqual(
        TableDim(rows=17, columns=3), list(pdf.iterate_tables())[0].get_dim()
    )

    second_table = list(pdf.iterate_tables())[1].get_table()
    # beginning of table
    tc.assertListEqual(
        ["Equity and Liabilities", "31/12/2024", "31/12/2023"], second_table[0]
    )
    tc.assertListEqual(["in €m", "", ""], second_table[1])
    tc.assertListEqual(["Share capital", "24.9", "24.9"], second_table[2])
    # end of table
    tc.assertListEqual(["Deferred tax —Liabilities", "46.4", "116.7"], second_table[-4])
    tc.assertListEqual(["Provisions", "0.2", "0.3"], second_table[-3])
    tc.assertListEqual(
        ["Current & Non-current liabilities", "1,504.1", "1,934.3"], second_table[-2]
    )
    tc.assertListEqual(["Overall total", "6,723.6", "8,882.3"], second_table[-1])

    tc.assertEqual(
        TableDim(rows=12, columns=3), list(pdf.iterate_tables())[1].get_dim()
    )


def test_read_pdf_6() -> None:
    path = (
        "sharepoint2text/tests/resources/pdf/vendor-creation-form-english-version.pdf"
    )
    pdf: PdfContent = next(
        read_pdf(file_like=_read_file_to_file_like(path=path), path=path)
    )

    full_text = pdf.get_full_text()
    tc.assertTrue(len(full_text) > 0)
    tc.assertIn("Supplier Registration Form", full_text)


def test_read_pdf_7() -> None:
    path = "sharepoint2text/tests/resources/pdf/wirecard-annual-report-2018-page190.pdf"
    pdf: PdfContent = next(
        read_pdf(file_like=_read_file_to_file_like(path=path), path=path)
    )

    units = list(pdf.iterate_units())
    tc.assertEqual(
        (
            " \n"
            " \n"
            "Changes in liabilities from financing activities  \n"
            "in EUR million   non cash-relevant  \n"
            "  \n"
            "1 Jan 2017 cash-relevant Additions\n"
            "Changes in \n"
            "measurement Other 31 Dec 2017\n"
            "Purchase price liabilities 77.2 –65.2 36.7 11.2 –0.5 59.4\n"
            "Lease liabilities 19.5 –11.7 18.3 0.0 0.0 26.1\n"
            "interest-bearing liabilities 593.4 467.1 0.0 0.0 0.4 1,060.9\n"
            "Total 690.1 390.2 55.0 11.2 –0.1 1,146.4\n"
            " \n"
            " \n"
            "6.4 Cash and cash equivalents at end of period \n"
            "After taking into account the total cash inflows and cash \n"
            "outflows reported above of EUR 821.3 million (previous \n"
            "year: EUR 563.3 million), exchange rate-related changes \n"
            "of EUR –14.6 million (previous year: EUR 1.1 million) as \n"
            "well as cash and cash equivalents at start of the reporting \n"
            "period of EUR 1.895,9 million (31 December 2017: \n"
            "EUR  1.331,5 million), cash and cash equivalents at the \n"
            "end of reporting period amounted to EUR 2,702.5 million \n"
            "(31 December 2017: EUR 1,895.9 million).  \n"
            "Alongside cash and cash equivalents, there are other cur-\n"
            "rent assets and liabilities that can have a significant effect \n"
            "on the availability of funds. Therefore, Wirecard has cre-\n"
            "ated an additional net cash calculation. The net cash cal-\n"
            "culation is based on the current availability of cash for the \n"
            "further development of the business and for investments.  \n"
            " \n"
            "6.5 Net cash items  \n"
            " \n"
            "Net Cash Position - Wirecard     \n"
            "in EUR million 31 Dec 2018 31 Dec 2017\n"
            "Cash and cash equivalents  2,719.8  1,901.3\n"
            "Interest-bearing securities and fixed-term deposits  2.3  1.8\n"
            "Receivables of the acquiring business and trade and \n"
            "other receivables  1,042.4  716.7\n"
            "Interest-bearing liabilities / other liabilities  –303.9  –463.1\n"
            "Customer deposits from banking operations –1,263.0 –1,098.7 –973.2 –819.4\n"
            "Non-current interest-bearing securities 24.7  44.6  \n"
            "Interest-bearing securities and fixed-term deposits 139.6  109.1  \n"
            "Liabilities of the acquiring business and trade payables  –715.3  –488.8\n"
            "Net Cash Position - Wirecard  1,646.6  848.6\n"
            " "
        ),
        units[0].get_text(),
    )

    tc.assertEqual(2, len(list(pdf.iterate_tables())))
    table_1 = list(pdf.iterate_tables())[0].get_table()
    tc.assertListEqual(
        [
            [
                "Changes in liabilities from financing activities",
                "",
                "",
                "",
                "",
                "",
                "",
            ],
            ["in EUR million", "", "", "", "non cash-relevant", "", ""],
            ["", "", "", "", "Changes in", "", ""],
            [
                "",
                "1 Jan 2017",
                "cash-relevant",
                "Additions measurement",
                "",
                "Other",
                "31 Dec 2017",
            ],
            [
                "Purchase price liabilities",
                "77.2",
                "-65.2",
                "36.7",
                "11.2",
                "-0.5",
                "59.4",
            ],
            ["Lease liabilities", "19.5", "-11.7", "18.3", "0.0", "0.0", "26.1"],
            [
                "interest-bearing liabilities",
                "593.4",
                "467.1",
                "0.0",
                "0.0",
                "0.4",
                "1,060.9",
            ],
            ["Total", "690.1", "390.2", "55.0", "11.2", "-0.1", "1,146.4"],
        ],
        table_1,
    )
    table_2 = list(pdf.iterate_tables())[1].get_table()
    tc.assertListEqual(
        [
            ["Net Cash Position - Wirecard", "", "", "", ""],
            ["in EUR million", "", "31 Dec 2018", "", "31 Dec 2017"],
            ["Cash and cash equivalents", "", "2,719.8", "", "1,901.3"],
            [
                "Interest-bearing securities and fixed-term deposits",
                "",
                "2.3",
                "",
                "1.8",
            ],
            ["Receivables of the acquiring business and trade and", "", "", "", ""],
            ["other receivables", "", "1,042.4", "", "716.7"],
            [
                "Interest-bearing liabilities / other liabilities",
                "",
                "-303.9",
                "",
                "-463.1",
            ],
            [
                "Customer deposits from banking operations",
                "-1,263.0",
                "-1,098.7",
                "-973.2",
                "-819.4",
            ],
            ["Non-current interest-bearing securities", "24.7", "", "44.6", ""],
            [
                "Interest-bearing securities and fixed-term deposits",
                "139.6",
                "",
                "109.1",
                "",
            ],
            [
                "Liabilities of the acquiring business and trade payables",
                "",
                "-715.3",
                "",
                "-488.8",
            ],
            ["Net Cash Position - Wirecard", "", "1,646.6", "", "848.6"],
        ],
        table_2,
    )


def test_read_html() -> None:
    path = "sharepoint2text/tests/resources/sample.html"
    html: HtmlContent = next(
        read_html(file_like=_read_file_to_file_like(path=path), path=path)
    )

    full_text = "Welcome on my website\n\n\nParticipants\n\n\nName  | Age\nAlice | 25\nBob   | 30\n\n\nThis is a simple example of an HTML page with a table and links.\n\n\nVisit:\nWikipedia |\nGoogle"
    tc.assertEqual(full_text, html.get_full_text())
    tc.assertListEqual([[["Name", "Age"], ["Alice", "25"], ["Bob", "30"]]], html.tables)
    tc.assertListEqual(
        [
            {"text": "Wikipedia", "href": "https://www.wikipedia.org"},
            {"text": "Google", "href": "https://www.google.com"},
        ],
        html.links,
    )
    tc.assertEqual(0, len(list(html.iterate_images())))
    tc.assertEqual(1, len(list(html.iterate_tables())))
    tc.assertListEqual(
        [["Name", "Age"], ["Alice", "25"], ["Bob", "30"]],
        list(html.iterate_tables())[0].get_table(),
    )

    tc.assertEqual(
        HtmlUnitMetadata(unit_number=1), list(html.iterate_units())[0].get_metadata()
    )


def test_read_epub__1() -> None:
    """Test EPUB extraction with a sample EPUB file."""
    path = "sharepoint2text/tests/resources/epub/sample.epub"
    epub: EpubContent = next(
        read_epub(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Check metadata
    tc.assertEqual("Test EPUB Book", epub.metadata.title)
    tc.assertEqual("Test Author", epub.metadata.creator)
    tc.assertEqual("en", epub.metadata.language)
    tc.assertEqual("Test Publisher", epub.metadata.publisher)
    tc.assertEqual("2024-01-15", epub.metadata.date)
    tc.assertEqual("A test EPUB file for sharepoint-to-text", epub.metadata.description)
    tc.assertEqual("Testing", epub.metadata.subject)
    tc.assertEqual("3.0", epub.metadata.epub_version)

    # Check chapters
    tc.assertEqual(2, len(epub.chapters))

    # Chapter 1
    chapter1 = epub.chapters[0]
    tc.assertEqual(1, chapter1.chapter_number)
    tc.assertIn("Chapter 1: Introduction", chapter1.title)
    tc.assertIn("Welcome to the test EPUB book", chapter1.text)
    tc.assertIn("sample text for extraction testing", chapter1.text)
    tc.assertIn("Section 1.1", chapter1.text)

    # Chapter 1 table
    tc.assertEqual(1, len(chapter1.tables))
    tc.assertListEqual(
        [["Name", "Value"], ["Item A", "100"], ["Item B", "200"]],
        chapter1.tables[0],
    )

    # Chapter 2
    chapter2 = epub.chapters[1]
    tc.assertEqual(2, chapter2.chapter_number)
    tc.assertIn("Chapter 2: Getting Started", chapter2.title)
    tc.assertIn("second chapter", chapter2.text)
    tc.assertIn("First item in the list", chapter2.text)

    # Test iterate_units
    units = list(epub.iterate_units())
    tc.assertEqual(2, len(units))
    tc.assertEqual(
        EpubUnitMetadata(
            unit_number=1, href="OEBPS/chapter1.xhtml", title=chapter1.title
        ),
        units[0].get_metadata(),
    )

    # Test get_full_text
    full_text = epub.get_full_text()
    tc.assertIn("Chapter 1: Introduction", full_text)
    tc.assertIn("Chapter 2: Getting Started", full_text)
    tc.assertIn("Welcome to the test EPUB book", full_text)

    # Test iterate_tables
    tables = list(epub.iterate_tables())
    tc.assertEqual(1, len(tables))
    tc.assertEqual(TableDim(rows=3, columns=2), tables[0].get_dim())

    # Test table of contents
    tc.assertEqual(2, len(epub.toc))
    tc.assertEqual("Chapter 1: Introduction", epub.toc[0]["title"])
    tc.assertEqual("Chapter 2: Getting Started", epub.toc[1]["title"])


def test_read_epub__2() -> None:
    """Test EPUB extraction with a sample EPUB file."""
    path = "sharepoint2text/tests/resources/epub/BJNR274910013.epub"
    epub: EpubContent = next(
        read_epub(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # general
    tc.assertEqual(31, len(epub.chapters))
    tc.assertEqual(3, len(list(epub.iterate_tables())))
    tc.assertEqual("Gesetz zur Förderung der elektronischen", epub.get_full_text()[:39])
    tc.assertListEqual(
        [
            ["", ""],
            [
                "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - "
                "EGovG)"
            ],
            [
                "E-Government-Gesetz vom 25. Juli 2013 (BGBl. I S. 2749), das zuletzt durch "
                "Artikel 11 des Gesetzes vom 2. Dezember 2025 (BGBl. 2025 I Nr. 301) "
                "geändert worden ist"
            ],
            [
                "Gesetze im Internet - ePub herausgegeben vom Bundesministerium der Justiz "
                "und für Verbraucherschutz"
            ],
            ["erzeugt am: 05.12.2025"],
        ],
        list(epub.iterate_tables())[0].get_table(),
    )

    # metadata
    tc.assertEqual("BJNR274910013.epub", epub.get_metadata().filename)
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung "
        "(E-Government-Gesetz - EGovG)",
        epub.get_metadata().title,
    )
    tc.assertEqual("2025-12-05", epub.get_metadata().date)

    # units
    tc.assertEqual(31, len(list(epub.iterate_units())))
    units = list(epub.iterate_units())
    # 0
    tc.assertEqual("", units[0].get_text())
    tc.assertEqual(1, len(list(units[0].get_tables())))
    tc.assertListEqual(
        [
            ["", ""],
            [
                "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - "
                "EGovG)"
            ],
            [
                "E-Government-Gesetz vom 25. Juli 2013 (BGBl. I S. 2749), das zuletzt durch "
                "Artikel 11 des Gesetzes vom 2. Dezember 2025 (BGBl. 2025 I Nr. 301) "
                "geändert worden ist"
            ],
            [
                "Gesetze im Internet - ePub herausgegeben vom Bundesministerium der Justiz "
                "und für Verbraucherschutz"
            ],
            ["erzeugt am: 05.12.2025"],
        ],
        units[0].get_tables()[0].get_table(),
    )
    tc.assertEqual(
        EpubUnitMetadata(
            unit_number=2,
            href="BJNR274910013.html",
            title="Gesetz zur Förderung der elektronischen Verwaltung "
            "(E-Government-Gesetz - EGovG)",
        ),
        units[1].get_metadata(),
    )
    # 1
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - EGovG)",
        units[1].get_text()[:80],
    )
    # 2
    tc.assertEqual("Inhaltsübersicht", units[2].get_text())
    # 3
    tc.assertEqual("§ 1\n\nGeltungsbereich\n\n(1)", units[3].get_text()[:25])
    # last page
    tc.assertEqual("§ 19\n\nÜbergangsvorschriften", units[-1].get_text()[:27])


def test_read_macro_enabled_docm() -> None:
    """Test .docm (macro-enabled Word) extraction - same structure as .docx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.docm"
    result: DocxContent = next(
        read_docx(file_like=_read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as DocxContent (same as .docx)
    tc.assertIsInstance(result, DocxContent)
    tc.assertTrue(len(result.get_full_text()) > 0)


def test_read_macro_enabled_xlsm() -> None:
    """Test .xlsm (macro-enabled Excel) extraction - same structure as .xlsx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.xlsm"
    result: XlsxContent = next(
        read_xlsx(file_like=_read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as XlsxContent (same as .xlsx)
    tc.assertIsInstance(result, XlsxContent)
    tc.assertTrue(len(result.sheets) > 0)


def test_read_macro_enabled_pptm() -> None:
    """Test .pptm (macro-enabled PowerPoint) extraction - same structure as .pptx."""
    path = "sharepoint2text/tests/resources/modern_ms/sample.pptm"
    result: PptxContent = next(
        read_pptx(file_like=_read_file_to_file_like(path=path), path=path)
    )
    # Verify it extracts as PptxContent (same as .pptx)
    tc.assertIsInstance(result, PptxContent)
    tc.assertTrue(len(result.slides) > 0)


def test_read_mhtml() -> None:
    """Test MHTML (web archive) extraction."""
    path = "sharepoint2text/tests/resources/sample.mhtml"
    result: HtmlContent = next(
        read_mhtml(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Verify it returns HtmlContent
    tc.assertIsInstance(result, HtmlContent)

    # Check metadata
    tc.assertEqual("Test MHTML Page", result.metadata.title)

    # Check content extraction
    tc.assertIn("Welcome to the Test Page", result.content)
    tc.assertIn("test MHTML document", result.content)
    tc.assertIn("More Content", result.content)

    # Check table extraction
    tc.assertEqual(1, len(result.tables))
    tc.assertListEqual(
        [["Product", "Price"], ["Widget", "$10.00"], ["Gadget", "$25.00"]],
        result.tables[0],
    )

    # Check link extraction
    tc.assertEqual(1, len(result.links))
    tc.assertEqual("link to example.com", result.links[0]["text"])
    tc.assertEqual("https://example.com", result.links[0]["href"])


############
# Archives #
############


def test_read_zip_archive_1() -> None:
    """Test ZIP archive extraction with multiple supported files."""
    path = "sharepoint2text/tests/resources/archives/test_archive.zip"
    results = list(
        read_archive(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Should extract 2 text files from the archive
    tc.assertEqual(2, len(results))

    # All results should be PlainTextContent
    for result in results:
        tc.assertIsInstance(result, PlainTextContent)

    # Check that we got the expected content
    texts = [r.get_full_text() for r in results]
    tc.assertTrue(any("This is a test document" in t for t in texts))
    tc.assertTrue(any("Another file in the archive" in t for t in texts))

    # Check that metadata includes archive path
    for result in results:
        tc.assertIn("test_archive.zip!/", result.get_metadata().file_path)


def test_read_zip_archive_2() -> None:
    """Test ZIP archive extraction with multiple supported files."""

    # three files - of which two are supported
    path = "sharepoint2text/tests/resources/archives/sample.zip"
    results = list(
        read_archive(file_like=_read_file_to_file_like(path=path), path=path)
    )
    tc.assertEqual(2, len(results))
    tc.assertTrue(isinstance(results[0], PlainTextContent))
    tc.assertTrue(isinstance(results[1], EpubContent))


def test_read_tar_archive() -> None:
    """Test TAR archive extraction."""
    path = "sharepoint2text/tests/resources/archives/test_archive.tar"
    results = list(
        read_archive(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Should extract 2 text files from the archive
    tc.assertEqual(2, len(results))

    # All results should be PlainTextContent
    for result in results:
        tc.assertIsInstance(result, PlainTextContent)

    # Check that we got the expected content
    texts = [r.get_full_text() for r in results]
    tc.assertTrue(any("This is a test document" in t for t in texts))
    tc.assertTrue(any("Another file in the tar archive" in t for t in texts))


def test_read_tar_gz_archive() -> None:
    """Test compressed TAR.GZ archive extraction."""
    path = "sharepoint2text/tests/resources/archives/test_archive.tar.gz"
    results = list(
        read_archive(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Should extract 1 text file from the archive
    tc.assertEqual(1, len(results))

    result = results[0]
    tc.assertIsInstance(result, PlainTextContent)
    tc.assertIn("This is a test document", result.get_full_text())


def test_archive_skips_nested_archives() -> None:
    """Test that nested archives are skipped to prevent zip bombs."""
    # Create a ZIP with a nested ZIP inside
    nested_content = b"nested content"
    inner_zip = std_io.BytesIO()
    with zipfile.ZipFile(inner_zip, "w") as zf:
        zf.writestr("inner.txt", nested_content)
    inner_zip.seek(0)

    outer_zip = std_io.BytesIO()
    with zipfile.ZipFile(outer_zip, "w") as zf:
        zf.writestr("outer.txt", b"outer content")
        zf.writestr("nested.zip", inner_zip.read())
    outer_zip.seek(0)

    results = list(read_archive(outer_zip, path="test.zip"))

    # Should only extract the outer.txt, not the nested.zip
    tc.assertEqual(1, len(results))
    tc.assertIn("outer content", results[0].get_full_text())


def test_archive_skips_hidden_files() -> None:
    """Test that hidden files (starting with .) are skipped."""

    zip_buffer = std_io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        zf.writestr("visible.txt", b"visible content")
        zf.writestr(".hidden.txt", b"hidden content")
        zf.writestr("__MACOSX/file.txt", b"macos resource fork")
    zip_buffer.seek(0)

    results = list(read_archive(zip_buffer, path="test.zip"))

    # Should only extract visible.txt
    tc.assertEqual(1, len(results))
    tc.assertIn("visible content", results[0].get_full_text())
