import io
import logging
import typing
from unittest import TestCase

from sharepoint2text.extractors.data_types import (
    DocContent,
    DocxComment,
    DocxContent,
    DocxFormula,
    DocxNote,
    EmailContent,
    FileMetadataInterface,
    HtmlContent,
    ImageMetadata,
    OdpAnnotation,
    OdpContent,
    OdsAnnotation,
    OdsContent,
    OdtAnnotation,
    OdtContent,
    OdtHeaderFooter,
    OdtNote,
    PdfContent,
    PlainTextContent,
    PptContent,
    PptxContent,
    RtfContent,
    XlsContent,
    XlsxContent,
)
from sharepoint2text.extractors.html_extractor import read_html
from sharepoint2text.extractors.mail.eml_email_extractor import read_eml_format_mail
from sharepoint2text.extractors.mail.mbox_email_extractor import read_mbox_format_mail
from sharepoint2text.extractors.mail.msg_email_extractor import read_msg_format_mail
from sharepoint2text.extractors.ms_legacy.doc_extractor import read_doc
from sharepoint2text.extractors.ms_legacy.ppt_extractor import read_ppt
from sharepoint2text.extractors.ms_legacy.rtf_extractor import read_rtf
from sharepoint2text.extractors.ms_legacy.xls_extractor import read_xls
from sharepoint2text.extractors.ms_modern.docx_extractor import read_docx
from sharepoint2text.extractors.ms_modern.pptx_extractor import read_pptx
from sharepoint2text.extractors.ms_modern.xlsx_extractor import read_xlsx
from sharepoint2text.extractors.open_office.odp_extractor import read_odp
from sharepoint2text.extractors.open_office.ods_extractor import read_ods
from sharepoint2text.extractors.open_office.odt_extractor import read_odt
from sharepoint2text.extractors.pdf_extractor import read_pdf
from sharepoint2text.extractors.plain_extractor import read_plain_text

logger = logging.getLogger(__name__)

tc = TestCase()

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
        },
        meta.to_dict(),
    )


#########
# Plain #
#########


def test_read_text() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.txt"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain: PlainTextContent = next(read_plain_text(file_like=file_like))

    tc.assertEqual("Hello World\n", plain.content)


def test_read_plain_csv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.csv"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain: PlainTextContent = next(read_plain_text(file_like=file_like, path=path))

    tc.assertEqual('Text; Date\n"Hello World"; "2025-12-25"\n', plain.content)

    tc.assertEqual(1, len(list(plain.iterator())))


def test_read_plain_tsv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.tsv"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain = next(read_plain_text(file_like=file_like, path=path))

    tc.assertEqual("Text\tDate\nHello World\t2025-12-25\n", plain.content)


def test_read_plain_markdown() -> None:
    path = "sharepoint2text/tests/resources/plain_text/document.md"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain = next(read_plain_text(file_like=file_like, path=path))

    tc.assertEqual("# Markdown file\n\nThis is a text", plain.get_full_text())


####################
# Modern Microsoft #
####################
def test_read_xlsx_1() -> None:
    filename = "sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    xlsx: XlsxContent = next(read_xlsx(file_like=file_like))

    tc.assertEqual("2006-09-16T00:00:00", xlsx.metadata.created)
    tc.assertEqual("2015-05-06T11:46:24", xlsx.metadata.modified)

    tc.assertEqual(3, len(xlsx.sheets))
    tc.assertListEqual(
        sorted(["Sheet1", "Sheet2", "Sheet3"]), sorted([s.name for s in xlsx.sheets])
    )

    # check raw data and table interface
    # check that the first row in the first sheet is the headline
    tc.assertListEqual(["AREA", "CODE", "COUNTRY NAME"], xlsx.sheets[0].data[0])
    tc.assertListEqual(["AREA", "CODE", "COUNTRY NAME"], xlsx.sheets[0].get_table()[0])
    tc.assertListEqual(
        ["European Union (EU)", "EU-28 ", "European Union (28 countries)"],
        xlsx.sheets[0].get_table()[1],
    )

    tc.assertEqual(3, len(list(xlsx.iterator())))

    tc.assertEqual("Sheet1\nAREA     CODE", xlsx.get_full_text()[:20])

    tc.assertDictEqual(
        {
            "filename": None,
            "file_extension": None,
            "file_path": None,
            "folder_path": None,
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


def test_read_xlsx_2() -> None:
    filename = "sharepoint2text/tests/resources/modern_ms/mwe.xlsx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    xlsx: XlsxContent = next(read_xlsx(file_like=file_like))
    tc.assertEqual(
        "Blatt 1\nTabelle 1 Unnamed: 1\n     ColA       ColB\n        1          2",
        xlsx.get_full_text(),
    )
    tc.assertListEqual([["ColA", "ColB"], [1, 2]], xlsx.sheets[0].data)


def test_read_xlsx_3() -> None:
    """Verifies the treatment of empty rows and columns in a sheet

    We want that the list of rows is easily processable with Pandas or Polars to create
    dataframes. This requires that None/Nulls are not accidentally pruned. The rows must have
    the same number of columns for this to work
    """
    filename = "sharepoint2text/tests/resources/modern_ms/empty_row_columns.xlsx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    xlsx: XlsxContent = next(read_xlsx(file_like=file_like))
    tc.assertListEqual(
        [
            [None, "Name", None, "Age"],
            [None, "A", None, 25],
            [None, None, None, None],
            [None, "B", None, 28],
        ],
        xlsx.sheets[0].data,
    )


def test_read_xlsx_4__image_extraction() -> None:
    filename = "sharepoint2text/tests/resources/modern_ms/image_in_excel.xlsx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    xlsx: XlsxContent = next(read_xlsx(file_like=file_like))
    tc.assertEqual("Image Sheet", xlsx.sheets[0].name)
    tc.assertEqual(1, len(xlsx.sheets[0].images))

    image = xlsx.sheets[0].images[0]
    tc.assertEqual(7280, len(image.get_bytes().getvalue()))
    tc.assertEqual("Image 1", image.get_caption())
    tc.assertEqual("Picture", image.get_description())
    tc.assertEqual(600, image.width)
    tc.assertEqual(300, image.height)


def test_read_pptx_1() -> None:
    filename = "sharepoint2text/tests/resources/modern_ms/eu-visibility_rules_00704232-AF9F-1A18-BD782C469454ADAD_68401.pptx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    pptx: PptxContent = next(read_pptx(file_like))

    # metadata
    tc.assertEqual("IVAN Anda-Otilia", pptx.metadata.author)
    tc.assertEqual("MAGLI Mia (JUST)", pptx.metadata.last_modified_by)
    tc.assertEqual("2011-10-28T10:25:18", pptx.metadata.created)
    tc.assertEqual("2020-07-12T09:25:35", pptx.metadata.modified)

    tc.assertEqual(3, len(pptx.slides))

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

    # iterator
    tc.assertEqual(3, len(list(pptx.iterator())))

    # full text
    expected = (
        "EU-funding visibility - art. 22 GA"
        + "\n"
        + "To be applied on all materials and communica"
    )
    tc.assertEqual(expected, pptx.get_full_text()[:79])


def test_read_pptx_2() -> None:
    filename = "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    pptx: PptxContent = next(read_pptx(file_like))

    # Test default get_full_text() - without formulas, comments, or image captions
    # Note: "A beach" is a regular textbox, not an image caption
    base_text = pptx.get_full_text()
    tc.assertEqual(
        "The slide title\nThe first text line\n\n\n\n\nThe last text line\nA beach",
        base_text,
    )

    # images
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
        ImageMetadata(unit_index=1, image_index=1, content_type="image/jpeg"),
        pptx.slides[0].images[0].get_metadata(),
    )

    # Test with formulas included
    text_with_formulas = pptx.get_full_text(include_formulas=True)
    tc.assertIn("$$f(x)=\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}", text_with_formulas)

    # Test with comments included
    text_with_comments = pptx.get_full_text(include_comments=True)
    tc.assertIn("[Comment: 0@2025-12-28T11:15:49.694: Not second?]", text_with_comments)

    # Test with all included (formulas + comments)
    full_text = pptx.get_full_text(include_formulas=True, include_comments=True)
    tc.assertIn("The slide title", full_text)
    tc.assertIn("$$f(x)=\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}", full_text)
    tc.assertIn("[Comment: 0@2025-12-28T11:15:49.694: Not second?]", full_text)


def test_read_docx_1() -> None:
    # An actual document from the web - this is likely created on a Windows client
    filename = (
        "sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx"
    )
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    docx: DocxContent = next(read_docx(file_like))

    # text is long. Verify only beginning
    tc.assertEqual("Welcome to the Government", docx.full_text[:25].strip())

    tc.assertEqual(230, len(docx.paragraphs))

    tc.assertEqual(17, docx.metadata.revision)
    # Raw XML format uses 'Z' for UTC timezone
    tc.assertEqual("2023-01-20T16:07:00Z", docx.metadata.modified)
    tc.assertEqual("2022-04-19T14:03:00Z", docx.metadata.created)

    # test iterator
    tc.assertEqual(1, len(list(docx.iterator())))

    # test full text
    tc.assertEqual("Welcome to the Government", docx.get_full_text()[:25].strip())


def test_read_docx_2() -> None:
    # A converted docx from OSX pages - may not populate like a true MS client .docx
    # dedicated test for comment, table and footnote extraction
    filename = (
        "sharepoint2text/tests/resources/modern_ms/sample_with_comment_and_table.docx"
    )
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    docx: DocxContent = next(read_docx(file_like))
    # Formula with properly converted multiplication sign
    tc.assertEqual(
        "Hello World!\nAn image of space\nIncome\ntax\n119\n19\nAnother sentence after the table.\n$$\\frac{3}{4}\\times4=\\sqrt{9}$$",
        docx.full_text,
    )
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
    tc.assertGreater(len(image_bytes.getvalue()), 0)
    tc.assertEqual(docx.images[0].size_bytes, len(image_bytes.getvalue()))
    # get_metadata returns ImageMetadata
    img_meta = docx.images[0].get_metadata()
    tc.assertEqual(
        ImageMetadata(unit_index=0, image_index=1, content_type="image/png"),
        img_meta,
    )


def test_read_docx__image_extraction_1() -> None:
    # Test for caption extraction from following paragraph with caption style
    filename = "sharepoint2text/tests/resources/modern_ms/vorlage-abschlussarbeit.docx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    docx: DocxContent = next(read_docx(file_like))

    tc.assertEqual(1, len(docx.images))

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
    filename = "sharepoint2text/tests/resources/modern_ms/thesis-template.docx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    docx: DocxContent = next(read_docx(file_like))

    tc.assertEqual(2, len(docx.images))
    tc.assertEqual("Illustration 1: [Figure title]", docx.images[1].get_caption())
    tc.assertEqual(
        """Ein Bild, das Zeichnung "Marketing" enthält.""",
        docx.images[1].get_description(),
    )


####################
# Legacy Microsoft #
####################


def test_read_xls_1() -> None:
    filename = "sharepoint2text/tests/resources/legacy_ms/pb_2011_1_gen_web.xls"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    xls: XlsContent = next(read_xls(file_like=file_like))

    tc.assertEqual(13, len(xls.sheets))

    tc.assertEqual("2007-09-19T14:21:02", xls.metadata.created)
    tc.assertEqual("2011-06-01T13:54:08", xls.metadata.modified)
    tc.assertEqual("European Commission", xls.metadata.company)

    # iterator
    tc.assertEqual(13, len(list(xls.iterator())))

    xls_it = xls.iterator()
    # test first page
    s1 = next(xls_it)
    expected = (
        "EUROPEAN UNION\n"
        "                             European Commission\n"
        "  Directorate-General for Mobility and Transport\n"
    )
    tc.assertEqual(expected, s1[:113])

    # test second page
    s2 = next(xls_it)
    tc.assertIn(
        "The content of this pocketbook is based on a range of sources including Eurostat",
        s2,
    )

    # all text
    tc.assertIsNotNone(xls.get_full_text())


def test_read_xls_2() -> None:
    filename = "sharepoint2text/tests/resources/legacy_ms/mwe.xls"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    xls: XlsContent = next(read_xls(file_like=file_like))
    tc.assertEqual(
        "colA  colB\n   1     2",
        xls.get_full_text(),
    )


def test_read_ppt() -> None:
    filename = "sharepoint2text/tests/resources/legacy_ms/eurouni2.ppt"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    ppt: PptContent = next(read_ppt(file_like))

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
    tc.assertEqual(48, len(list(ppt.iterator())))

    # test full text
    tc.assertEqual("European Union", ppt.get_full_text()[:14])


def test_read_doc() -> None:
    with open(
        "sharepoint2text/tests/resources/legacy_ms/Speech_Prime_Minister_of_The_Netherlands_EN.doc",
        mode="rb",
    ) as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    doc: DocContent = next(read_doc(file_like=file_like))

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
    tc.assertEqual(1, len(list(doc.iterator())))

    # test full text
    tc.assertEqual(
        "Short dinner speech by the Prime Minister of the Kingdom of the Netherlands, Dr Jan Peter Balkenende"
        + "\n"
        + "Welcome by the Prime Minister of the Kingdom",
        doc.get_full_text()[:145],
    )


def test_read_rtf() -> None:
    with open(
        "sharepoint2text/tests/resources/legacy_ms/2025.144.un.rtf", mode="rb"
    ) as file:
        file_like = io.BytesIO(file.read())
        rtf_gen: typing.Generator[RtfContent] = read_rtf(file_like=file_like)

        rtfs = list(rtf_gen)
    tc.assertEqual(1, len(rtfs))

    rtf = rtfs[0]
    full_text = rtf.get_full_text()
    tc.assertEqual("c1\nSouth Australia", full_text[:18])
    tc.assertEqual("\non 18 December 2025\nNo 144 of 2025", full_text[-35:])

    tc.assertEqual(1, len(list(rtf.iterator())))


#################
# Email formats #
#################
def test_email__eml_format() -> None:
    with open(
        "sharepoint2text/tests/resources/mails/basic_email.eml", mode="rb"
    ) as file:
        file_like = io.BytesIO(file.read())
        mail_gen: typing.Generator[EmailContent, None, None] = read_eml_format_mail(
            file_like=file_like,
            path="sharepoint2text/tests/resources/mails/basic_email.eml",
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
        "Plain email.\n\nHope it works well!\n\nMikel", list(mail.iterator())[0]
    )

    # metadata
    mail_meta = mail.get_metadata()
    tc.assertEqual("basic_email.eml", mail_meta.filename)
    tc.assertEqual(".eml", mail_meta.file_extension)
    tc.assertEqual("2008-11-22T04:04:59+00:00", mail_meta.date)
    tc.assertEqual(
        "<6B7EC235-5B17-4CA8-B2B8-39290DEB43A3@test.lindsaar.net>", mail_meta.message_id
    )


def test_email__msg_format() -> None:
    with open(
        "sharepoint2text/tests/resources/mails/basic_email.msg", mode="rb"
    ) as file:
        file_like = io.BytesIO(file.read())
        mail_gen: typing.Generator[EmailContent, None, None] = read_msg_format_mail(
            file_like=file_like,
            path="sharepoint2text/tests/resources/mails/basic_email.msg",
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


def test_email__mbox_format() -> None:
    with open(
        "sharepoint2text/tests/resources/mails/basic_email.mbox", mode="rb"
    ) as file:
        file_like = io.BytesIO(file.read())
        mail_gen: typing.Generator[EmailContent, None, None] = read_mbox_format_mail(
            file_like=file_like,
            path="sharepoint2text/tests/resources/mails/basic_email.mbox",
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


###############
# Open Office #
###############


def test_read_open_office__document() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_document.odt"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    odt: OdtContent = next(read_odt(file_like=file_like, path=path))

    tc.assertEqual(".odt", odt.get_metadata().to_dict().get("file_extension"))
    tc.assertEqual("sample_document.odt", odt.get_metadata().to_dict().get("filename"))

    # comments
    tc.assertListEqual(
        [
            OdtAnnotation(
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
    tc.assertListEqual([[["Header 1", "Header 2"], ["Cell A", "Cell B"]]], odt.tables)

    # iterator items
    tc.assertEqual(1, len(list(odt.iterator())))

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
        "End of document.\n"
        "[Annotation: User@2025-12-28T12:00:00: This is a comment by User on the "
        "sample text.]",
        odt.get_full_text(include_annotations=True),
    )


def test_read_open_office__presentation() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_presentation.odp"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    odp: OdpContent = next(read_odp(file_like=file_like, path=path))

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
    # Annotation on slide 2
    tc.assertEqual(1, len(odp.slides[1].annotations))
    tc.assertEqual(
        OdpAnnotation(
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
    tc.assertEqual(3, len(list(odp.iterator())))

    # Full text (default - no annotations, no notes)
    full_text = odp.get_full_text()
    tc.assertIn("Hello World Presentation", full_text)
    tc.assertIn("Content Slide", full_text)
    tc.assertIn("Thank You!", full_text)
    tc.assertNotIn("[Annotation:", full_text)
    tc.assertNotIn("[Note:", full_text)

    # Full text with annotations
    full_text_with_annotations = odp.get_full_text(include_annotations=True)
    tc.assertIn(
        "[Annotation: User@2025-12-28T12:00:00: This is a comment by User on the presentation content.]",
        full_text_with_annotations,
    )

    # Full text with notes
    full_text_with_notes = odp.get_full_text(include_notes=True)
    tc.assertIn("[Note: Speaker notes for Slide 1:", full_text_with_notes)


def test_read_open_office__spreadsheet() -> None:
    path = "sharepoint2text/tests/resources/open_office/sample_spreadsheet.ods"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    ods: OdsContent = next(read_ods(file_like=file_like, path=path))

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
    tc.assertGreater(len(ods.sheets[0].data), 0)
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
        OdsAnnotation(
            creator="User",
            date="2025-12-28T12:00:00",
            text="This is our best-selling product line.",
        ),
        ods.sheets[0].annotations[0],
    )
    # Second annotation: on the notes row
    tc.assertEqual(
        OdsAnnotation(
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
        OdsAnnotation(
            creator="User",
            date="2025-12-28T15:00:00",
            text="These formulas reference the Sales Data sheet. Update source data to refresh.",
        ),
        ods.sheets[1].annotations[0],
    )

    # Iterator yields 2 items (one per sheet)
    tc.assertEqual(2, len(list(ods.iterator())))

    # check length of full text with length of all sheets
    total_length_iteration = sum([len(e) for e in ods.iterator()])
    # one line break is added
    length_total = len(ods.get_full_text()) - 1
    tc.assertEqual(total_length_iteration, length_total)

    # Full text contains data from both sheets
    full_text = ods.get_full_text()
    tc.assertEqual(
        "Sales Data\n" "Product\tQ1\tQ2\tQ3\tQ4\tTotal\nWidget",
        full_text[:44].strip(),
    )


def test_read_open_office__spreadsheet_2() -> None:
    """Verifies the treatment of empty rows and columns in a sheet

    We want that the list of rows is easily processable with Pandas or Polars to create
    dataframes. This requires that None/Nulls are not accidentally pruned. The rows must have
    the same number of columns for this to work
    """
    filename = "sharepoint2text/tests/resources/modern_ms/empty_row_columns.ods"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    ods: OdsContent = next(read_ods(file_like=file_like))
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


def test_open_office__document_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odt"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    odt: OdtContent = next(read_odt(file_like=file_like, path=path))

    tc.assertEqual(2, len(odt.images))
    tc.assertEqual(
        "Illustration 1: Screenshot from the Open Office download website",
        odt.images[0].get_caption(),
    )
    tc.assertEqual(
        ImageMetadata(unit_index=1, image_index=1, content_type="image/png"),
        odt.images[0].get_metadata(),
    )
    tc.assertEqual(90038, len(odt.images[0].get_bytes().getvalue()))
    tc.assertEqual(
        "Illustration 2: Another Image from the download website",
        odt.images[1].get_caption(),
    )
    tc.assertEqual(
        ImageMetadata(unit_index=1, image_index=2, content_type="image/png"),
        odt.images[1].get_metadata(),
    )
    tc.assertEqual(82881, len(odt.images[1].get_bytes().getvalue()))


def test_open_office__presentation_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.odp"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    odp: OdpContent = next(read_odp(file_like=file_like, path=path))
    tc.assertEqual(1, len(odp.slides[0].images))
    tc.assertEqual(
        "",
        odp.slides[0].images[0].get_caption(),
    )
    tc.assertEqual(
        "Screenshot test image\nA test image from the Internet",
        odp.slides[0].images[0].get_description(),
    )


def test_open_office__spreadsheet_image_interface() -> None:
    """Test that OpenDocumentImage correctly implements ImageInterface."""
    # Create an OpenDocumentImage with test data
    path = "sharepoint2text/tests/resources/open_office/image_extraction.ods"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    ods: OdsContent = next(read_ods(file_like=file_like, path=path))
    tc.assertEqual(1, len(ods.sheets[0].images))
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


def test_read_pdf() -> None:
    with open(
        "sharepoint2text/tests/resources/sample.pdf",
        mode="rb",
    ) as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)
    pdf: PdfContent = next(read_pdf(file_like=file_like))

    tc.assertEqual(2, pdf.metadata.total_pages)
    tc.assertListEqual(sorted([1, 2]), sorted(pdf.pages.keys()))

    # Text page 1
    expected = (
        "This is a test sentence" + "\n"
        "This is a table" + "\n"
        "C1 C2" + "\n"
        "R1 V1" + "\n"
        "R2 V2"
    )
    page_1_text = pdf.pages[1].text
    tc.assertEqual(
        expected.strip().replace("\n", " "), page_1_text.strip().replace("\n", " ")
    )

    # Text page 2
    expected = "This is page 2" "\n" "An image of the Google landing page"
    page_2_text = pdf.pages[2].text
    tc.assertEqual(
        expected.strip().replace("\n", " "), page_2_text.strip().replace("\n", " ")
    )

    # Image data
    tc.assertEqual(0, len(pdf.pages[1].images))
    tc.assertEqual(1, len(pdf.pages[2].images))

    # test iterator
    tc.assertEqual(2, len(list(pdf.iterator())))

    # test full text
    tc.assertEqual("This is a test sentence", pdf.get_full_text()[:23])


def test_read_html() -> None:
    path = "sharepoint2text/tests/resources/sample.html"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    html: HtmlContent = next(read_html(file_like=file_like, path=path))

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
