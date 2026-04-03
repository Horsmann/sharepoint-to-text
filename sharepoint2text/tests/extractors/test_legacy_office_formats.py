import logging
import typing
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    DocContent,
    DocImage,
    ImageInterface,
    ImageMetadata,
    PptContent,
    PptImage,
    RtfContent,
    RtfUnitMetadata,
    TableDim,
    XlsContent,
    XlsUnitMetadata,
)
from sharepoint2text.parsing.extractors.ms_legacy.doc_extractor import read_doc
from sharepoint2text.parsing.extractors.ms_legacy.ppt_extractor import read_ppt
from sharepoint2text.parsing.extractors.ms_legacy.rtf_extractor import read_rtf
from sharepoint2text.parsing.extractors.ms_legacy.xls_extractor import read_xls
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


####################
# Legacy Microsoft #
####################


def test_read_xls_1() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/pb_2011_1_gen_web.xls"

    xls: XlsContent = next(read_xls(file_like=read_file_to_file_like(path=path)))

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
    xls: XlsContent = next(read_xls(file_like=read_file_to_file_like(path=path)))
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
    xls: XlsContent = next(read_xls(file_like=read_file_to_file_like(path=path)))

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
    ppt: PptContent = next(read_ppt(read_file_to_file_like(path=path)))

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
        read_ppt(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertListEqual(
        ["This is an example text in the notes section"], ppt.slides[0].notes
    )


def test_read_ppt__image_extraction() -> None:
    """Test image extraction from legacy PPT files."""
    path = "sharepoint2text/tests/resources/legacy_ms/ppt_with_images.ppt"
    ppt: PptContent = next(read_ppt(read_file_to_file_like(path=path)))

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
    units = list(ppt.iterate_units())
    tc.assertEqual(2, len(units))
    tc.assertEqual("", units[0].get_text())
    tc.assertEqual(1, len(units[0].get_images()))
    tc.assertEqual(1, len(units[1].get_images()))


def test_read_ppt__image_flag() -> None:
    """Test legacy .ppt image extraction can be disabled with ignore_images."""
    path = "sharepoint2text/tests/resources/legacy_ms/ppt_with_images.ppt"

    result_with_images: PptContent = next(
        read_ppt(
            file_like=read_file_to_file_like(path=path),
            path=path,
            ignore_images=False,
        ),
    )
    tc.assertEqual(2, len(list(result_with_images.iterate_images())))
    tc.assertEqual(2, sum(len(slide.images) for slide in result_with_images.slides))

    result_without_images: PptContent = next(
        read_ppt(
            file_like=read_file_to_file_like(path=path),
            path=path,
            ignore_images=True,
        ),
    )
    tc.assertEqual(0, len(list(result_without_images.iterate_images())))
    tc.assertEqual(0, sum(len(slide.images) for slide in result_without_images.slides))


def test_read_doc() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/Speech_Prime_Minister_of_The_Netherlands_EN.doc"
    doc: DocContent = next(read_doc(file_like=read_file_to_file_like(path=path)))

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

    doc: DocContent = next(read_doc(file_like=read_file_to_file_like(path=path)))

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

    units = list(doc.iterate_units())
    tc.assertEqual(1, len(units))
    tc.assertEqual(1, len(units[0].get_images()))
    tc.assertEqual(
        ImageMetadata(
            unit_number=1,
            image_number=1,
            content_type="image/bmp",
            width=1304,
            height=660,
        ),
        units[0].get_images()[0].get_metadata(),
    )


def test_read_doc__image_extraction_2() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/legacy_doc_multi_image.doc"
    doc: DocContent = next(read_doc(file_like=read_file_to_file_like(path=path)))
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
        read_doc(file_like=read_file_to_file_like(path=path), path=path)
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
        file_like=read_file_to_file_like(path=path)
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
        file_like=read_file_to_file_like(path=path), path=path
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
        file_like=read_file_to_file_like(path=path)
    )

    rtfs = list(rtf_gen)
    tc.assertEqual(1, len(rtfs))

    tables = list(rtfs[0].iterate_tables())
    tc.assertEqual(23, len(tables))


def test_password_protected__xls() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/xls-password-protected-pw123.xls"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_xls(file_like=read_file_to_file_like(path=path), path=path))


def test_password_protected__doc() -> None:
    path = "sharepoint2text/tests/resources/legacy_ms/password_protected/doc-password-protected-pw123.doc"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_doc(file_like=read_file_to_file_like(path=path), path=path))
