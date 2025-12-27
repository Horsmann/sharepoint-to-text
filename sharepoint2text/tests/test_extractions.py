import io
import logging
import typing
from unittest import TestCase

from sharepoint2text.extractors.data_types import (
    DocContent,
    DocxContent,
    EmailContent,
    FileMetadataInterface,
    PdfContent,
    PlainTextContent,
    PptContent,
    PptxContent,
    XlsContent,
    XlsxContent,
)
from sharepoint2text.extractors.doc_extractor import read_doc
from sharepoint2text.extractors.docx_extractor import read_docx
from sharepoint2text.extractors.mail.eml_email_extractor import read_eml_format_mail
from sharepoint2text.extractors.mail.mbox_email_extractor import read_mbox_format_mail
from sharepoint2text.extractors.mail.msg_email_extractor import read_msg_format_mail
from sharepoint2text.extractors.pdf_extractor import read_pdf
from sharepoint2text.extractors.plain_extractor import read_plain_text
from sharepoint2text.extractors.ppt_extractor import read_ppt
from sharepoint2text.extractors.pptx_extractor import read_pptx
from sharepoint2text.extractors.xls_extractor import read_xls
from sharepoint2text.extractors.xlsx_extractor import read_xlsx

logger = logging.getLogger(__name__)

tc = TestCase()


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


def test_read_text() -> None:
    path = "sharepoint2text/tests/resources/plain.txt"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain: PlainTextContent = next(read_plain_text(file_like=file_like))

    tc.assertEqual("Hello World\n", plain.content)

    # csv file
    path = "sharepoint2text/tests/resources/plain.csv"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain: PlainTextContent = next(read_plain_text(file_like=file_like, path=path))

    tc.assertEqual('Text; Date\n"Hello World"; "2025-12-25"\n', plain.content)

    tc.assertEqual(1, len(list(plain.iterator())))

    # tsv file
    path = "sharepoint2text/tests/resources/plain.tsv"
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    plain = next(read_plain_text(file_like=file_like, path=path))

    tc.assertEqual("Text\tDate\nHello World\t2025-12-25\n", plain.content)


def test_read_xlsx() -> None:
    filename = "sharepoint2text/tests/resources/Country_Codes_and_Names.xlsx"
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


def test_read_xls() -> None:
    filename = "sharepoint2text/tests/resources/pb_2011_1_gen_web.xls"
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

    # all text
    tc.assertIsNotNone(xls.get_full_text())


def test_read_ppt() -> None:
    filename = "sharepoint2text/tests/resources/eurouni2.ppt"
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


def test_read_pptx() -> None:
    filename = "sharepoint2text/tests/resources/eu-visibility_rules_00704232-AF9F-1A18-BD782C469454ADAD_68401.pptx"
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

    expected = ["This is the wrong EU emblem", "2", "Don’t use this emblem!"]
    tc.assertListEqual(expected, pptx.slides[1].other_textboxes)

    # images
    tc.assertEqual(5, len(pptx.slides[1].images))
    # test presence of metadata for first image only
    tc.assertEqual(1, pptx.slides[1].images[0].image_index)
    tc.assertEqual("image.jpeg", pptx.slides[1].images[0].filename)
    tc.assertEqual("image/jpeg", pptx.slides[1].images[0].content_type)
    tc.assertEqual(8947, pptx.slides[1].images[0].size_bytes)
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


def test_read_docx() -> None:
    filename = "sharepoint2text/tests/resources/GKIM_Skills_Framework_-_static.docx"
    with open(filename, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)

    docx: DocxContent = next(read_docx(file_like))

    # text is long. Verify only beginning and ending of the combined text
    tc.assertEqual("Welcome to the Government", docx.full_text[:25].strip())
    tc.assertEqual("improved records systems.", docx.full_text[-25:])

    tc.assertEqual(230, len(docx.paragraphs))

    tc.assertEqual(17, docx.metadata.revision)
    tc.assertEqual("2023-01-20T16:07:00+00:00", docx.metadata.modified)
    tc.assertEqual("2022-04-19T14:03:00+00:00", docx.metadata.created)

    # test iterator
    tc.assertEqual(1, len(list(docx.iterator())))

    # test full text
    tc.assertEqual("Welcome to the Government", docx.get_full_text()[:25].strip())


def test_read_doc() -> None:
    with open(
        "sharepoint2text/tests/resources/Speech_Prime_Minister_of_The_Netherlands_EN.doc",
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
