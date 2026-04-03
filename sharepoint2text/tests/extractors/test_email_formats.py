import io
import logging
import typing
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFailedError,
)
from sharepoint2text.parsing.extractors.data_types import (
    EmailAddress,
    EmailAttachment,
    EmailContent,
    EmailUnitMetadata,
    PdfContent,
    PptxContent,
)
from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
    read_eml_format_mail,
)
from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
    read_mbox_format_mail,
)
from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
    read_msg_format_mail,
)
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


#################
# Email formats #
#################
def test_email__eml_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.eml"
    mail_gen: typing.Generator[EmailContent, None, None] = read_eml_format_mail(
        file_like=read_file_to_file_like(path=path),
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
        file_like=read_file_to_file_like(path=path),
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


def test_email__msg_format_reply_to_is_normalized_list() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.msg"
    mail = next(
        read_msg_format_mail(
            file_like=read_file_to_file_like(path=path),
            path=path,
        )
    )

    tc.assertIsInstance(mail.reply_to, list)
    tc.assertListEqual([], mail.reply_to)


def test_email__msg_format_with_attachment() -> None:
    path = "sharepoint2text/tests/resources/mails/msg_with_attachment.msg"
    mail_gen: typing.Generator[EmailContent, None, None] = read_msg_format_mail(
        file_like=read_file_to_file_like(path=path),
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


def test_email_iterate_supported_attachments_can_raise_or_skip_failures() -> None:
    broken_mail = EmailContent(
        from_email=EmailAddress(name="Sender", address="sender@example.com"),
        subject="broken attachments",
        attachments=[
            EmailAttachment(
                filename="broken.pdf",
                mime_type="application/pdf",
                data=io.BytesIO(b"not-a-real-pdf"),
                is_supported_mime_type=True,
            )
        ],
    )

    with tc.assertRaises(ExtractionFailedError):
        list(broken_mail.iterate_supported_attachments())

    tc.assertEqual(
        [],
        list(broken_mail.iterate_supported_attachments(skip_failed=True)),
    )


def test_email__eml_format_with_attachment() -> None:
    path = "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
    mail_gen: typing.Generator[EmailContent, None, None] = read_eml_format_mail(
        file_like=read_file_to_file_like(path=path),
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


def test_email__eml_format_missing_from_header_is_tolerated() -> None:
    payload = b"Subject: Missing From header\n\nBody"
    mail = next(read_eml_format_mail(file_like=io.BytesIO(payload)))

    tc.assertEqual("", mail.from_email.name)
    tc.assertEqual("", mail.from_email.address)
    tc.assertEqual("Missing From header", mail.subject)
    tc.assertEqual("", mail.get_metadata().date)


def test_email__mbox_format_missing_date_and_from_headers_is_tolerated() -> None:
    payload = (
        b"From sender@example.com Mon Jan  1 00:00:00 2024\n"
        b"Subject: Missing date and from\n"
        b"\n"
        b"Body\n"
    )

    mails = list(read_mbox_format_mail(file_like=io.BytesIO(payload)))

    tc.assertEqual(1, len(mails))
    mail = mails[0]
    tc.assertEqual("", mail.from_email.name)
    tc.assertEqual("", mail.from_email.address)
    tc.assertEqual("", mail.get_metadata().date)
    tc.assertEqual("Missing date and from", mail.subject)
    tc.assertEqual("Body", mail.body_plain)
