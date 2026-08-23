import io
import logging
import typing
from pathlib import Path
from unittest import TestCase

from sharepoint2text import read_bytes
from sharepoint2text.parsing.exceptions import ExtractionFailedError
from sharepoint2text.parsing.extractors.mail.eml_email_extractor import (
    read_eml_format_mail,
)
from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
    read_mbox_format_mail,
)
from sharepoint2text.parsing.extractors.mail.msg_email_extractor import (
    read_msg_format_mail,
)
from sharepoint2text.parsing.models import Attachment, ExtractedDocument, JsonValue
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def _mail_addresses(
    document: ExtractedDocument, field: str
) -> list[dict[str, JsonValue]]:
    value = document.properties.get(f"{document.format}.{field}", [])
    return typing.cast(list[dict[str, JsonValue]], value)


def _sender(document: ExtractedDocument) -> dict[str, JsonValue]:
    value = document.properties[f"{document.format}.from_email"]
    return typing.cast(dict[str, JsonValue], value)


def _extract_attachment(attachment: Attachment) -> ExtractedDocument:
    return next(
        read_bytes(
            attachment.data or b"",
            mime_type=attachment.media_type,
            extension=Path(attachment.filename).suffix,
        )
    )


#################
# Email formats #
#################
def test_email__eml_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.eml"
    mail_gen: typing.Generator[ExtractedDocument, None, None] = read_eml_format_mail(
        file_like=read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    tc.assertEqual(1, len(mails))
    mail = mails[0]

    tc.assertEqual("Mikel Lindsaar", _sender(mail)["name"])
    tc.assertEqual("test@lindsaar.net", _sender(mail)["address"])

    recipients = _mail_addresses(mail, "to_emails")
    tc.assertEqual(1, len(recipients))
    tc.assertEqual("Mikel Lindsaar", recipients[0]["name"])
    tc.assertEqual("raasdnil@gmail.com", recipients[0]["address"])

    cc = _mail_addresses(mail, "to_cc")
    tc.assertEqual(2, len(cc))
    tc.assertEqual("Jane Doe", cc[0]["name"])
    tc.assertEqual("jane.doe@example.test", cc[0]["address"])
    tc.assertEqual("Bob Smith", cc[1]["name"])
    tc.assertEqual("bob.smith@example.test", cc[1]["address"])

    bcc = _mail_addresses(mail, "to_bcc")
    tc.assertEqual(2, len(bcc))
    tc.assertEqual("Hidden Tester", bcc[0]["name"])
    tc.assertEqual("hidden.tester@example.test", bcc[0]["address"])
    tc.assertEqual("Silent Observer", bcc[1]["name"])
    tc.assertEqual("silent.observer@example.test", bcc[1]["address"])

    expected_body = "Plain email.\n\nHope it works well!\n\nMikel"
    tc.assertEqual(expected_body, mail.properties["eml.body_plain"])
    tc.assertEqual("Testing 123", mail.metadata.title)
    tc.assertEqual(expected_body, mail.full_text)
    tc.assertEqual(expected_body, mail.units[0].text)
    tc.assertEqual(0, len(list(mail.iter_images())))
    tc.assertEqual(0, len(list(mail.iter_tables())))

    tc.assertEqual("basic_email.eml", mail.source.filename)
    tc.assertEqual(".eml", mail.source.extension)
    tc.assertEqual("2008-11-22T04:04:59+00:00", mail.metadata.created)
    tc.assertEqual(
        "<6B7EC235-5B17-4CA8-B2B8-39290DEB43A3@test.lindsaar.net>",
        mail.metadata.properties["eml.message_id"],
    )

    tc.assertEqual(1, len(mail.units))
    tc.assertEqual(1, mail.units[0].number)
    tc.assertEqual("message", mail.units[0].kind)
    tc.assertEqual("plain", mail.units[0].properties["eml.body_type"])


def test_email__msg_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.msg"
    mail = next(
        read_msg_format_mail(
            file_like=read_file_to_file_like(path=path),
            path=path,
        )
    )

    tc.assertEqual("Brian Zhou", _sender(mail)["name"])
    tc.assertEqual("brizhou@gmail.com", _sender(mail)["address"])

    recipients = _mail_addresses(mail, "to_emails")
    tc.assertEqual(1, len(recipients))
    tc.assertEqual("", recipients[0]["name"])
    tc.assertEqual("brianzhou@me.com", recipients[0]["address"])

    cc = _mail_addresses(mail, "to_cc")
    tc.assertEqual(1, len(cc))
    tc.assertEqual("Brian Zhou", cc[0]["name"])
    tc.assertEqual("brizhou@gmail.com", cc[0]["address"])
    tc.assertListEqual([], _mail_addresses(mail, "to_bcc"))

    tc.assertEqual("Test for TIF files", mail.metadata.title)
    tc.assertEqual(
        "This is a test email to experiment with",
        typing.cast(str, mail.properties["msg.body_plain"])[:39],
    )

    tc.assertEqual("basic_email.msg", mail.source.filename)
    tc.assertEqual(".msg", mail.source.extension)
    tc.assertEqual("2013-11-18T10:26:24+02:00", mail.metadata.created)
    tc.assertEqual(
        "<CADtJ4eNjQSkGcBtVteCiTF+YFG89+AcHxK3QZ=-Mt48xygkvdQ@mail.gmail.com>",
        mail.metadata.properties["msg.message_id"],
    )
    tc.assertEqual(0, len(list(mail.iter_images())))
    tc.assertEqual(0, len(list(mail.iter_tables())))


def test_email__msg_format_reply_to_is_normalized_list() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.msg"
    mail = next(
        read_msg_format_mail(
            file_like=read_file_to_file_like(path=path),
            path=path,
        )
    )

    tc.assertIsInstance(_mail_addresses(mail, "reply_to"), list)
    tc.assertListEqual([], _mail_addresses(mail, "reply_to"))


def test_email__msg_format_with_attachment() -> None:
    path = "sharepoint2text/tests/resources/mails/msg_with_attachment.msg"
    mail = next(
        read_msg_format_mail(
            file_like=read_file_to_file_like(path=path),
            path=path,
        )
    )

    tc.assertIsNotNone(_sender(mail)["name"])
    tc.assertIsNotNone(_sender(mail)["address"])
    recipients = _mail_addresses(mail, "to_emails")
    tc.assertEqual(1, len(recipients))
    tc.assertIsNotNone(recipients[0]["name"])
    tc.assertEqual("", recipients[0]["address"])
    tc.assertListEqual([], _mail_addresses(mail, "to_cc"))
    tc.assertListEqual([], _mail_addresses(mail, "to_bcc"))

    tc.assertEqual("Test .msg with attachment", mail.metadata.title)
    tc.assertEqual("", mail.properties.get("msg.body_plain", ""))
    tc.assertEqual(
        "<html><head>",
        typing.cast(str, mail.properties["msg.body_html"])[:12],
    )

    tc.assertEqual("msg_with_attachment.msg", mail.source.filename)
    tc.assertEqual(".msg", mail.source.extension)
    tc.assertEqual("2025-12-31T12:32:42+00:00", mail.metadata.created)
    tc.assertEqual(
        "<VE1PR10MB3790E964D9B988D177790593FABDA@VE1PR10MB3790.EURPRD10.PROD.OUTLOOK.COM>",
        mail.metadata.properties["msg.message_id"],
    )

    tc.assertEqual(2, len(mail.attachments))
    attachments_by_name = {
        attachment.filename: attachment for attachment in mail.attachments
    }
    tc.assertIn("sample.pdf", attachments_by_name)
    tc.assertIn("pptx_formula_image.pptx", attachments_by_name)

    pdf_attachment = attachments_by_name["sample.pdf"]
    tc.assertEqual("application/pdf", pdf_attachment.media_type)
    tc.assertIsInstance(pdf_attachment.data, bytes)
    tc.assertEqual(249095, len(pdf_attachment.data or b""))
    tc.assertTrue(pdf_attachment.properties["email.is_supported_mime_type"])

    attachments = [_extract_attachment(item) for item in mail.attachments]
    tc.assertEqual(2, len(attachments))
    tc.assertIsInstance(attachments[0], ExtractedDocument)
    tc.assertIsInstance(attachments[1], ExtractedDocument)
    tc.assertEqual(
        "This is a test sentence\n"
        "This is a table\n"
        "C1 C2\n"
        "R1 V1\n"
        "R2 V2\n"
        "This is page 2\n"
        "An image of the Google landing page",
        attachments[0].full_text,
    )
    tc.assertEqual(1, len(list(attachments[0].iter_images())))
    tc.assertEqual(
        "The slide title\nThe first text line\n\n\n\n\nThe last text line\nA beach\n$$f(x)=\\frac{1}{\\sqrt{2\\pi\\sigma^{2}}}e^{-\\frac{(x-\\mu)^{2}}{2\\sigma^{2}}}$$",
        attachments[1].full_text,
    )
    tc.assertEqual(1, len(list(attachments[1].iter_images())))

    pptx_attachment = attachments_by_name["pptx_formula_image.pptx"]
    tc.assertEqual(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        pptx_attachment.media_type,
    )
    tc.assertIsInstance(pptx_attachment.data, bytes)
    tc.assertEqual(1566612, len(pptx_attachment.data or b""))
    tc.assertTrue(pptx_attachment.properties["email.is_supported_mime_type"])
    tc.assertEqual(0, len(list(mail.iter_images())))
    tc.assertEqual(0, len(list(mail.iter_tables())))


def test_email_iterate_supported_attachments_can_raise_or_skip_failures() -> None:
    broken_attachment = Attachment(
        filename="broken.pdf",
        media_type="application/pdf",
        data=b"not-a-real-pdf",
        properties={"email.is_supported_mime_type": True},
    )

    with tc.assertRaises(ExtractionFailedError):
        _extract_attachment(broken_attachment)

    extracted: list[ExtractedDocument] = []
    try:
        extracted.append(_extract_attachment(broken_attachment))
    except ExtractionFailedError:
        pass
    tc.assertListEqual([], extracted)


def test_email__eml_format_with_attachment() -> None:
    path = "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
    mail = next(
        read_eml_format_mail(
            file_like=read_file_to_file_like(path=path),
            path=path,
        )
    )

    tc.assertIsNotNone(_sender(mail)["name"])
    tc.assertIsNotNone(_sender(mail)["address"])
    recipients = _mail_addresses(mail, "to_emails")
    tc.assertEqual(1, len(recipients))
    tc.assertIsNotNone(recipients[0]["name"])
    tc.assertIsNotNone(recipients[0]["address"])
    tc.assertListEqual([], _mail_addresses(mail, "to_cc"))
    tc.assertListEqual([], _mail_addresses(mail, "to_bcc"))

    tc.assertEqual("Test .msg with attachment", mail.metadata.title)
    tc.assertEqual(
        "<html><head>",
        typing.cast(str, mail.properties["eml.body_html"])[:12],
    )

    tc.assertEqual("msg_with_attachment.eml", mail.source.filename)
    tc.assertEqual(".eml", mail.source.extension)
    tc.assertEqual("2025-12-31T12:32:42+00:00", mail.metadata.created)
    tc.assertEqual(
        "<VE1PR10MB3790E964D9B988D177790593FABDA@VE1PR10MB3790.EURPRD10.PROD.OUTLOOK.COM>",
        mail.metadata.properties["eml.message_id"],
    )

    tc.assertEqual(2, len(mail.attachments))
    attachments_by_name = {
        attachment.filename: attachment for attachment in mail.attachments
    }
    tc.assertIn("sample.pdf", attachments_by_name)
    tc.assertIn("pptx_formula_image.pptx", attachments_by_name)

    pdf_attachment = attachments_by_name["sample.pdf"]
    tc.assertEqual("application/pdf", pdf_attachment.media_type)
    tc.assertIsInstance(pdf_attachment.data, bytes)
    tc.assertEqual(249095, len(pdf_attachment.data or b""))
    tc.assertTrue(pdf_attachment.properties["email.is_supported_mime_type"])

    attachments = [_extract_attachment(item) for item in mail.attachments]
    tc.assertEqual(2, len(attachments))
    tc.assertIsInstance(attachments[0], ExtractedDocument)
    tc.assertIsInstance(attachments[1], ExtractedDocument)

    pptx_attachment = attachments_by_name["pptx_formula_image.pptx"]
    tc.assertEqual(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        pptx_attachment.media_type,
    )
    tc.assertIsInstance(pptx_attachment.data, bytes)
    tc.assertEqual(1566612, len(pptx_attachment.data or b""))
    tc.assertTrue(pptx_attachment.properties["email.is_supported_mime_type"])
    tc.assertEqual(0, len(list(mail.iter_images())))
    tc.assertEqual(0, len(list(mail.iter_tables())))


def test_email__mbox_format_with_attachments() -> None:
    """MBOX extraction should retain and recursively extract MIME attachments."""
    eml_path = "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
    eml_payload = read_file_to_file_like(path=eml_path).getvalue()
    mbox_payload = (
        b"From sender@example.com Wed Dec 31 12:32:42 2025\n" + eml_payload + b"\n"
    )

    mail = next(
        read_mbox_format_mail(
            file_like=io.BytesIO(mbox_payload),
            path="mailbox.mbox",
        )
    )

    tc.assertListEqual(
        ["sample.pdf", "pptx_formula_image.pptx"],
        [attachment.filename for attachment in mail.attachments],
    )
    tc.assertTrue(
        all(
            item.properties["email.is_supported_mime_type"] for item in mail.attachments
        )
    )
    tc.assertTrue(all(isinstance(item.data, bytes) for item in mail.attachments))

    extracted_attachments = [_extract_attachment(item) for item in mail.attachments]
    tc.assertEqual(2, len(extracted_attachments))
    tc.assertIsInstance(extracted_attachments[0], ExtractedDocument)
    tc.assertIsInstance(extracted_attachments[1], ExtractedDocument)

    mail_without_attachments = next(
        read_mbox_format_mail(
            file_like=io.BytesIO(mbox_payload),
            include_attachments=False,
        )
    )
    tc.assertListEqual([], mail_without_attachments.attachments)


def test_email__eml_format_missing_from_header_is_tolerated() -> None:
    payload = b"Subject: Missing From header\n\nBody"
    mail = next(read_eml_format_mail(file_like=io.BytesIO(payload)))

    tc.assertEqual("", _sender(mail)["name"])
    tc.assertEqual("", _sender(mail)["address"])
    tc.assertEqual("Missing From header", mail.metadata.title)
    tc.assertIsNone(mail.metadata.created)


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
    tc.assertEqual("", _sender(mail)["name"])
    tc.assertEqual("", _sender(mail)["address"])
    tc.assertIsNone(mail.metadata.created)
    tc.assertEqual("Missing date and from", mail.metadata.title)
    tc.assertEqual("Body", mail.properties["mbox.body_plain"])
