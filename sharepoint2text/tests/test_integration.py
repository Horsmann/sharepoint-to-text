import glob
import io
import json
import logging
import os
import unittest

import sharepoint2text.parsing.exceptions
from sharepoint2text import (
    read_bytes,
    read_doc,
    read_docx,
    read_email__eml_format,
    read_email__mbox_format,
    read_email__msg_format,
    read_eml_email,
    read_file,
    read_html,
    read_mbox_email,
    read_msg_email,
    read_odf,
    read_odg,
    read_odp,
    read_ods,
    read_odt,
    read_pdf,
    read_plain_text,
    read_ppt,
    read_pptx,
    read_rtf,
    read_xls,
    read_xlsx,
)
from sharepoint2text.parsing.extractors.data_types import (
    DocContent,
    DocUnit,
    DocxContent,
    DocxUnit,
    EmailContent,
    EmailUnit,
    HtmlContent,
    HtmlUnit,
    OdfContent,
    OdfUnit,
    OdgContent,
    OdgUnit,
    OdpContent,
    OdpUnit,
    OdsContent,
    OdsUnit,
    OdtContent,
    OdtUnit,
    PdfContent,
    PdfUnit,
    PlainTextContent,
    PlainTextUnit,
    PptContent,
    PptUnit,
    PptxContent,
    PptxUnit,
    RtfContent,
    RtfUnit,
    XlsContent,
    XlsUnit,
    XlsxContent,
    XlsxUnit,
)
from sharepoint2text.parsing.extractors.serialization import deserialize_extraction

logger = logging.getLogger(__name__)

tc = unittest.TestCase()


def _load_as_bytes(path: str) -> io.BytesIO:
    with open(path, "rb") as fd:
        file_like = io.BytesIO(fd.read())
        file_like.seek(0)
        return file_like


def test_read_redirects_from_top_level():
    ############
    # legacy MS
    ############
    # powerpoint
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/legacy_ms/eurouni2.ppt")
    result = next(read_ppt(fl))
    tc.assertTrue(isinstance(result, PptContent))
    # Excel
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/legacy_ms/mwe.xls")
    result = next(read_xls(fl))
    tc.assertTrue(isinstance(result, XlsContent))
    # Word
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/legacy_ms/Speech_Prime_Minister_of_The_Netherlands_EN.doc"
    )
    result = next(read_doc(fl))
    tc.assertTrue(isinstance(result, DocContent))
    # rtf
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/legacy_ms/2025.144.un.rtf"
    )
    result = next(read_rtf(fl))
    tc.assertTrue(isinstance(result, RtfContent))

    ############
    # modern MS
    ############
    # xlsx
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx"
    )
    result = next(read_xlsx(fl))
    tc.assertTrue(isinstance(result, XlsxContent))

    # docx
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx"
    )
    result = next(read_docx(fl))
    tc.assertTrue(isinstance(result, DocxContent))

    # pptx
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx"
    )
    result = next(read_pptx(fl))
    tc.assertTrue(isinstance(result, PptxContent))

    ##############
    # open office
    ##############

    # odt - document
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/open_office/sample_document.odt"
    )
    result = next(read_odt(fl))
    tc.assertTrue(isinstance(result, OdtContent))

    # odp - presentation
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/open_office/sample_presentation.odp"
    )
    result = next(read_odp(fl))
    tc.assertTrue(isinstance(result, OdpContent))

    # ods - spreadsheet
    fl = _load_as_bytes(
        path="sharepoint2text/tests/resources/open_office/sample_spreadsheet.ods"
    )
    result = next(read_ods(fl))
    tc.assertTrue(isinstance(result, OdsContent))

    # odg - drawing
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/open_office/drawing.odg")
    result = next(read_odg(fl))
    tc.assertTrue(isinstance(result, OdgContent))

    # odf - drawing
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/open_office/formular.odf")
    result = next(read_odf(fl))
    tc.assertTrue(isinstance(result, OdfContent))

    ##############
    # Mail
    ##############
    # eml
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.eml")
    result = next(read_email__eml_format(fl))
    tc.assertTrue(isinstance(result, EmailContent))
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.eml")
    result = next(read_eml_email(fl))
    tc.assertTrue(isinstance(result, EmailContent))

    # msg
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.msg")
    result = next(read_email__msg_format(fl))
    tc.assertTrue(isinstance(result, EmailContent))
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.msg")
    result = next(read_msg_email(fl))
    tc.assertTrue(isinstance(result, EmailContent))

    # mbox
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.mbox")
    result = next(read_email__mbox_format(fl))
    tc.assertTrue(isinstance(result, EmailContent))
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.mbox")
    result = next(read_mbox_email(fl))
    tc.assertTrue(isinstance(result, EmailContent))

    ##############
    # Plain
    ##############
    # markdown
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/plain_text/document.md")
    result = next(read_plain_text(fl))
    tc.assertTrue(isinstance(result, PlainTextContent))

    # csv
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/plain_text/plain.csv")
    result = next(read_plain_text(fl))
    tc.assertTrue(isinstance(result, PlainTextContent))

    # tsv
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/plain_text/plain.tsv")
    result = next(read_plain_text(fl))
    tc.assertTrue(isinstance(result, PlainTextContent))

    # txt
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/plain_text/plain.tsv")
    result = next(read_plain_text(fl))
    tc.assertTrue(isinstance(result, PlainTextContent))

    ##############
    # other
    ##############
    # pdf
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/pdf/sample.pdf")
    result = next(read_pdf(fl))
    tc.assertTrue(isinstance(result, PdfContent))

    # html
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/html/sample.html")
    result = next(read_html(fl))
    tc.assertTrue(isinstance(result, HtmlContent))


def test_extract_serialize_deserialize_file():
    for path in glob.glob("sharepoint2text/tests/resources/**/*", recursive=True):
        if not os.path.isfile(path):
            continue
        logger.debug(f"Calling read_file with: [{path}]")
        try:
            for obj in read_file(path=path):
                # verify that all obj have the ExtractionInterface methods
                tc.assertTrue(hasattr(obj, "get_metadata"))
                tc.assertTrue(hasattr(obj, "iterate_units"))
                tc.assertTrue(hasattr(obj, "iterate_images"))
                tc.assertTrue(hasattr(obj, "iterate_tables"))
                tc.assertTrue(hasattr(obj, "get_full_text"))

                # call every method
                obj.get_full_text()
                obj.get_metadata()
                list(obj.iterate_units())
                list(obj.iterate_images())
                list(obj.iterate_tables())

                # make sure that the whole is json-dumpable to ensure we implemented all to_json()
                json.dumps(
                    {
                        "file": path,
                        "full": obj.get_full_text(),
                        "units": list(o.to_json() for o in obj.iterate_units()),
                        "tables": list(o.get_table() for o in obj.iterate_tables()),
                        "metadata": obj.get_metadata().to_dict(),
                    }
                )

                restored_obj = deserialize_extraction(obj.to_json())
                tc.assertEqual(type(restored_obj), type(obj))
        except sharepoint2text.parsing.exceptions.ExtractionFileEncryptedError:
            # silent ignore - we have encrypted test files
            logger.debug(f"File is encrypted: [{path}]")


def test_unit_serialize_deserialize_round_trip():
    cases = [
        ("sharepoint2text/tests/resources/mails/basic_email.eml", EmailUnit),
        (
            "sharepoint2text/tests/resources/legacy_ms/Speech_Prime_Minister_of_The_Netherlands_EN.doc",
            DocUnit,
        ),
        (
            "sharepoint2text/tests/resources/modern_ms/GKIM_Skills_Framework_-_static.docx",
            DocxUnit,
        ),
        ("sharepoint2text/tests/resources/pdf/sample.pdf", PdfUnit),
        ("sharepoint2text/tests/resources/plain_text/plain.txt", PlainTextUnit),
        ("sharepoint2text/tests/resources/html/sample.html", HtmlUnit),
        ("sharepoint2text/tests/resources/legacy_ms/eurouni2.ppt", PptUnit),
        (
            "sharepoint2text/tests/resources/modern_ms/pptx_formula_image.pptx",
            PptxUnit,
        ),
        ("sharepoint2text/tests/resources/legacy_ms/mwe.xls", XlsUnit),
        (
            "sharepoint2text/tests/resources/modern_ms/Country_Codes_and_Names.xlsx",
            XlsxUnit,
        ),
        (
            "sharepoint2text/tests/resources/open_office/sample_presentation.odp",
            OdpUnit,
        ),
        (
            "sharepoint2text/tests/resources/open_office/sample_spreadsheet.ods",
            OdsUnit,
        ),
        (
            "sharepoint2text/tests/resources/open_office/sample_document.odt",
            OdtUnit,
        ),
        (
            "sharepoint2text/tests/resources/open_office/formular.odf",
            OdfUnit,
        ),
        (
            "sharepoint2text/tests/resources/open_office/drawing.odg",
            OdgUnit,
        ),
        ("sharepoint2text/tests/resources/legacy_ms/2025.144.un.rtf", RtfUnit),
    ]

    for path, unit_type in cases:
        result = next(read_file(path))
        unit = next(result.iterate_units())
        tc.assertIsInstance(unit, unit_type)

        payload = unit.to_json()
        restored = deserialize_extraction(payload)
        tc.assertEqual(type(restored), type(unit))
        tc.assertEqual(restored.get_text(), unit.get_text())
        tc.assertEqual(restored.get_metadata(), unit.get_metadata())


def test_read_file_force_plain_text_unknown_extension(tmp_path) -> None:
    path = tmp_path / "custom.weirdformat"
    path.write_text("line one\nline two", encoding="utf-8")

    tc.assertRaises(
        sharepoint2text.parsing.exceptions.ExtractionFileFormatNotSupportedError,
        lambda: next(read_file(path)),
    )

    result = next(read_file(path, force_plain_text=True))
    tc.assertTrue(isinstance(result, PlainTextContent))
    tc.assertEqual("line one\nline two", result.get_full_text())


def test_read_bytes_with_extension_and_bytes_input() -> None:
    with open("sharepoint2text/tests/resources/plain_text/plain.txt", "rb") as fd:
        data = fd.read()

    result = next(read_bytes(data, extension=".txt"))
    tc.assertTrue(isinstance(result, PlainTextContent))
    tc.assertEqual("Hello World", result.get_full_text())

    result_no_dot = next(read_bytes(data, extension="txt"))
    tc.assertTrue(isinstance(result_no_dot, PlainTextContent))
    tc.assertEqual("Hello World", result_no_dot.get_full_text())


def test_read_bytes_with_mime_and_bytesio_input() -> None:
    file_like = _load_as_bytes("sharepoint2text/tests/resources/plain_text/plain.txt")
    result = next(read_bytes(file_like, mime_type="text/plain"))
    tc.assertTrue(isinstance(result, PlainTextContent))
    tc.assertEqual("Hello World", result.get_full_text())


def test_read_bytes_requires_mime_or_extension() -> None:
    tc.assertRaises(ValueError, lambda: next(read_bytes(b"Hello World")))
