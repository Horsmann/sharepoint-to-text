import glob
import io
import logging
import os
import unittest

import sharepoint2text.exceptions
from sharepoint2text import (
    read_doc,
    read_docx,
    read_email__eml_format,
    read_email__mbox_format,
    read_email__msg_format,
    read_file,
    read_html,
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
from sharepoint2text.extractors.data_types import (
    DocContent,
    DocxContent,
    EmailContent,
    HtmlContent,
    OdpContent,
    OdsContent,
    OdtContent,
    PdfContent,
    PlainTextContent,
    PptContent,
    PptxContent,
    RtfContent,
    XlsContent,
    XlsxContent,
)
from sharepoint2text.extractors.serialization import deserialize_extraction

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

    ##############
    # Mail
    ##############
    # eml
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.eml")
    result = next(read_email__eml_format(fl))
    tc.assertTrue(isinstance(result, EmailContent))

    # msg
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.msg")
    result = next(read_email__msg_format(fl))
    tc.assertTrue(isinstance(result, EmailContent))

    # mbox
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/mails/basic_email.mbox")
    result = next(read_email__mbox_format(fl))
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
    fl = _load_as_bytes(path="sharepoint2text/tests/resources/sample.html")
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

                restored_obj = deserialize_extraction(obj.to_json())
                tc.assertEqual(type(restored_obj), type(obj))
        except sharepoint2text.exceptions.ExtractionFileEncryptedError:
            # silent ignore - we have encrypted test files
            logger.debug(f"File is encrypted: [{path}]")
            pass
