import logging
import unittest

from sharepoint2text.exceptions import ExtractionFileFormatNotSupportedError
from sharepoint2text.extractors.epub_extractor import read_epub
from sharepoint2text.extractors.html_extractor import read_html
from sharepoint2text.extractors.mail.eml_email_extractor import read_eml_format_mail
from sharepoint2text.extractors.mail.mbox_email_extractor import read_mbox_format_mail
from sharepoint2text.extractors.mail.msg_email_extractor import read_msg_format_mail
from sharepoint2text.extractors.mhtml_extractor import read_mhtml
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
from sharepoint2text.extractors.pdf.pdf_extractor import read_pdf
from sharepoint2text.extractors.plain_extractor import read_plain_text
from sharepoint2text.router import get_extractor, is_supported_file

logger = logging.getLogger(__name__)

tc = unittest.TestCase()


def test_is_supported():
    # supported
    tc.assertTrue(is_supported_file("myfile.ppt"))
    tc.assertTrue(is_supported_file("myfile.pptx"))
    tc.assertTrue(is_supported_file("myfile.xls"))
    tc.assertTrue(is_supported_file("myfile.xlsx"))
    tc.assertTrue(is_supported_file("myfile.doc"))
    tc.assertTrue(is_supported_file("myfile.docx"))
    tc.assertTrue(is_supported_file("myfile.pdf"))
    tc.assertTrue(is_supported_file("myfile.eml"))
    tc.assertTrue(is_supported_file("myfile.msg"))
    tc.assertTrue(is_supported_file("myfile.mbox"))
    tc.assertTrue(is_supported_file("myfile.txt"))
    tc.assertTrue(is_supported_file("myfile.md"))
    tc.assertTrue(is_supported_file("myfile.csv"))
    tc.assertTrue(is_supported_file("myfile.tsv"))
    tc.assertTrue(is_supported_file("myfile.rtf"))
    tc.assertTrue(is_supported_file("myfile.html"))
    tc.assertTrue(is_supported_file("myfile.htm"))
    tc.assertTrue(is_supported_file("myfile.HTML"))
    tc.assertTrue(is_supported_file("myfile.HTM"))
    tc.assertTrue(is_supported_file("myfile.odt"))
    tc.assertTrue(is_supported_file("myfile.odp"))
    tc.assertTrue(is_supported_file("myfile.ods"))
    tc.assertTrue(is_supported_file("myfile.epub"))
    tc.assertTrue(is_supported_file("myfile.mhtml"))
    tc.assertTrue(is_supported_file("myfile.mht"))
    # Macro-enabled Office formats
    tc.assertTrue(is_supported_file("myfile.docm"))
    tc.assertTrue(is_supported_file("myfile.xlsm"))
    tc.assertTrue(is_supported_file("myfile.pptm"))
    # not supported
    tc.assertFalse(is_supported_file("myfile.zip"))
    tc.assertFalse(is_supported_file("myfile.rar"))
    tc.assertFalse(is_supported_file("myfile.exe"))
    tc.assertFalse(is_supported_file("myfile.bat"))


def test_router():

    # xls
    func = get_extractor("myfile.xls")
    tc.assertEqual(read_xls, func)

    # xlsx
    func = get_extractor("myfile.xlsx")
    tc.assertEqual(read_xlsx, func)

    # pdf
    func = get_extractor("myfile.pdf")
    tc.assertEqual(read_pdf, func)

    # ppt
    func = get_extractor("myfile.ppt")
    tc.assertEqual(read_ppt, func)

    # pptx
    func = get_extractor("myfile.pptx")
    tc.assertEqual(read_pptx, func)

    # doc
    func = get_extractor("myfile.doc")
    tc.assertEqual(read_doc, func)

    # docx
    func = get_extractor("myfile.docx")
    tc.assertEqual(read_docx, func)

    # json
    func = get_extractor("myfile.json")
    tc.assertEqual(read_plain_text, func)

    # txt
    func = get_extractor("myfile.txt")
    tc.assertEqual(read_plain_text, func)

    # md
    func = get_extractor("myfile.md")
    tc.assertEqual(read_plain_text, func)

    # csv
    func = get_extractor("myfile.csv")
    tc.assertEqual(read_plain_text, func)

    # tsv
    func = get_extractor("myfile.tsv")
    tc.assertEqual(read_plain_text, func)

    # rtf
    func = get_extractor("myfile.rtf")
    tc.assertEqual(read_rtf, func)

    # open office - document - odt
    func = get_extractor("myfile.odt")
    tc.assertEqual(read_odt, func)

    # open office - presentation - odp
    func = get_extractor("myfile.odp")
    tc.assertEqual(read_odp, func)

    # open office - spreadsheet - ods
    func = get_extractor("myfile.ods")
    tc.assertEqual(read_ods, func)

    # html
    func = get_extractor("myfile.html")
    tc.assertEqual(read_html, func)

    # htm
    func = get_extractor("myfile.htm")
    tc.assertEqual(read_html, func)

    # msg
    func = get_extractor("myfile.msg")
    tc.assertEqual(read_msg_format_mail, func)

    # eml
    func = get_extractor("myfile.eml")
    tc.assertEqual(read_eml_format_mail, func)

    # mbox
    func = get_extractor("myfile.mbox")
    tc.assertEqual(read_mbox_format_mail, func)

    # epub
    func = get_extractor("myfile.epub")
    tc.assertEqual(read_epub, func)

    # mhtml
    func = get_extractor("myfile.mhtml")
    tc.assertEqual(read_mhtml, func)

    # mht (alias for mhtml)
    func = get_extractor("myfile.mht")
    tc.assertEqual(read_mhtml, func)

    # Macro-enabled Office formats (use same extractors as non-macro)
    # docm -> read_docx
    func = get_extractor("myfile.docm")
    tc.assertEqual(read_docx, func)

    # xlsm -> read_xlsx
    func = get_extractor("myfile.xlsm")
    tc.assertEqual(read_xlsx, func)

    # pptm -> read_pptx
    func = get_extractor("myfile.pptm")
    tc.assertEqual(read_pptx, func)

    tc.assertRaises(
        ExtractionFileFormatNotSupportedError,
        get_extractor,
        "not_supported.misc",
    )

    tc.assertRaises(
        ExtractionFileFormatNotSupportedError,
        get_extractor,
        "i-have-no-file-type",
    )
