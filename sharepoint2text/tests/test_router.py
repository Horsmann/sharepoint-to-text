import logging
import unittest

from sharepoint2text.extractors.doc_extractor import read_doc
from sharepoint2text.extractors.docx_extractor import read_docx
from sharepoint2text.extractors.pdf_extractor import read_pdf
from sharepoint2text.extractors.plain_extractor import read_plain_text
from sharepoint2text.extractors.ppt_extractor import read_ppt
from sharepoint2text.extractors.pptx_extractor import read_pptx
from sharepoint2text.extractors.xls_extractor import read_xls
from sharepoint2text.extractors.xlsx_extractor import read_xlsx
from sharepoint2text.router import get_extractor

logger = logging.getLogger(__name__)


def test_router():
    test_case_obj = unittest.TestCase()

    # xls
    func = get_extractor("myfile.xls")
    test_case_obj.assertEqual(read_xls, func)

    # xlsx
    func = get_extractor("myfile.xlsx")
    test_case_obj.assertEqual(read_xlsx, func)

    # pdf
    func = get_extractor("myfile.pdf")
    test_case_obj.assertEqual(read_pdf, func)

    # ppt
    func = get_extractor("myfile.ppt")
    test_case_obj.assertEqual(read_ppt, func)

    # pptx
    func = get_extractor("myfile.pptx")
    test_case_obj.assertEqual(read_pptx, func)

    # doc
    func = get_extractor("myfile.doc")
    test_case_obj.assertEqual(read_doc, func)

    # docx
    func = get_extractor("myfile.docx")
    test_case_obj.assertEqual(read_docx, func)

    # json
    func = get_extractor("myfile.json")
    test_case_obj.assertEqual(read_plain_text, func)

    # txt
    func = get_extractor("myfile.txt")
    test_case_obj.assertEqual(read_plain_text, func)

    # csv
    func = get_extractor("myfile.csv")
    test_case_obj.assertEqual(read_plain_text, func)

    # tsv
    func = get_extractor("myfile.tsv")
    test_case_obj.assertEqual(read_plain_text, func)

    test_case_obj.assertRaises(
        RuntimeError,
        get_extractor,
        "not_supported.misc",
    )
