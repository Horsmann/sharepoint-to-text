import logging
import unittest

from sharepoint2text.extractors.doc_extractor import read_doc
from sharepoint2text.extractors.docx_extractor import read_docx
from sharepoint2text.extractors.pdf_extractor import read_pdf
from sharepoint2text.extractors.ppt_extractor import read_ppt
from sharepoint2text.extractors.pptx_extractor import read_pptx
from sharepoint2text.extractors.xls_extractor import read_xls
from sharepoint2text.extractors.xlsx_extractor import read_xlsx
from sharepoint2text.router import get_extractor

logger = logging.getLogger(__name__)


def test_router():
    test_case_obj = unittest.TestCase()

    # xls
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/pb_2011_1_gen_web.xls"
    )
    test_case_obj.assertEqual(read_xls, func)

    # xlsx
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/Country_Codes_and_Names.xlsx"
    )
    test_case_obj.assertEqual(read_xlsx, func)

    # pdf
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/sample.pdf"
    )
    test_case_obj.assertEqual(read_pdf, func)

    # ppt
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/eurouni2.ppt"
    )
    test_case_obj.assertEqual(read_ppt, func)

    # pptx
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/eu-visibility_rules_00704232-AF9F-1A18-BD782C469454ADAD_68401.pptx"
    )
    test_case_obj.assertEqual(read_pptx, func)

    # doc
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/Speech_Prime_Minister_of_The_Netherlands_EN.doc"
    )
    test_case_obj.assertEqual(read_doc, func)

    # docx
    func = get_extractor(
        "sharepoint-to-text/sharepoint2text/tests/resources/GKIM_Skills_Framework_-_static.docx"
    )
    test_case_obj.assertEqual(read_docx, func)

    test_case_obj.assertRaises(
        RuntimeError,
        get_extractor,
        "sharepoint-to-text/sharepoint2text/tests/resources/not_supported.misc",
    )
