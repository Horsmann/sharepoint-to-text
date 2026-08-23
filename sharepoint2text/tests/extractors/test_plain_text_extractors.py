import logging
from unittest import TestCase

from sharepoint2text.parsing.extractors.plain_extractor import read_plain_text
from sharepoint2text.parsing.models import ExtractedDocument
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


#########
# Plain #
#########


def test_read_text() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.txt"
    plain: ExtractedDocument = next(
        read_plain_text(file_like=read_file_to_file_like(path), path=path)
    )

    tc.assertEqual("Hello World", plain.full_text)
    tc.assertEqual("Hello World", plain.units[0].text)
    tc.assertEqual(1, len(list(plain.units)))
    tc.assertEqual(0, len(list(plain.iter_images())))
    tc.assertEqual(0, len(list(plain.iter_tables())))

    units = list(plain.units)
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("document", units[0].kind)

    tc.assertEqual("ascii", plain.source.encoding)
    tc.assertEqual("plain.txt", plain.source.filename)
    tc.assertEqual(".txt", plain.source.extension)


def test_read_plain_csv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.csv"
    plain: ExtractedDocument = next(
        read_plain_text(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual('Text; Date\n"Hello World"; "2025-12-25"', plain.full_text)

    tc.assertEqual(
        'Text; Date\n"Hello World"; "2025-12-25"',
        "\n".join(unit.text for unit in plain.units),
    )
    tc.assertEqual(0, len(list(plain.iter_images())))
    tc.assertEqual(0, len(list(plain.iter_tables())))


def test_read_plain_tsv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.tsv"
    plain = next(
        read_plain_text(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Text\tDate\nHello World\t2025-12-25", plain.full_text)
    tc.assertEqual("Text\tDate\nHello World\t2025-12-25", plain.full_text)
    tc.assertEqual(
        "Text\tDate\nHello World\t2025-12-25",
        "\n".join(unit.text for unit in plain.units),
    )
    tc.assertEqual(0, len(list(plain.iter_images())))
    tc.assertEqual(0, len(list(plain.iter_tables())))


def test_read_plain_markdown() -> None:
    path = "sharepoint2text/tests/resources/plain_text/document.md"
    plain = next(
        read_plain_text(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("# Markdown file\n\nThis is a text", plain.full_text)
    tc.assertEqual("# Markdown file\n\nThis is a text", plain.full_text)
    tc.assertEqual(
        "# Markdown file\n\nThis is a text",
        "\n".join(unit.text for unit in plain.units),
    )
    tc.assertEqual(0, len(list(plain.iter_images())))
    tc.assertEqual(0, len(list(plain.iter_tables())))
