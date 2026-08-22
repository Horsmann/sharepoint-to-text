import logging
from unittest import TestCase

from sharepoint2text.parsing.extractors._legacy_types import (
    PlainTextContent,
    PlainUnitMetadata,
)
from sharepoint2text.parsing.extractors.plain_extractor import read_plain_text
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


#########
# Plain #
#########


def test_read_text() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.txt"
    plain: PlainTextContent = next(
        read_plain_text(file_like=read_file_to_file_like(path), path=path)
    )

    tc.assertEqual("Hello World", plain.content)
    tc.assertEqual("Hello World", plain.get_full_text())
    tc.assertEqual(1, len(list(plain.iterate_units())))
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))

    units = list(plain.iterate_units())
    tc.assertTrue(isinstance(units[0].get_metadata(), PlainUnitMetadata))
    tc.assertEqual(PlainUnitMetadata(unit_number=1), units[0].get_metadata())

    meta = plain.get_metadata()
    tc.assertEqual("ascii", meta.detected_encoding)
    tc.assertEqual("plain.txt", meta.filename)
    tc.assertEqual(".txt", meta.file_extension)


def test_read_plain_csv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.csv"
    plain: PlainTextContent = next(
        read_plain_text(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual('Text; Date\n"Hello World"; "2025-12-25"', plain.content)

    tc.assertEqual(
        'Text; Date\n"Hello World"; "2025-12-25"',
        "\n".join(unit.get_text() for unit in plain.iterate_units()),
    )
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))


def test_read_plain_tsv() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.tsv"
    plain = next(
        read_plain_text(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("Text\tDate\nHello World\t2025-12-25", plain.content)
    tc.assertEqual("Text\tDate\nHello World\t2025-12-25", plain.get_full_text())
    tc.assertEqual(
        "Text\tDate\nHello World\t2025-12-25",
        "\n".join(unit.get_text() for unit in plain.iterate_units()),
    )
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))


def test_read_plain_markdown() -> None:
    path = "sharepoint2text/tests/resources/plain_text/document.md"
    plain = next(
        read_plain_text(file_like=read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual("# Markdown file\n\nThis is a text", plain.content)
    tc.assertEqual("# Markdown file\n\nThis is a text", plain.get_full_text())
    tc.assertEqual(
        "# Markdown file\n\nThis is a text",
        "\n".join(unit.get_text() for unit in plain.iterate_units()),
    )
    tc.assertEqual(0, len(list(plain.iterate_images())))
    tc.assertEqual(0, len(list(plain.iterate_tables())))
