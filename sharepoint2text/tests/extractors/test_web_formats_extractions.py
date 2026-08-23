import io
import logging
import typing
import zipfile
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors.archive_extractor import read_archive
from sharepoint2text.parsing.extractors.csv_extractor import read_csv
from sharepoint2text.parsing.extractors.epub_extractor import read_epub
from sharepoint2text.parsing.extractors.html_extractor import read_html
from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
    read_mbox_format_mail,
)
from sharepoint2text.parsing.extractors.mhtml_extractor import read_mhtml
from sharepoint2text.parsing.models import ExtractedDocument, SourceMetadata

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def _read_file_to_file_like(path: str) -> io.BytesIO:
    with open(path, mode="rb") as file:
        file_like = io.BytesIO(file.read())
        file_like.seek(0)
        return file_like


def _zip_bytes_to_file_like(files: dict[str, str]) -> io.BytesIO:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, text in files.items():
            zf.writestr(name, text)
    buffer.seek(0)
    return buffer


#############
# Interface #
#############


def test_file_metadata_extraction() -> None:
    meta = SourceMetadata(
        filename="path.txt",
        extension=".txt",
        path="my/dummy/path.txt",
        folder="my/dummy",
    )

    tc.assertEqual("path.txt", meta.filename)
    tc.assertEqual(".txt", meta.extension)
    tc.assertEqual("my/dummy/path.txt", meta.path)
    tc.assertEqual("my/dummy", meta.folder)
    tc.assertIsNone(meta.encoding)


def test_password_protected__zip() -> None:
    path = "sharepoint2text/tests/resources/archives/password_protected/sample-password-protected-pw123.zip"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_archive(file_like=_read_file_to_file_like(path=path), path=path))


def test_email__mbox_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.mbox"

    mail_gen: typing.Generator[ExtractedDocument, None, None] = read_mbox_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    # number of mails
    tc.assertEqual(2, len(mails))

    # 1st mail
    # subject
    tc.assertEqual("Test Email 1", mails[0].metadata.title)
    # body
    tc.assertEqual("This is the body", mails[0].units[0].text[:16])
    # sender
    sender = typing.cast(dict[str, str], mails[0].properties["mbox.from_email"])
    tc.assertEqual("John Doe", sender["name"])
    tc.assertEqual("john@example.com", sender["address"])

    # receiver
    recipients = typing.cast(
        list[dict[str, str]], mails[0].properties["mbox.to_emails"]
    )
    tc.assertEqual(1, len(recipients))
    tc.assertEqual("Jane Smith", recipients[0]["name"])
    tc.assertEqual("jane@example.com", recipients[0]["address"])

    # cc
    tc.assertNotIn("mbox.to_cc", mails[0].properties)

    # bcc
    tc.assertNotIn("mbox.to_bcc", mails[0].properties)

    # metadata
    tc.assertEqual("basic_email.mbox", mails[0].source.filename)
    tc.assertEqual(".mbox", mails[0].source.extension)
    tc.assertEqual("2025-12-27T10:00:00+00:00", mails[0].metadata.created)
    tc.assertEqual(
        "<msg001@example.com>", mails[0].metadata.properties["mbox.message_id"]
    )

    tc.assertEqual(0, len(list(mails[0].iter_images())))
    tc.assertEqual(0, len(list(mails[0].iter_tables())))


#########
# Other #
#########


def test_read_html__1() -> None:
    path = "sharepoint2text/tests/resources/html/sample.html"
    html: ExtractedDocument = next(
        read_html(file_like=_read_file_to_file_like(path=path), path=path)
    )

    full_text = "Welcome on my website\n\n\nParticipants\n\n\nName  | Age\nAlice | 25\nBob   | 30\n\n\nThis is a simple example of an HTML page with a table and links.\n\n\nVisit:\nWikipedia |\nGoogle"
    tc.assertEqual(full_text, html.full_text)
    tc.assertListEqual(
        [["Name", "Age"], ["Alice", "25"], ["Bob", "30"]],
        list(html.iter_tables())[0].rows,
    )
    tc.assertListEqual(
        ["Wikipedia", "Google"],
        [annotation.text for annotation in html.document_annotations],
    )
    tc.assertEqual(0, len(list(html.iter_images())))
    tc.assertEqual(1, len(list(html.iter_tables())))
    tc.assertListEqual(
        [["Name", "Age"], ["Alice", "25"], ["Bob", "30"]],
        list(html.iter_tables())[0].rows,
    )

    tc.assertListEqual(
        ["https://www.wikipedia.org", "https://www.google.com"],
        [annotation.target for annotation in html.document_annotations],
    )
    tc.assertEqual(1, html.units[0].number)
    tc.assertEqual("document", html.units[0].kind)


def test_read_html__2() -> None:
    path = "sharepoint2text/tests/resources/html/large_complex.html"
    html: ExtractedDocument = next(
        read_html(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(12, len(list(html.iter_tables())))
    tables = list(html.iter_tables())
    tc.assertListEqual(
        [
            ["ID", "Name", "Email", "Department"],
            ["001", "Alice Johnson", "alice@company.com", "Engineering"],
            ["002", "Bob Smith", "bob@company.com", "Marketing"],
            ["003", "Carol Davis", "carol@company.com", "Sales"],
        ],
        tables[0].rows,
    )

    tc.assertListEqual(
        [
            ["Property", "Value", "Type"],
            ["Size", "Large", "String"],
            ["Count", "150", "Integer"],
        ],
        tables[1].rows,
    )

    tc.assertListEqual(
        [
            ["Component", "Throughput", "Latency", "Error Rate"],
            ["Parser", "1000 docs/sec", "5ms", "0.1%"],
            ["Extractor", "800 docs/sec", "8ms", "0.2%"],
        ],
        tables[2].rows,
    )

    tc.assertListEqual(
        [
            ["Category", "Q1", "Q2", "Q3", "Q4", "Total"],
            ["Revenue", "$1.2M", "$1.5M", "$1.8M", "$2.1M", "$6.6M"],
            ["Expenses", "$0.8M", "$0.9M", "$1.0M", "$1.1M", "$3.8M"],
            ["Profit", "$0.4M", "$0.6M", "$0.8M", "$1.0M", "$2.8M"],
        ],
        tables[3].rows,
    )

    tc.assertListEqual(
        [
            ["Deep Property", "Value"],
            ["Nesting Level", "6"],
            ["Extraction Complexity", "High"],
        ],
        tables[4].rows,
    )

    tc.assertListEqual(
        [
            ["User Type", "Count", "Percentage"],
            ["Active", "15,000", "75%"],
            ["Inactive", "3,000", "15%"],
            ["New", "2,000", "10%"],
        ],
        tables[5].rows,
    )

    tc.assertListEqual(
        [
            ["Metric", "Value", "Target", "Status"],
            ["Response Time", "2.1s", "<3.0s", "✓ Good"],
            ["Throughput", "850 req/s", ">800 req/s", "✓ Good"],
            ["Error Rate", "0.15%", "<0.5%", "✓ Excellent"],
        ],
        tables[6].rows,
    )

    tc.assertListEqual(
        [
            ["Feature", "Daily Users", "Weekly Users", "Monthly Users"],
            ["Search", "12,000", "45,000", "120,000"],
            ["Export", "3,500", "12,000", "35,000"],
            ["Sharing", "8,000", "28,000", "85,000"],
        ],
        tables[7].rows,
    )

    tc.assertListEqual(
        [
            ["Tool", "Link", "Description"],
            ["Parser", "Parser Tool", "HTML parsing utility"],
            ["Extractor", "Extractor Tool", "Content extraction utility"],
        ],
        tables[8].rows,
    )

    tc.assertListEqual(
        [
            ["Test ID", "Iterations", "Time (ms)"],
            ["PERF-001", "1000", "2.5"],
            ["PERF-002", "5000", "12.3"],
        ],
        tables[9].rows,
    )

    tc.assertListEqual(
        [
            ["Test ID", "Iterations", "Time (ms)"],
            ["PERF-003", "10000", "24.7"],
            ["PERF-004", "50000", "125.1"],
        ],
        tables[10].rows,
    )

    tc.assertListEqual(
        [
            ["Parameter", "Value", "Unit"],
            ["Cache Hit Rate", "95.2", "%"],
            ["Traversal Reduction", "94.8", "%"],
            ["Speed Improvement", "12.0", "%"],
        ],
        tables[11].rows,
    )


def test_read_epub__1() -> None:
    """Test EPUB extraction with a sample EPUB file."""
    path = "sharepoint2text/tests/resources/epub/sample.epub"
    epub: ExtractedDocument = next(
        read_epub(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Check metadata
    tc.assertEqual("Test EPUB Book", epub.metadata.title)
    tc.assertEqual("Test Author", epub.metadata.author)
    tc.assertEqual("en", epub.metadata.language)
    tc.assertEqual("Test Publisher", epub.metadata.properties["epub.publisher"])
    tc.assertEqual("2024-01-15", epub.metadata.created)
    tc.assertEqual(
        "A test EPUB file for sharepoint-to-text",
        epub.metadata.properties["epub.description"],
    )
    tc.assertEqual("Testing", epub.metadata.subject)
    tc.assertEqual("3.0", epub.metadata.properties["epub.epub_version"])

    # Check chapters
    tc.assertEqual(2, len(epub.units))

    # Chapter 1
    chapter1 = epub.units[0]
    tc.assertEqual(1, chapter1.number)
    tc.assertIn("Chapter 1: Introduction", chapter1.title)
    tc.assertIn("Welcome to the test EPUB book", chapter1.text)
    tc.assertIn("sample text for extraction testing", chapter1.text)
    tc.assertIn("Section 1.1", chapter1.text)

    # Chapter 1 table
    tc.assertEqual(1, len(chapter1.tables))
    tc.assertListEqual(
        [["Name", "Value"], ["Item A", "100"], ["Item B", "200"]],
        chapter1.tables[0].rows,
    )

    # Chapter 2
    chapter2 = epub.units[1]
    tc.assertEqual(2, chapter2.number)
    tc.assertIn("Chapter 2: Getting Started", chapter2.title)
    tc.assertIn("second chapter", chapter2.text)
    tc.assertIn("First item in the list", chapter2.text)

    # Test canonical content units
    units = list(epub.units)
    tc.assertEqual(2, len(units))
    tc.assertEqual(1, units[0].number)
    tc.assertEqual("chapter", units[0].kind)
    tc.assertEqual("OEBPS/chapter1.xhtml", units[0].properties["epub.href"])
    tc.assertEqual(chapter1.title, units[0].title)

    # Test canonical aggregate text
    full_text = epub.full_text
    tc.assertIn("Chapter 1: Introduction", full_text)
    tc.assertIn("Chapter 2: Getting Started", full_text)
    tc.assertIn("Welcome to the test EPUB book", full_text)

    # Test canonical table iteration
    tables = list(epub.iter_tables())
    tc.assertEqual(1, len(tables))
    tc.assertEqual((3, 2), tables[0].dimensions)

    # Test table of contents
    toc = typing.cast(list[dict[str, str]], epub.properties["epub.toc"])
    tc.assertEqual(2, len(toc))
    tc.assertEqual("Chapter 1: Introduction", toc[0]["title"])
    tc.assertEqual("Chapter 2: Getting Started", toc[1]["title"])


def test_read_epub__2() -> None:
    """Test EPUB extraction with a sample EPUB file."""
    path = "sharepoint2text/tests/resources/epub/BJNR274910013.epub"
    epub: ExtractedDocument = next(
        read_epub(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # general
    tc.assertEqual(31, len(epub.units))
    tc.assertEqual(3, len(list(epub.iter_tables())))
    tc.assertEqual("Gesetz zur Förderung der elektronischen", epub.full_text[:39])
    tc.assertListEqual(
        [
            ["", ""],
            [
                "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - "
                "EGovG)"
            ],
            [
                "E-Government-Gesetz vom 25. Juli 2013 (BGBl. I S. 2749), das zuletzt durch "
                "Artikel 11 des Gesetzes vom 2. Dezember 2025 (BGBl. 2025 I Nr. 301) "
                "geändert worden ist"
            ],
            [
                "Gesetze im Internet - ePub herausgegeben vom Bundesministerium der Justiz "
                "und für Verbraucherschutz"
            ],
            ["erzeugt am: 05.12.2025"],
        ],
        list(epub.iter_tables())[0].rows,
    )

    # metadata
    tc.assertEqual("BJNR274910013.epub", epub.source.filename)
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung "
        "(E-Government-Gesetz - EGovG)",
        epub.metadata.title,
    )
    tc.assertEqual("2025-12-05", epub.metadata.created)

    # units
    tc.assertEqual(31, len(list(epub.units)))
    units = list(epub.units)
    # 0
    tc.assertEqual("", units[0].text)
    tc.assertEqual(1, len(list(units[0].tables)))
    tc.assertListEqual(
        [
            ["", ""],
            [
                "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - "
                "EGovG)"
            ],
            [
                "E-Government-Gesetz vom 25. Juli 2013 (BGBl. I S. 2749), das zuletzt durch "
                "Artikel 11 des Gesetzes vom 2. Dezember 2025 (BGBl. 2025 I Nr. 301) "
                "geändert worden ist"
            ],
            [
                "Gesetze im Internet - ePub herausgegeben vom Bundesministerium der Justiz "
                "und für Verbraucherschutz"
            ],
            ["erzeugt am: 05.12.2025"],
        ],
        units[0].tables[0].rows,
    )
    tc.assertEqual(2, units[1].number)
    tc.assertEqual("BJNR274910013.html", units[1].properties["epub.href"])
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung "
        "(E-Government-Gesetz - EGovG)",
        units[1].title,
    )
    # 1
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - EGovG)",
        units[1].text[:80],
    )
    # 2
    tc.assertEqual("Inhaltsübersicht", units[2].text)
    # 3
    tc.assertEqual("§ 1\n\nGeltungsbereich\n\n(1)", units[3].text[:25])
    # last page
    tc.assertEqual("§ 19\n\nÜbergangsvorschriften", units[-1].text[:27])


def test_read_mhtml() -> None:
    """Test MHTML (web archive) extraction."""
    path = "sharepoint2text/tests/resources/html/sample.mhtml"
    result: ExtractedDocument = next(
        read_mhtml(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Verify it returns ExtractedDocument
    tc.assertIsInstance(result, ExtractedDocument)

    # Check metadata
    tc.assertEqual("Test MHTML Page", result.metadata.title)

    # Check content extraction
    tc.assertIn("Welcome to the Test Page", result.full_text)
    tc.assertIn("test MHTML document", result.full_text)
    tc.assertIn("More Content", result.full_text)

    # Check table extraction
    tc.assertEqual(1, len(list(result.iter_tables())))
    tc.assertListEqual(
        [["Product", "Price"], ["Widget", "$10.00"], ["Gadget", "$25.00"]],
        list(result.iter_tables())[0].rows,
    )

    # Check link extraction
    tc.assertEqual(1, len(result.document_annotations))
    tc.assertEqual("link to example.com", result.document_annotations[0].text)
    tc.assertEqual("https://example.com", result.document_annotations[0].target)


def test_read_csv_2() -> None:
    path = "sharepoint2text/tests/resources/plain_text/plain.csv"
    results = list(read_csv(file_like=_read_file_to_file_like(path=path), path=path))

    tc.assertListEqual(
        [["Text", "Date"], ["Hello World", "2025-12-25"]],
        list(results[0].iter_tables())[0].rows,
    )
