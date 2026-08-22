import io
import logging
import typing
import zipfile
from unittest import TestCase

from sharepoint2text.parsing.exceptions import (
    ExtractionFileEncryptedError,
)
from sharepoint2text.parsing.extractors._records import (
    EmailParserOutput,
    EpubParserOutput,
    EpubUnitMetadata,
    HtmlParserOutput,
    HtmlUnitMetadata,
    SourceRecord,
    TableDim,
)
from sharepoint2text.parsing.extractors.archive_extractor import read_archive
from sharepoint2text.parsing.extractors.csv_extractor import read_csv
from sharepoint2text.parsing.extractors.epub_extractor import read_epub
from sharepoint2text.parsing.extractors.html_extractor import read_html
from sharepoint2text.parsing.extractors.mail.mbox_email_extractor import (
    read_mbox_format_mail,
)
from sharepoint2text.parsing.extractors.mhtml_extractor import read_mhtml

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
    meta = SourceRecord()
    meta.populate_from_path("my/dummy/path.txt")

    tc.assertEqual("path.txt", meta.filename)
    tc.assertEqual(".txt", meta.file_extension)
    tc.assertEqual("my/dummy/path.txt", meta.file_path)
    tc.assertEqual("my/dummy", meta.folder_path)

    tc.assertDictEqual(
        {
            "filename": "path.txt",
            "file_extension": ".txt",
            "file_path": "my/dummy/path.txt",
            "folder_path": "my/dummy",
            "detected_encoding": None,
        },
        meta.to_dict(),
    )


def test_password_protected__zip() -> None:
    path = "sharepoint2text/tests/resources/archives/password_protected/sample-password-protected-pw123.zip"
    with tc.assertRaises(ExtractionFileEncryptedError):
        list(read_archive(file_like=_read_file_to_file_like(path=path), path=path))


def test_email__mbox_format() -> None:
    path = "sharepoint2text/tests/resources/mails/basic_email.mbox"

    mail_gen: typing.Generator[EmailParserOutput, None, None] = read_mbox_format_mail(
        file_like=_read_file_to_file_like(path=path),
        path=path,
    )
    mails = list(mail_gen)

    # number of mails
    tc.assertEqual(2, len(mails))

    # 1st mail
    # subject
    tc.assertEqual("Test Email 1", mails[0].subject)
    # body
    tc.assertEqual("This is the body", mails[0].body_plain[:16])
    # sender
    tc.assertEqual("John Doe", mails[0].from_email.name)
    tc.assertEqual("john@example.com", mails[0].from_email.address)

    # receiver
    tc.assertEqual(1, len(mails[0].to_emails))
    tc.assertEqual("Jane Smith", mails[0].to_emails[0].name)
    tc.assertEqual("jane@example.com", mails[0].to_emails[0].address)

    # cc
    tc.assertEqual(0, len(mails[0].to_cc))

    # bcc
    tc.assertEqual(0, len(mails[0].to_bcc))

    # metadata
    mail_meta = mails[0].get_metadata()
    tc.assertEqual("basic_email.mbox", mail_meta.filename)
    tc.assertEqual(".mbox", mail_meta.file_extension)
    tc.assertEqual("2025-12-27T10:00:00+00:00", mail_meta.date)
    tc.assertEqual("<msg001@example.com>", mail_meta.message_id)

    tc.assertEqual(0, len(list(mails[0].iterate_images())))
    tc.assertEqual(0, len(list(mails[0].iterate_tables())))


#########
# Other #
#########


def test_read_html__1() -> None:
    path = "sharepoint2text/tests/resources/html/sample.html"
    html: HtmlParserOutput = next(
        read_html(file_like=_read_file_to_file_like(path=path), path=path)
    )

    full_text = "Welcome on my website\n\n\nParticipants\n\n\nName  | Age\nAlice | 25\nBob   | 30\n\n\nThis is a simple example of an HTML page with a table and links.\n\n\nVisit:\nWikipedia |\nGoogle"
    tc.assertEqual(full_text, html.get_full_text())
    tc.assertListEqual([[["Name", "Age"], ["Alice", "25"], ["Bob", "30"]]], html.tables)
    tc.assertListEqual(
        [
            {"text": "Wikipedia", "href": "https://www.wikipedia.org"},
            {"text": "Google", "href": "https://www.google.com"},
        ],
        html.links,
    )
    tc.assertEqual(0, len(list(html.iterate_images())))
    tc.assertEqual(1, len(list(html.iterate_tables())))
    tc.assertListEqual(
        [["Name", "Age"], ["Alice", "25"], ["Bob", "30"]],
        list(html.iterate_tables())[0].get_table(),
    )

    tc.assertEqual(
        HtmlUnitMetadata(unit_number=1), list(html.iterate_units())[0].get_metadata()
    )


def test_read_html__2() -> None:
    path = "sharepoint2text/tests/resources/html/large_complex.html"
    html: HtmlParserOutput = next(
        read_html(file_like=_read_file_to_file_like(path=path), path=path)
    )

    tc.assertEqual(12, len(list(html.iterate_tables())))
    tables = list(html.iterate_tables())
    tc.assertListEqual(
        [
            ["ID", "Name", "Email", "Department"],
            ["001", "Alice Johnson", "alice@company.com", "Engineering"],
            ["002", "Bob Smith", "bob@company.com", "Marketing"],
            ["003", "Carol Davis", "carol@company.com", "Sales"],
        ],
        tables[0].get_table(),
    )

    tc.assertListEqual(
        [
            ["Property", "Value", "Type"],
            ["Size", "Large", "String"],
            ["Count", "150", "Integer"],
        ],
        tables[1].get_table(),
    )

    tc.assertListEqual(
        [
            ["Component", "Throughput", "Latency", "Error Rate"],
            ["Parser", "1000 docs/sec", "5ms", "0.1%"],
            ["Extractor", "800 docs/sec", "8ms", "0.2%"],
        ],
        tables[2].get_table(),
    )

    tc.assertListEqual(
        [
            ["Category", "Q1", "Q2", "Q3", "Q4", "Total"],
            ["Revenue", "$1.2M", "$1.5M", "$1.8M", "$2.1M", "$6.6M"],
            ["Expenses", "$0.8M", "$0.9M", "$1.0M", "$1.1M", "$3.8M"],
            ["Profit", "$0.4M", "$0.6M", "$0.8M", "$1.0M", "$2.8M"],
        ],
        tables[3].get_table(),
    )

    tc.assertListEqual(
        [
            ["Deep Property", "Value"],
            ["Nesting Level", "6"],
            ["Extraction Complexity", "High"],
        ],
        tables[4].get_table(),
    )

    tc.assertListEqual(
        [
            ["User Type", "Count", "Percentage"],
            ["Active", "15,000", "75%"],
            ["Inactive", "3,000", "15%"],
            ["New", "2,000", "10%"],
        ],
        tables[5].get_table(),
    )

    tc.assertListEqual(
        [
            ["Metric", "Value", "Target", "Status"],
            ["Response Time", "2.1s", "<3.0s", "✓ Good"],
            ["Throughput", "850 req/s", ">800 req/s", "✓ Good"],
            ["Error Rate", "0.15%", "<0.5%", "✓ Excellent"],
        ],
        tables[6].get_table(),
    )

    tc.assertListEqual(
        [
            ["Feature", "Daily Users", "Weekly Users", "Monthly Users"],
            ["Search", "12,000", "45,000", "120,000"],
            ["Export", "3,500", "12,000", "35,000"],
            ["Sharing", "8,000", "28,000", "85,000"],
        ],
        tables[7].get_table(),
    )

    tc.assertListEqual(
        [
            ["Tool", "Link", "Description"],
            ["Parser", "Parser Tool", "HTML parsing utility"],
            ["Extractor", "Extractor Tool", "Content extraction utility"],
        ],
        tables[8].get_table(),
    )

    tc.assertListEqual(
        [
            ["Test ID", "Iterations", "Time (ms)"],
            ["PERF-001", "1000", "2.5"],
            ["PERF-002", "5000", "12.3"],
        ],
        tables[9].get_table(),
    )

    tc.assertListEqual(
        [
            ["Test ID", "Iterations", "Time (ms)"],
            ["PERF-003", "10000", "24.7"],
            ["PERF-004", "50000", "125.1"],
        ],
        tables[10].get_table(),
    )

    tc.assertListEqual(
        [
            ["Parameter", "Value", "Unit"],
            ["Cache Hit Rate", "95.2", "%"],
            ["Traversal Reduction", "94.8", "%"],
            ["Speed Improvement", "12.0", "%"],
        ],
        tables[11].get_table(),
    )


def test_read_epub__1() -> None:
    """Test EPUB extraction with a sample EPUB file."""
    path = "sharepoint2text/tests/resources/epub/sample.epub"
    epub: EpubParserOutput = next(
        read_epub(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Check metadata
    tc.assertEqual("Test EPUB Book", epub.metadata.title)
    tc.assertEqual("Test Author", epub.metadata.creator)
    tc.assertEqual("en", epub.metadata.language)
    tc.assertEqual("Test Publisher", epub.metadata.publisher)
    tc.assertEqual("2024-01-15", epub.metadata.date)
    tc.assertEqual("A test EPUB file for sharepoint-to-text", epub.metadata.description)
    tc.assertEqual("Testing", epub.metadata.subject)
    tc.assertEqual("3.0", epub.metadata.epub_version)

    # Check chapters
    tc.assertEqual(2, len(epub.chapters))

    # Chapter 1
    chapter1 = epub.chapters[0]
    tc.assertEqual(1, chapter1.chapter_number)
    tc.assertIn("Chapter 1: Introduction", chapter1.title)
    tc.assertIn("Welcome to the test EPUB book", chapter1.text)
    tc.assertIn("sample text for extraction testing", chapter1.text)
    tc.assertIn("Section 1.1", chapter1.text)

    # Chapter 1 table
    tc.assertEqual(1, len(chapter1.tables))
    tc.assertListEqual(
        [["Name", "Value"], ["Item A", "100"], ["Item B", "200"]],
        chapter1.tables[0],
    )

    # Chapter 2
    chapter2 = epub.chapters[1]
    tc.assertEqual(2, chapter2.chapter_number)
    tc.assertIn("Chapter 2: Getting Started", chapter2.title)
    tc.assertIn("second chapter", chapter2.text)
    tc.assertIn("First item in the list", chapter2.text)

    # Test iterate_units
    units = list(epub.iterate_units())
    tc.assertEqual(2, len(units))
    tc.assertEqual(
        EpubUnitMetadata(
            unit_number=1, href="OEBPS/chapter1.xhtml", title=chapter1.title
        ),
        units[0].get_metadata(),
    )

    # Test get_full_text
    full_text = epub.get_full_text()
    tc.assertIn("Chapter 1: Introduction", full_text)
    tc.assertIn("Chapter 2: Getting Started", full_text)
    tc.assertIn("Welcome to the test EPUB book", full_text)

    # Test iterate_tables
    tables = list(epub.iterate_tables())
    tc.assertEqual(1, len(tables))
    tc.assertEqual(TableDim(rows=3, columns=2), tables[0].get_dim())

    # Test table of contents
    tc.assertEqual(2, len(epub.toc))
    tc.assertEqual("Chapter 1: Introduction", epub.toc[0]["title"])
    tc.assertEqual("Chapter 2: Getting Started", epub.toc[1]["title"])


def test_read_epub__2() -> None:
    """Test EPUB extraction with a sample EPUB file."""
    path = "sharepoint2text/tests/resources/epub/BJNR274910013.epub"
    epub: EpubParserOutput = next(
        read_epub(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # general
    tc.assertEqual(31, len(epub.chapters))
    tc.assertEqual(3, len(list(epub.iterate_tables())))
    tc.assertEqual("Gesetz zur Förderung der elektronischen", epub.get_full_text()[:39])
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
        list(epub.iterate_tables())[0].get_table(),
    )

    # metadata
    tc.assertEqual("BJNR274910013.epub", epub.get_metadata().filename)
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung "
        "(E-Government-Gesetz - EGovG)",
        epub.get_metadata().title,
    )
    tc.assertEqual("2025-12-05", epub.get_metadata().date)

    # units
    tc.assertEqual(31, len(list(epub.iterate_units())))
    units = list(epub.iterate_units())
    # 0
    tc.assertEqual("", units[0].get_text())
    tc.assertEqual(1, len(list(units[0].get_tables())))
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
        units[0].get_tables()[0].get_table(),
    )
    tc.assertEqual(
        EpubUnitMetadata(
            unit_number=2,
            href="BJNR274910013.html",
            title="Gesetz zur Förderung der elektronischen Verwaltung "
            "(E-Government-Gesetz - EGovG)",
        ),
        units[1].get_metadata(),
    )
    # 1
    tc.assertEqual(
        "Gesetz zur Förderung der elektronischen Verwaltung (E-Government-Gesetz - EGovG)",
        units[1].get_text()[:80],
    )
    # 2
    tc.assertEqual("Inhaltsübersicht", units[2].get_text())
    # 3
    tc.assertEqual("§ 1\n\nGeltungsbereich\n\n(1)", units[3].get_text()[:25])
    # last page
    tc.assertEqual("§ 19\n\nÜbergangsvorschriften", units[-1].get_text()[:27])


def test_read_mhtml() -> None:
    """Test MHTML (web archive) extraction."""
    path = "sharepoint2text/tests/resources/html/sample.mhtml"
    result: HtmlParserOutput = next(
        read_mhtml(file_like=_read_file_to_file_like(path=path), path=path)
    )

    # Verify it returns HtmlParserOutput
    tc.assertIsInstance(result, HtmlParserOutput)

    # Check metadata
    tc.assertEqual("Test MHTML Page", result.metadata.title)

    # Check content extraction
    tc.assertIn("Welcome to the Test Page", result.content)
    tc.assertIn("test MHTML document", result.content)
    tc.assertIn("More Content", result.content)

    # Check table extraction
    tc.assertEqual(1, len(result.tables))
    tc.assertListEqual(
        [["Product", "Price"], ["Widget", "$10.00"], ["Gadget", "$25.00"]],
        result.tables[0],
    )

    # Check link extraction
    tc.assertEqual(1, len(result.links))
    tc.assertEqual("link to example.com", result.links[0]["text"])
    tc.assertEqual("https://example.com", result.links[0]["href"])


def test_read_csv_2():
    path = "sharepoint2text/tests/resources/plain_text/plain.csv"
    results = list(read_csv(file_like=_read_file_to_file_like(path=path), path=path))

    tc.assertListEqual(
        [["Text", "Date"], ["Hello World", "2025-12-25"]],
        list(results[0].iterate_tables())[0].get_table(),
    )
