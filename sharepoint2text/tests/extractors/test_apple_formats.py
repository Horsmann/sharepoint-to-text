from unittest import TestCase

from sharepoint2text.parsing.extractors.apple.pages_extractor import (
    DocumentFlow,
    _infer_outline_levels,
    align_document_flow_to_tables,
    read_apple_pages,
)
from sharepoint2text.parsing.extractors.data_types import ApplePagesContent
from sharepoint2text.tests.extractors.utils import read_file_to_file_like

tc = TestCase()
tc.maxDiff = None


def test_apple_pages_1():
    """Test apple pages extractor."""

    path = "sharepoint2text/tests/resources/apple/mwe.pages"

    page_obj: ApplePagesContent = next(
        read_apple_pages(read_file_to_file_like(path=path))
    )

    tc.assertEqual(0, len(page_obj.tables))
    tc.assertEqual("This is a test document.", page_obj.get_full_text())


def test_apple_pages_2():
    """Test apple pages extractor."""

    path = "sharepoint2text/tests/resources/apple/with_tables_image.pages"

    page_obj: ApplePagesContent = next(
        read_apple_pages(read_file_to_file_like(path=path))
    )

    tc.assertEqual(
        "\n".join(
            (
                "This is a test document. A new journey!",
                "",
                "A | B | C | D | Z",
                "1 | 2 | 3 | 4 | ",
                "5 | 6 | 7 | 8 | ",
                "9 | 10 | 11 | 12 | Ü",
                "13 | 14 | 15 | 16 | ",
                "",
                "A | B",
                "John | long",
                "White | Gray",
                "Red | Blue",
                "8 | 9",
                "",
                "Headerline",
                "^^ okay, das ist jetzt mal etwas mehr Text :P",
            )
        ),
        page_obj.get_full_text(),
    )

    tc.assertEqual(1, len(list(page_obj.iterate_images())))
    tc.assertEqual(
        71354, len(list(page_obj.iterate_images())[0].get_bytes().getvalue())
    )
    tc.assertEqual("Space Image", list(page_obj.iterate_images())[0].get_caption())

    tc.assertEqual(2, len(page_obj.tables))
    tc.assertEqual(
        [
            ["A", "B", "C", "D", "Z"],
            ["1", "2", "3", "4", ""],
            ["5", "6", "7", "8", ""],
            ["9", "10", "11", "12", "Ü"],
            ["13", "14", "15", "16", ""],
        ],
        list(page_obj.iterate_tables())[0],
    )
    tc.assertEqual(
        [
            ["A", "B"],
            ["John", "long"],
            ["White", "Gray"],
            ["Red", "Blue"],
            ["8", "9"],
            ["", ""],
        ],
        list(page_obj.iterate_tables())[1],
    )


def test_apple_pages_3():
    """Test apple pages extractor."""

    path = "sharepoint2text/tests/resources/apple/pages_text_only.pages"
    expected_text = "\n".join(
        (
            "My Title",
            "",
            "Chapter 1",
            "",
            "Paragraph 1",
            "",
            "The document outlines a series of standard procedures that are to be followed during routine operational checks. Each procedure has been reviewed and approved by the relevant oversight committee and is intended to ensure consistency across all applicable scenarios. Staff members are expected to familiarize themselves with these procedures and apply them as described, without deviation unless otherwise instructed. Any discrepancies observed during implementation should be recorded in the appropriate log and submitted for further evaluation at the end of the reporting period.",
            "",
            "Paragraph 2",
            "",
            "In addition to the procedural guidelines, the document includes a summary of general expectations regarding documentation and communication. All entries must be completed in a clear and legible manner, using the designated formats provided in the appendix. Communication between departments should adhere to the established channels to avoid unnecessary delays or misunderstandings. Periodic reviews will be conducted to confirm compliance with these expectations, and any required adjustments will be communicated through standard administrative updates.",
        )
    )
    expected_markdown = "\n".join(
        (
            "# My Title",
            "",
            "## Chapter 1",
            "",
            "### Paragraph 1",
            "",
            "The document outlines a series of standard procedures that are to be followed during routine operational checks. Each procedure has been reviewed and approved by the relevant oversight committee and is intended to ensure consistency across all applicable scenarios. Staff members are expected to familiarize themselves with these procedures and apply them as described, without deviation unless otherwise instructed. Any discrepancies observed during implementation should be recorded in the appropriate log and submitted for further evaluation at the end of the reporting period.",
            "",
            "### Paragraph 2",
            "",
            "In addition to the procedural guidelines, the document includes a summary of general expectations regarding documentation and communication. All entries must be completed in a clear and legible manner, using the designated formats provided in the appendix. Communication between departments should adhere to the established channels to avoid unnecessary delays or misunderstandings. Periodic reviews will be conducted to confirm compliance with these expectations, and any required adjustments will be communicated through standard administrative updates.",
        )
    )

    page_obj: ApplePagesContent = next(
        read_apple_pages(read_file_to_file_like(path=path))
    )

    tc.assertEqual(expected_text, page_obj.get_full_text())
    tc.assertEqual(expected_markdown, page_obj.get_full_markdown())


def test_apple_pages_heading_inference_reuses_family_across_style_mismatch():
    """Keep repeated heading families stable even when Pages style ids drift."""

    levels = _infer_outline_levels(
        [
            ("My Title", 1),
            ("Chapter 1", 2),
            ("Paragraph 1", 3),
            (
                "The document outlines a series of standard procedures that are to be "
                "followed during routine operational checks and reporting windows.",
                None,
            ),
            ("Paragraph 2", 2),
            (
                "In addition to the procedural guidelines, the document includes a "
                "summary of general expectations regarding documentation.",
                None,
            ),
        ]
    )

    tc.assertEqual([1, 2, 3, None, 3, None], levels)


def test_apple_pages_heading_inference_ignores_unstyled_short_labels_after_body():
    """Avoid promoting arbitrary short labels to headings once body prose has started."""

    levels = _infer_outline_levels(
        [
            ("My Title", 1),
            (
                "This introductory paragraph explains the purpose of the document and "
                "sets expectations for the remaining sections.",
                None,
            ),
            ("Note", None),
            (
                "This follow-up paragraph continues the prose and should remain body "
                "content instead of becoming a heading.",
                None,
            ),
        ]
    )

    tc.assertEqual([1, None, None, None], levels)


def test_apple_pages_heading_inference_supports_unstyled_front_matter():
    """Infer a simple heading ladder from unstyled front matter before body prose."""

    levels = _infer_outline_levels(
        [
            ("Project Atlas", None),
            ("Section 1", None),
            (
                "This opening section contains enough descriptive prose to count as "
                "body text and anchor later family-based heading reuse.",
                None,
            ),
            ("Section 2", None),
            (
                "This later section should inherit the same heading depth from the "
                "normalized family key instead of becoming body text.",
                None,
            ),
        ]
    )

    tc.assertEqual([1, 2, None, 2, None], levels)


def test_apple_pages_flow_alignment_uses_paragraph_evidence():
    """Prefer merges that recreate known paragraph text when placeholders exceed tables."""

    flow = DocumentFlow(
        segments=[
            "This is a test document. A new",
            "journey!",
            "",
            "Headerline\n^^ okay, das ist jetzt mal etwas mehr Text :P",
        ],
        placeholder_count=3,
        source_object_id=1732514,
    )

    aligned = align_document_flow_to_tables(
        flow,
        2,
        [
            "This is a test document. A new journey!",
            "Headerline\n^^ okay, das ist jetzt mal etwas mehr Text :P",
        ],
    )

    tc.assertEqual(
        DocumentFlow(
            segments=[
                "This is a test document. A new journey!",
                "",
                "Headerline\n^^ okay, das ist jetzt mal etwas mehr Text :P",
            ],
            placeholder_count=2,
            source_object_id=1732514,
        ),
        aligned,
    )


def test_apple_pages_flow_alignment_keeps_ambiguous_boundaries():
    """Avoid collapsing extra placeholders when no merge looks like real prose."""

    flow = DocumentFlow(
        segments=["Overview.", "Appendix", "Summary"],
        placeholder_count=2,
        source_object_id=42,
    )

    tc.assertEqual(flow, align_document_flow_to_tables(flow, 1))
