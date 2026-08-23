"""Tests for the read_many() batch extraction function."""

import glob
import logging
import tempfile
import unittest
from pathlib import Path
from typing import Any, Iterator, cast

import pytest

from sharepoint2text import (
    BatchFileResult,
    ExtractedDocument,
    InvalidConfigurationError,
    read_many,
)

tc = unittest.TestCase()
tc.maxDiff = None

# Path to test resources
RESOURCES_PATH = Path("sharepoint2text/tests/resources")


def _is_normalized_document(obj: Any) -> bool:
    """Return whether an object is a normalized public extraction result."""
    return isinstance(obj, ExtractedDocument)


def test_read_many_with_specific_suffixes() -> None:
    """read_many should extract only files matching the specified suffixes."""
    # Use the plain_text folder which has .txt files
    results = list(
        read_many(
            RESOURCES_PATH / "plain_text",
            suffixes=[".txt"],
            recursive=True,
        )
    )

    tc.assertGreater(len(results), 0)
    for result in results:
        tc.assertTrue(
            _is_normalized_document(result),
            f"Result should be an ExtractedDocument: {type(result)}",
        )
        tc.assertTrue(
            result.source.path is not None and result.source.path.endswith(".txt"),
            f"Expected .txt file, got: {result.source.path}",
        )


def test_read_many_with_multiple_suffixes() -> None:
    """read_many should handle multiple suffixes."""
    results = list(
        read_many(
            RESOURCES_PATH / "modern_ms",
            suffixes=[".docx", ".xlsx"],
            recursive=True,
        )
    )

    tc.assertGreater(len(results), 0)
    for result in results:
        file_path = result.source.path or ""
        tc.assertTrue(
            file_path.endswith(".docx") or file_path.endswith(".xlsx"),
            f"Expected .docx or .xlsx file, got: {file_path}",
        )


def test_read_many_matches_compound_archive_suffix() -> None:
    """Match a compressed TAR when its complete compound suffix is requested."""
    results = list(
        read_many(
            RESOURCES_PATH / "archives",
            suffixes=[".tar.gz"],
            recursive=False,
        )
    )

    tc.assertEqual(1, len(results))
    tc.assertIn("test_archive.tar.gz!/", results[0].source.path or "")


def test_read_many_extract_all_supported() -> None:
    """read_many with extract_all_supported should extract all supported files."""
    results = list(
        read_many(
            RESOURCES_PATH / "plain_text",
            extract_all_supported=True,
            recursive=True,
        )
    )

    tc.assertGreater(len(results), 0)
    for result in results:
        tc.assertTrue(
            _is_normalized_document(result),
            f"Result should be an ExtractedDocument: {type(result)}",
        )


def test_read_many_invalid_configuration_both_options() -> None:
    """read_many should raise InvalidConfigurationError when both options are set."""
    with tc.assertRaises(InvalidConfigurationError) as ctx:
        read_many(
            RESOURCES_PATH,
            suffixes=[".txt"],
            extract_all_supported=True,
        )

    tc.assertIn("Cannot specify both", str(ctx.exception))


def test_read_many_invalid_configuration_no_options() -> None:
    """read_many should raise ValueError when neither option is set."""
    with tc.assertRaises(ValueError) as ctx:
        read_many(RESOURCES_PATH)

    tc.assertIn("Must specify either", str(ctx.exception))


def test_read_many_rejects_invalid_result_callback_eagerly() -> None:
    """Reject a non-callable reporting callback when the API is called."""
    with pytest.raises(TypeError, match="on_file_result must be callable or None"):
        read_many(
            RESOURCES_PATH,
            suffixes=[".txt"],
            on_file_result=cast(Any, object()),
        )


def test_read_many_nonexistent_folder() -> None:
    """read_many should raise FileNotFoundError for non-existent folder."""
    with tc.assertRaises(FileNotFoundError):
        read_many("/nonexistent/path/to/folder", suffixes=[".txt"])


def test_read_many_file_instead_of_folder() -> None:
    """read_many should raise NotADirectoryError when path is a file."""
    # Use an actual file path
    file_path = RESOURCES_PATH / "plain_text" / "lorem_ipsum.txt"
    if file_path.exists():
        with tc.assertRaises(NotADirectoryError):
            read_many(file_path, suffixes=[".txt"])


def test_read_many_non_recursive() -> None:
    """read_many with recursive=False should only look in the top folder."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Create files at root level
        (tmpdir_path / "root.txt").write_text("root content")

        # Create files in subdirectory
        subdir = tmpdir_path / "subdir"
        subdir.mkdir()
        (subdir / "nested.txt").write_text("nested content")

        # Non-recursive should only find root.txt
        results_non_recursive = list(
            read_many(tmpdir_path, suffixes=[".txt"], recursive=False)
        )
        tc.assertEqual(len(results_non_recursive), 1)
        tc.assertIn("root content", results_non_recursive[0].full_text)

        # Recursive should find both
        results_recursive = list(
            read_many(tmpdir_path, suffixes=[".txt"], recursive=True)
        )
        tc.assertEqual(len(results_recursive), 2)


def test_read_many_suffix_normalization() -> None:
    """read_many should normalize suffixes (with/without leading dot)."""
    # Test with suffix without leading dot
    results_without_dot = list(
        read_many(
            RESOURCES_PATH / "plain_text",
            suffixes=["txt"],  # No leading dot
            recursive=True,
        )
    )

    # Test with suffix with leading dot
    results_with_dot = list(
        read_many(
            RESOURCES_PATH / "plain_text",
            suffixes=[".txt"],  # With leading dot
            recursive=True,
        )
    )

    # Should find the same files
    tc.assertEqual(len(results_without_dot), len(results_with_dot))


def test_read_many_empty_folder() -> None:
    """read_many should handle empty folders gracefully."""
    with tempfile.TemporaryDirectory() as tmpdir:
        results = list(read_many(tmpdir, suffixes=[".txt"]))
        tc.assertEqual(0, len(results))


def test_read_many_logs_only_batch_lifecycle_at_info(
    tmp_path: Path, caplog: pytest.LogCaptureFixture
) -> None:
    """Keep INFO output bounded to one batch start and one batch summary."""
    (tmp_path / "first.txt").write_text("first")
    (tmp_path / "second.txt").write_text("second")

    with caplog.at_level(logging.INFO, logger="sharepoint2text"):
        results = list(read_many(tmp_path, suffixes=[".txt"]))

    info_messages = [
        record.getMessage()
        for record in caplog.records
        if record.levelno == logging.INFO and record.name.startswith("sharepoint2text")
    ]

    assert len(results) == 2
    assert len(info_messages) == 2
    assert info_messages[0].startswith("Starting batch extraction")
    assert info_messages[1].endswith(
        "files_found=2, documents_extracted=2, files_skipped=0"
    )


def test_read_many_no_matching_files() -> None:
    """read_many should return empty when no files match the suffixes."""
    results = list(
        read_many(
            RESOURCES_PATH / "plain_text",
            suffixes=[".nonexistent_extension"],
            recursive=True,
        )
    )
    tc.assertEqual(0, len(results))


def test_read_many_ignores_unsupported_in_extract_all_mode() -> None:
    """read_many with extract_all_supported should skip unsupported files."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Create a supported file
        (tmpdir_path / "supported.txt").write_text("supported content")

        # Create an unsupported file
        (tmpdir_path / "unsupported.xyz").write_text("unsupported content")

        results = list(read_many(tmpdir_path, extract_all_supported=True))

        # Should only extract the .txt file
        tc.assertEqual(len(results), 1)
        tc.assertIn("supported content", results[0].full_text)


def test_read_many_force_plain_text_extracts_unknown_extensions() -> None:
    """Forced plain-text mode should include files with unknown extensions."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        (tmpdir_path / "unknown.xyz").write_text("unknown plain-text content")

        results = list(
            read_many(
                tmpdir_path,
                extract_all_supported=True,
                force_plain_text=True,
            )
        )

        tc.assertEqual(1, len(results))
        tc.assertEqual("unknown plain-text content", results[0].full_text)


def test_read_many_with_path_object() -> None:
    """read_many should accept both string and Path objects."""
    # Test with Path object
    results_path = list(
        read_many(
            Path(RESOURCES_PATH / "plain_text"),
            suffixes=[".txt"],
            recursive=True,
        )
    )

    # Test with string
    results_str = list(
        read_many(
            str(RESOURCES_PATH / "plain_text"),
            suffixes=[".txt"],
            recursive=True,
        )
    )

    tc.assertEqual(len(results_path), len(results_str))


def test_read_many_continues_on_extraction_error(monkeypatch: Any) -> None:
    """read_many should continue processing other files when one fails."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Create two valid text files
        (tmpdir_path / "file1.txt").write_text("content 1")
        (tmpdir_path / "file2.txt").write_text("content 2")

        # Patch read_file to fail on file1.txt but succeed on file2.txt
        original_read_file = read_many.__globals__["read_file"]
        call_count = {"value": 0}

        def patched_read_file(path: Any, **kwargs: Any) -> Any:
            call_count["value"] += 1
            if "file1.txt" in str(path):
                raise OSError("Simulated IO error")
            return original_read_file(path, **kwargs)

        monkeypatch.setattr(
            "sharepoint2text._api.read_file",
            patched_read_file,
        )

        results = list(read_many(tmpdir_path, suffixes=[".txt"]))

        # Should have processed at least one file successfully
        tc.assertGreaterEqual(len(results), 1)


def test_read_many_reports_each_selected_file(tmp_path: Path) -> None:
    """Report one successful structured result per selected file."""
    (tmp_path / "first.txt").write_text("first")
    (tmp_path / "second.txt").write_text("second")
    reports: list[BatchFileResult] = []

    documents = list(
        read_many(
            tmp_path,
            suffixes=[".txt"],
            on_file_result=reports.append,
        )
    )

    assert len(reports) == 2
    assert {report.path.name for report in reports} == {"first.txt", "second.txt"}
    assert all(report.succeeded for report in reports)
    assert all(report.error is None for report in reports)
    assert sum(report.documents_extracted for report in reports) == len(documents)


def test_read_many_reports_recoverable_failure_and_continues(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Report a recoverable file error while continuing later extraction."""
    (tmp_path / "failed.txt").write_text("failed")
    (tmp_path / "successful.txt").write_text("successful")
    original_read_file = read_many.__globals__["read_file"]
    reports: list[BatchFileResult] = []

    def failing_read_file(path: Path, **kwargs: Any) -> Iterator[ExtractedDocument]:
        """Fail one selected file and delegate all others."""
        if path.name == "failed.txt":
            raise OSError("simulated failure")
        return original_read_file(path, **kwargs)

    monkeypatch.setattr("sharepoint2text._api.read_file", failing_read_file)

    documents = list(
        read_many(
            tmp_path,
            suffixes=[".txt"],
            on_file_result=reports.append,
        )
    )

    reports_by_name = {report.path.name: report for report in reports}
    assert len(documents) == 1
    assert reports_by_name["successful.txt"].succeeded
    assert not reports_by_name["failed.txt"].succeeded
    assert isinstance(reports_by_name["failed.txt"].error, OSError)


def test_read_many_reports_partial_documents_before_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Include documents emitted before a selected file fails."""
    (tmp_path / "partial.txt").write_text("partial")
    reports: list[BatchFileResult] = []

    def partially_failing_read_file(
        path: Path, **kwargs: Any
    ) -> Iterator[ExtractedDocument]:
        """Yield one document before raising a recoverable error."""
        del path, kwargs
        yield ExtractedDocument(format="txt")
        raise OSError("failure after document")

    monkeypatch.setattr(
        "sharepoint2text._api.read_file",
        partially_failing_read_file,
    )

    documents = list(
        read_many(
            tmp_path,
            suffixes=[".txt"],
            on_file_result=reports.append,
        )
    )

    assert len(documents) == 1
    assert len(reports) == 1
    assert reports[0].documents_extracted == 1
    assert isinstance(reports[0].error, OSError)


def test_read_many_propagates_callback_failure(tmp_path: Path) -> None:
    """Stop batch iteration when the reporting callback raises."""
    (tmp_path / "document.txt").write_text("content")

    def reject_result(result: BatchFileResult) -> None:
        """Reject the completed per-file result."""
        del result
        raise RuntimeError("callback failed")

    with pytest.raises(RuntimeError, match="callback failed"):
        list(
            read_many(
                tmp_path,
                suffixes=[".txt"],
                on_file_result=reject_result,
            )
        )


def test_read_many_reports_lazily_without_accumulation(tmp_path: Path) -> None:
    """Deliver reports as files complete without retaining a batch report."""
    (tmp_path / "first.txt").write_text("first")
    (tmp_path / "second.txt").write_text("second")
    reports: list[BatchFileResult] = []
    documents = read_many(
        tmp_path,
        suffixes=[".txt"],
        on_file_result=reports.append,
    )

    next(documents)
    assert reports == []

    next(documents)
    assert len(reports) == 1

    assert list(documents) == []
    assert len(reports) == 2


def test_read_many_case_insensitive_suffix_matching() -> None:
    """read_many should match suffixes case-insensitively."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Create files with different case extensions
        (tmpdir_path / "lower.txt").write_text("lower")
        (tmpdir_path / "upper.TXT").write_text("upper")
        (tmpdir_path / "mixed.TxT").write_text("mixed")

        results = list(read_many(tmpdir_path, suffixes=[".txt"]))

        # Should find all three files
        tc.assertEqual(3, len(results))


def test_read_many_enumerates_folder_lazily(monkeypatch: Any) -> None:
    """Use lazy glob iteration instead of materializing every folder path."""
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        (tmpdir_path / "document.txt").write_text("content")

        def reject_eager_glob(*args: Any, **kwargs: Any) -> list[str]:
            """Fail if read_many calls the eager glob implementation."""
            raise AssertionError("read_many must not materialize glob results")

        traversal_started = False
        original_iglob = glob.iglob

        def recording_iglob(*args: Any, **kwargs: Any) -> Iterator[str]:
            """Record when iteration requests filesystem traversal."""
            nonlocal traversal_started
            traversal_started = True
            return original_iglob(*args, **kwargs)

        monkeypatch.setattr("glob.glob", reject_eager_glob)
        monkeypatch.setattr("glob.iglob", recording_iglob)

        documents = read_many(tmpdir_path, suffixes=[".txt"])
        tc.assertFalse(traversal_started)

        results = list(documents)

    tc.assertTrue(traversal_started)
    tc.assertEqual(1, len(results))
    tc.assertEqual("content", results[0].full_text)


def test_read_many_with_ignore_images() -> None:
    """Retain image metadata but omit bytes when read_many ignores images."""
    results = list(
        read_many(
            RESOURCES_PATH / "modern_ms",
            suffixes=[".docx"],
            ignore_images=True,
            recursive=True,
        )
    )

    # Should still extract files
    tc.assertGreater(len(results), 0)

    images = [image for result in results for image in result.iter_images()]
    tc.assertGreater(len(images), 0)
    tc.assertTrue(all(image.data is None for image in images))
    for image in images:
        if image.width is not None and image.height is not None:
            tc.assertEqual(image.width / image.height, image.ratio)
