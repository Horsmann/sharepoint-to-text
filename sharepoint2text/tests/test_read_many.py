"""Tests for the read_many() batch extraction function."""

import tempfile
import unittest
from pathlib import Path
from typing import Any

from sharepoint2text import (
    InvalidConfigurationError,
    read_many,
)

tc = unittest.TestCase()
tc.maxDiff = None

# Path to test resources
RESOURCES_PATH = Path("sharepoint2text/tests/resources")


def _has_extraction_interface(obj: Any) -> bool:
    """Check if an object implements the ExtractionInterface protocol."""
    return (
        hasattr(obj, "get_metadata")
        and hasattr(obj, "get_full_text")
        and hasattr(obj, "iterate_units")
        and callable(obj.get_metadata)
        and callable(obj.get_full_text)
    )


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
            _has_extraction_interface(result),
            f"Result should implement ExtractionInterface: {type(result)}",
        )
        metadata = result.get_metadata()
        tc.assertTrue(
            metadata.file_path is not None and metadata.file_path.endswith(".txt"),
            f"Expected .txt file, got: {metadata.file_path}",
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
        metadata = result.get_metadata()
        file_path = metadata.file_path or ""
        tc.assertTrue(
            file_path.endswith(".docx") or file_path.endswith(".xlsx"),
            f"Expected .docx or .xlsx file, got: {file_path}",
        )


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
            _has_extraction_interface(result),
            f"Result should implement ExtractionInterface: {type(result)}",
        )


def test_read_many_invalid_configuration_both_options() -> None:
    """read_many should raise InvalidConfigurationError when both options are set."""
    with tc.assertRaises(InvalidConfigurationError) as ctx:
        list(
            read_many(
                RESOURCES_PATH,
                suffixes=[".txt"],
                extract_all_supported=True,
            )
        )

    tc.assertIn("Cannot specify both", str(ctx.exception))


def test_read_many_invalid_configuration_no_options() -> None:
    """read_many should raise ValueError when neither option is set."""
    with tc.assertRaises(ValueError) as ctx:
        list(read_many(RESOURCES_PATH))

    tc.assertIn("Must specify either", str(ctx.exception))


def test_read_many_nonexistent_folder() -> None:
    """read_many should raise FileNotFoundError for non-existent folder."""
    with tc.assertRaises(FileNotFoundError):
        list(read_many("/nonexistent/path/to/folder", suffixes=[".txt"]))


def test_read_many_file_instead_of_folder() -> None:
    """read_many should raise NotADirectoryError when path is a file."""
    # Use an actual file path
    file_path = RESOURCES_PATH / "plain_text" / "lorem_ipsum.txt"
    if file_path.exists():
        with tc.assertRaises(NotADirectoryError):
            list(read_many(file_path, suffixes=[".txt"]))


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
        tc.assertIn("root content", results_non_recursive[0].get_full_text())

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
        tc.assertIn("supported content", results[0].get_full_text())


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
        tc.assertEqual("unknown plain-text content", results[0].get_full_text())


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
            "sharepoint2text.read_file",
            patched_read_file,
        )

        results = list(read_many(tmpdir_path, suffixes=[".txt"]))

        # Should have processed at least one file successfully
        tc.assertGreaterEqual(len(results), 1)


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


def test_read_many_with_ignore_images() -> None:
    """read_many should pass ignore_images flag to extractors."""
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

    # Check that no images were extracted
    for result in results:
        images = list(result.iterate_images())
        tc.assertEqual(len(images), 0, "Images should be ignored")
