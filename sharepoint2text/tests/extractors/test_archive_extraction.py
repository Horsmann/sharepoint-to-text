import io as std_io
import logging
import zipfile
from unittest import TestCase

import sharepoint2text.parsing.extractors.archive_extractor as archive_module
from sharepoint2text.parsing.exceptions import (
    ExtractionFailedError,
    ExtractionFileTooLargeError,
    ExtractionZipBombError,
)
from sharepoint2text.parsing.extractors.archive_extractor import read_archive
from sharepoint2text.parsing.extractors.data_types import (
    EpubContent,
    PlainTextContent,
)
from sharepoint2text.tests.extractors.utils import (
    read_file_to_file_like,
    tar_bytes_to_file_like,
)

logger = logging.getLogger(__name__)

tc = TestCase()
tc.maxDiff = None


def test_read_zip_archive_1() -> None:
    """Test ZIP archive extraction with multiple supported files."""
    path = "sharepoint2text/tests/resources/archives/test_archive.zip"
    results = list(read_archive(file_like=read_file_to_file_like(path=path), path=path))

    # Should extract 2 text files from the archive
    tc.assertEqual(2, len(results))

    # All results should be PlainTextContent
    for result in results:
        tc.assertIsInstance(result, PlainTextContent)

    # Check that we got the expected content
    texts = [r.get_full_text() for r in results]
    tc.assertTrue(any("This is a test document" in t for t in texts))
    tc.assertTrue(any("Another file in the archive" in t for t in texts))

    # Check that metadata includes archive path
    for result in results:
        tc.assertIn("test_archive.zip!/", result.get_metadata().file_path)


def test_read_zip_archive_2() -> None:
    """Test ZIP archive extraction with multiple supported files."""

    # three files - of which two are supported
    path = "sharepoint2text/tests/resources/archives/sample.zip"
    results = list(read_archive(file_like=read_file_to_file_like(path=path), path=path))
    tc.assertEqual(2, len(results))
    tc.assertTrue(isinstance(results[0], PlainTextContent))
    tc.assertTrue(isinstance(results[1], EpubContent))


def test_read_zip_archive_rejects_zip_bomb_ratio() -> None:
    """ZIP archives with extreme compression ratio should be rejected."""
    zip_buffer = std_io.BytesIO()
    with zipfile.ZipFile(
        zip_buffer,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as zf:
        zf.writestr("bomb.txt", b"A" * 5_000_000)
    zip_buffer.seek(0)

    with tc.assertRaises(ExtractionZipBombError):
        list(read_archive(zip_buffer, path="bomb.zip"))


def test_read_tar_archive() -> None:
    """Test TAR archive extraction."""
    path = "sharepoint2text/tests/resources/archives/test_archive.tar"
    results = list(read_archive(file_like=read_file_to_file_like(path=path), path=path))

    # Should extract 2 text files from the archive
    tc.assertEqual(2, len(results))

    # All results should be PlainTextContent
    for result in results:
        tc.assertIsInstance(result, PlainTextContent)

    # Check that we got the expected content
    texts = [r.get_full_text() for r in results]
    tc.assertTrue(any("This is a test document" in t for t in texts))
    tc.assertTrue(any("Another file in the tar archive" in t for t in texts))


def test_tar_archive_entry_limit() -> None:
    """TAR archives exceeding configured entry limit should fail fast."""
    import sharepoint2text.parsing.extractors.archive_extractor as archive_module

    original_limit = archive_module.MAX_TAR_ENTRIES
    archive_module.MAX_TAR_ENTRIES = 1

    try:
        tar_buffer = tar_bytes_to_file_like({"a.txt": b"a", "b.txt": b"b"})
        with tc.assertRaises(ExtractionFailedError) as cm:
            list(read_archive(tar_buffer, path="test.tar"))
        tc.assertIn("too many entries", str(cm.exception))
    finally:
        archive_module.MAX_TAR_ENTRIES = original_limit


def test_read_7zip_archive() -> None:
    """Test TAR archive extraction."""
    path = "sharepoint2text/tests/resources/archives/test_archive.7z"
    results = list(read_archive(file_like=read_file_to_file_like(path=path), path=path))

    tc.assertEqual(2, len(results))
    tc.assertTrue(isinstance(results[0], PlainTextContent))
    tc.assertTrue(isinstance(results[1], EpubContent))


def test_7zip_file_size_limit() -> None:
    """Test that 7z archives exceeding size limit raise appropriate exception."""
    test_max_size = 1024  # 1KB for testing

    original_max_7z_file_size = archive_module.MAX_7Z_FILE_SIZE
    archive_module.MAX_7Z_FILE_SIZE = test_max_size

    try:
        path = "sharepoint2text/tests/resources/archives/test_archive.7z"
        # This should raise ExtractionFileTooLargeError
        with tc.assertRaises(ExtractionFileTooLargeError) as cm:
            list(read_archive(file_like=read_file_to_file_like(path=path), path=path))

        # Verify the exception details
        error = cm.exception
        tc.assertEqual(test_max_size, error.max_size)
        tc.assertGreater(error.actual_size, test_max_size)
        tc.assertIn("exceeds maximum allowed size", str(error))

    finally:
        # Restore original size limit
        archive_module.MAX_7Z_FILE_SIZE = original_max_7z_file_size


def test_7zip_total_uncompressed_size_limit() -> None:
    """7z archives should be rejected when total uncompressed bytes exceed limit."""
    import sharepoint2text.parsing.extractors.archive_extractor as archive_module

    original_total_limit = archive_module.MAX_7Z_MEMORY_USAGE
    archive_module.MAX_7Z_MEMORY_USAGE = 1

    try:
        path = "sharepoint2text/tests/resources/archives/test_archive.7z"
        with tc.assertRaises(ExtractionFileTooLargeError) as cm:
            list(read_archive(file_like=read_file_to_file_like(path=path), path=path))

        error = cm.exception
        tc.assertEqual(1, error.max_size)
        tc.assertGreater(error.actual_size, 1)
    finally:
        archive_module.MAX_7Z_MEMORY_USAGE = original_total_limit


def test_read_tar_gz_archive() -> None:
    """Test compressed TAR.GZ archive extraction."""
    path = "sharepoint2text/tests/resources/archives/test_archive.tar.gz"
    results = list(read_archive(file_like=read_file_to_file_like(path=path), path=path))

    # Should extract 1 text file from the archive
    tc.assertEqual(1, len(results))

    result = results[0]
    tc.assertIsInstance(result, PlainTextContent)
    tc.assertIn("This is a test document", result.get_full_text())


def test_archive_skips_nested_archives() -> None:
    """Test that nested archives are skipped to prevent zip bombs."""
    # Create a ZIP with a nested ZIP inside
    nested_content = b"nested content"
    inner_zip = std_io.BytesIO()
    with zipfile.ZipFile(inner_zip, "w") as zf:
        zf.writestr("inner.txt", nested_content)
    inner_zip.seek(0)

    outer_zip = std_io.BytesIO()
    with zipfile.ZipFile(outer_zip, "w") as zf:
        zf.writestr("outer.txt", b"outer content")
        zf.writestr("nested.zip", inner_zip.read())
    outer_zip.seek(0)

    results = list(read_archive(outer_zip, path="test.zip"))

    # Should only extract the outer.txt, not the nested.zip
    tc.assertEqual(1, len(results))
    tc.assertIn("outer content", results[0].get_full_text())


def test_archive_skips_hidden_files() -> None:
    """Test that hidden files (starting with .) are skipped."""

    zip_buffer = std_io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        zf.writestr("visible.txt", b"visible content")
        zf.writestr(".hidden.txt", b"hidden content")
        zf.writestr("__MACOSX/file.txt", b"macos resource fork")
    zip_buffer.seek(0)

    results = list(read_archive(zip_buffer, path="test.zip"))

    # Should only extract visible.txt
    tc.assertEqual(1, len(results))
    tc.assertIn("visible content", results[0].get_full_text())


def test_archive_skips_images() -> None:

    path = "sharepoint2text/tests/resources/archives/with_images.zip"
    results = list(
        read_archive(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=True
        )
    )
    tc.assertEqual(2, len(results))
    tc.assertEqual("#\nApache sample", results[0].get_full_text())
    tc.assertEqual("Hello World", results[1].get_full_text())
    tc.assertEqual([], list(results[0].iterate_images()))
    tc.assertEqual([], list(results[1].iterate_images()))

    path = "sharepoint2text/tests/resources/archives/with_images.zip"
    results = list(
        read_archive(
            file_like=read_file_to_file_like(path=path), path=path, ignore_images=False
        )
    )
    tc.assertEqual(2, len(results))
    tc.assertEqual(1, len(list(results[0].iterate_images())))
    tc.assertEqual(1, len(list(results[1].iterate_images())))
