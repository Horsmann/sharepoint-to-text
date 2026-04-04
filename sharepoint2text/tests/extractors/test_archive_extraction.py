import io as std_io
import logging
import stat
import tarfile
import zipfile
from typing import Any, List
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
from sharepoint2text.parsing.extractors.util.sevenzip import FileInfo
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
        tc.assertIn("test_archive.zip!/", result.get_metadata().file_path or "")


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


def test_archive_skips_zip_symbolic_links() -> None:
    """ZIP symbolic-link entries should be ignored."""
    zip_buffer = std_io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        zf.writestr("visible.txt", b"visible content")
        symlink = zipfile.ZipInfo("link.txt")
        symlink.create_system = 3
        symlink.external_attr = (stat.S_IFLNK | 0o777) << 16
        zf.writestr(symlink, "../etc/passwd")
    zip_buffer.seek(0)

    results = list(read_archive(zip_buffer, path="symlinks.zip"))

    tc.assertEqual(1, len(results))
    tc.assertIn("visible content", results[0].get_full_text())


def test_archive_skips_tar_symbolic_links() -> None:
    """TAR symbolic-link entries should be ignored."""
    tar_buffer = std_io.BytesIO()
    with tarfile.open(fileobj=tar_buffer, mode="w") as tf:
        visible_data = b"visible content"
        visible_info = tarfile.TarInfo("visible.txt")
        visible_info.size = len(visible_data)
        tf.addfile(visible_info, std_io.BytesIO(visible_data))

        symlink_info = tarfile.TarInfo("link.txt")
        symlink_info.type = tarfile.SYMTYPE
        symlink_info.linkname = "/etc/passwd"
        tf.addfile(symlink_info)
    tar_buffer.seek(0)

    results = list(read_archive(tar_buffer, path="symlinks.tar"))

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


def test_archive_spools_large_entry_instead_of_skipping(monkeypatch: Any) -> None:
    """ZIP entries larger than the memory threshold should roll to disk."""
    original_config = archive_module._config
    archive_module.configure_archive_extraction(max_memory_size=64)

    payload = "\n".join(f"Line {index}" for index in range(128)).encode("utf-8")
    zip_buffer = std_io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_STORED) as zf:
        zf.writestr("large.txt", payload)
    zip_buffer.seek(0)

    rolled_states: list[bool] = []
    real_spooled_file = archive_module.tempfile.SpooledTemporaryFile

    class TrackingSpooledFile:
        """Wrap SpooledTemporaryFile and record whether rollover occurred."""

        def __init__(self, *args: Any, **kwargs: Any) -> None:
            self._wrapped = real_spooled_file(*args, **kwargs)

        def __enter__(self) -> "TrackingSpooledFile":
            self._wrapped.__enter__()
            return self

        def __exit__(
            self,
            exc_type: type[BaseException] | None,
            exc_val: BaseException | None,
            exc_tb: Any,
        ) -> None:
            rolled_states.append(bool(getattr(self._wrapped, "_rolled", False)))
            self._wrapped.__exit__(exc_type, exc_val, exc_tb)

        def __getattr__(self, name: str) -> Any:
            return getattr(self._wrapped, name)

    monkeypatch.setattr(
        archive_module.tempfile,
        "SpooledTemporaryFile",
        TrackingSpooledFile,
    )

    try:
        results = list(read_archive(zip_buffer, path="large.zip"))
    finally:
        archive_module.configure_archive_extraction(
            buffer_size=original_config.buffer_size,
            max_memory_size=original_config.max_memory_size,
            max_workers=original_config.max_workers,
            enable_parallel=original_config.enable_parallel,
            enable_caching=original_config.enable_caching,
            enable_streaming=original_config.enable_streaming,
        )

    tc.assertEqual(1, len(results))
    tc.assertIsInstance(results[0], PlainTextContent)
    tc.assertIn("Line 0", results[0].get_full_text())
    tc.assertTrue(any(rolled_states))


def test_archive_skips_7z_symbolic_links(monkeypatch: Any) -> None:
    """7z symbolic-link entries should be ignored before extraction."""
    visible_data = b"visible content"

    class FakeSevenZipFile:
        """Provide a minimal 7z interface for archive-extractor tests."""

        def __init__(self, file_like: Any, mode: str) -> None:
            self._mode = mode

        def __enter__(self) -> "FakeSevenZipFile":
            return self

        def __exit__(
            self,
            exc_type: type[BaseException] | None,
            exc_val: BaseException | None,
            exc_tb: Any,
        ) -> None:
            return None

        def needs_password(self) -> bool:
            return False

        def list(self) -> List[FileInfo]:
            return [
                FileInfo(
                    filename="link.txt",
                    uncompressed=len(b"../etc/passwd"),
                    is_directory=False,
                    is_symlink=True,
                    attributes=(stat.S_IFLNK | 0o777) << 16,
                ),
                FileInfo(
                    filename="visible.txt",
                    uncompressed=len(visible_data),
                    is_directory=False,
                    attributes=(stat.S_IFREG | 0o644) << 16,
                ),
            ]

        def extract(self, path: str, targets: List[str]) -> None:
            tc.assertEqual(["visible.txt"], targets)
            with open(f"{path}/visible.txt", "wb") as extracted_file:
                extracted_file.write(visible_data)

    monkeypatch.setattr(archive_module, "SevenZipFile", FakeSevenZipFile)

    seven_zip_buffer = std_io.BytesIO(b"7z\xbc\xaf\x27\x1c" + b"payload")
    results = list(read_archive(seven_zip_buffer, path="symlinks.7z"))

    tc.assertEqual(1, len(results))
    tc.assertIn("visible content", results[0].get_full_text())
