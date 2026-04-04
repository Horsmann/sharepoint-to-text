import json
from pathlib import Path
from typing import Any

import sharepoint2text
from sharepoint2text.cli import _serialize_results, main
from sharepoint2text.parsing.extractors.serialization import serialize_extraction

EMAIL_WITH_ATTACHMENT_PATH = Path(
    "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
).resolve()
BASIC_EMAIL_PATH = Path(
    "sharepoint2text/tests/resources/mails/basic_email.eml"
).resolve()


def test_cli_outputs_full_text_by_default(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = next(sharepoint2text.read_file(path)).get_full_text()

    exit_code = main(["--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == f"{expected}\n"


def test_cli_outputs_json_with_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = [
        serialize_extraction(
            next(sharepoint2text.read_file(path)), include_binary=False
        )
    ]

    exit_code = main(["--json", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def test_cli_outputs_json_with_short_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = [
        serialize_extraction(
            next(sharepoint2text.read_file(path)), include_binary=False
        )
    ]

    exit_code = main(["-j", "-f", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def test_serialize_results_returns_list_for_multiple_results() -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    result = next(sharepoint2text.read_file(path))

    payload = _serialize_results([result, result], include_binary=False)

    assert isinstance(payload, list)
    assert len(payload) == 2


def test_cli_outputs_json_unit_with_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    result = next(sharepoint2text.read_file(path))
    file_metadata = serialize_extraction(result.get_metadata(), include_binary=False)
    expected = []
    for unit in result.iterate_units():
        unit_dict = serialize_extraction(unit, include_binary=False)
        unit_dict["unit_metadata"] = serialize_extraction(
            unit.get_metadata(), include_binary=False
        )
        unit_dict["file_metadata"] = file_metadata
        expected.append(unit_dict)

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def test_cli_outputs_json_unit_with_short_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    result = next(sharepoint2text.read_file(path))
    file_metadata = serialize_extraction(result.get_metadata(), include_binary=False)
    expected = []
    for unit in result.iterate_units():
        unit_dict = serialize_extraction(unit, include_binary=False)
        unit_dict["unit_metadata"] = serialize_extraction(
            unit.get_metadata(), include_binary=False
        )
        unit_dict["file_metadata"] = file_metadata
        expected.append(unit_dict)

    exit_code = main(["-u", "-f", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def test_cli_plain_text_extracts_email_content(capsys) -> None:
    expected = next(sharepoint2text.read_file(BASIC_EMAIL_PATH)).get_full_text()

    exit_code = main(["--file", str(BASIC_EMAIL_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == f"{expected}\n"


def test_cli_plain_text_extracts_supported_email_attachments(capsys) -> None:
    exit_code = main(["--file", str(EMAIL_WITH_ATTACHMENT_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "This is a test sentence" in captured.out
    assert "The slide title" in captured.out


def test_cli_json_extracts_supported_email_attachments(capsys) -> None:
    exit_code = main(["--json", "--file", str(EMAIL_WITH_ATTACHMENT_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert {item["_type"] for item in payload} == {
        "EmailContent",
        "PdfContent",
        "PptxContent",
    }


def test_cli_json_no_attachments_excludes_email_attachments(capsys) -> None:
    exit_code = main(
        ["--json", "--no-attachments", "--file", str(EMAIL_WITH_ATTACHMENT_PATH)]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert {item["_type"] for item in payload} == {"EmailContent"}
    assert payload[0]["attachments"] == []


def test_cli_json_no_attachments_excludes_email_attachments_with_short_flag(
    capsys,
) -> None:
    exit_code = main(["-j", "-n", "-f", str(EMAIL_WITH_ATTACHMENT_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert {item["_type"] for item in payload} == {"EmailContent"}
    assert payload[0]["attachments"] == []


def test_cli_json_unit_extracts_supported_email_attachments(capsys) -> None:
    exit_code = main(["--json-unit", "--file", str(EMAIL_WITH_ATTACHMENT_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)

    unit_types = {
        unit["_type"] for unit in payload if isinstance(unit, dict) and "_type" in unit
    }
    assert "EmailUnit" in unit_types
    assert "PdfUnit" in unit_types
    assert "PptxUnit" in unit_types


def _contains_binary_markers(value: object) -> bool:
    if isinstance(value, dict):
        if "_bytes" in value or "_bytesio" in value:
            return True
        return any(_contains_binary_markers(v) for v in value.values())
    if isinstance(value, list):
        return any(_contains_binary_markers(v) for v in value)
    return False


def test_cli_outputs_json_without_images(capsys) -> None:
    """Test that by default (without --include-images), images are not extracted."""
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert payload[0]["_type"] == "PdfContent"
    assert _contains_binary_markers(payload) is False

    # Images are not extracted by default
    images = payload[0]["pages"][0]["images"]
    assert len(images) == 0


def test_cli_outputs_json_unit_without_images(capsys) -> None:
    """Test that by default (without --include-images), images are not extracted."""
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0
    assert payload[0]["_type"] == "PdfUnit"
    assert _contains_binary_markers(payload) is False

    # Images are not extracted by default
    images = payload[0]["images"]
    assert len(images) == 0


def test_cli_outputs_json_with_binary_payloads_when_requested(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json", "--include-images", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert payload[0]["_type"] == "PdfContent"
    assert _contains_binary_markers(payload) is True

    images = payload[0]["pages"][0]["images"]
    assert len(images) > 0
    assert isinstance(images[0]["data"], dict)
    assert "_bytesio" in images[0]["data"] or "_bytes" in images[0]["data"]


def test_cli_outputs_json_unit_with_binary_payloads_when_requested(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json-unit", "--include-images", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0
    assert payload[0]["_type"] == "PdfUnit"
    assert _contains_binary_markers(payload) is True

    images = payload[0]["images"]
    assert len(images) > 0
    assert isinstance(images[0]["data"], dict)
    assert "_bytesio" in images[0]["data"] or "_bytes" in images[0]["data"]


def test_cli_json_unit_includes_unit_metadata(capsys) -> None:
    """Test that --json-unit output includes unit_metadata for each unit."""
    path = Path("sharepoint2text/tests/resources/modern_ms/headings.docx").resolve()

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0

    # Each unit should have unit_metadata
    for unit in payload:
        assert "unit_metadata" in unit, "unit_metadata should be present"
        metadata = unit["unit_metadata"]
        assert "_type" in metadata
        assert "unit_number" in metadata


def test_cli_json_unit_includes_file_metadata(capsys) -> None:
    """Test that --json-unit output includes file_metadata for each unit."""
    path = Path("sharepoint2text/tests/resources/modern_ms/headings.docx").resolve()

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0

    # Each unit should have file_metadata with required fields
    for unit in payload:
        assert "file_metadata" in unit, "file_metadata should be present"
        metadata = unit["file_metadata"]
        assert "_type" in metadata
        assert "filename" in metadata
        assert metadata["filename"] == "headings.docx"
        assert "file_path" in metadata
        assert "file_extension" in metadata
        assert metadata["file_extension"] == ".docx"


def test_cli_json_unit_metadata_contains_docx_details(capsys) -> None:
    """Test that DOCX file_metadata includes author and creation info."""
    path = Path("sharepoint2text/tests/resources/modern_ms/headings.docx").resolve()

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())

    # Check DOCX-specific metadata fields
    file_metadata = payload[0]["file_metadata"]
    assert file_metadata["_type"] == "DocxMetadata"
    assert "author" in file_metadata
    assert "created" in file_metadata
    assert "modified" in file_metadata


def test_cli_json_unit_metadata_heading_path(capsys) -> None:
    """Test that DOCX unit_metadata includes heading information."""
    path = Path("sharepoint2text/tests/resources/modern_ms/headings.docx").resolve()

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())

    # DOCX units should have heading information in unit_metadata
    unit_metadata = payload[0]["unit_metadata"]
    assert unit_metadata["_type"] == "DocxUnitMetadata"
    assert "heading_level" in unit_metadata
    assert "heading_path" in unit_metadata
    assert "location" in unit_metadata


def test_cli_json_includes_metadata(capsys) -> None:
    """Test that --json output includes metadata in the extraction result."""
    path = Path("sharepoint2text/tests/resources/modern_ms/headings.docx").resolve()

    exit_code = main(["--json", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0

    # The extraction result should have metadata
    result = payload[0]
    assert "metadata" in result
    metadata = result["metadata"]
    assert "_type" in metadata
    assert metadata["_type"] == "DocxMetadata"
    assert "filename" in metadata
    assert "author" in metadata
    assert "created" in metadata


def test_cli_passes_timeout_to_read_file(capsys: Any, monkeypatch: Any) -> None:
    """CLI should forward --timeout to the single-file API."""
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    result = next(sharepoint2text.read_file(path, timeout_seconds=0))
    observed_timeout: dict[str, float] = {}

    def patched_read_file(file_path: Any, **kwargs: Any) -> Any:
        observed_timeout["value"] = kwargs["timeout_seconds"]
        return iter([result])

    monkeypatch.setattr(sharepoint2text, "read_file", patched_read_file)

    exit_code = main(["--file", str(path), "--timeout", "12.5"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert observed_timeout["value"] == 12.5
    assert captured.out == f"{result.get_full_text()}\n"


def test_cli_warns_on_removed_no_binary_argument(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--no-binary", "--json", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "warning: unsupported arguments" in captured.err


def test_cli_rejects_include_images_without_json(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--include-images", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "requires --json or --json-unit" in captured.err


def test_cli_warns_on_unsupported_argument(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--json", "--not-a-real-flag", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "warning: unsupported arguments" in captured.err


def test_cli_respects_max_file_size_override(capsys, tmp_path) -> None:
    path = tmp_path / "small.txt"
    path.write_text("hello", encoding="utf-8")

    # 0.000001 MiB ~= 1 byte, should fail for a 5-byte file.
    exit_code = main(["--file", str(path), "--max-file-size-mb", "0.000001"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "exceeds CLI maximum" in captured.err

    # 0 disables size checks.
    exit_code = main(["--file", str(path), "--max-file-size-mb", "0"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == "hello\n"


def test_cli_respects_max_file_size_short_flag(capsys, tmp_path) -> None:
    path = tmp_path / "small.txt"
    path.write_text("hello", encoding="utf-8")

    # 0.000001 MiB ~= 1 byte, should fail for a 5-byte file.
    exit_code = main(["--file", str(path), "-m", "0.000001"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "exceeds CLI maximum" in captured.err


def test_cli_rejects_negative_max_file_size_mb(capsys, tmp_path) -> None:
    path = tmp_path / "small.txt"
    path.write_text("hello", encoding="utf-8")

    exit_code = main(["--file", str(path), "--max-file-size-mb", "-1"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "--max-file-size-mb must be >= 0" in captured.err


# =============================================================================
# Folder extraction tests (--folder / -d)
# =============================================================================


def test_cli_folder_extracts_all_supported_by_default(capsys) -> None:
    """--folder without --suffixes should extract all supported files."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder)])
    captured = capsys.readouterr()

    assert exit_code == 0
    # Should contain content from multiple files
    assert len(captured.out) > 0


def test_cli_folder_with_short_flag(capsys) -> None:
    """-d should work as shorthand for --folder."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["-d", str(folder)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert len(captured.out) > 0


def test_cli_folder_with_suffixes_filter(capsys) -> None:
    """--folder with --suffixes should only extract matching files."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--suffixes", ".txt"])
    captured = capsys.readouterr()

    assert exit_code == 0
    # plain.txt content should be present
    assert "Hello" in captured.out or len(captured.out) > 0


def test_cli_folder_with_multiple_suffixes(capsys) -> None:
    """--suffixes should accept comma-separated values."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--suffixes", ".txt,.md"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert len(captured.out) > 0


def test_cli_folder_with_suffixes_short_flag(capsys) -> None:
    """-s should work as shorthand for --suffixes."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["-d", str(folder), "-s", ".txt"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert len(captured.out) > 0


def test_cli_folder_suffixes_without_leading_dot(capsys) -> None:
    """--suffixes should work without leading dot."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--suffixes", "txt,md"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert len(captured.out) > 0


def test_cli_folder_json_output(capsys) -> None:
    """--folder should work with --json output."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--json", "--suffixes", ".txt"])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0


def test_cli_folder_json_unit_output(capsys) -> None:
    """--folder should work with --json-unit output."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--json-unit", "--suffixes", ".txt"])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0


def test_cli_folder_no_recursive(capsys, tmp_path) -> None:
    """--no-recursive should only extract from top-level folder."""
    # Create folder structure
    (tmp_path / "top.txt").write_text("top level", encoding="utf-8")
    subdir = tmp_path / "subdir"
    subdir.mkdir()
    (subdir / "nested.txt").write_text("nested level", encoding="utf-8")

    # With --no-recursive, should only find top.txt
    exit_code = main(
        ["--folder", str(tmp_path), "--no-recursive", "--suffixes", ".txt"]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "top level" in captured.out
    assert "nested level" not in captured.out


def test_cli_folder_recursive_by_default(capsys, tmp_path) -> None:
    """Folder extraction should be recursive by default."""
    # Create folder structure
    (tmp_path / "top.txt").write_text("top level", encoding="utf-8")
    subdir = tmp_path / "subdir"
    subdir.mkdir()
    (subdir / "nested.txt").write_text("nested level", encoding="utf-8")

    # Without --no-recursive, should find both files
    exit_code = main(["--folder", str(tmp_path), "--suffixes", ".txt"])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "top level" in captured.out
    assert "nested level" in captured.out


def test_cli_folder_nonexistent_raises_error(capsys) -> None:
    """--folder with non-existent path should return error."""
    exit_code = main(["--folder", "/nonexistent/path/to/folder"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "Folder not found" in captured.err


def test_cli_folder_file_path_raises_error(capsys) -> None:
    """--folder with a file path (not directory) should return error."""
    file_path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()

    exit_code = main(["--folder", str(file_path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "not a directory" in captured.err


def test_cli_suffixes_without_folder_raises_error(capsys) -> None:
    """--suffixes without --folder should return error."""
    file_path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()

    exit_code = main(["--file", str(file_path), "--suffixes", ".txt"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "--suffixes can only be used with --folder" in captured.err


def test_cli_no_recursive_without_folder_raises_error(capsys) -> None:
    """--no-recursive without --folder should return error."""
    file_path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()

    exit_code = main(["--file", str(file_path), "--no-recursive"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "--no-recursive can only be used with --folder" in captured.err


def test_cli_file_and_folder_mutually_exclusive(capsys) -> None:
    """--file and --folder should be mutually exclusive."""
    file_path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    folder_path = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--file", str(file_path), "--folder", str(folder_path)])

    # argparse should reject this combination
    assert exit_code != 0


def test_cli_folder_empty_suffixes_raises_error(capsys) -> None:
    """--suffixes with empty/whitespace value should return error."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--suffixes", "   "])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "at least one valid suffix" in captured.err


def test_cli_folder_no_matches_raises_error(capsys) -> None:
    """--folder with suffixes that match no files should return error."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--suffixes", ".nonexistent"])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "No extraction results" in captured.err


def test_cli_folder_with_max_file_size(capsys, tmp_path) -> None:
    """--folder should respect --max-file-size-mb."""
    # Create a file
    (tmp_path / "test.txt").write_text("hello world", encoding="utf-8")

    # With very small limit, should fail
    exit_code = main(
        [
            "--folder",
            str(tmp_path),
            "--suffixes",
            ".txt",
            "--max-file-size-mb",
            "0.000001",
        ]
    )
    captured = capsys.readouterr()

    # File should be skipped due to size, resulting in no results
    assert exit_code == 1
    assert "No extraction results" in captured.err


# =============================================================================
# Folder output tests (--output with folder extraction)
# =============================================================================


def test_cli_folder_output_to_existing_folder(capsys, tmp_path) -> None:
    """--folder with --output to existing folder should write separate files."""
    # Create input folder with files
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file1.txt").write_text("content one", encoding="utf-8")
    (input_folder / "file2.txt").write_text("content two", encoding="utf-8")

    # Create output folder
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--output",
            str(output_folder),
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Successfully extracted 2 file(s)" in captured.err

    # Check output files exist with correct content
    assert (output_folder / "file1.txt").exists()
    assert (output_folder / "file2.txt").exists()
    assert "content one" in (output_folder / "file1.txt").read_text()
    assert "content two" in (output_folder / "file2.txt").read_text()


def test_cli_folder_output_preserves_subdirectory_structure(capsys, tmp_path) -> None:
    """--folder output should preserve subdirectory structure."""
    # Create input folder with nested structure
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "root.txt").write_text("root content", encoding="utf-8")

    subdir = input_folder / "subdir"
    subdir.mkdir()
    (subdir / "nested.txt").write_text("nested content", encoding="utf-8")

    deep_subdir = subdir / "deep"
    deep_subdir.mkdir()
    (deep_subdir / "deep.txt").write_text("deep content", encoding="utf-8")

    # Create output folder
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--output",
            str(output_folder),
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Successfully extracted 3 file(s)" in captured.err

    # Check output structure mirrors input
    assert (output_folder / "root.txt").exists()
    assert (output_folder / "subdir" / "nested.txt").exists()
    assert (output_folder / "subdir" / "deep" / "deep.txt").exists()

    assert "root content" in (output_folder / "root.txt").read_text()
    assert "nested content" in (output_folder / "subdir" / "nested.txt").read_text()
    assert (
        "deep content" in (output_folder / "subdir" / "deep" / "deep.txt").read_text()
    )


def test_cli_folder_output_creates_new_folder(capsys, tmp_path) -> None:
    """--folder output should create new folder if path has no extension."""
    # Create input folder
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file.txt").write_text("test content", encoding="utf-8")

    # Output folder does not exist (no extension = treated as folder)
    output_folder = tmp_path / "new_output"

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--output",
            str(output_folder),
        ]
    )
    assert exit_code == 0
    assert output_folder.is_dir()
    assert (output_folder / "file.txt").exists()


def test_cli_folder_output_to_file_combines_results(capsys, tmp_path) -> None:
    """--folder with --output to file should combine all results."""
    # Create input folder
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file1.txt").write_text("content one", encoding="utf-8")
    (input_folder / "file2.txt").write_text("content two", encoding="utf-8")

    # Output file (has extension = treated as single file)
    output_file = tmp_path / "combined.txt"

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--output",
            str(output_file),
        ]
    )

    assert exit_code == 0
    assert output_file.exists()

    content = output_file.read_text()
    assert "content one" in content
    assert "content two" in content


def test_cli_folder_output_json_format(capsys, tmp_path) -> None:
    """--folder output to folder with --json should create .json files."""
    # Create input folder
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file.txt").write_text("test content", encoding="utf-8")

    # Create output folder
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--json",
            "--output",
            str(output_folder),
        ]
    )
    assert exit_code == 0

    # Should create .json file instead of .txt
    assert (output_folder / "file.json").exists()
    assert not (output_folder / "file.txt").exists()

    # Verify it's valid JSON
    content = json.loads((output_folder / "file.json").read_text())
    assert isinstance(content, list)


def test_cli_folder_output_json_unit_format(capsys, tmp_path) -> None:
    """--folder output with --json-unit should create per-unit .json files."""
    # Create input folder
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file.txt").write_text("test content", encoding="utf-8")

    # Create output folder
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--json-unit",
            "--output",
            str(output_folder),
        ]
    )
    assert exit_code == 0
    assert (output_folder / "file.json").exists()


def test_cli_folder_output_trailing_slash_creates_folder(capsys, tmp_path) -> None:
    """--output with trailing slash should create folder."""
    # Create input folder
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file.txt").write_text("test content", encoding="utf-8")

    # Output path with trailing slash
    output_path = str(tmp_path / "output_dir") + "/"

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--output",
            output_path,
        ]
    )

    assert exit_code == 0
    assert (tmp_path / "output_dir").is_dir()
    assert (tmp_path / "output_dir" / "file.txt").exists()


def test_cli_folder_output_prints_progress(capsys, tmp_path) -> None:
    """--folder output should print progress to stderr."""
    # Create input folder
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "file.txt").write_text("test content", encoding="utf-8")

    # Create output folder
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--output",
            str(output_folder),
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Extracted:" in captured.err
    assert "file.txt" in captured.err


def test_cli_single_file_output_unchanged(capsys, tmp_path) -> None:
    """Single file extraction with --output should still work as before."""
    # Create input file
    input_file = tmp_path / "input.txt"
    input_file.write_text("test content", encoding="utf-8")

    # Output file
    output_file = tmp_path / "output.txt"

    exit_code = main(
        [
            "--file",
            str(input_file),
            "--output",
            str(output_file),
        ]
    )

    assert exit_code == 0
    assert output_file.exists()
    assert "test content" in output_file.read_text()


def test_cli_folder_output_with_no_recursive(capsys, tmp_path) -> None:
    """--folder output with --no-recursive should only extract top-level files."""
    # Create input folder with nested structure
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "top.txt").write_text("top content", encoding="utf-8")

    subdir = input_folder / "subdir"
    subdir.mkdir()
    (subdir / "nested.txt").write_text("nested content", encoding="utf-8")

    # Create output folder
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(input_folder),
            "--suffixes",
            ".txt",
            "--no-recursive",
            "--output",
            str(output_folder),
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 0
    assert "Successfully extracted 1 file(s)" in captured.err

    # Only top-level file should be extracted
    assert (output_folder / "top.txt").exists()
    assert not (output_folder / "subdir").exists()
