"""Tests for version-2 CLI text and JSON output."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

import sharepoint2text
from sharepoint2text.cli import (
    _build_parser,
    _build_zip_bomb_limits,
    _parse_zip_bomb_limit_multiplier,
    _serialize_results,
    _serialize_unit_results,
    main,
)
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    DEFAULT_ZIP_BOMB_LIMITS,
)
from sharepoint2text.parsing.models import (
    Annotation,
    Attachment,
    ContentUnit,
    ExtractedDocument,
    ImageAsset,
    Table,
    document_to_dict,
)

PLAIN_PATH = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
MISLABELED_DOCX_PATH = Path(
    "sharepoint2text/tests/resources/legacy_ms/ECE-TRANS-2021-24e.DOC"
).resolve()
IMAGE_PDF_PATH = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()
EMAIL_PATH = Path(
    "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
).resolve()
ENCRYPTED_PDF_PATH = Path(
    "sharepoint2text/tests/resources/legacy_ms/password_protected/"
    "pdf-password-protected-pw123.pdf"
).resolve()
MAILBOX_FOLDER = Path("sharepoint2text/tests/resources/mails").resolve()

EXPECTED_CLI_OPTIONS = {
    "-h",
    "--help",
    "-v",
    "--version",
    "-f",
    "--file",
    "-d",
    "--folder",
    "-s",
    "--suffixes",
    "--no-recursive",
    "-j",
    "--json",
    "-u",
    "--json-unit",
    "-o",
    "--output",
    "-i",
    "--include-binary",
    "--no-images",
    "-n",
    "--no-attachments",
    "-a",
    "--extract-annotations",
    "--force-plain-text",
    "-m",
    "--max-file-size-mb",
    "--zip-bomb-limit-multiplier",
    "--zblm",
}


def _body(envelope: dict[str, Any]) -> dict[str, Any]:
    """Return the document body from one version-2 envelope."""
    body = envelope["document"]
    assert isinstance(body, dict)
    return body


def test_cli_help_lists_every_current_option() -> None:
    """Verify rendered help contains the complete current option set."""
    parser = _build_parser()
    help_text = parser.format_help()
    actual_options = {
        option for action in parser._actions for option in action.option_strings
    }

    assert actual_options == EXPECTED_CLI_OPTIONS
    assert all(option in help_text for option in EXPECTED_CLI_OPTIONS)


def test_cli_help_explains_modes_constraints_and_defaults() -> None:
    """Verify help gives actionable details for non-obvious CLI behavior."""
    help_text = " ".join(_build_parser().format_help().split())

    assert "input selection:" in help_text
    assert "output format and destination:" in help_text
    assert "extraction options:" in help_text
    assert "resource limits:" in help_text
    assert "Requires --json or --json-unit" in help_text
    assert "image and attachment payloads" in help_text
    assert "new extensionless path" in help_text
    assert "preserves subdirectories" in help_text
    assert "default: 100" in help_text
    assert "Omit this option to keep the defaults" in help_text
    assert "examples:" in help_text


def test_cli_outputs_full_text_by_default(capsys: Any) -> None:
    """Verify the default output is normalized full text."""
    expected = next(sharepoint2text.read_file(PLAIN_PATH)).full_text

    exit_code = main(["--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == (expected if expected.endswith("\n") else f"{expected}\n")


def test_cli_extracts_docx_package_with_doc_extension(capsys: Any) -> None:
    """Verify content detection handles a DOCX package named with .DOC."""
    exit_code = main(["--file", str(MISLABELED_DOCX_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out.startswith("United Nations\nECE/TRANS/2021/24")
    assert captured.err == ""


def test_cli_outputs_version_two_json(capsys: Any) -> None:
    """Verify JSON output is the public version-2 wire schema."""
    document = next(sharepoint2text.read_file(PLAIN_PATH))

    exit_code = main(["--json", "--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert json.loads(captured.out) == [document_to_dict(document)]


def test_cli_short_json_flag_uses_version_two(capsys: Any) -> None:
    """Verify the short JSON flags produce the same stable schema."""
    exit_code = main(["-j", "-f", str(PLAIN_PATH)])
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert exit_code == 0
    assert payload[0]["schema"] == "sharepoint2text.extraction"
    assert payload[0]["version"] == 2


def test_serialize_results_keeps_stable_list_shape() -> None:
    """Verify helper output remains a list for multiple documents."""
    document = next(sharepoint2text.read_file(PLAIN_PATH))

    payload = _serialize_results([document, document], include_binary=False)

    assert len(payload) == 2
    assert all(item["version"] == 2 for item in payload)


def test_cli_json_unit_emits_one_v2_envelope_per_unit(capsys: Any) -> None:
    """Verify unit mode keeps schema and document metadata intact."""
    document = next(sharepoint2text.read_file(PLAIN_PATH))

    exit_code = main(["-u", "-f", str(PLAIN_PATH)])
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert exit_code == 0
    assert len(payload) == len(document.units)
    assert all(item["version"] == 2 for item in payload)
    assert all(len(_body(item)["units"]) == 1 for item in payload)
    assert all(_body(item)["source"]["filename"] == "plain.txt" for item in payload)


def test_json_unit_preserves_document_level_content() -> None:
    """Verify every self-contained unit envelope retains parent-level records."""
    document = ExtractedDocument(
        format="test",
        units=[
            ContentUnit(
                number=1,
                kind="document",
                text="body",
                images=[ImageAsset(number=1, filename="image.png")],
                tables=[Table(rows=[["cell"]])],
                annotations=[Annotation(kind="note", text="context")],
            )
        ],
        attachments=[Attachment(filename="attachment.txt")],
    )

    payload = _serialize_unit_results([document], include_binary=False)
    body = _body(payload[0])
    unit = body["units"][0]

    assert unit["images"]
    assert unit["tables"]
    assert unit["annotations"]
    assert body["attachments"]


def test_cli_json_unit_retains_email_attachments(capsys: Any) -> None:
    """Verify streaming unit output retains parent email attachment records."""
    exit_code = main(["--json-unit", "--file", str(EMAIL_PATH)])
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert exit_code == 0
    assert payload
    assert all(_body(envelope)["attachments"] for envelope in payload)


def test_cli_omits_binary_payloads_by_default(capsys: Any) -> None:
    """Verify default JSON retains image dimensions but omits image bytes."""
    exit_code = main(["--json", "--file", str(IMAGE_PDF_PATH)])
    captured = capsys.readouterr()
    body = _body(json.loads(captured.out)[0])
    images = [image for unit in body["units"] for image in unit["images"]]

    assert exit_code == 0
    assert images
    assert all("data" not in image for image in images)
    assert all(image["width"] > 0 for image in images)
    assert all(image["height"] > 0 for image in images)
    assert all(image["ratio"] == image["width"] / image["height"] for image in images)


def test_cli_encodes_binary_payloads_when_requested(capsys: Any) -> None:
    """Verify requested images use plain base64 strings in v2 JSON."""
    exit_code = main(["--json", "--include-binary", "--file", str(IMAGE_PDF_PATH)])
    captured = capsys.readouterr()
    body = _body(json.loads(captured.out)[0])
    images = [image for unit in body["units"] for image in unit["images"]]

    assert exit_code == 0
    assert images
    assert all(isinstance(image["data"], str) for image in images)


def test_cli_include_binary_encodes_attachment_payloads(capsys: Any) -> None:
    """Verify the accurately named binary option includes attachment bytes."""
    exit_code = main(["--json", "--include-binary", "--file", str(EMAIL_PATH)])
    captured = capsys.readouterr()
    attachments = _body(json.loads(captured.out)[0])["attachments"]

    assert exit_code == 0
    assert attachments
    assert all(isinstance(attachment["data"], str) for attachment in attachments)


def test_cli_no_attachments_removes_attachment_records(capsys: Any) -> None:
    """Verify attachment suppression reaches normalized JSON output."""
    exit_code = main(["--json", "--no-attachments", "--file", str(EMAIL_PATH)])
    captured = capsys.readouterr()
    body = _body(json.loads(captured.out)[0])

    assert exit_code == 0
    assert body["attachments"] == []


def test_cli_rejects_removed_include_images_option(capsys: Any) -> None:
    """Verify the removed version 1 binary option is no longer accepted."""
    exit_code = main(["--include-images", "--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "unsupported arguments" in captured.err


def test_cli_rejects_unknown_arguments(capsys: Any) -> None:
    """Verify removed or misspelled flags fail explicitly."""
    exit_code = main(["--no-binary", "--json", "--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "warning: unsupported arguments" in captured.err


def test_cli_reports_extraction_errors_without_tracebacks(capsys: Any) -> None:
    """Verify library extraction errors become concise CLI failures."""
    exit_code = main(["--file", str(ENCRYPTED_PDF_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "encrypted or password-protected" in captured.err
    assert "Traceback" not in captured.err


def test_cli_respects_max_file_size_override(capsys: Any, tmp_path: Path) -> None:
    """Verify CLI byte limits can be enforced and disabled."""
    path = tmp_path / "small.txt"
    path.write_text("hello", encoding="utf-8")

    exit_code = main(["--file", str(path), "--max-file-size-mb", "0.000001"])
    captured = capsys.readouterr()
    assert exit_code == 1
    assert "exceeds CLI maximum" in captured.err

    exit_code = main(["--file", str(path), "--max-file-size-mb", "0"])
    captured = capsys.readouterr()
    assert exit_code == 0
    assert captured.out == "hello\n"


def test_cli_zip_bomb_multiplier_scales_every_default() -> None:
    """Verify one multiplier is blindly applied to every ZIP-bomb threshold."""
    multiplier = 3

    limits = _build_zip_bomb_limits(multiplier)

    assert limits.max_entries == DEFAULT_ZIP_BOMB_LIMITS.max_entries * multiplier
    assert (
        limits.max_total_uncompressed_bytes
        == DEFAULT_ZIP_BOMB_LIMITS.max_total_uncompressed_bytes * multiplier
    )
    assert (
        limits.max_single_uncompressed_bytes
        == DEFAULT_ZIP_BOMB_LIMITS.max_single_uncompressed_bytes * multiplier
    )
    assert (
        limits.max_total_compression_ratio
        == DEFAULT_ZIP_BOMB_LIMITS.max_total_compression_ratio * multiplier
    )
    assert (
        limits.max_entry_compression_ratio
        == DEFAULT_ZIP_BOMB_LIMITS.max_entry_compression_ratio * multiplier
    )


def test_cli_zip_bomb_multiplier_default_preserves_limits() -> None:
    """Verify an omitted option keeps every default threshold unchanged."""
    assert _build_zip_bomb_limits(1) == DEFAULT_ZIP_BOMB_LIMITS


def test_cli_zip_bomb_multiplier_accepts_none_case_insensitively() -> None:
    """Verify the literal disable value maps to the disabled-check marker."""
    assert _parse_zip_bomb_limit_multiplier("None") is None


@pytest.mark.parametrize(("value", "expected"), [("2", 2), ("10", 10)])
def test_cli_accepts_zip_bomb_multiplier_boundaries(
    value: str,
    expected: int,
) -> None:
    """Verify both inclusive multiplier boundaries are accepted."""
    assert _parse_zip_bomb_limit_multiplier(value) == expected


@pytest.mark.parametrize("value", ["1", "11", "2.5", "disabled"])
def test_cli_rejects_invalid_zip_bomb_multipliers(
    value: str,
    capsys: Any,
) -> None:
    """Verify only whole multipliers from 2 through 10 or none are accepted."""
    exit_code = main(
        [
            "--file",
            str(PLAIN_PATH),
            "--zip-bomb-limit-multiplier",
            value,
        ]
    )
    captured = capsys.readouterr()

    assert exit_code == 2
    assert "whole integer from 2 through 10, or 'none'" in captured.err


@pytest.mark.parametrize(
    ("option_name", "cli_value", "expected_multiplier"),
    [
        ("--zip-bomb-limit-multiplier", "3", 3),
        ("--zip-bomb-limit-multiplier", "None", None),
        ("--zblm", "4", 4),
    ],
)
def test_cli_forwards_zip_bomb_setting_to_single_file(
    option_name: str,
    cli_value: str,
    expected_multiplier: int | None,
    capsys: Any,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Verify the CLI passes scaled or disabled limits into file extraction."""
    received_limits: list[sharepoint2text.ZipBombLimits | None] = []
    original_read_file = sharepoint2text.read_file

    def recording_read_file(path: Any, **kwargs: Any) -> Any:
        received_limits.append(kwargs.get("zip_bomb_limits"))
        return original_read_file(path, **kwargs)

    monkeypatch.setattr(sharepoint2text, "read_file", recording_read_file)

    exit_code = main(
        [
            "--file",
            str(PLAIN_PATH),
            option_name,
            cli_value,
        ]
    )
    capsys.readouterr()

    assert exit_code == 0
    assert received_limits == [_build_zip_bomb_limits(expected_multiplier)]


@pytest.mark.parametrize("mirror_output", [False, True])
def test_cli_forwards_zip_bomb_multiplier_to_folder_modes(
    mirror_output: bool,
    capsys: Any,
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    """Verify combined and mirrored folder modes pass the same scaled limits."""
    input_folder = tmp_path / "input"
    input_folder.mkdir()
    (input_folder / "sample.txt").write_text("hello", encoding="utf-8")
    received_limits: list[sharepoint2text.ZipBombLimits | None] = []
    original_read_many = sharepoint2text.read_many

    def recording_read_many(path: Any, **kwargs: Any) -> Any:
        received_limits.append(kwargs.get("zip_bomb_limits"))
        return original_read_many(path, **kwargs)

    monkeypatch.setattr(sharepoint2text, "read_many", recording_read_many)
    argv = [
        "--folder",
        str(input_folder),
        "--zip-bomb-limit-multiplier",
        "4",
    ]
    if mirror_output:
        output_folder = tmp_path / "output"
        output_folder.mkdir()
        argv.extend(["--output", str(output_folder)])

    exit_code = main(argv)
    capsys.readouterr()

    assert exit_code == 0
    assert received_limits == [_build_zip_bomb_limits(4)]


def test_cli_folder_json_is_a_streamed_v2_array(capsys: Any) -> None:
    """Verify combined folder output contains only version-2 envelopes."""
    folder = Path("sharepoint2text/tests/resources/plain_text").resolve()

    exit_code = main(["--folder", str(folder), "--json", "--suffixes", ".txt"])
    captured = capsys.readouterr()
    payload = json.loads(captured.out)

    assert exit_code == 0
    assert payload
    assert all(item["schema"] == "sharepoint2text.extraction" for item in payload)


def test_cli_folder_to_folder_writes_version_two_json(tmp_path: Path) -> None:
    """Verify per-file folder output also uses the version-2 schema."""
    input_folder = tmp_path / "input"
    output_folder = tmp_path / "output"
    input_folder.mkdir()
    output_folder.mkdir()
    (input_folder / "sample.txt").write_text("hello", encoding="utf-8")

    exit_code = main(
        ["--folder", str(input_folder), "--output", str(output_folder), "--json"]
    )
    payload = json.loads((output_folder / "sample.json").read_text())

    assert exit_code == 0
    assert payload[0]["version"] == 2


def test_cli_folder_output_preserves_every_mailbox_message(tmp_path: Path) -> None:
    """Verify mirrored output aggregates all documents yielded by one source."""
    output_folder = tmp_path / "output"
    output_folder.mkdir()

    exit_code = main(
        [
            "--folder",
            str(MAILBOX_FOLDER),
            "--suffixes",
            ".mbox",
            "--output",
            str(output_folder),
            "--json",
        ]
    )
    payload = json.loads((output_folder / "basic_email.json").read_text())

    assert exit_code == 0
    assert len(payload) == 2
