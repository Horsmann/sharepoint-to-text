"""Tests for version-2 CLI text and JSON output."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import pytest

import sharepoint2text
from sharepoint2text.cli import (
    _build_zip_bomb_limits,
    _parse_zip_bomb_limit_multiplier,
    _serialize_results,
    main,
)
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    DEFAULT_ZIP_BOMB_LIMITS,
)
from sharepoint2text.parsing.models import document_to_dict

PLAIN_PATH = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
IMAGE_PDF_PATH = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()
EMAIL_PATH = Path(
    "sharepoint2text/tests/resources/mails/msg_with_attachment.eml"
).resolve()


def _body(envelope: dict[str, Any]) -> dict[str, Any]:
    """Return the document body from one version-2 envelope."""
    body = envelope["document"]
    assert isinstance(body, dict)
    return body


def test_cli_outputs_full_text_by_default(capsys: Any) -> None:
    """Verify the default output is normalized full text."""
    expected = next(sharepoint2text.read_file(PLAIN_PATH)).full_text

    exit_code = main(["--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == f"{expected}\n"


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


def test_cli_omits_binary_payloads_by_default(capsys: Any) -> None:
    """Verify image data is neither extracted nor serialized by default."""
    exit_code = main(["--json", "--file", str(IMAGE_PDF_PATH)])
    captured = capsys.readouterr()
    body = _body(json.loads(captured.out)[0])

    assert exit_code == 0
    assert all(unit["images"] == [] for unit in body["units"])


def test_cli_encodes_binary_payloads_when_requested(capsys: Any) -> None:
    """Verify requested images use plain base64 strings in v2 JSON."""
    exit_code = main(["--json", "--include-images", "--file", str(IMAGE_PDF_PATH)])
    captured = capsys.readouterr()
    body = _body(json.loads(captured.out)[0])
    images = [image for unit in body["units"] for image in unit["images"]]

    assert exit_code == 0
    assert images
    assert all(isinstance(image["data"], str) for image in images)


def test_cli_no_attachments_removes_attachment_records(capsys: Any) -> None:
    """Verify attachment suppression reaches normalized JSON output."""
    exit_code = main(["--json", "--no-attachments", "--file", str(EMAIL_PATH)])
    captured = capsys.readouterr()
    body = _body(json.loads(captured.out)[0])

    assert exit_code == 0
    assert body["attachments"] == []


def test_cli_rejects_include_images_without_json(capsys: Any) -> None:
    """Verify binary extraction requires a structured output mode."""
    exit_code = main(["--include-images", "--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "requires --json or --json-unit" in captured.err


def test_cli_rejects_unknown_arguments(capsys: Any) -> None:
    """Verify removed or misspelled flags fail explicitly."""
    exit_code = main(["--no-binary", "--json", "--file", str(PLAIN_PATH)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "warning: unsupported arguments" in captured.err


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
    ("cli_value", "expected_multiplier"),
    [("3", 3), ("None", None)],
)
def test_cli_forwards_zip_bomb_setting_to_single_file(
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
            "--zip-bomb-limit-multiplier",
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
