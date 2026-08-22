"""Tests for version-2 CLI text and JSON output."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import sharepoint2text
from sharepoint2text.cli import _serialize_results, main
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
