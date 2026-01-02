import json
from pathlib import Path

import sharepoint2text
from sharepoint2text.cli import main


def test_cli_outputs_full_text_by_default(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = next(sharepoint2text.read_file(path)).get_full_text()

    exit_code = main([str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == f"{expected}\n"


def test_cli_outputs_json_with_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = next(sharepoint2text.read_file(path)).to_json()

    exit_code = main(["--json", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def _contains_binary_markers(value: object) -> bool:
    if isinstance(value, dict):
        if "_bytes" in value or "_bytesio" in value:
            return True
        return any(_contains_binary_markers(v) for v in value.values())
    if isinstance(value, list):
        return any(_contains_binary_markers(v) for v in value)
    return False


def test_cli_outputs_json_without_binary_payloads(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json", "--no-binary", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload["_type"] == "PdfContent"
    assert _contains_binary_markers(payload) is False

    images = payload["pages"][0]["images"]
    assert len(images) > 0
    assert images[0]["data"] is None


def test_cli_rejects_no_binary_without_json(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--no-binary", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "requires --json" in captured.err
