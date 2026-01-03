import json
from pathlib import Path

import sharepoint2text
from sharepoint2text.cli import main
from sharepoint2text.extractors.serialization import serialize_extraction


def test_cli_outputs_full_text_by_default(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = next(sharepoint2text.read_file(path)).get_full_text()

    exit_code = main([str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    assert captured.out == f"{expected}\n"


def test_cli_outputs_json_with_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = serialize_extraction(
        next(sharepoint2text.read_file(path)), include_binary=False
    )

    exit_code = main(["--json", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def test_cli_outputs_json_unit_with_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    result = next(sharepoint2text.read_file(path))
    expected = [
        serialize_extraction(unit, include_binary=False)
        for unit in result.iterate_units()
    ]

    exit_code = main(["--json-unit", str(path)])
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

    exit_code = main(["--json", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload["_type"] == "PdfContent"
    assert _contains_binary_markers(payload) is False

    images = payload["pages"][0]["images"]
    assert len(images) > 0
    assert images[0]["data"] is None


def test_cli_outputs_json_unit_without_binary_payloads(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json-unit", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert isinstance(payload, list)
    assert len(payload) > 0
    assert payload[0]["_type"] == "PdfUnit"
    assert _contains_binary_markers(payload) is False

    images = payload[0]["images"]
    assert len(images) > 0
    assert images[0]["data"] is None


def test_cli_outputs_json_with_binary_payloads_when_requested(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json", "--binary", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload["_type"] == "PdfContent"
    assert _contains_binary_markers(payload) is True

    images = payload["pages"][0]["images"]
    assert len(images) > 0
    assert isinstance(images[0]["data"], dict)
    assert "_bytesio" in images[0]["data"] or "_bytes" in images[0]["data"]


def test_cli_outputs_json_unit_with_binary_payloads_when_requested(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/pdf/multi_image.pdf").resolve()

    exit_code = main(["--json-unit", "--binary", str(path)])
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


def test_cli_warns_on_removed_no_binary_argument(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--no-binary", "--json", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "warning: unsupported arguments" in captured.err


def test_cli_rejects_binary_without_json(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--binary", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "requires --json or --json-unit" in captured.err


def test_cli_warns_on_unsupported_argument(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    exit_code = main(["--json", "--not-a-real-flag", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 1
    assert "warning: unsupported arguments" in captured.err
