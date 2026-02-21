import json
from pathlib import Path

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
    expected = [
        serialize_extraction(unit, include_binary=False)
        for unit in result.iterate_units()
    ]

    exit_code = main(["--json-unit", "--file", str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected


def test_cli_outputs_json_unit_with_short_flag(capsys) -> None:
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    result = next(sharepoint2text.read_file(path))
    expected = [
        serialize_extraction(unit, include_binary=False)
        for unit in result.iterate_units()
    ]

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
