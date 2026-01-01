import json
from pathlib import Path

import sharepoint2text
from sharepoint2text.cli import main


def test_cli_outputs_json_for_single_file(capsys) -> None:
    """Use pytest's capsys fixture to capture stdout/stderr from the CLI call."""
    path = Path("sharepoint2text/tests/resources/plain_text/plain.txt").resolve()
    expected = next(sharepoint2text.read_file(path)).to_json()

    exit_code = main([str(path)])
    captured = capsys.readouterr()

    assert exit_code == 0
    payload = json.loads(captured.out.strip())
    assert payload == expected
