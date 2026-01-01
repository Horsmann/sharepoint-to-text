from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Sequence

import sharepoint2text
from sharepoint2text.extractors.data_types import ExtractionInterface


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="sharepoint2text",
        description="Extract file content and emit full text to stdout (or JSON with --json).",
    )
    parser.add_argument(
        "path",
        type=Path,
        help="Path to the file to extract.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit structured JSON (result.to_json()) instead of plain full text.",
    )
    return parser


def _serialize_results(results: list[ExtractionInterface]) -> dict | list[dict]:
    if len(results) == 1:
        return results[0].to_json()
    return [result.to_json() for result in results]


def _serialize_full_text(results: list[ExtractionInterface]) -> str:
    return "\n\n".join(result.get_full_text().rstrip() for result in results).rstrip()


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    try:
        results = list(sharepoint2text.read_file(args.path))
        if not results:
            raise RuntimeError(f"No extraction results for {args.path}")
        if args.json:
            payload = _serialize_results(results)
            json.dump(payload, sys.stdout)
            sys.stdout.write("\n")
        else:
            sys.stdout.write(_serialize_full_text(results))
            sys.stdout.write("\n")
        return 0
    except Exception as exc:
        print(f"sharepoint2text: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
