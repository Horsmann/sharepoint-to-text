from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Sequence

import sharepoint2text
from sharepoint2text.extractors.data_types import ExtractionInterface
from sharepoint2text.extractors.serialization import serialize_extraction


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
        help="Emit structured JSON instead of plain full text (omits binary payloads by default).",
    )
    parser.add_argument(
        "--binary",
        action="store_true",
        help="With --json, include binary payloads (images/attachments) as base64 blobs.",
    )
    return parser


def _serialize_results(
    results: list[ExtractionInterface], *, include_binary: bool
) -> dict | list[dict]:
    if len(results) == 1:
        return serialize_extraction(results[0], include_binary=include_binary)
    return [
        serialize_extraction(result, include_binary=include_binary)
        for result in results
    ]


def _serialize_full_text(results: list[ExtractionInterface]) -> str:
    return "\n\n".join(result.get_full_text().rstrip() for result in results).rstrip()


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    try:
        args, unknown = parser.parse_known_args(argv)
    except SystemExit as exc:
        code = exc.code if isinstance(exc.code, int) else 1
        return code

    if unknown:
        unknown_str = " ".join(unknown)
        print(
            f"sharepoint2text: warning: unsupported arguments: {unknown_str}",
            file=sys.stderr,
        )
        return 1

    try:
        if args.binary and not args.json:
            raise ValueError("--binary requires --json")
        results = list(sharepoint2text.read_file(args.path))
        if not results:
            raise RuntimeError(f"No extraction results for {args.path}")
        if args.json:
            payload = _serialize_results(results, include_binary=bool(args.binary))
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
