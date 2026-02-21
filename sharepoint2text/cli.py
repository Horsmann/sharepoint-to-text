from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Sequence, TextIO

import sharepoint2text
from sharepoint2text.parsing.extractors.data_types import ExtractionInterface
from sharepoint2text.parsing.extractors.serialization import serialize_extraction


def _build_parser() -> argparse.ArgumentParser:
    """Build and return the command-line argument parser for the CLI."""
    parser = argparse.ArgumentParser(
        prog="sharepoint2text",
        description="Extract file content and emit full text to stdout (or JSON with --json/--json-unit).",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {sharepoint2text.__version__}",
        help="Show the version and exit.",
    )
    parser.add_argument(
        "--file",
        type=Path,
        required=True,
        help="Path to the file to extract.",
    )
    output_group = parser.add_mutually_exclusive_group()
    output_group.add_argument(
        "--json",
        action="store_true",
        help="Emit structured JSON instead of plain full text (omits binary payloads by default).",
    )
    output_group.add_argument(
        "--json-unit",
        dest="json_unit",
        action="store_true",
        help="Emit JSON for extracted text units instead of full extraction objects (omits binary payloads by default).",
    )
    parser.add_argument(
        "--include-images",
        dest="include_images",
        action="store_true",
        help="Extract images from the file and include image data as base64 blobs in JSON output (default: images are ignored for faster processing).",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        help="Output file path (default: stdout).",
    )
    return parser


def _serialize_results(
    results: list[ExtractionInterface], *, include_binary: bool
) -> dict | list[dict]:
    """Serialize extraction results for ``--json`` output.

    Returns a single JSON object for one extraction result, otherwise a list of
    JSON objects for multi-result inputs (for example, ``.mbox`` or archives).
    """
    if len(results) == 1:
        return serialize_extraction(results[0], include_binary=include_binary)
    return [
        serialize_extraction(result, include_binary=include_binary)
        for result in results
    ]


def _serialize_unit_results(
    results: list[ExtractionInterface], *, include_binary: bool
) -> list[dict] | list[list[dict]]:
    """Serialize per-unit output for ``--json-unit`` mode.

    For a single extraction result, returns a flat list of unit objects.
    For multiple extraction results, returns a list of unit-lists preserving the
    result boundaries.
    """
    if len(results) == 1:
        return [
            serialize_extraction(unit, include_binary=include_binary)
            for unit in results[0].iterate_units()
        ]
    return [
        [
            serialize_extraction(unit, include_binary=include_binary)
            for unit in result.iterate_units()
        ]
        for result in results
    ]


def _serialize_full_text(results: list[ExtractionInterface]) -> str:
    """Join ``get_full_text()`` from all results using blank-line separators."""
    return "\n\n".join(result.get_full_text().rstrip() for result in results).rstrip()


def main(argv: Sequence[str] | None = None) -> int:
    """Run the CLI and return a process-style exit code.

    Args:
        argv: Optional argument list. If ``None``, arguments are read from
            ``sys.argv`` by ``argparse``.

    Returns:
        ``0`` on success, ``1`` on validation/extraction/serialization errors.
        Parser-driven early exits (for example ``--help`` / ``--version``) return
        the exit code produced by ``argparse``.
    """
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
        if args.include_images and not (args.json or args.json_unit):
            raise ValueError("--include-images requires --json or --json-unit")

        # Validate file path and size before processing
        file_path = Path(args.file)
        if not file_path.exists():
            raise FileNotFoundError(f"File not found: {args.file}")

        # Check file size (100MB limit for CLI to prevent memory issues)
        MAX_CLI_FILE_SIZE = 100 * 1024 * 1024  # 100MB
        file_size = file_path.stat().st_size
        if file_size > MAX_CLI_FILE_SIZE:
            raise ValueError(
                f"File size {file_size} bytes exceeds CLI maximum of {MAX_CLI_FILE_SIZE} bytes"
            )

        # Determine output stream
        output_stream: TextIO = sys.stdout
        output_file: TextIO | None = None
        if args.output:
            output_file = open(args.output, "w", encoding="utf-8")
            output_stream = output_file

        try:
            results = list(
                sharepoint2text.read_file(
                    args.file, ignore_images=not args.include_images
                )
            )
            if not results:
                raise RuntimeError(f"No extraction results for {args.file}")

            if args.json or args.json_unit:
                include_binary = bool(args.include_images)
                payload = (
                    _serialize_unit_results(results, include_binary=include_binary)
                    if args.json_unit
                    else _serialize_results(results, include_binary=include_binary)
                )
                json.dump(payload, output_stream)
                output_stream.write("\n")
            else:
                output_stream.write(_serialize_full_text(results))
                output_stream.write("\n")
        finally:
            if output_file is not None:
                output_file.close()

        return 0
    except (
        FileNotFoundError,
        PermissionError,
        ValueError,
        RuntimeError,
        OSError,
        TypeError,
    ) as exc:
        print(f"sharepoint2text: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
