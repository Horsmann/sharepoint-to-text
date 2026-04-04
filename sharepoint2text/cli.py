from __future__ import annotations

import argparse
import itertools
import json
import os
import sys
from pathlib import Path
from typing import Iterator, Sequence, TextIO

import sharepoint2text
from sharepoint2text.parsing.extractors.data_types import (
    EmailContent,
    ExtractionInterface,
)
from sharepoint2text.parsing.extractors.serialization import serialize_extraction


def _build_parser() -> argparse.ArgumentParser:
    """Build and return the command-line argument parser for the CLI."""
    parser = argparse.ArgumentParser(
        prog="sharepoint2text",
        description="Extract file content and emit full text to stdout (or JSON with --json/--json-unit).",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {sharepoint2text.__version__}",
        help="Show the version and exit.",
    )

    # Input source: either a single file or a folder
    input_group = parser.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        "-f",
        "--file",
        type=Path,
        help="Path to a single file to extract.",
    )
    input_group.add_argument(
        "-d",
        "--folder",
        type=Path,
        help="Path to a folder to extract files from (recursive by default).",
    )
    parser.add_argument(
        "-s",
        "--suffixes",
        type=str,
        help=(
            "Comma-separated list of file suffixes to extract when using --folder "
            "(e.g., '.docx,.pdf,.txt'). If omitted, all supported file types are extracted."
        ),
    )
    parser.add_argument(
        "--no-recursive",
        dest="no_recursive",
        action="store_true",
        help="When using --folder, only extract files in the top-level directory (no subdirectories).",
    )

    output_group = parser.add_mutually_exclusive_group()
    output_group.add_argument(
        "-j",
        "--json",
        action="store_true",
        help="Emit structured JSON instead of plain full text (omits binary payloads by default).",
    )
    output_group.add_argument(
        "-u",
        "--json-unit",
        dest="json_unit",
        action="store_true",
        help="Emit JSON for extracted text units instead of full extraction objects (omits binary payloads by default).",
    )
    parser.add_argument(
        "-i",
        "--include-images",
        dest="include_images",
        action="store_true",
        help="Extract images from the file and include image data as base64 blobs in JSON output (default: images are ignored for faster processing).",
    )
    parser.add_argument(
        "-n",
        "--no-attachments",
        dest="no_attachments",
        action="store_true",
        help="For email files, exclude supported attachments from CLI extraction output.",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        help=(
            "Output path (default: stdout). "
            "For single file extraction: specify a file path. "
            "For folder extraction: specify either a file (all results combined) "
            "or a folder (each file written separately, mirroring input structure)."
        ),
    )
    parser.add_argument(
        "-m",
        "--max-file-size-mb",
        type=float,
        default=100.0,
        help=(
            "Maximum input file size in MiB (default: 100). "
            "Use 0 to disable size checks."
        ),
    )
    parser.add_argument(
        "-t",
        "--timeout",
        type=float,
        default=60.0,
        help=(
            "Maximum extraction time per file in seconds (default: 60). "
            "Use 0 to disable timeout enforcement."
        ),
    )
    return parser


def _serialize_results(
    results: list[ExtractionInterface], *, include_binary: bool
) -> list[dict]:
    """Serialize extraction results for ``--json`` output.

    Always returns a list so output shape is stable regardless of result
    cardinality (single file, archives, mbox, etc.).
    """
    return [
        serialize_extraction(result, include_binary=include_binary)
        for result in results
    ]


def _serialize_unit_results(
    results: list[ExtractionInterface],
    *,
    include_binary: bool,
    include_email_attachments: bool = False,
) -> list[dict]:
    """Serialize per-unit output for ``--json-unit`` mode.

    Returns a flat ``list[dict]`` with one dictionary per extracted unit.
    """
    return [
        serialize_extraction(unit, include_binary=include_binary)
        for result in results
        for extraction in _iter_result_tree(
            result, include_email_attachments=include_email_attachments
        )
        for unit in extraction.iterate_units()
    ]


def _serialize_full_text(results: list[ExtractionInterface]) -> str:
    """Join ``get_full_text()`` from all results using blank-line separators."""
    return "\n\n".join(result.get_full_text().rstrip() for result in results).rstrip()


def _expand_email_results(
    results: list[ExtractionInterface],
) -> list[ExtractionInterface]:
    """Expand email results with any supported extracted attachments."""
    expanded: list[ExtractionInterface] = []
    for result in results:
        expanded.extend(_iter_result_tree(result, include_email_attachments=True))
    return expanded


def _iter_result_tree(
    result: ExtractionInterface, *, include_email_attachments: bool
) -> Iterator[ExtractionInterface]:
    """Yield a root result and optionally nested supported email attachments."""
    yield result
    if not include_email_attachments or not isinstance(result, EmailContent):
        return
    for attachment in result.iterate_supported_attachments():
        yield from _iter_result_tree(attachment, include_email_attachments=True)


def _strip_email_attachments(results: list[ExtractionInterface]) -> None:
    """Remove parsed attachment metadata/payloads from email results in-place."""
    for result in results:
        if isinstance(result, EmailContent):
            result.attachments = []


def _iter_expanded_results(
    results: Iterator[ExtractionInterface], *, include_email_attachments: bool
) -> Iterator[ExtractionInterface]:
    for result in results:
        yield from _iter_result_tree(
            result, include_email_attachments=include_email_attachments
        )


def _iter_serialized_results(
    results: Iterator[ExtractionInterface], *, include_binary: bool
) -> Iterator[dict]:
    for result in results:
        yield serialize_extraction(result, include_binary=include_binary)


def _iter_serialized_unit_results(
    results: Iterator[ExtractionInterface],
    *,
    include_binary: bool,
    include_email_attachments: bool = False,
) -> Iterator[dict]:
    for result in results:
        for extraction in _iter_result_tree(
            result, include_email_attachments=include_email_attachments
        ):
            for unit in extraction.iterate_units():
                yield serialize_extraction(unit, include_binary=include_binary)


def _write_json_array(items: Iterator[dict], output_stream: TextIO) -> None:
    output_stream.write("[")
    first = True
    for item in items:
        if not first:
            output_stream.write(",")
        output_stream.write("\n")
        output_stream.write(json.dumps(item, indent=4))
        first = False
    if not first:
        output_stream.write("\n")
    output_stream.write("]\n")


def _write_full_text(
    results: Iterator[ExtractionInterface], output_stream: TextIO
) -> None:
    first = True
    for result in results:
        if not first:
            output_stream.write("\n\n")
        output_stream.write(result.get_full_text().rstrip())
        first = False
    output_stream.write("\n")


def _get_output_extension(args: argparse.Namespace) -> str:
    """Determine the output file extension based on output format."""
    if args.json or args.json_unit:
        return ".json"
    return ".txt"


def _compute_output_path(
    source_path: Path,
    input_folder: Path,
    output_folder: Path,
    extension: str,
) -> Path:
    """Compute the output path for a file, mirroring the input folder structure.

    Args:
        source_path: The original source file path.
        input_folder: The input folder being extracted.
        output_folder: The output folder to write to.
        extension: The output file extension (.txt or .json).

    Returns:
        The computed output path within the output folder.
    """
    # Get the relative path from the input folder
    try:
        relative_path = source_path.relative_to(input_folder)
    except ValueError:
        # If source_path is not relative to input_folder, use just the filename
        relative_path = Path(source_path.name)

    # Change the extension
    output_name = relative_path.with_suffix(extension)

    return output_folder / output_name


def _write_single_result_to_file(
    result: ExtractionInterface,
    output_path: Path,
    args: argparse.Namespace,
) -> None:
    """Write a single extraction result to a file.

    Args:
        result: The extraction result to write.
        output_path: The path to write to.
        args: CLI arguments for formatting options.
    """
    # Ensure parent directories exist
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        if args.json or args.json_unit:
            include_binary = bool(args.include_images)
            if args.json_unit:
                # Write units as JSON array
                units = list(result.iterate_units())
                payload = [
                    serialize_extraction(unit, include_binary=include_binary)
                    for unit in units
                ]
            else:
                # Write full extraction as JSON (wrapped in array for consistency)
                expanded = list(
                    _iter_result_tree(
                        result, include_email_attachments=not args.no_attachments
                    )
                )
                payload = [
                    serialize_extraction(r, include_binary=include_binary)
                    for r in expanded
                ]
            json.dump(payload, f, indent=4)
            f.write("\n")
        else:
            # Write plain text
            expanded = list(
                _iter_result_tree(
                    result, include_email_attachments=not args.no_attachments
                )
            )
            text_parts = [r.get_full_text().rstrip() for r in expanded]
            f.write("\n\n".join(text_parts))
            f.write("\n")


def _process_folder_to_folder(
    args: argparse.Namespace,
    max_file_size_bytes: int,
    input_folder: Path,
    output_folder: Path,
) -> int:
    """Process folder extraction with per-file output.

    Args:
        args: CLI arguments.
        max_file_size_bytes: Maximum file size limit.
        input_folder: The input folder to extract from.
        output_folder: The output folder to write to.

    Returns:
        Number of files successfully extracted.
    """
    extension = _get_output_extension(args)
    files_written = 0

    # Parse suffixes if provided
    suffixes: list[str] | None = None
    extract_all_supported = True
    if args.suffixes:
        suffixes = _parse_suffixes(args.suffixes)
        if not suffixes:
            raise ValueError("--suffixes must contain at least one valid suffix")
        extract_all_supported = False

    # Use read_many but process each result individually
    for result in sharepoint2text.read_many(
        input_folder,
        suffixes=suffixes,
        extract_all_supported=extract_all_supported,
        max_file_size=max_file_size_bytes,
        ignore_images=not args.include_images,
        include_attachments=not args.no_attachments,
        recursive=not args.no_recursive,
    ):
        # Get the source file path from metadata
        metadata = result.get_metadata()
        source_path_str = metadata.file_path
        if not source_path_str:
            # Fallback to filename if full path not available
            source_path_str = metadata.filename or "unknown"

        source_path = Path(source_path_str)

        # Compute output path
        output_path = _compute_output_path(
            source_path, input_folder, output_folder, extension
        )

        # Write the result
        _write_single_result_to_file(result, output_path, args)
        files_written += 1
        print(f"Extracted: {source_path} -> {output_path}", file=sys.stderr)

    return files_written


def _parse_suffixes(suffixes_str: str) -> list[str]:
    """Parse comma-separated suffixes string into a list of normalized suffixes."""
    suffixes = []
    for suffix in suffixes_str.split(","):
        suffix = suffix.strip().lower()
        if suffix:
            if not suffix.startswith("."):
                suffix = f".{suffix}"
            suffixes.append(suffix)
    return suffixes


def _get_file_results(
    args: argparse.Namespace, max_file_size_bytes: int
) -> Iterator[ExtractionInterface]:
    """Get extraction results for a single file."""
    file_path = Path(args.file)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {args.file}")

    file_size = file_path.stat().st_size
    if max_file_size_bytes > 0 and file_size > max_file_size_bytes:
        raise ValueError(
            "File size "
            f"{file_size} bytes exceeds CLI maximum of {max_file_size_bytes} bytes"
        )

    return iter(
        sharepoint2text.read_file(
            args.file,
            max_file_size=max_file_size_bytes,
            timeout_seconds=args.timeout,
            ignore_images=not args.include_images,
            include_attachments=not args.no_attachments,
        )
    )


def _get_folder_results(
    args: argparse.Namespace, max_file_size_bytes: int
) -> Iterator[ExtractionInterface]:
    """Get extraction results for a folder."""
    folder_path = Path(args.folder)
    if not folder_path.exists():
        raise FileNotFoundError(f"Folder not found: {args.folder}")
    if not folder_path.is_dir():
        raise NotADirectoryError(f"Path is not a directory: {args.folder}")

    # Parse suffixes if provided
    suffixes: list[str] | None = None
    extract_all_supported = True
    if args.suffixes:
        suffixes = _parse_suffixes(args.suffixes)
        if not suffixes:
            raise ValueError("--suffixes must contain at least one valid suffix")
        extract_all_supported = False

    return iter(
        sharepoint2text.read_many(
            folder_path,
            suffixes=suffixes,
            extract_all_supported=extract_all_supported,
            max_file_size=max_file_size_bytes,
            timeout_seconds=args.timeout,
            ignore_images=not args.include_images,
            include_attachments=not args.no_attachments,
            recursive=not args.no_recursive,
        )
    )


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
        if args.max_file_size_mb < 0:
            raise ValueError("--max-file-size-mb must be >= 0")
        if args.timeout < 0:
            raise ValueError("--timeout must be >= 0")
        if args.suffixes and not args.folder:
            raise ValueError("--suffixes can only be used with --folder")
        if args.no_recursive and not args.folder:
            raise ValueError("--no-recursive can only be used with --folder")

        max_file_size_bytes = int(args.max_file_size_mb * 1024 * 1024)

        # Check for folder-to-folder output mode
        if args.folder and args.output and args.output.is_dir():
            # Folder extraction with folder output: write each file separately
            files_written = _process_folder_to_folder(
                args, max_file_size_bytes, args.folder, args.output
            )
            if files_written == 0:
                raise RuntimeError(f"No extraction results for folder: {args.folder}")
            print(
                f"Successfully extracted {files_written} file(s) to {args.output}",
                file=sys.stderr,
            )
            return 0

        # Standard mode: single output stream (stdout, file, or combined folder output)
        output_stream: TextIO = sys.stdout
        output_file: TextIO | None = None
        if args.output:
            # If output path doesn't exist and we're in folder mode,
            # check if user wants folder output (path ends with separator or has no extension)
            if args.folder and not args.output.exists():
                # Heuristic: if path has no extension or ends with separator, treat as folder
                if (
                    not args.output.suffix
                    or str(args.output).endswith(os.sep)
                    or str(args.output).endswith("/")
                ):
                    # Create as folder and use folder-to-folder mode
                    args.output.mkdir(parents=True, exist_ok=True)
                    files_written = _process_folder_to_folder(
                        args, max_file_size_bytes, args.folder, args.output
                    )
                    if files_written == 0:
                        raise RuntimeError(
                            f"No extraction results for folder: {args.folder}"
                        )
                    print(
                        f"Successfully extracted {files_written} file(s) to {args.output}",
                        file=sys.stderr,
                    )
                    return 0

            # Otherwise treat as file output
            output_file = open(args.output, "w", encoding="utf-8")
            output_stream = output_file

        try:
            # Get extraction results based on input type (file or folder)
            if args.folder:
                results = _get_folder_results(args, max_file_size_bytes)
            else:
                results = _get_file_results(args, max_file_size_bytes)

            first_result = next(results, None)
            if first_result is None:
                if args.folder:
                    raise RuntimeError(
                        f"No extraction results for folder: {args.folder}"
                    )
                raise RuntimeError(f"No extraction results for {args.file}")
            results = itertools.chain([first_result], results)

            if args.json or args.json_unit:
                include_binary = bool(args.include_images)
                payload_items = (
                    _iter_serialized_unit_results(
                        results,
                        include_binary=include_binary,
                        include_email_attachments=not args.no_attachments,
                    )
                    if args.json_unit
                    else _iter_serialized_results(
                        _iter_expanded_results(
                            results,
                            include_email_attachments=not args.no_attachments,
                        ),
                        include_binary=include_binary,
                    )
                )
                _write_json_array(payload_items, output_stream)
            else:
                _write_full_text(
                    _iter_expanded_results(
                        results,
                        include_email_attachments=not args.no_attachments,
                    ),
                    output_stream,
                )
        finally:
            if output_file is not None:
                output_file.close()

        return 0
    except (
        FileNotFoundError,
        NotADirectoryError,
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
