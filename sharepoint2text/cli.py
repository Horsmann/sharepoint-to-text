from __future__ import annotations

import argparse
import itertools
import json
import os
import sys
from dataclasses import replace
from pathlib import Path
from typing import Iterator, Sequence, TextIO

import sharepoint2text
from sharepoint2text.parsing.extractors.util.zip_bomb import (
    _ZIP_BOMB_CHECKS_DISABLED,
    DEFAULT_ZIP_BOMB_LIMITS,
    ZipBombLimits,
)
from sharepoint2text.parsing.models import (
    BinaryMode,
    ContentUnit,
    ExtractedDocument,
    JsonValue,
    document_to_dict,
)

_MIN_ZIP_BOMB_LIMIT_MULTIPLIER = 2
_MAX_ZIP_BOMB_LIMIT_MULTIPLIER = 10
_DEFAULT_ZIP_BOMB_LIMIT_MULTIPLIER = 1
_DISABLED_ZIP_BOMB_LIMIT_VALUE = "none"

_CLI_DESCRIPTION = """\
Extract normalized text and structure from supported files.

Choose exactly one input source. Plain text is written by default; use --json
or --json-unit for the stable version-2 JSON schema."""

_CLI_EPILOG = """\
examples:
  sharepoint2text --file report.pdf
  sharepoint2text --file report.pdf --json --include-binary
  sharepoint2text --folder ./documents --suffixes .docx,.pdf
  sharepoint2text --folder ./documents --output ./extracted

Use --output with a file path to combine results, or with a directory path to
write one .txt or .json file per input file while preserving subdirectories."""


def _parse_zip_bomb_limit_multiplier(value: str) -> int | None:
    """Parse a ZIP-bomb limit multiplier supplied on the command line.

    Args:
        value: Integer text from 2 through 10, or ``none`` to disable checks.

    Returns:
        The validated multiplier, or ``None`` when checks should be disabled.

    Raises:
        argparse.ArgumentTypeError: If the value is outside the accepted forms.
    """
    if value.casefold() == _DISABLED_ZIP_BOMB_LIMIT_VALUE:
        return None

    try:
        multiplier = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            "must be a whole integer from 2 through 10, or 'none'"
        ) from exc

    if not (
        _MIN_ZIP_BOMB_LIMIT_MULTIPLIER <= multiplier <= _MAX_ZIP_BOMB_LIMIT_MULTIPLIER
    ):
        raise argparse.ArgumentTypeError(
            "must be a whole integer from 2 through 10, or 'none'"
        )
    return multiplier


def _build_zip_bomb_limits(multiplier: int | None) -> ZipBombLimits:
    """Build the CLI's uniformly scaled ZIP-bomb limits.

    Args:
        multiplier: Scale factor, or ``None`` to disable ZIP-bomb checks.

    Returns:
        Default limits with every threshold multiplied, or the internal
        disabled-checks marker.
    """
    if multiplier is None:
        return _ZIP_BOMB_CHECKS_DISABLED

    defaults = DEFAULT_ZIP_BOMB_LIMITS
    return ZipBombLimits(
        max_entries=defaults.max_entries * multiplier,
        max_total_uncompressed_bytes=(
            defaults.max_total_uncompressed_bytes * multiplier
        ),
        max_single_uncompressed_bytes=(
            defaults.max_single_uncompressed_bytes * multiplier
        ),
        max_total_compression_ratio=(defaults.max_total_compression_ratio * multiplier),
        max_entry_compression_ratio=(defaults.max_entry_compression_ratio * multiplier),
    )


def _build_parser() -> argparse.ArgumentParser:
    """Build the command-line argument parser.

    Returns:
        Parser containing every supported CLI option and its help text.
    """
    parser = argparse.ArgumentParser(
        prog="sharepoint2text",
        description=_CLI_DESCRIPTION,
        epilog=_CLI_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    _add_general_arguments(parser)
    _add_input_arguments(parser)
    _add_output_arguments(parser)
    _add_extraction_arguments(parser)
    _add_safety_arguments(parser)
    return parser


def _add_general_arguments(parser: argparse.ArgumentParser) -> None:
    """Add help and version options to a CLI parser.

    Args:
        parser: Parser that receives the general options.
    """
    group = parser.add_argument_group("general options")
    group.add_argument(
        "-h",
        "--help",
        action="help",
        help="Show this help information and exit.",
    )
    group.add_argument(
        "-v",
        "--version",
        action="version",
        version=f"%(prog)s {sharepoint2text.__version__}",
        help="Show the installed sharepoint2text version and exit.",
    )


def _add_input_arguments(parser: argparse.ArgumentParser) -> None:
    """Add input selection and folder traversal options to a CLI parser.

    Args:
        parser: Parser that receives the input options.
    """
    group = parser.add_argument_group("input selection")
    _add_input_source_arguments(group)
    _add_folder_filter_arguments(group)


def _add_input_source_arguments(group: argparse._ArgumentGroup) -> None:
    """Add the mutually exclusive file and folder arguments.

    Args:
        group: Argument group that receives the input-source options.
    """
    input_group = group.add_mutually_exclusive_group(required=True)
    input_group.add_argument(
        "-f",
        "--file",
        type=Path,
        metavar="FILE",
        help="Extract one existing supported file.",
    )
    input_group.add_argument(
        "-d",
        "--folder",
        type=Path,
        metavar="FOLDER",
        help=(
            "Extract supported files from a directory. Descends into "
            "subdirectories by default."
        ),
    )


def _add_folder_filter_arguments(group: argparse._ArgumentGroup) -> None:
    """Add options that control folder traversal.

    Args:
        group: Argument group that receives the folder-only options.
    """
    group.add_argument(
        "-s",
        "--suffixes",
        type=str,
        metavar="SUFFIX[,...]",
        help=(
            "Folder input only. Limit extraction to comma-separated suffixes "
            "such as .docx,.pdf,txt; leading dots are optional. When omitted, "
            "all supported file types are considered."
        ),
    )
    group.add_argument(
        "--no-recursive",
        dest="no_recursive",
        action="store_true",
        help=(
            "Folder input only. Inspect the selected directory without "
            "descending into subdirectories."
        ),
    )


def _add_output_arguments(parser: argparse.ArgumentParser) -> None:
    """Add output format and destination options to a CLI parser.

    Args:
        parser: Parser that receives the output options.
    """
    group = parser.add_argument_group("output format and destination")
    _add_json_output_arguments(group)
    _add_output_destination_argument(group)


def _add_json_output_arguments(group: argparse._ArgumentGroup) -> None:
    """Add mutually exclusive structured-output options.

    Args:
        group: Argument group that receives the JSON output options.
    """
    output_group = group.add_mutually_exclusive_group()
    output_group.add_argument(
        "-j",
        "--json",
        action="store_true",
        help=(
            "Write a JSON array with one version-2 extraction envelope per "
            "document. Binary payloads are omitted unless --include-binary is set."
        ),
    )
    output_group.add_argument(
        "-u",
        "--json-unit",
        dest="json_unit",
        action="store_true",
        help=(
            "Write a JSON array with one version-2 extraction envelope per "
            "content unit. Binary payloads are omitted unless --include-binary is set."
        ),
    )


def _add_output_destination_argument(group: argparse._ArgumentGroup) -> None:
    """Add the output destination option.

    Args:
        group: Argument group that receives the destination option.
    """
    group.add_argument(
        "-o",
        "--output",
        type=Path,
        metavar="PATH",
        help=(
            "Write to PATH instead of stdout. For folder input, a file path "
            "combines results; an existing directory or new extensionless path "
            "receives one .txt or .json file per input and preserves subdirectories."
        ),
    )


def _add_extraction_arguments(parser: argparse.ArgumentParser) -> None:
    """Add content-selection options to a CLI parser.

    Args:
        parser: Parser that receives the extraction options.
    """
    group = parser.add_argument_group("extraction options")
    _add_binary_extraction_argument(group)
    _add_attachment_extraction_argument(group)


def _add_binary_extraction_argument(group: argparse._ArgumentGroup) -> None:
    """Add the binary-payload extraction option.

    Args:
        group: Argument group that receives the binary-payload option.
    """
    group.add_argument(
        "-i",
        "--include-binary",
        dest="include_binary",
        action="store_true",
        help=(
            "Encode image and attachment payloads as base64. Image metadata is "
            "included even when binary payloads are omitted. "
            "Requires --json or --json-unit and can increase processing time and "
            "output size."
        ),
    )


def _add_attachment_extraction_argument(group: argparse._ArgumentGroup) -> None:
    """Add the email-attachment suppression option.

    Args:
        group: Argument group that receives the attachment option.
    """
    group.add_argument(
        "-n",
        "--no-attachments",
        dest="no_attachments",
        action="store_true",
        help=(
            "Do not extract supported email attachments or include their "
            "records in the output."
        ),
    )


def _add_safety_arguments(parser: argparse.ArgumentParser) -> None:
    """Add resource-limit options to a CLI parser.

    Args:
        parser: Parser that receives the safety options.
    """
    group = parser.add_argument_group("resource limits")
    _add_file_size_limit_argument(group)
    _add_zip_bomb_limit_argument(group)


def _add_file_size_limit_argument(group: argparse._ArgumentGroup) -> None:
    """Add the input file-size limit option.

    Args:
        group: Argument group that receives the file-size option.
    """
    group.add_argument(
        "-m",
        "--max-file-size-mb",
        type=float,
        default=100.0,
        metavar="MIB",
        help=(
            "Reject each input larger than this many MiB (default: 100). Use 0 "
            "to disable only the input file-size check."
        ),
    )


def _add_zip_bomb_limit_argument(group: argparse._ArgumentGroup) -> None:
    """Add the ZIP-bomb threshold scaling option.

    Args:
        group: Argument group that receives the archive safety option.
    """
    group.add_argument(
        "--zip-bomb-limit-multiplier",
        "--zblm",
        type=_parse_zip_bomb_limit_multiplier,
        default=_DEFAULT_ZIP_BOMB_LIMIT_MULTIPLIER,
        metavar="2..10|none",
        help=(
            "Scale every default ZIP-bomb threshold by a whole number from 2 "
            "through 10. Omit this option to keep the defaults, or use 'none' "
            "to disable the checks for trusted input only."
        ),
    )


def _serialize_results(
    results: list[ExtractedDocument], *, include_binary: bool
) -> list[dict[str, JsonValue]]:
    """Serialize documents for ``--json`` output.

    Args:
        results: Normalized documents to encode.
        include_binary: Encode binary payloads as base64 when true.

    Returns:
        Version-2 schema envelopes in source order.
    """
    binary: BinaryMode = "base64" if include_binary else "omit"
    return [document_to_dict(result, binary=binary) for result in results]


def _serialize_unit_results(
    results: list[ExtractedDocument], *, include_binary: bool
) -> list[dict[str, JsonValue]]:
    """Serialize one version-2 document envelope per content unit.

    Args:
        results: Normalized documents whose units should be emitted.
        include_binary: Encode binary payloads as base64 when true.

    Returns:
        Version-2 schema envelopes containing exactly one unit each.
    """
    binary: BinaryMode = "base64" if include_binary else "omit"
    serialized_units: list[dict[str, JsonValue]] = []
    for result in results:
        for unit in result.units:
            unit_document = _document_for_unit(result, unit)
            serialized_units.append(document_to_dict(unit_document, binary=binary))
    return serialized_units


def _document_for_unit(
    document: ExtractedDocument, unit: ContentUnit
) -> ExtractedDocument:
    """Return a self-contained document containing one selected unit.

    Args:
        document: Parent document whose document-level records are retained.
        unit: Single content unit to place in the returned document.

    Returns:
        A shallow copy containing only ``unit`` in its unit collection.
    """
    return replace(document, units=[unit])


def _serialize_full_text(results: list[ExtractedDocument]) -> str:
    """Join normalized document text using blank-line separators.

    Args:
        results: Normalized documents to render.

    Returns:
        Combined plain text with trailing whitespace removed.
    """
    return "\n\n".join(result.full_text.rstrip() for result in results).rstrip()


def _iter_serialized_results(
    results: Iterator[ExtractedDocument], *, include_binary: bool
) -> Iterator[dict[str, JsonValue]]:
    """Yield version-2 document envelopes.

    Args:
        results: Normalized documents to encode.
        include_binary: Encode binary payloads as base64 when true.

    Yields:
        Version-2 schema envelopes in source order.
    """
    binary: BinaryMode = "base64" if include_binary else "omit"
    for result in results:
        yield document_to_dict(result, binary=binary)


def _iter_serialized_unit_results(
    results: Iterator[ExtractedDocument], *, include_binary: bool
) -> Iterator[dict[str, JsonValue]]:
    """Yield one version-2 document envelope per content unit.

    Args:
        results: Normalized documents whose units should be emitted.
        include_binary: Encode binary payloads as base64 when true.

    Yields:
        Version-2 schema envelopes containing exactly one unit each.
    """
    binary: BinaryMode = "base64" if include_binary else "omit"
    for result in results:
        for unit in result.units:
            unit_document = _document_for_unit(result, unit)
            yield document_to_dict(unit_document, binary=binary)


def _write_json_array(
    items: Iterator[dict[str, JsonValue]], output_stream: TextIO
) -> None:
    """Write JSON objects as one streaming array.

    Args:
        items: JSON-compatible objects to write.
        output_stream: Text stream receiving the array.
    """
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
    results: Iterator[ExtractedDocument], output_stream: TextIO
) -> None:
    """Write normalized document text with blank-line separators.

    Args:
        results: Normalized documents to render.
        output_stream: Text stream receiving the output.
    """
    first = True
    for result in results:
        if not first:
            output_stream.write("\n\n")
        output_stream.write(result.full_text.rstrip())
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


def _write_results_to_file(
    results: list[ExtractedDocument],
    output_path: Path,
    args: argparse.Namespace,
) -> None:
    """Write every document from one source to a single output file.

    Args:
        results: Documents yielded from one source, in source order.
        output_path: The path to write to.
        args: CLI arguments for formatting options.

    Raises:
        ValueError: If ``results`` is empty.
    """
    if not results:
        raise ValueError("Cannot write an empty extraction result group")

    # Ensure parent directories exist
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        if args.json or args.json_unit:
            include_binary = bool(args.include_binary)
            if args.json_unit:
                payload = _serialize_unit_results(
                    results, include_binary=include_binary
                )
            else:
                payload = _serialize_results(results, include_binary=include_binary)
            json.dump(payload, f, indent=4)
            f.write("\n")
        else:
            f.write(_serialize_full_text(results))
            f.write("\n")


def _group_results_by_source(
    results: Iterator[ExtractedDocument],
) -> Iterator[tuple[Path, list[ExtractedDocument]]]:
    """Group adjacent extraction documents that came from the same source.

    Args:
        results: Documents yielded in source order by ``read_many``.

    Yields:
        Source paths paired with all documents yielded from that source.
    """

    def source_path(document: ExtractedDocument) -> Path:
        value = document.source.path or document.source.filename or "unknown"
        return Path(value)

    for path, grouped_results in itertools.groupby(results, key=source_path):
        yield path, list(grouped_results)


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

    results = sharepoint2text.read_many(
        input_folder,
        suffixes=suffixes,
        extract_all_supported=extract_all_supported,
        max_file_size=max_file_size_bytes,
        ignore_images=not args.include_binary,
        include_attachments=not args.no_attachments,
        recursive=not args.no_recursive,
        zip_bomb_limits=_build_zip_bomb_limits(args.zip_bomb_limit_multiplier),
    )
    for source_path, source_results in _group_results_by_source(results):
        # Compute output path
        output_path = _compute_output_path(
            source_path, input_folder, output_folder, extension
        )

        # Write every document yielded from the current source together.
        _write_results_to_file(source_results, output_path, args)
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
) -> Iterator[ExtractedDocument]:
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
            ignore_images=not args.include_binary,
            include_attachments=not args.no_attachments,
            zip_bomb_limits=_build_zip_bomb_limits(args.zip_bomb_limit_multiplier),
        )
    )


def _get_folder_results(
    args: argparse.Namespace, max_file_size_bytes: int
) -> Iterator[ExtractedDocument]:
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
            ignore_images=not args.include_binary,
            include_attachments=not args.no_attachments,
            recursive=not args.no_recursive,
            zip_bomb_limits=_build_zip_bomb_limits(args.zip_bomb_limit_multiplier),
        )
    )


def main(argv: Sequence[str] | None = None) -> int:
    """Run the CLI and return a process-style exit code.

    Args:
        argv: Optional argument list. If ``None``, arguments are read from
            ``sys.argv`` by ``argparse``.

    Returns:
        ``0`` on success, ``1`` on validation/extraction/serialization errors.
        Parser-driven early exits return the code produced by ``argparse``; invalid
        command syntax uses ``2``.
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
        if args.include_binary and not (args.json or args.json_unit):
            raise ValueError("--include-binary requires --json or --json-unit")
        if args.max_file_size_mb < 0:
            raise ValueError("--max-file-size-mb must be >= 0")
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
                include_binary = bool(args.include_binary)
                payload_items = (
                    _iter_serialized_unit_results(
                        results,
                        include_binary=include_binary,
                    )
                    if args.json_unit
                    else _iter_serialized_results(
                        results,
                        include_binary=include_binary,
                    )
                )
                _write_json_array(payload_items, output_stream)
            else:
                _write_full_text(results, output_stream)
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
        sharepoint2text.ExtractionError,
    ) as exc:
        print(f"sharepoint2text: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
