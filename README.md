# sharepoint-to-text

`sharepoint-to-text` is a typed, pure-Python library for extracting text and structured content from file types commonly found in SharePoint and document-management workflows.

It is built for software engineers who need one extraction interface across modern Microsoft Office files, legacy Office files, OpenDocument, PDF, email, HTML-like content, plain-text formats, and archives.

## Why This Package Exists

Document ingestion pipelines usually fail in one of two ways:

- they only support a narrow set of office formats
- they require heavyweight external runtimes such as LibreOffice, Java, or platform-specific tooling

`sharepoint-to-text` takes a different approach:

- Pure Python library API and CLI
- One routing layer for many file types
- Works with file paths and in-memory bytes
- Typed extraction objects with metadata, units, images, and tables
- Suitable for indexing, RAG, ETL, compliance review, and migration tooling

## At A Glance

| Item | Details |
|---|---|
| Python | `>=3.10` |
| Install | `uv add sharepoint-to-text` |
| Runtime model | Pure Python |
| Primary interfaces | `read_file(...)`, `read_bytes(...)`, `read_many(...)`, CLI |
| Output model | Generator of typed extraction objects |
| SharePoint access | Optional `sharepoint_io` helper for Graph-backed listing/download |

## Who This Is For

This package is a good fit if you need to:

- normalize text extraction across many enterprise document formats
- process documents from disk, APIs, queues, or object storage
- preserve some document structure for downstream chunking or citations
- run extraction in Python-only environments such as services, workers, or serverless jobs

It is not a full document rendering engine, OCR system, or layout-preserving conversion tool.

## Installation

### Package install

```bash
uv add sharepoint-to-text
```

With `pip`:

```bash
pip install sharepoint-to-text
```

### Development install

```bash
git clone https://github.com/Horsmann/sharepoint-to-text.git
cd sharepoint-to-text
uv sync --all-groups
```

## Quick Start

### Extract from a local file

```python
import sharepoint2text

result = next(sharepoint2text.read_file("document.docx"))
print(result.get_full_text())
```

### Extract from in-memory bytes

```python
import sharepoint2text

payload = b"hello from memory"
result = next(sharepoint2text.read_bytes(payload, extension="txt"))
print(result.get_full_text())
```

### Use structural units for chunking

```python
import sharepoint2text

result = next(sharepoint2text.read_file("report.pdf", ignore_images=True))

for unit in result.iterate_units():
    print(unit.get_text())
    print(unit.get_metadata())
```

### Batch extraction from a folder

```python
import sharepoint2text

# Extract only Word and PDF files from a folder
for result in sharepoint2text.read_many("docs", suffixes=[".docx", ".pdf"]):
    print(result.get_metadata().file_path)
    print(result.get_full_text()[:200])
```

### Extract all supported files from a folder

```python
import sharepoint2text

# Extract all supported file formats recursively
for result in sharepoint2text.read_many("docs", extract_all_supported=True):
    for unit in result.iterate_units(ignore_images=True):
        text = unit.get_text().strip()
        if text:
            print(result.get_metadata().file_path, text[:120])
```

### Parse an email and its attachments

```python
import sharepoint2text
from sharepoint2text import EmailContent

# .eml and .msg return an EmailContent object
email: EmailContent = next(sharepoint2text.read_file("message.msg"))

print(email.subject)
print(email.from_email.address)
print([recipient.address for recipient in email.to_emails])
print(email.get_full_text())

# Recursively extract text from supported attachments (docx, pdf, xlsx, ...)
for attachment in email.iterate_supported_attachments(skip_failed=True):
    print(attachment.get_metadata().file_path, attachment.get_full_text()[:200])
```

> `.mbox` files instead yield one `EmailContent` per message, so iterate them
> with a `for` loop rather than `next(...)`.

## Core API

### Main entry points

```python
import sharepoint2text

sharepoint2text.read_file(
    path,
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
    force_plain_text=False,
    include_attachments=True,
)

sharepoint2text.read_bytes(
    data,
    extension="pdf",      # or ".pdf"
    mime_type=None,        # for example "application/pdf"
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
    force_plain_text=False,
    include_attachments=True,
)

sharepoint2text.read_many(
    folder_path,
    suffixes=[".docx", ".pdf"],  # list of extensions to extract
    extract_all_supported=False,  # or True to extract all supported formats
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
    force_plain_text=False,
    include_attachments=True,
    recursive=True,               # traverse subdirectories
)

sharepoint2text.is_supported_file(path)
sharepoint2text.get_extractor(path)
```

### Batch extraction with `read_many`

The `read_many` function extracts content from multiple files in a folder:

| Parameter | Description |
|---|---|
| `folder_path` | Path to the folder to traverse |
| `suffixes` | List of file extensions to extract (e.g., `[".docx", ".pdf"]`) |
| `extract_all_supported` | If `True`, extract all supported formats, or every file when combined with `force_plain_text=True` (mutually exclusive with `suffixes`) |
| `force_plain_text` | Treat selected files as plain text, including unknown extensions in `extract_all_supported` mode |
| `recursive` | If `True` (default), traverse subdirectories |

Configuration rules:
- You must specify either `suffixes` or `extract_all_supported=True`
- Specifying both raises `InvalidConfigurationError`
- Suffixes are normalized (with or without leading dot)
- Extraction continues on errors, logging warnings for failed files

### Result model

All extracted results implement a common interface:

- `get_full_text()`
- `iterate_units()`
- `iterate_images()`
- `iterate_tables()`
- `get_metadata()`
- `to_json()` / `from_json(...)`

Use `get_full_text()` when you want one string per extraction result.

Use `iterate_units()` when you want coarse structural chunks such as:

- one page per PDF unit
- one slide per presentation unit
- one sheet per spreadsheet unit
- one document-level unit for most text-document formats

### Generator semantics matter

The API returns generators because some inputs can produce multiple results:

- archives can yield one result per supported member file
- `.mbox` can yield one result per email
- email extraction can recursively expose supported attachments

For single-document formats, `next(...)` is usually the simplest call pattern.

## CLI

The package installs a `sharepoint2text` command.

### Single file extraction

Plain text output:

```bash
sharepoint2text --file /path/to/file.docx
```

Full extraction objects as JSON:

```bash
sharepoint2text --file /path/to/file.docx --json
```

Per-unit JSON:

```bash
sharepoint2text --file /path/to/file.pdf --json-unit
```

### Folder extraction

Extract all supported files from a folder:

```bash
sharepoint2text --folder /path/to/folder
```

Extract only specific file types:

```bash
sharepoint2text --folder /path/to/folder --suffixes .docx,.pdf,.txt
```

Non-recursive (top-level only):

```bash
sharepoint2text --folder /path/to/folder --no-recursive
```

### Folder output (mirrored structure)

When extracting from a folder, output to another folder to preserve the directory structure:

```bash
# Write each file separately to output folder
sharepoint2text --folder /input/docs --output /output/extracted/

# The output structure mirrors the input:
# /input/docs/report.docx      -> /output/extracted/report.txt
# /input/docs/sub/data.xlsx    -> /output/extracted/sub/data.txt
```

Output path behavior:
- If `--output` is an existing directory, files are written separately
- If `--output` is a new path without extension, it's created as a directory
- If `--output` has a file extension, all results are combined into that file

### CLI options

| Option | Description |
|---|---|
| `--file FILE`, `-f FILE` | Path to a single file to extract |
| `--folder FOLDER`, `-d FOLDER` | Path to a folder to extract files from (recursive by default) |
| `--suffixes SUFFIXES`, `-s SUFFIXES` | Comma-separated file suffixes to filter (e.g., `.docx,.pdf`). Only with `--folder`. If omitted, extracts all supported types. |
| `--no-recursive` | Only extract top-level files (no subdirectories). Only with `--folder`. |
| `--output PATH`, `-o PATH` | Output path: file (combined) or folder (separate files mirroring input structure) |
| `--json`, `-j` | Emit `list[extraction_object]` |
| `--json-unit`, `-u` | Emit `list[unit_object]` |
| `--include-images`, `-i` | Include base64 image payloads in JSON output |
| `--no-attachments`, `-n` | Skip expanding supported email attachments |
| `--max-file-size-mb`, `-m` | Maximum input size in MiB, default `100`, use `0` to disable |
| `--version`, `-v` | Print CLI version |

Important CLI rules:

- `--file` and `--folder` are mutually exclusive (one is required)
- `--suffixes` and `--no-recursive` only work with `--folder`
- `--json` and `--json-unit` are mutually exclusive
- `--include-images` requires `--json` or `--json-unit`
- the CLI enforces the same file-size guard as the Python API

## Supported Formats

### Microsoft Office

- Modern: `.docx`, `.docm`, `.xlsx`, `.xlsm`, `.xlsb`, `.pptx`, `.pptm`
- Legacy: `.doc`, `.dot`, `.xls`, `.xlt`, `.ppt`, `.pot`, `.pps`, `.rtf`
- Alias mapping: `.dotx`, `.dotm`, `.xltx`, `.xltm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`

### OpenDocument

- `.odt`, `.ods`, `.odp`, `.odg`, `.odf`
- Alias mapping: `.ott`, `.ots`, `.otp`

### Email

- `.eml`, `.msg`, `.mbox`
- `.eml` and `.msg` can parse and expose supported attachments
- `.mbox` yields one result per message

### Plain text and data-like formats

- `.txt`, `.md`, `.csv`, `.tsv`, `.json`
- `.yaml`, `.yml`, `.xml`, `.log`, `.ini`, `.cfg`, `.conf`, `.properties`

### Web and ebook

- `.html`, `.htm`, `.mhtml`, `.mht`, `.epub`

### PDF

- `.pdf`

### Archives

- `.zip`, `.tar`, `.7z`
- `.tar.gz`, `.tgz`, `.tar.bz2`, `.tbz2`, `.tar.xz`, `.txz`

For a behavior-focused view of units, attachments, and caveats by format, see [doc/format-matrix.md](doc/format-matrix.md).

## Increasing the ZIP Bomb Limit

Every ZIP-based format (OOXML `.docx`/`.xlsx`/`.pptx`, OpenDocument, and `.zip`
archives) is checked against ZIP-bomb heuristics. Legitimate but very large
SharePoint exports can occasionally trip these guards and raise
`ExtractionZipBombError`. When you trust the input, raise the relevant
threshold **once at application startup** with `set_zip_bomb_limits`.

```python
import sharepoint2text
from sharepoint2text import ZipBombLimits

# Raise the limits once, before any extraction runs.
# Only set the fields you need — unspecified fields keep the library defaults.
sharepoint2text.set_zip_bomb_limits(
    ZipBombLimits(
        max_total_uncompressed_bytes=16 * 1024 * 1024 * 1024,  # 16 GiB (default 4 GiB)
        max_single_uncompressed_bytes=4 * 1024 * 1024 * 1024,  # 4 GiB (default 1 GiB)
        max_entry_compression_ratio=1500.0,                    # default 500.0
    )
)

# Subsequent extractions use the raised limits automatically.
result = next(sharepoint2text.read_file("large_trusted_export.zip"))
print(result.get_full_text()[:200])
```

Available `ZipBombLimits` fields and their defaults:

| Field | Default | Meaning |
|---|---|---|
| `max_entries` | `50_000` | Maximum number of entries in a container |
| `max_total_uncompressed_bytes` | `4 GiB` | Maximum combined uncompressed size |
| `max_single_uncompressed_bytes` | `1 GiB` | Maximum uncompressed size of one entry |
| `max_total_compression_ratio` | `200.0` | Maximum overall compression ratio |
| `max_entry_compression_ratio` | `500.0` | Maximum compression ratio of one entry |

Related helpers, all importable from `sharepoint2text`:

```python
from sharepoint2text import (
    ZipBombLimits,          # the limits dataclass
    get_zip_bomb_limits,    # inspect the currently active limits
    reset_zip_bomb_limits,  # restore the library defaults
    set_zip_bomb_limits,    # override the process-wide limits
)
```

> The limits are process-wide and apply to every subsequent ZIP-based
> extraction. Only raise them for input you trust, since these guards are a
> deliberate denial-of-service protection.

## SharePoint Integration

The extraction library works independently of SharePoint. The optional `sharepoint_io` module is a separate helper layer for listing and downloading files through Microsoft Graph before extraction.

```python
import io

import sharepoint2text
from sharepoint2text.sharepoint_io import (
    EntraIDAppCredentials,
    SharePointRestClient,
)

credentials = EntraIDAppCredentials(
    tenant_id="your-tenant-id",
    client_id="your-client-id",
    client_secret="your-client-secret",
)

client = SharePointRestClient(
    site_url="https://contoso.sharepoint.com/sites/Documents",
    credentials=credentials,
)

for file_meta in client.list_all_files():
    data = client.download_file(file_meta.id)
    extractor = sharepoint2text.get_extractor(file_meta.name)
    for result in extractor(io.BytesIO(data), path=file_meta.name):
        print(result.get_full_text()[:200])
```

Setup details live in [sharepoint2text/sharepoint_io/SETUP.md](sharepoint2text/sharepoint_io/SETUP.md).

## Operational Constraints

These are the points an engineering team usually needs before adopting the package:

- No OCR: scanned-image PDFs will often produce little or no text
- No external office renderer: output is extraction-oriented, not fidelity-oriented
- Word-like formats do not expose reliable page boundaries
- Nested archives are intentionally skipped
- Password-protected or encrypted inputs raise extraction errors
- Large files and highly compressed archives are guarded by size limits and zip-bomb protections

### Archive behavior

- archives are processed one level deep
- supported non-archive files inside an archive can yield extraction results
- nested archives are skipped as a safety measure
- 7z extraction is capped at 100 MB internally

### Performance guidance

- set `ignore_images=True` when image payloads are not needed
- use `iterate_units()` for chunk-wise downstream processing instead of materializing one large string when structure matters
- keep size limits enabled unless you trust the input source

## Failure Modes and Exceptions

Common exceptions:

- `ExtractionFileFormatNotSupportedError`
- `ExtractionFileEncryptedError`
- `ExtractionFileTooLargeError`
- `ExtractionLegacyMicrosoftParsingError`
- `ExtractionZipBombError`
- `ExtractionPathTraversalError`
- `ExtractionFailedError`
- `InvalidConfigurationError` (for `read_many` with conflicting options)

If you are integrating this into a service, see [doc/integration-guide.md](doc/integration-guide.md) and [doc/troubleshooting.md](doc/troubleshooting.md).

## Serialization

```python
import json

import sharepoint2text

result = next(sharepoint2text.read_file("document.docx"))
payload = result.to_json()

print(json.dumps(payload))
```

Restore from JSON:

```python
from sharepoint2text.parsing.extractors.data_types import ExtractionInterface

restored = ExtractionInterface.from_json(payload)
```

## Additional Documentation

- [doc/cli.md](doc/cli.md): complete CLI reference with examples
- [doc/direct-extractors.md](doc/direct-extractors.md): call format-specific extractors directly and work with concrete result attributes
- [doc/format-matrix.md](doc/format-matrix.md): per-format behavior, units, and caveats
- [doc/improvements.md](doc/improvements.md): roadmap and improvement ideas
- [CONTRIBUTING.md](CONTRIBUTING.md): contributor workflow
- [CHANGELOG.md](CHANGELOG.md): release history

## License

Apache 2.0. See [LICENSE](LICENSE).

## Disclaimer

This project is not affiliated with, endorsed by, or sponsored by Microsoft.
