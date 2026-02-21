# sharepoint-to-text

A pure-Python library for extracting text and structured content from files commonly found in SharePoint ecosystems:

- Microsoft Office (modern and legacy)
- OpenDocument
- PDF
- Email formats
- Plain text and config formats
- HTML/EPUB/MHTML
- Archives containing supported files

It also includes an optional SharePoint Graph client (`sharepoint_io`) for listing/downloading files before extraction.

## Table of Contents

- [Why Use This Library](#why-use-this-library)
- [Install](#install)
- [Quick Start](#quick-start)
- [Core Interface](#core-interface)
- [CLI](#cli)
- [Optional SharePoint Integration](#optional-sharepoint-integration)
- [Supported Formats](#supported-formats)
- [Archive Processing and Security](#archive-processing-and-security)
- [Limitations and Caveats](#limitations-and-caveats)
- [API Cheat Sheet](#api-cheat-sheet)
- [Exceptions](#exceptions)
- [License](#license)
- [Disclaimer](#disclaimer)
- [More Usage Examples](#more-usage-examples)

## Why Use This Library

- Pure Python (no Java runtime, no LibreOffice subprocesses)
- Unified extraction interface across many file types
- Works with file paths and in-memory bytes
- Suitable for RAG/indexing pipelines where chunking and metadata matter
- Handles both modern and legacy Office formats in one API

## Install

```bash
uv add sharepoint-to-text
```

Optional PDF crypto acceleration:

```bash
uv add "sharepoint-to-text[pdf-crypto]"
```

From source:

```bash
git clone https://github.com/Horsmann/sharepoint-to-text.git
cd sharepoint-to-text
uv sync --all-groups
```

## Quick Start

### 1) Read any supported local file

```python
import sharepoint2text

result = next(sharepoint2text.read_file("document.docx"))
print(result.get_full_text())
```

`read_file(...)` returns a generator. Most files produce one result, but archives and `.mbox` can produce multiple.

### 2) Read bytes already in memory

```python
import sharepoint2text

payload = b"hello from memory"
result = next(sharepoint2text.read_bytes(payload, extension="txt"))
print(result.get_full_text())
```

### 3) Choose chunking strategy

```python
import sharepoint2text

result = next(sharepoint2text.read_file("report.pdf"))

# Single text blob
full_text = result.get_full_text()

# Structured chunks (page/slide/sheet depending on format)
for unit in result.iterate_units():
    print(unit.get_text())
    print(unit.get_metadata())
```

### 4) Serialize results

```python
import json
import sharepoint2text

result = next(sharepoint2text.read_file("document.docx"))
print(json.dumps(result.to_json()))
```

Restore from JSON:

```python
from sharepoint2text.parsing.extractors.data_types import ExtractionInterface

restored = ExtractionInterface.from_json(result.to_json())
```

## Core Interface

All extracted results implement a common interface (`ExtractionInterface`):

- `get_full_text()`
- `iterate_units()`
- `iterate_images()`
- `iterate_tables()`
- `get_metadata()`
- `to_json()` / `from_json(...)`

Use this interface when you want one pipeline that works across formats.

### Which text method should you use?

| Goal | Method |
|---|---|
| One string per document | `get_full_text()` |
| Chunk by structure (RAG/citations) | `iterate_units()` |
| All images in a file | `iterate_images()` |
| All tables in a file | `iterate_tables()` |

### What `iterate_units()` means by format

| Format family | Units yielded |
|---|---|
| Word / text docs (`.docx`, `.doc`, `.odt`, plain text, config files) | Usually one unit |
| Spreadsheets (`.xlsx`, `.xls`, `.ods`) | One unit per sheet |
| Presentations (`.pptx`, `.ppt`, `.odp`) | One unit per slide |
| PDF | One unit per page |
| Email (`.eml`, `.msg`) | One unit per email |
| Mailbox (`.mbox`) | Multiple extraction results (one per email) |

Notes:

- Word formats do not store reliable page boundaries, so units are document-level.
- `iterate_units(ignore_images=True)` skips image payloads in unit objects for better performance.

## CLI

After installation, `sharepoint2text` is available.

Plain text output:

```bash
sharepoint2text --file /path/to/file.docx > extraction.txt
```

JSON output:

```bash
sharepoint2text --file /path/to/file.docx --json > extraction.json
```

### Options

| Option | Description |
|---|---|
| `--file FILE` | Required input file |
| `--output FILE`, `-o FILE` | Write output to file (default: stdout) |
| `--json` | Emit `list[extraction_object]` |
| `--json-unit` | Emit `list[unit_object]` |
| `--include-images` | Include binary image payloads as base64 in JSON output |
| `--version` | Print CLI version |

Rules:

- `--json` and `--json-unit` are mutually exclusive.
- `--include-images` requires `--json` or `--json-unit`.
- CLI enforces a 100MB input file limit.

## Optional SharePoint Integration

`sharepoint_io` is optional. It helps list/download files from SharePoint, while extraction still runs through `sharepoint2text`.

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

Setup details: [`sharepoint2text/sharepoint_io/SETUP.md`](sharepoint2text/sharepoint_io/SETUP.md)

## Supported Formats

### Microsoft Office

- Modern: `.docx`, `.docm`, `.xlsx`, `.xlsm`, `.xlsb`, `.pptx`, `.pptm`
- Legacy: `.doc`, `.dot`, `.xls`, `.xlt`, `.ppt`, `.pot`, `.pps`, `.rtf`
- Template/show aliases are auto-mapped (for example `.dotx` -> `.docx`, `.ppsx` -> `.pptx`)

### OpenDocument

- `.odt`, `.ods`, `.odp`, `.odg`, `.odf`
- Template aliases supported (`.ott`, `.ots`, `.otp`)

### Email

- `.eml`, `.msg`, `.mbox`
- Email extraction includes sender/recipient metadata, subject, and body (`body_plain` / `body_html`).
- `.eml` and `.msg` parse attachments and store them on `EmailContent.attachments`.
- `.mbox` extraction currently focuses on message headers/body and does not parse/store attachments.
- Parsed supported attachments can be extracted via `EmailContent.iterate_supported_attachments()`.
- If supported-attachment extraction fails, the default behavior is to raise; use `skip_failed=True` to continue.

### Plain text and config/data

- `.txt`, `.md`, `.csv`, `.tsv`, `.json`
- `.yaml`, `.yml`, `.xml`, `.log`, `.ini`, `.cfg`, `.conf`, `.properties`

### Web and ebook

- `.html`, `.htm`, `.mhtml`, `.mht`, `.epub`

### PDF

- `.pdf`

### Archives

- `.zip`, `.tar`, `.7z`
- Compressed tar aliases: `.tar.gz`/`.tgz`, `.tar.bz2`/`.tbz2`, `.tar.xz`/`.txz`
- `.gz`, `.bz2`, `.xz` are routed as compressed tar variants

## Archive Processing and Security

Archives are processed one level deep. Supported non-archive files inside the archive can yield extraction results.
Nested archives are intentionally skipped as a safety guard.

Built-in safeguards include zip-bomb protections and file size limits. For 7z, extraction is limited to 100MB archives.
Archive entries may also be skipped when they exceed internal per-entry size limits or fail extraction.

## Limitations and Caveats

### PDF

- No OCR. Scanned-image PDFs may return empty text.
- Structured table extraction is not implemented for PDF (`iterate_tables()` is empty).
- Password-protected PDFs (non-empty password) raise `ExtractionFileEncryptedError`.
- Some JBIG2 images need `jbig2dec` installed for image decoding.

### General

- Inputs are expected to be already decrypted. If a file has encryption, DRM, password protection, or similar security controls, remove/unlock those before calling `sharepoint2text`.
- Very large or highly compressed files may hit protection limits.
- Raise limits only for trusted inputs.

## API Cheat Sheet

### Main entry points

```python
import sharepoint2text

sharepoint2text.read_file(
    path,
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
    force_plain_text=False,
)

sharepoint2text.read_bytes(
    data,
    extension="pdf",      # or ".pdf"
    mime_type=None,        # e.g. "application/pdf"
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
)

sharepoint2text.is_supported_file(path)
sharepoint2text.get_extractor(path)
```

### Format-specific extractors (selected)

- Office/OpenDocument: `read_docx`, `read_doc`, `read_xlsx`, `read_xls`, `read_pptx`, `read_ppt`, `read_rtf`, `read_odt`, `read_ods`, `read_odp`, `read_odg`, `read_odf`
- Other documents: `read_pdf`, `read_html`, `read_epub`, `read_mhtml`, `read_plain_text`
- Email: `read_eml_email`, `read_msg_email`, `read_mbox_email` (legacy aliases also exported)

All extractor functions accept a binary stream plus optional `path` and return generators.

Email helper API:

- `EmailContent.iterate_supported_attachments(skip_failed=False)` extracts supported parsed attachments on demand (primarily from `.eml`/`.msg`).

## Exceptions

Common exceptions:

- `ExtractionFileFormatNotSupportedError`
- `ExtractionFileEncryptedError`
- `ExtractionFileTooLargeError`
- `ExtractionLegacyMicrosoftParsingError`
- `ExtractionZipBombError`
- `ExtractionFailedError`

## License

Apache 2.0. See [`LICENSE`](LICENSE).

## Disclaimer

This project is not affiliated with, endorsed by, or sponsored by Microsoft.

## More Usage Examples

### Extract email body plus supported attachments

```python
import sharepoint2text

email = next(sharepoint2text.read_file("message-with-attachments.eml"))

print(email.subject)
print(email.get_full_text())  # plain body if available, otherwise HTML body
print(f"Attachment count: {len(email.attachments)}")

# Extract supported attachment types (pdf, docx, pptx, etc.)
for attachment_result in email.iterate_supported_attachments():
    print(type(attachment_result).__name__)
    print(attachment_result.get_full_text()[:200])
```

### Continue even if a supported attachment fails to extract

```python
import sharepoint2text

email = next(sharepoint2text.read_file("message-with-attachments.msg"))

for attachment_result in email.iterate_supported_attachments(skip_failed=True):
    print(attachment_result.get_metadata().filename)
```

### Process a mailbox (`.mbox`) and read message bodies

```python
import sharepoint2text

for email in sharepoint2text.read_file("team-archive.mbox"):
    print(f"Subject: {email.subject}")
    print(email.get_full_text()[:200])
```

### Batch-extract units for RAG-style chunking

```python
from pathlib import Path
import sharepoint2text

for path in Path("docs").rglob("*"):
    if not path.is_file() or not sharepoint2text.is_supported_file(path):
        continue
    for result in sharepoint2text.read_file(path):
        meta = result.get_metadata()
        for unit in result.iterate_units(ignore_images=True):
            chunk = unit.get_text().strip()
            if chunk:
                payload = {
                    "text": chunk,
                    "source": str(path),
                    "filename": meta.filename,
                    "unit_number": getattr(unit.get_metadata(), "unit_number", None),
                }
                # store payload in your index/vector DB
```

### Extract from API bytes when you only know MIME type

```python
import sharepoint2text

# Example: bytes from HTTP response
data = get_file_bytes_somehow()

result = next(
    sharepoint2text.read_bytes(
        data,
        mime_type="application/pdf",
        ignore_images=True,
    )
)
print(result.get_full_text()[:500])
```
