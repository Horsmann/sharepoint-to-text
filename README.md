# sharepoint-to-text

`sharepoint-to-text` is a typed, pure-Python library for extracting text and
structured content from formats commonly found in SharePoint and document
management systems.

Every supported format returns the same `ExtractedDocument` model. The public
API does not expose format-specific result classes or direct extractor
functions.

## Installation

```bash
uv add sharepoint-to-text
```

With `pip`:

```bash
pip install sharepoint-to-text
```

For development:

```bash
git clone https://github.com/Horsmann/sharepoint-to-text.git
cd sharepoint-to-text
uv sync --all-groups
```

## Quick Start

### Read from disk

```python
import sharepoint2text

document = next(sharepoint2text.read_file("document.docx"))
print(document.full_text)
print(document.source.path)
```

### Read from bytes

```python
import sharepoint2text

document = next(
    sharepoint2text.read_bytes(b"hello from memory", extension="txt")
)
print(document.full_text)
```

`read_bytes` accepts either an `extension` or a `mime_type` routing hint.

### Read a folder

```python
import sharepoint2text

for document in sharepoint2text.read_many(
    "docs",
    suffixes=[".docx", ".pdf"],
):
    print(document.source.path, len(document.full_text))
```

Use `extract_all_supported=True` instead of `suffixes` to process every
supported file. Folder traversal is recursive by default.

### Process structural units

```python
import sharepoint2text

document = next(sharepoint2text.read_file("report.pdf"))

for unit in document.units:
    print(unit.kind, unit.number, unit.title)
    print(unit.text)
```

Typical units are pages, slides, sheets, chapters, messages, sections, or a
single document unit.

### Read or disable images

Image extraction is enabled by default. `iter_images()` yields images owned by
structural units followed by document-level images:

```python
import sharepoint2text

document = next(sharepoint2text.read_file("illustrated.docx"))

for image in document.iter_images():
    print(image.filename, image.media_type, len(image.data or b""))

document_without_images = next(
    sharepoint2text.read_file("illustrated.docx", ignore_images=True)
)
assert not list(document_without_images.iter_images())
```

`ignore_images=True` is also available on `read_bytes` and `read_many`.

## Public Data Model

All three extraction entry points yield `ExtractedDocument`:

```text
ExtractedDocument
├── source: SourceMetadata
├── metadata: DocumentMetadata
├── units: list[ContentUnit]
│   ├── images: list[ImageAsset]
│   ├── tables: list[Table]
│   └── annotations: list[Annotation]
├── document_images: list[ImageAsset]
├── document_tables: list[Table]
├── document_annotations: list[Annotation]
└── attachments: list[Attachment]
```

The public records are:

- `ExtractedDocument`
- `ContentUnit`
- `SourceMetadata`
- `DocumentMetadata`
- `ImageAsset`
- `Table`
- `Annotation`
- `Attachment`

Fields are accessed directly. For example, use `unit.text`, `unit.images`, and
`document.metadata.author`. Binary payloads are immutable `bytes`.

`document.full_text` joins non-empty unit text in source order.
`document.iter_images()` and `document.iter_tables()` traverse unit-owned
assets first and then unassigned document-level assets. An asset has one
canonical owner and is not duplicated between those collections.

Format-specific scalar details live in namespaced `properties` dictionaries,
such as `unit.properties["xlsx.hidden"]`. Parser trees and mutable parser state
are not part of the public model.

## API Reference

```python
sharepoint2text.read_file(
    path,
    max_file_size=100 * 1024 * 1024,
    *,
    ignore_images=False,
    force_plain_text=False,
    include_attachments=True,
    zip_bomb_limits=None,
)

sharepoint2text.read_bytes(
    data,
    *,
    extension=None,
    mime_type=None,
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
    force_plain_text=False,
    include_attachments=True,
    zip_bomb_limits=None,
)

sharepoint2text.read_many(
    folder_path,
    suffixes=None,
    *,
    extract_all_supported=False,
    max_file_size=100 * 1024 * 1024,
    ignore_images=False,
    force_plain_text=False,
    include_attachments=True,
    recursive=True,
    zip_bomb_limits=None,
)

sharepoint2text.is_supported_file(path)
```

The extraction functions return generators because archives and `.mbox` files
can produce multiple documents. `next(...)` is convenient for ordinary
single-document formats; iterate the generator when cardinality is not known.

`read_many` requires exactly one of `suffixes` or
`extract_all_supported=True`. It logs and skips individual extraction errors so
the rest of a folder can continue.

## JSON and Markdown

Serialization is centralized and independent of Python class names. Use
`document_to_dict` with the standard library's `json.dump` to write a readable
file:

```python
import json
from pathlib import Path

import sharepoint2text
from sharepoint2text import (
    document_from_json,
    document_to_dict,
    document_to_json,
    render_markdown,
)

document = next(sharepoint2text.read_file("report.pdf"))

mapping = document_to_dict(document)
with Path("report.json").open("w", encoding="utf-8") as output_file:
    json.dump(mapping, output_file, ensure_ascii=False, indent=2)

payload = document_to_json(document)
payload_with_binary = document_to_json(document, binary="base64")
restored = document_from_json(payload_with_binary)
markdown = render_markdown(document)
```

Every JSON envelope contains:

```json
{
  "schema": "sharepoint2text.extraction",
  "version": 2,
  "document": {
    "format": "pdf",
    "source": {},
    "metadata": {"keywords": [], "properties": {}},
    "units": [],
    "document_images": [],
    "document_tables": [],
    "document_annotations": [],
    "attachments": [],
    "properties": {}
  }
}
```

The outer `schema` and `version` fields identify the wire format. `document`
contains source and descriptive metadata, ordered content units, images,
tables, annotations, attachments, and namespaced format-specific properties.
An attachment's `extracted_document`, when present, is serialized recursively
using the same document shape.

Binary image and attachment data is omitted by default. Pass
`binary="base64"` to `document_to_dict` or `document_to_json` to include it as
base64 text. `document_to_json` returns compact JSON directly; use
`document_to_dict` with `json.dump`, as above, when indentation or a file-like
object is needed. For archives and `.mbox` inputs, serialize each yielded
`ExtractedDocument` as its own envelope or put those envelopes in a JSON list.

Decoding rejects unknown schema versions, malformed input, invalid base64, and
cumulative binary data above 100 MiB by default. Use `max_binary_bytes` to
impose a stricter boundary.

## Attachments

Email attachment records and their immutable byte payloads are available
through `document.attachments`:

```python
import sharepoint2text

document = next(sharepoint2text.read_file("message.eml"))

if document.attachments:
    for attachment in document.attachments:
        print(
            attachment.filename,
            attachment.media_type,
            len(attachment.data or b""),
        )

document_without_attachments = next(
    sharepoint2text.read_file("message.eml", include_attachments=False)
)
assert not document_without_attachments.attachments
```

`include_attachments=False` is also available on `read_bytes` and `read_many`.

### Recursively read mail attachments

`.mbox` yields one `ExtractedDocument` per message. Attachments are retained as
records but are not eagerly parsed, so feed `attachment.data` back into
`read_bytes` to extract attached documents or attached email messages. The
following depth-first loop handles attachments nested at any depth and skips
only unsupported attachment formats:

```python
import sharepoint2text

pending = list(sharepoint2text.read_file("mailbox.mbox"))

while pending:
    document = pending.pop()
    print(document.source.filename, document.full_text)

    if not document.attachments:
        continue

    for attachment in document.attachments:
        print("attachment:", attachment.filename, attachment.media_type)
        if attachment.data is None:
            continue

        try:
            child_documents = list(
                sharepoint2text.read_bytes(
                    attachment.data,
                    extension=attachment.filename,
                    mime_type=attachment.media_type,
                )
            )
        except sharepoint2text.ExtractionFileFormatNotSupportedError:
            continue

        pending.extend(child_documents)
```

Passing the complete attachment filename as `extension` preserves compound
suffix routing such as `.tar.gz`; `mime_type` acts as a fallback. Other
extraction failures are intentionally not suppressed by this example.

## Supported Formats

| Family | Formats |
|---|---|
| Modern Office | `.docx`, `.docm`, `.xlsx`, `.xlsm`, `.xlsb`, `.pptx`, `.pptm` and template/show aliases |
| Legacy Office | `.doc`, `.dot`, `.xls`, `.xlt`, `.ppt`, `.pot`, `.pps`, `.rtf` |
| OpenDocument | `.odt`, `.ods`, `.odp`, `.odg`, `.odf`, `.ott`, `.ots`, `.otp` |
| Email | `.eml`, `.msg`, `.mbox` |
| Plain/data | `.txt`, `.md`, `.csv`, `.tsv`, `.json`, `.yaml`, `.yml`, `.xml`, `.log`, `.ini`, `.cfg`, `.conf`, `.properties` |
| Web/ebook | `.html`, `.htm`, `.mhtml`, `.mht`, `.epub` |
| PDF | `.pdf` |
| Archives | `.zip`, `.tar`, `.7z`, `.tar.gz`, `.tgz`, `.tar.bz2`, `.tbz2`, `.tar.xz`, `.txz` |

See [doc/format-matrix.md](doc/format-matrix.md) for result cardinality and
unit behavior.

## CLI

```bash
# Plain text
sharepoint2text --file document.docx

# Version-2 JSON; binary payloads omitted
sharepoint2text --file document.docx --json

# One version-2 envelope per unit
sharepoint2text --file report.pdf --json-unit

# Include image payloads as base64
sharepoint2text --file report.pdf --json --include-images

# Folder extraction
sharepoint2text --folder ./docs --suffixes .docx,.pdf --output ./extracted/
```

See [doc/cli.md](doc/cli.md) for the complete option reference.

## ZIP-Bomb Limits

ZIP-based Office/OpenDocument files and archives are guarded by configurable
limits. Raise them only for a trusted input and only on that extraction call:

```python
from sharepoint2text import ZipBombLimits, read_file

documents = read_file(
    "large_trusted_export.zip",
    zip_bomb_limits=ZipBombLimits(
        max_total_uncompressed_bytes=16 * 1024 * 1024 * 1024,
        max_single_uncompressed_bytes=4 * 1024 * 1024 * 1024,
        max_entry_compression_ratio=1500.0,
    ),
)

for document in documents:
    print(document.full_text)
```

The override applies only while that generator is actively extracting. It is
automatically reset before a result is yielded and after completion, failure,
or early generator closure, so later and concurrent calls keep the defaults.
The same `zip_bomb_limits` keyword is available on `read_bytes` and
`read_many`; for `read_many`, it is applied independently to each selected
file.

## SharePoint Integration

The extraction API works independently of SharePoint. The optional
`sharepoint_io` helper can download Microsoft Graph content and pass its bytes
to the same normalized API:

```python
import sharepoint2text

data = client.download_file(file_meta.id)
for document in sharepoint2text.read_bytes(data, extension=file_meta.name):
    print(document.full_text)
```

See [sharepoint2text/sharepoint_io/SETUP.md](sharepoint2text/sharepoint_io/SETUP.md).

## Operational Constraints

- OCR is not included, so image-only PDFs can return sparse or empty text.
- Word-like formats generally cannot provide reliable page boundaries.
- Output is extraction-oriented, not layout-preserving rendering.
- Nested archives are skipped.
- Encrypted inputs raise extraction errors.
- Size and decompression safety limits remain enabled by default.

Set `ignore_images=True` when binary assets are unnecessary. Process
`document.units` incrementally when structure matters.

## Development Validation

```bash
uv run pytest
uv run mypy .
```

## Additional Documentation

- [doc/cli.md](doc/cli.md): CLI reference and JSON shape
- [doc/format-matrix.md](doc/format-matrix.md): behavior by format family
- [CONTRIBUTING.md](CONTRIBUTING.md): contributor workflow
- [CHANGELOG.md](CHANGELOG.md): release history

## License

Apache 2.0. See [LICENSE](LICENSE).

This project is not affiliated with, endorsed by, or sponsored by Microsoft.
