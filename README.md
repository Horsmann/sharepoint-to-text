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

### Read images or skip their extraction

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
assert list(document_without_images.iter_images()) == []
```

`ignore_images=True` skips image extraction, so image records, dimensions, and
binary data are unavailable. It is also available on `read_bytes` and `read_many`.

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
├── document_images: list[ImageAsset]       (all unit images)
├── document_tables: list[Table]            (all unit tables)
├── document_annotations: list[Annotation]  (all unit annotations)
└── attachments: list[Attachment]
```

The normalized document records are:

- `ExtractedDocument`
- `ContentUnit`
- `SourceMetadata`
- `DocumentMetadata`
- `ImageAsset`
- `Table`
- `Annotation`
- `Attachment`

`BatchFileResult` is a separate operational record used by the optional
`read_many(..., on_file_result=...)` callback described below.

Fields are accessed directly. For example, use `unit.text`, `unit.images`, and
`document.metadata.author`. Binary payloads are immutable `bytes`.

`document.full_text` joins non-empty unit text in source order.
`document.document_images`, `document.document_tables`, and
`document.document_annotations` provide document-wide aggregates of the same
objects canonically owned by units. `document.iter_images()` and
`document.iter_tables()` traverse those aggregates. A document without a
structural unit receives one default unit with `kind="document"`.

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
    extract_annotations=False,
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
    extract_annotations=False,
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
    extract_annotations=False,
    recursive=True,
    on_file_result=None,
    zip_bomb_limits=None,
)

sharepoint2text.is_supported_file(path)
```

### File-size limits

`read_file`, `read_bytes`, and `read_many` default to a 100 MiB input-size
limit (`100 * 1024 * 1024`, or 104,857,600 bytes). For `read_many`, the limit
is applied independently to each selected file; it is not a cumulative limit
for the folder.

Pass a larger `max_file_size` value when processing trusted files that exceed
the default. The value is always expressed in bytes:

```python
from pathlib import Path

import sharepoint2text

MAX_INPUT_SIZE = 500 * 1024 * 1024  # 500 MiB
file_data = Path("large-report.pdf").read_bytes()

file_documents = sharepoint2text.read_file(
    "large-report.pdf",
    max_file_size=MAX_INPUT_SIZE,
)
memory_documents = sharepoint2text.read_bytes(
    file_data,
    extension="pdf",
    max_file_size=MAX_INPUT_SIZE,
)
folder_documents = sharepoint2text.read_many(
    "large-documents",
    extract_all_supported=True,
    max_file_size=MAX_INPUT_SIZE,
)
```

Set `max_file_size=0` to disable the input-size check. For archives, this
setting limits the top-level archive file or byte buffer; decompressed ZIP
members remain subject to the separate [ZIP-bomb limits](#zip-bomb-limits).

### Iterator semantics

The extraction functions return lazy, synchronous, single-pass
`Iterator[ExtractedDocument]` values. Archives and `.mbox` files can produce
multiple documents, so even `read_file` and `read_bytes` return iterators.

- Consume each returned iterator only once. An exhausted iterator stays
  exhausted; call the API again to repeat extraction.
- Documents are produced on demand. Advancing the iterator can perform file
  I/O, decompression, parsing, and normalization.
- `read_many` traverses lazily and processes one top-level file at a time. All
  documents from one source are adjacent, but filesystem traversal order is
  not a portable sorting guarantee.
- A yielded `ExtractedDocument` is fully materialized. Calling
  `list(read_many(...))` retains every result and removes the iterator's
  result-memory benefit.

Static validation happens when the public function is called, while source
processing remains deferred:

| API | Validated when called | Deferred until iteration |
|---|---|---|
| `read_file` | Path, source size, routing, and ZIP-bomb configuration | Opening, reading, parsing, and normalization |
| `read_bytes` | Input type and size, routing hints, and ZIP-bomb configuration | Stream positioning, parsing, and normalization |
| `read_many` | Folder, selection, result-callback, and ZIP-bomb configuration | Traversal and each selected file's validation and extraction |

For a source that may yield any number of documents, iterate to exhaustion:

```python
for document in sharepoint2text.read_file("mailbox.mbox"):
    consume(document)
```

When exactly one result is an application invariant, iterable unpacking checks
that invariant and exhausts the iterator:

```python
document, = sharepoint2text.read_file("report.pdf")
```

This raises `ValueError` for zero or multiple documents. By comparison,
`next(...)` returns only the first document and does not verify that the source
is exhausted.

`read_file` and `read_bytes` propagate lazy extraction failures. `read_many`
logs expected per-file extraction and I/O failures, skips those files, and
continues. Other exceptions stop iteration.

Fully exhausting an iterator releases its open file, archive, and parser
resources. Breaking early can retain those resources until the concrete
iterator is closed or finalized. The current implementation uses generators
with `close()` at runtime, but the stable public contract currently promises
only `Iterator`; deterministic context-managed cleanup after partial
consumption is not yet part of the API.

See [Iterator Semantics](ITERATOR_SEMANTIC.md) for the detailed contract,
including ordering, partial-result, memory, and non-goal definitions.

`read_many` requires exactly one of `suffixes` or
`extract_all_supported=True`.

### Per-file batch reporting

Pass `on_file_result` to receive one structured `BatchFileResult` after each
selected top-level file completes or encounters a recoverable extraction or
I/O error. Reporting is callback-based so `read_many` does not need to retain
an arbitrarily large result collection:

```python
import sharepoint2text
from sharepoint2text import BatchFileResult


def report_file(result: BatchFileResult) -> None:
    """Report one completed batch file."""
    if result.succeeded:
        print(result.path, result.documents_extracted)
    else:
        print(result.path, type(result.error).__name__, result.error)


for document in sharepoint2text.read_many(
    "documents",
    extract_all_supported=True,
    on_file_result=report_file,
):
    index_document(document)
```

Each result contains the selected top-level `path`, the number of
`documents_extracted` before completion or failure, the recoverable `error`
(`None` on success), and the convenience property `succeeded`.

The callback runs only after a selected file finishes. It does not run for
filtered files or for the current file when iteration is abandoned early. If
the callback raises an exception, that exception stops batch iteration. An
application may collect callback values when its own memory bounds permit.

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
    "units": [{
      "number": 1,
      "kind": "document",
      "text": "",
      "heading_path": [],
      "images": [],
      "tables": [],
      "annotations": [],
      "properties": {}
    }],
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

Decoding rejects unknown schema versions, malformed input, and invalid base64.
Binary payload size is unlimited by default so complete serialized extraction
results can be restored. When decoding untrusted input, pass `max_binary_bytes`
to impose a cumulative allocation boundary.

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

# One self-contained version-2 envelope per unit
sharepoint2text --file report.pdf --json-unit

# Include image and attachment payloads as base64
sharepoint2text --file report.pdf --json --include-binary

# Folder extraction
sharepoint2text --folder ./docs --suffixes .docx,.pdf --output ./extracted/
```

Mirrored folder output writes one output file per input source. When an input
such as `.mbox` yields multiple documents, they are kept together in that
file.

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
import os

import sharepoint2text
from sharepoint2text.sharepoint_io import (
    EntraIDAppCredentials,
    SharePointRestClient,
)

credentials = EntraIDAppCredentials(
    tenant_id=os.environ["sp_tenant_id"],
    client_id=os.environ["sp_client_id"],
    client_secret=os.environ["sp_client_secret"],
)
client = SharePointRestClient(
    site_url=os.environ["sp_site_url"],
    credentials=credentials,
)

file_path = "Documents/report.pdf"
data = client.download_file_by_path(file_path)
for document in sharepoint2text.read_bytes(data, extension=file_path):
    print(document.full_text)
```

The client itself uses the Python standard library. The optional `.env` setup
helper additionally requires `python-dotenv`; it is included by the repository's
development dependency group but not by a normal package installation. See
[sharepoint2text/sharepoint_io/SETUP.md](sharepoint2text/sharepoint_io/SETUP.md)
for permission setup, environment configuration, and installation options.

## Operational Constraints

- OCR is not included, so image-only PDFs can return sparse or empty text.
- Word-like formats generally cannot provide reliable page boundaries.
- Output is extraction-oriented, not layout-preserving rendering.
- Nested archives are skipped.
- Encrypted inputs raise extraction errors.
- Size and decompression safety limits remain enabled by default.

Set `ignore_images=True` when image content is unnecessary. Process
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
