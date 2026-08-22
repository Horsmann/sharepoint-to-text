# Migrating from 1.2 to 2.0

Version 2.0 replaces the format-specific public extraction objects from 1.2
with one normalized `ExtractedDocument` model. This is a breaking release:
existing extraction calls generally keep the same inputs, but their yielded
objects, import paths, serialization format, and CLI JSON output have changed.

This guide covers migration from the final 1.2 release, 1.2.1.

## Upgrade

Update the dependency constraint to permit the new major version:

```bash
uv add "sharepoint-to-text>=2.0,<3"
```

Applications that cannot migrate immediately should remain pinned to 1.2.1.

## What remains stable

The main extraction entry points remain available:

- `read_file(...)`
- `read_bytes(...)`
- `read_many(...)`
- `is_supported_file(...)`

They still return generators, because archives and mailboxes can produce more
than one document. Options such as `ignore_images`, `force_plain_text`,
`include_attachments`, file-size limits, and recursive folder traversal remain
available.

The important difference is that every yielded value is now an
`ExtractedDocument`, regardless of the source format.

## Result objects

### Basic extraction

In 1.2, extraction returned a format-specific object:

```python
import sharepoint2text

result = next(sharepoint2text.read_file("report.pdf"))
print(result.get_full_text())
print(result.get_metadata().file_path)
```

In 2.0, access normalized fields directly:

```python
import sharepoint2text

document = next(sharepoint2text.read_file("report.pdf"))
print(document.full_text)
print(document.source.path)
```

### Common method replacements

| 1.2 | 2.0 |
|---|---|
| `result.get_full_text()` | `document.full_text` |
| `result.get_full_markdown()` | `render_markdown(document)` |
| `result.get_metadata()` | `document.source` and `document.metadata` |
| `result.iterate_units()` | `iter(document.units)` |
| `result.iterate_images()` | `document.iter_images()` |
| `result.iterate_tables()` | `document.iter_tables()` |
| `unit.get_text()` | `unit.text` |
| `unit.get_metadata()` | `unit.number`, `unit.kind`, `unit.title`, `unit.heading_path`, and `unit.properties` |
| `image.get_bytes()` | `image.data` |
| `image.get_content_type()` | `image.media_type` |
| `table.get_table()` | `table.rows` |
| `table.get_dim()` | `table.dimensions` |
| `result.to_json()` | `document_to_dict(document)` or `document_to_json(document)` |
| `ExtractionInterface.from_json(data)` | `document_from_dict(data)` or `document_from_json(value)` |

`document_to_dict` returns a dictionary. `document_to_json` returns a JSON
string, so it should not be passed through `json.dumps` again.

### The normalized model

All source formats use these eight public records:

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

Use `document.format` and `unit.kind` instead of checking a format-specific
Python class. Typical unit kinds include `document`, `section`, `page`,
`slide`, `sheet`, `chapter`, and `message`.

For example, replace:

```python
from sharepoint2text import PdfContent

result = next(sharepoint2text.read_file("report.pdf"))
if isinstance(result, PdfContent):
    for page in result.pages:
        print(page.text)
```

with:

```python
document = next(sharepoint2text.read_file("report.pdf"))
if document.format == "pdf":
    for unit in document.units:
        if unit.kind == "page":
            print(unit.text)
```

## Metadata mapping

Source identity and descriptive metadata are now separated:

| 1.2 metadata field | 2.0 field |
|---|---|
| `filename` | `document.source.filename` |
| `file_extension` | `document.source.extension` |
| `file_path` | `document.source.path` |
| `folder_path` | `document.source.folder` |
| `detected_encoding` | `document.source.encoding` |
| MIME or content type | `document.source.media_type` |
| `title` | `document.metadata.title` |
| `author` / `creator` | `document.metadata.author` |
| `subject` / `description` | `document.metadata.subject` |
| `keywords` | `document.metadata.keywords` |
| `language` | `document.metadata.language` |
| creation timestamp | `document.metadata.created` |
| modification timestamp | `document.metadata.modified` |

Format-specific scalar values that remain useful are stored under namespaced
keys in `properties`, for example `unit.properties["xlsx.hidden"]`. Callers
must not assume that every old field has a direct property equivalent.

## Format-specific structures

The following public result classes have been removed:

- `CsvContent`, `DocContent`, `DocxContent`, `EmailContent`, `EpubContent`,
  `HtmlContent`, `OdfContent`, `OdgContent`, `OdpContent`, `OdsContent`,
  `OdtContent`, `PdfContent`, `PlainTextContent`, `PptContent`, `PptxContent`,
  `RtfContent`, `XlsContent`, and `XlsxContent`

Their associated page, slide, sheet, paragraph, run, metadata, image, and table
classes are no longer public. Imports from
`sharepoint2text.parsing.extractors.data_types` therefore fail in 2.0.

Common content is normalized as follows:

- pages, slides, sheets, chapters, messages, and document sections become
  `ContentUnit` records;
- tables become `Table.rows`;
- images become `ImageAsset` records with immutable `bytes` payloads;
- comments, notes, formulas, hyperlinks, headers, footers, and similar records
  become `Annotation` records when supported;
- unassigned assets live in the corresponding `document_*` collection;
- scalar format-specific details may be retained in namespaced `properties`.

Parser-specific trees and mutable parser state are intentionally not part of
the 2.0 public model. Consumers that relied on detailed paragraph/run trees,
format-specific constructors, mutable `io.BytesIO` image payloads, or custom
arguments on format-specific `get_full_text(...)` implementations need to
redesign that part of their integration.

The new public dataclasses use `slots=True`. Applications can modify declared
fields, but cannot attach arbitrary new attributes to model instances.

## Removed direct readers and routing access

The top-level direct format readers have been removed, including `read_pdf`,
`read_docx`, `read_xlsx`, `read_pptx`, `read_doc`, `read_xls`, `read_ppt`,
`read_rtf`, the OpenDocument readers, email readers, and plain/web readers.
`get_extractor` is no longer a public top-level export either.

Route through `read_file` or `read_bytes` instead:

```python
# 1.2
result = next(sharepoint2text.read_pdf(file_like, path="report.pdf"))

# 2.0
result = next(
    sharepoint2text.read_bytes(file_like, extension=".pdf")
)
```

Internal extractor modules may still contain parser functions, but their
underscored records and module paths are implementation details and are not a
compatibility surface.

## Images and binary data

Image and attachment payloads are immutable `bytes | None` in 2.0 rather than
mutable `io.BytesIO` values:

```python
for image in document.iter_images():
    if image.data is not None:
        consume(image.data)
```

JSON serialization omits binary data by default. Request base64 explicitly
when payload preservation is required:

```python
from sharepoint2text import document_to_json

without_binary = document_to_json(document)
with_binary = document_to_json(document, binary="base64")
```

Decoding applies a cumulative 100 MiB binary limit by default. Set
`max_binary_bytes` to a lower application-specific limit when reading
untrusted payloads.

## Email attachments

Email attachments are represented by `document.attachments`. The 1.2
`EmailContent.iterate_supported_attachments(...)` convenience method and the
CLI's recursive attachment expansion are no longer available.

Version 2.0 retains attachment metadata and, unless `include_attachments=False`
is used, its binary payload. It does not eagerly extract attached documents.
Applications can route supported attachment bytes explicitly:

```python
from pathlib import Path

import sharepoint2text

message = next(sharepoint2text.read_file("message.eml"))
for attachment in message.attachments:
    if attachment.data is None:
        continue
    extension = Path(attachment.filename).suffix or None
    for attached_document in sharepoint2text.read_bytes(
        attachment.data,
        extension=extension,
        mime_type=attachment.media_type,
    ):
        print(attached_document.full_text)
```

Unsupported attachments raise `ExtractionFileFormatNotSupportedError`; handle
that exception if an email may contain arbitrary file types.

## JSON schema migration

The 1.2 serializer encoded Python class names using `_type` markers and could
reconstruct format-specific object graphs. Version 2.0 uses an explicit,
class-independent wire envelope:

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

The 2.0 decoder deliberately rejects 1.2 `_type` payloads. There is no
in-process compatibility decoder. If persisted 1.2 data must be retained:

1. keep a migration environment pinned to `sharepoint-to-text==1.2.1`;
2. deserialize each old payload there;
3. export the text, metadata, units, tables, images, annotations, and
   attachments needed by the application into an intermediate format;
4. construct `ExtractedDocument` records and encode them with the 2.0 codec;
5. verify representative documents before replacing the original data.

Keep the original payloads until the migrated data has been validated. A
custom adapter is required if the intermediate export must preserve
format-specific information that has no normalized 2.0 equivalent.

## CLI JSON output

Plain-text CLI usage remains broadly similar. JSON consumers must be updated:

- `--json` now emits version-2 document envelopes;
- `--json-unit` emits one complete version-2 envelope containing one unit,
  rather than a raw unit object with separate `unit_metadata` and
  `file_metadata` keys;
- `--include-images` controls base64 encoding of binary image and attachment
  payloads;
- attachments remain records on their parent document and are not recursively
  emitted as additional documents;
- `--no-attachments` omits attachment records and payloads.

Downstream validators should check both the `schema` and `version` fields
before consuming an envelope.

## Migration checklist

- Replace format-specific imports and `isinstance` checks.
- Replace method access with normalized fields and iterators.
- Split source metadata from descriptive document metadata.
- Decide which namespaced `properties` the application needs.
- Update image and attachment handling from `io.BytesIO` to `bytes`.
- Replace instance serialization calls with the centralized codec functions.
- Migrate or retain access to persisted 1.2 JSON before upgrading production.
- Update CLI JSON parsers to consume the version-2 envelope.
- Add regression tests for each source format and structured field the
  application depends on.
