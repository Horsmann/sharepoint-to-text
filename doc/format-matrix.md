# Format Matrix

Every row below produces `ExtractedDocument`; differences are represented by
`format`, unit `kind`, metadata, annotations, and namespaced properties.

| Format family | Example extensions | Typical documents yielded | Unit shape | Notes |
|---|---|---:|---|---|
| Word-like documents | `.docx`, `.doc`, `.odt`, `.rtf`, `.txt`, `.md`, `.json` | 1 | Document or section units | Page boundaries are generally unavailable |
| Spreadsheets | `.xlsx`, `.xls`, `.xlsb`, `.xlsm`, `.ods` | 1 | One `sheet` unit per sheet | Sheet names use `unit.title` |
| Presentations | `.pptx`, `.ppt`, `.pptm`, `.odp` | 1 | One `slide` unit per slide | Notes may be included when supported |
| PDF | `.pdf` | 1 | One `page` unit per page | OCR is not included |
| Email | `.eml`, `.msg` | 1 | One `message` unit | Attachment records use `document.attachments` |
| Mailbox | `.mbox` | Many | One `message` unit per yielded document | Iterate the result generator |
| HTML-like | `.html`, `.htm`, `.mhtml`, `.mht` | 1 | Usually one document unit | Extraction-oriented, not browser rendering |
| Ebook | `.epub` | 1 | `chapter` units | Reading order follows the source spine |
| Archives | `.zip`, `.tar`, `.7z`, `.tgz`, `.tbz2`, `.txz` | Many | Depends on member format | Nested archives are skipped |

## Common Operations

```python
document.full_text
document.source.filename
document.metadata.title

for unit in document.units:
    print(unit.kind, unit.number, unit.text)

for image in document.iter_images():
    consume(image.data or b"")

for table in document.iter_tables():
    print(table.rows)
```

Assets normally belong to a unit. `document_images`, `document_tables`, and
`document_annotations` contain only items that cannot be assigned reliably.

## Attachments and Archives

- Email attachment records are in `document.attachments`.
- `include_attachments=False` omits attachment records and payloads.
- `.mbox` yields one normalized document per message.
- Archives yield one normalized document per supported member.
- Archive member paths are retained in `document.source.path`.
- Zip-bomb and path-traversal protections apply.

## Caveats

- Image-only PDFs often have little or no text because OCR is not included.
- Password-protected PDFs raise `ExtractionFileEncryptedError`.
- Binary legacy Office files are less predictable than OOXML and may raise
  `ExtractionLegacyMicrosoftParsingError`.
- `force_plain_text=True` can route unknown text-based extensions through the
  plain-text extractor.
