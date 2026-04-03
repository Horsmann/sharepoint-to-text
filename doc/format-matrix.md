# Format Matrix

This document summarizes how `sharepoint-to-text` behaves by format family. It is meant to answer the practical engineering question: "What do I actually get back?"

## Result Shape By Format

| Format family | Example extensions | Typical result count | `iterate_units()` shape | Notable details |
|---|---|---:|---|---|
| Word-like documents | `.docx`, `.doc`, `.odt`, `.rtf`, `.txt`, `.md`, `.json` | 1 | Usually one document-level unit | Page boundaries are generally not reliable |
| Spreadsheets | `.xlsx`, `.xls`, `.xlsb`, `.xlsm`, `.ods` | 1 | One unit per sheet | Useful for chunking by worksheet |
| Presentations | `.pptx`, `.ppt`, `.pptm`, `.odp` | 1 | One unit per slide | Slide notes may also be represented in extracted text depending on format support |
| PDF | `.pdf` | 1 | One unit per page | No OCR for scanned pages |
| Email messages | `.eml`, `.msg` | 1 | One unit per email | Attachments can be parsed and recursively exposed |
| Mailbox archives | `.mbox` | Many | One extraction result per email | Treat as a stream, not a single document |
| HTML-like content | `.html`, `.htm`, `.mhtml`, `.mht` | 1 | Usually one document-level unit | Output is extraction-oriented, not browser-rendered |
| Ebook | `.epub` | 1 | Format-defined document units | Good for text extraction, not ebook rendering |
| Archives | `.zip`, `.tar`, `.7z`, `.tgz`, `.tbz2`, `.txz` | Many | Depends on contained files | Only one archive level is processed |

## Attachment Behavior

### `.eml` and `.msg`

- message metadata is extracted
- body text is available through the email result
- attachments are stored on `EmailContent.attachments`
- supported attachments can be recursively extracted with `iterate_supported_attachments()`

### `.mbox`

- one extraction result is yielded per message
- body and header extraction are supported
- attachment handling is intentionally more limited than `.eml` and `.msg`

## Archive Behavior

- supported member files can be extracted
- nested archives are skipped
- zip-bomb protections apply
- archive members may be skipped if they exceed safety limits or fail extraction

## Practical Caveats

### PDF

- scanned-image PDFs often return empty or sparse text because OCR is not included
- `iterate_tables()` is currently empty for PDF
- password-protected PDFs raise `ExtractionFileEncryptedError`
- some JBIG2 images may require `jbig2dec` to decode image data

### Legacy Microsoft formats

- coverage is broad, but binary legacy formats are inherently less predictable than OOXML formats
- parser failures may surface as `ExtractionLegacyMicrosoftParsingError`

### Word-like formats

- expect text extraction, metadata, images, and tables where supported
- do not expect stable page-level units

### Plain-text formats

- these are routed through the plain extractor
- `force_plain_text=True` can be useful for unknown internal file extensions that are actually text
