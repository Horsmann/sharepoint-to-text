# Direct Extractor Usage

This guide is for the cases where you already know the file format and do not
want to go through `read_file(...)`, `read_bytes(...)`, or `get_extractor(...)`.

Direct extractors are useful when you want:

- static typing against a concrete result dataclass such as `DocxContent`
- access to format-specific attributes such as `slides`, `chapters`, `tables`,
  `comments`, `attachments`, or `paragraphs`
- explicit control over which extractor runs for alias formats like `.docm`,
  `.dotx`, `.ppsx`, or `.ots`

The top-level `sharepoint2text.read_*` functions are thin wrappers around the
underlying extractor modules, so they are the preferred direct-call API.

## Common Pattern

For single-result formats, call the direct extractor and use `next(...)`:

```python
from pathlib import Path

from sharepoint2text import DocxContent, read_docx

path = Path("report.docx")

with path.open("rb") as handle:
    result: DocxContent = next(read_docx(handle, str(path)))

print(result.metadata.title)
print(result.get_full_text())
```

For in-memory data, wrap the payload in `io.BytesIO`:

```python
import io

from sharepoint2text import PdfContent, read_pdf

payload = io.BytesIO(pdf_bytes)
result: PdfContent = next(read_pdf(payload, path="report.pdf"))
```

Notes:

- Pass the real filename as `path=...` when possible. It is used to populate
  metadata such as `filename`, `file_extension`, and `file_path`.
- Most direct extractors accept `ignore_images=True`.
- The email extractors also accept `include_attachments=False`.
- `.mbox` is still a direct extractor, but it yields one `EmailContent` per
  message, so use a `for` loop instead of `next(...)`.

## Shared Methods On All Concrete Results

Every concrete `*Content` type implements the `ExtractionInterface` methods:

- `get_full_text()`
- `get_full_markdown()`
- `iterate_units()`
- `iterate_images()`
- `iterate_tables()`
- `get_metadata()`
- `to_json()` / `from_json(...)`

The value of the direct extractor API is that you can combine those shared
methods with concrete attributes from the specific result type.

## Extractor Map

| Direct function | Use for these file types | Concrete result type | Typical unit shape | Useful concrete attributes |
|---|---|---|---|---|
| `read_docx` | `.docx`, `.docm`, `.dotx`, `.dotm` | `DocxContent` | heading-based document units | `paragraphs`, `tables`, `headers`, `footers`, `images`, `hyperlinks`, `footnotes`, `endnotes`, `comments`, `sections`, `styles`, `formulas` |
| `read_doc` | `.doc`, `.dot` | `DocContent` | document or heading-based units | `main_text`, `footnotes`, `headers_footers`, `annotations`, `images`, `tables`, `metadata` |
| `read_rtf` | `.rtf` | `RtfContent` | page units if page breaks exist | `paragraphs`, `headers`, `footers`, `hyperlinks`, `bookmarks`, `fields`, `images`, `tables`, `footnotes`, `annotations`, `fonts`, `colors`, `styles`, `pages` |
| `read_odt` | `.odt`, `.ott` | `OdtContent` | heading-based document units | `paragraphs`, `tables`, `headers`, `footers`, `images`, `hyperlinks`, `footnotes`, `endnotes`, `annotations`, `bookmarks`, `styles`, `full_text` |
| `read_xlsx` | `.xlsx`, `.xlsm`, `.xlsb`, `.xltx`, `.xltm` | `XlsxContent` | one unit per sheet | `sheets`, `metadata` |
| `read_xls` | `.xls`, `.xlt` | `XlsContent` | one unit per sheet | `sheets`, `images`, `full_text`, `metadata` |
| `read_ods` | `.ods`, `.ots` | `OdsContent` | one unit per sheet | `sheets`, `metadata` |
| `read_csv` | `.csv`, `.tsv` | `CsvContent` | one document-level unit | `content`, `table`, `metadata` |
| `read_pptx` | `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm` | `PptxContent` | one unit per slide | `slides`, `metadata` |
| `read_ppt` | `.ppt`, `.pot`, `.pps` | `PptContent` | one unit per slide | `slides`, `master_text`, `all_text`, `streams`, `metadata` |
| `read_odp` | `.odp`, `.otp` | `OdpContent` | one unit per slide | `slides`, `metadata` |
| `read_pdf` | `.pdf` | `PdfContent` | one unit per page | `pages`, `metadata` |
| `read_html` | `.html`, `.htm` | `HtmlContent` | one document-level unit | `content`, `tables`, `headings`, `links`, `metadata` |
| `read_mhtml` | `.mhtml`, `.mht` | `HtmlContent` | one document-level unit | `content`, `tables`, `headings`, `links`, `metadata` |
| `read_epub` | `.epub` | `EpubContent` | one unit per chapter/content document | `chapters`, `images`, `toc`, `metadata` |
| `read_plain_text` | `.txt`, `.md`, `.json`, `.yaml`, `.yml`, `.xml`, `.log`, `.ini`, `.cfg`, `.conf`, `.properties` | `PlainTextContent` | one document-level unit | `content`, `metadata` |
| `read_msg_email` | `.msg` | `EmailContent` | one unit per email body | `from_email`, `subject`, `to_emails`, `to_cc`, `to_bcc`, `reply_to`, `body_plain`, `body_html`, `attachments`, `metadata` |
| `read_eml_email` | `.eml` | `EmailContent` | one unit per email body | `from_email`, `subject`, `to_emails`, `to_cc`, `to_bcc`, `reply_to`, `body_plain`, `body_html`, `attachments`, `metadata` |
| `read_mbox_email` | `.mbox` | `EmailContent` | one result per message | `from_email`, `subject`, `body_plain`, `body_html`, `attachments`, `metadata` |
| `read_odg` | `.odg` | `OdgContent` | one document-level unit | `full_text`, `images`, `metadata` |
| `read_odf` | `.odf` | `OdfContent` | one document-level unit | `full_text`, `metadata` |

`read_apple_pages` is experimental. It is useful for text-first `.pages`
documents, but its heading reconstruction, layout ordering, and table/image
placement are currently less stable than the mature OOXML and OpenDocument
extractors.

## Working With Concrete Types

### Word-like documents

Use the direct extractor when you need document-structure attributes instead of
just flattened text.

```python
from pathlib import Path

from sharepoint2text import DocxContent, read_docx

path = Path("proposal.docm")

with path.open("rb") as handle:
    doc: DocxContent = next(read_docx(handle, str(path)))

print(doc.metadata.title)
print([comment.author for comment in doc.comments])
print([link.url for link in doc.hyperlinks])

for paragraph in doc.paragraphs:
    if paragraph.style and paragraph.style.startswith("Heading"):
        print(paragraph.style, paragraph.text)

for unit in doc.iterate_units():
    meta = unit.get_metadata()
    print(meta.heading_path, unit.get_text())
```

Equivalent patterns for other word-like results:

- `DocContent`: use `main_text`, `footnotes`, `headers_footers`, `annotations`
- `RtfContent`: use `paragraphs`, `hyperlinks`, `tables`, `footnotes`,
  `annotations`, `pages`
- `OdtContent`: use `paragraphs`, `hyperlinks`, `annotations`, `bookmarks`,
  `footnotes`, `endnotes`
- `ApplePagesContent` (experimental): use `paragraphs`, `tables`, `images`,
  `full_text`

### Spreadsheets

Spreadsheet extractors expose real sheet objects, which is usually more useful
than treating the workbook as a single text blob.

```python
from pathlib import Path

from sharepoint2text import XlsxContent, read_xlsx

path = Path("finance.xlsm")

with path.open("rb") as handle:
    workbook: XlsxContent = next(read_xlsx(handle, str(path)))

for sheet in workbook.sheets:
    print(sheet.name)
    print(sheet.text)
    print(sheet.get_dim())
    for image in sheet.images:
        print(image.filename, image.get_metadata().unit_number)
```

Use the same idea for:

- `XlsContent.sheets`
- `OdsContent.sheets`
- `CsvContent.table`

### Presentations

Presentation extractors make slide-level fields available directly.

```python
from pathlib import Path

from sharepoint2text import PptxContent, read_pptx

path = Path("deck.ppsx")

with path.open("rb") as handle:
    deck: PptxContent = next(read_pptx(handle, str(path)))

for slide in deck.slides:
    print(slide.slide_number, slide.title)
    print(slide.content_placeholders)
    print(slide.other_textboxes)
    print([comment.text for comment in slide.comments])
    print([formula.latex for formula in slide.formulas])
```

Other presentation-specific attributes:

- `PptContent.slides[*].notes`
- `PptContent.slides[*].all_text`
- `OdpContent.slides[*].notes`
- `OdpContent.slides[*].annotations`

### Email

The email extractors return `EmailContent`, which is more useful than the
generic interface when you need addresses, MIME attachments, or recursive
attachment extraction.

```python
from pathlib import Path

from sharepoint2text import EmailContent, read_msg_email

path = Path("message.msg")

with path.open("rb") as handle:
    email: EmailContent = next(read_msg_email(handle, str(path)))

print(email.subject)
print(email.from_email.address)
print([recipient.address for recipient in email.to_emails])
print([attachment.filename for attachment in email.attachments])

for attachment_result in email.iterate_supported_attachments(skip_failed=True):
    print(type(attachment_result).__name__, attachment_result.get_full_text()[:120])
```

For mailbox files:

```python
from pathlib import Path

from sharepoint2text import read_mbox_email

path = Path("mailbox.mbox")

with path.open("rb") as handle:
    for email in read_mbox_email(handle, str(path), include_attachments=False):
        print(email.subject)
```

### PDF, HTML, EPUB, and plain text

These formats also benefit from the concrete result objects:

```python
from pathlib import Path

from sharepoint2text import (
    EpubContent,
    HtmlContent,
    PdfContent,
    read_epub,
    read_html,
    read_pdf,
)

with Path("paper.pdf").open("rb") as handle:
    pdf: PdfContent = next(read_pdf(handle, "paper.pdf"))
    print(pdf.metadata.total_pages)
    print(len(pdf.pages))

with Path("page.html").open("rb") as handle:
    html: HtmlContent = next(read_html(handle, "page.html"))
    print(html.metadata.title)
    print(html.headings)
    print(html.links[:5])

with Path("book.epub").open("rb") as handle:
    book: EpubContent = next(read_epub(handle, "book.epub"))
    print(book.metadata.creator)
    print(book.toc[:5])
    print([chapter.title for chapter in book.chapters[:5]])
```

Useful attributes by type:

- `PdfContent.pages[*].text`, `PdfContent.pages[*].images`,
  `PdfContent.pages[*].tables`
- `HtmlContent.headings`, `HtmlContent.links`, `HtmlContent.tables`
- `EpubContent.chapters`, `EpubContent.images`, `EpubContent.toc`
- `PlainTextContent.content`
- `OdgContent.images`
- `OdfContent.full_text`

## Choosing Direct Extractors Vs Routing

Prefer direct extractors when:

- the file type is already known
- you want static typing for downstream code
- you want format-specific fields without type narrowing or casts

Prefer `read_file(...)`, `read_bytes(...)`, or `get_extractor(...)` when:

- the input is heterogeneous
- you are building a general ingestion pipeline
- file type is only known at runtime

## Archive Note

This guide intentionally focuses on extractors that return a concrete,
format-specific result type.

Archive extraction is different: archive members can resolve to many different
result types, so archive processing is usually better handled through
`read_file(...)`, `read_bytes(...)`, or the router layer. The archive extractor
does not give you a single concrete `ArchiveContent` dataclass to program
against.
