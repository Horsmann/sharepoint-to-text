# sharepoint-to-text

A **pure Python** library for extracting text, metadata, and structured elements from Microsoft Office files—both modern (`.docx`, `.xlsx`, `.pptx`) and legacy (`.doc`, `.xls`, `.ppt`) formats—plus PDF, email formats, and plain text.

## Why This Library?

### Pure Python, No External Dependencies

Unlike popular alternatives that shell out to **LibreOffice** or **Apache Tika** (requiring Java), `sharepoint-to-text` is a **native Python implementation** with no system-level dependencies:

| Approach | Requirements | Cross-platform | Container-friendly |
|----------|-------------|----------------|-------------------|
| **sharepoint-to-text** | `pip install` only | Yes | Yes (minimal image) |
| LibreOffice-based | LibreOffice install, X11/headless setup | Complex | Large images (~1GB+) |
| Apache Tika | Java runtime, Tika server | Complex | Heavy (~500MB+) |
| subprocess-based | Shell access, security concerns | No | Risky |

This library parses Office binary formats (OLE2) and XML-based formats (OOXML) directly in Python, making it ideal for:

- **RAG pipelines** and LLM document ingestion
- **Serverless functions** (AWS Lambda, Google Cloud Functions)
- **Containerized deployments** with minimal footprint
- **Secure environments** where shell execution is restricted
- **Cross-platform** applications (Windows, macOS, Linux)

### Enterprise SharePoint Reality

Enterprise SharePoints contain decades of accumulated documents. While modern `.docx`, `.xlsx`, and `.pptx` files are well-supported, legacy `.doc`, `.xls`, and `.ppt` files remain common. This library provides a **unified interface** for all formats—no conditional logic needed.

## Supported Formats

| Format            | Extension | Description                      |
|-------------------|-----------|----------------------------------|
| Modern Word       | `.docx`   | Word 2007+ documents             |
| Legacy Word       | `.doc`    | Word 97-2003 documents           |
| Modern Excel      | `.xlsx`   | Excel 2007+ spreadsheets         |
| Legacy Excel      | `.xls`    | Excel 97-2003 spreadsheets       |
| Modern PowerPoint | `.pptx`   | PowerPoint 2007+ presentations   |
| Legacy PowerPoint | `.ppt`    | PowerPoint 97-2003 presentations |
| PDF               | `.pdf`    | PDF documents                    |
| EML Email         | `.eml`    | RFC 822 email format             |
| MSG Email         | `.msg`    | Microsoft Outlook email format   |
| MBOX Email        | `.mbox`   | Unix mailbox format (multiple emails) |
| Plain Text        | `.txt`    | Plain text files                 |
| CSV               | `.csv`    | Comma-separated values           |
| TSV               | `.tsv`    | Tab-separated values             |
| JSON              | `.json`   | JSON files                       |

## Installation

```bash
pip install sharepoint-to-text
```

Or install from source:

```bash
git clone https://github.com/Horsmann/sharepoint-to-text.git
cd sharepoint-to-text
pip install -e .
```

## Quick Start

### The Unified Interface

All extractors return **generators** that yield content objects implementing a common interface. This design enables memory-efficient processing and supports formats that may contain multiple items (like `.mbox` mailboxes with multiple emails).

```python
import sharepoint2text

# Works identically for ANY supported format
# Most formats yield a single item, so use next() for convenience
for result in sharepoint2text.read_file("document.docx"):  # or .doc, .pdf, .pptx, etc.
    # Three methods available on ALL content types:
    text = result.get_full_text()       # Complete text as a single string
    metadata = result.get_metadata()    # File metadata (author, dates, etc.)

    # Iterate over logical units (varies by format - see below)
    for unit in result.iterator():
        print(unit)

# For single-item formats, you can use next() directly:
result = next(sharepoint2text.read_file("document.docx"))
print(result.get_full_text())
```

### Understanding `iterator()` Output by Format

Different file formats have different natural structural units:

| Format | `iterator()` yields | Notes |
|--------|-------------------|-------|
| `.docx`, `.doc` | 1 item (full text) | Word documents have no page structure in the file format |
| `.xlsx`, `.xls` | 1 item per **sheet** | Each yield contains sheet content |
| `.pptx`, `.ppt` | 1 item per **slide** | Each yield contains slide text |
| `.pdf` | 1 item per **page** | Each yield contains page text |
| `.eml`, `.msg` | 1 item (email body) | Plain text or HTML body |
| `.mbox` | 1 item per **email** | Mailboxes can contain multiple emails |
| `.txt`, `.csv`, `.json`, `.tsv` | 1 item (full content) | Single unit |

**Note on Word documents:** The `.doc` and `.docx` file formats do not store page boundaries—pages are a rendering artifact determined by fonts, margins, and printer settings. The library returns the full document as a single text unit.

**Note on generators:** All extractors return generators. Most formats yield a single content object, but `.mbox` files can yield multiple `EmailContent` objects (one per email in the mailbox). Use `next()` for single-item formats or iterate with `for` to handle all cases.

### Basic Usage Examples

```python
import sharepoint2text

# Extract from any file - format auto-detected (use next() for single-item formats)
result = next(sharepoint2text.read_file("quarterly_report.docx"))
print(result.get_full_text())

# Check format support before processing
if sharepoint2text.is_supported_file("document.xyz"):
    for result in sharepoint2text.read_file("document.xyz"):
        print(result.get_full_text())

# Access metadata
result = next(sharepoint2text.read_file("presentation.pptx"))
meta = result.get_metadata()
print(f"Author: {meta.author}, Modified: {meta.modified}")
print(meta.to_dict())  # Convert to dictionary

# Process emails (mbox can contain multiple emails)
for email in sharepoint2text.read_file("mailbox.mbox"):
    print(f"From: {email.from_email.address}")
    print(f"Subject: {email.subject}")
    print(email.get_full_text())
```

### Working with Structured Content

```python
import sharepoint2text

# Excel: iterate over sheets
result = next(sharepoint2text.read_file("budget.xlsx"))
for sheet in result.sheets:
    print(f"Sheet: {sheet.name}")
    print(f"Rows: {len(sheet.data)}")  # List of row dictionaries
    print(sheet.text)                   # Text representation

# PowerPoint: iterate over slides
result = next(sharepoint2text.read_file("deck.pptx"))
for slide in result.slides:
    print(f"Slide {slide.slide_number}: {slide.title}")
    print(slide.content_placeholders)  # Body text
    print(slide.images)                # Image metadata

# PDF: iterate over pages
result = next(sharepoint2text.read_file("report.pdf"))
for page_num, page in result.pages.items():
    print(f"Page {page_num}: {page.text[:100]}...")
    print(f"Images: {len(page.images)}")

# Email: access email-specific fields
email = next(sharepoint2text.read_file("message.eml"))
print(f"From: {email.from_email.name} <{email.from_email.address}>")
print(f"To: {', '.join(e.address for e in email.to_emails)}")
print(f"Subject: {email.subject}")
print(f"Body: {email.body_plain or email.body_html}")
```

### Using Format-Specific Extractors with BytesIO

For API responses or in-memory data:

```python
import sharepoint2text
import io

# Direct extractor usage with BytesIO (returns generator, use next() for single items)
with open("document.docx", "rb") as f:
    result = next(sharepoint2text.read_docx(io.BytesIO(f.read()), path="document.docx"))

# Get extractor dynamically based on filename
def extract_from_api(filename: str, content: bytes):
    extractor = sharepoint2text.get_extractor(filename)
    # Returns a generator - iterate or use next()
    return list(extractor(io.BytesIO(content), path=filename))

results = extract_from_api("report.pdf", pdf_bytes)
for result in results:
    print(result.get_full_text())
```

## API Reference

### Main Functions

```python
import sharepoint2text

# Read any supported file (recommended entry point)
# Returns a generator - use next() for single-item formats or iterate for all
for result in sharepoint2text.read_file(path: str | Path):
    ...

# Check if a file extension is supported
supported = sharepoint2text.is_supported_file(path: str) -> bool

# Get extractor function for a file type
extractor = sharepoint2text.get_extractor(path: str) -> Callable[[io.BytesIO, str | None], Generator[ContentType, Any, None]]
```

### Format-Specific Extractors

All accept `io.BytesIO` and optional `path` for metadata population. All return generators:

```python
sharepoint2text.read_docx(file: io.BytesIO, path: str | None = None) -> Generator[DocxContent, Any, None]
sharepoint2text.read_doc(file: io.BytesIO, path: str | None = None) -> Generator[DocContent, Any, None]
sharepoint2text.read_xlsx(file: io.BytesIO, path: str | None = None) -> Generator[XlsxContent, Any, None]
sharepoint2text.read_xls(file: io.BytesIO, path: str | None = None) -> Generator[XlsContent, Any, None]
sharepoint2text.read_pptx(file: io.BytesIO, path: str | None = None) -> Generator[PptxContent, Any, None]
sharepoint2text.read_ppt(file: io.BytesIO, path: str | None = None) -> Generator[PptContent, Any, None]
sharepoint2text.read_pdf(file: io.BytesIO, path: str | None = None) -> Generator[PdfContent, Any, None]
sharepoint2text.read_plain_text(file: io.BytesIO, path: str | None = None) -> Generator[PlainTextContent, Any, None]
sharepoint2text.read_email__eml_format(file: io.BytesIO, path: str | None = None) -> Generator[EmailContent, Any, None]
sharepoint2text.read_email__msg_format(file: io.BytesIO, path: str | None = None) -> Generator[EmailContent, Any, None]
sharepoint2text.read_email__mbox_format(file: io.BytesIO, path: str | None = None) -> Generator[EmailContent, Any, None]
```

### Return Types

All content types implement the common interface:

```python
class ExtractionInterface(Protocol):
    def iterator() -> Iterator[str]           # Iterate over logical units
    def get_full_text() -> str                # Complete text as string
    def get_metadata() -> FileMetadataInterface  # Metadata with to_dict()
```

#### DocxContent (.docx)

```python
result.metadata       # DocxMetadata (title, author, created, modified, ...)
result.paragraphs     # List[DocxParagraph] (text, style, runs with formatting)
result.tables         # List[List[List[str]]] (cell data)
result.images         # List[DocxImage] (filename, content_type, data, size_bytes)
result.headers        # List[DocxHeaderFooter]
result.footers        # List[DocxHeaderFooter]
result.hyperlinks     # List[DocxHyperlink] (text, url)
result.footnotes      # List[DocxNote] (id, text)
result.endnotes       # List[DocxNote]
result.comments       # List[DocxComment] (author, date, text)
result.sections       # List[DocxSection] (page dimensions, margins)
result.full_text      # str (pre-computed full text)
```

#### DocContent (.doc)

```python
result.metadata         # DocMetadata (title, author, num_pages, num_words, num_chars, ...)
result.main_text        # str (main document body)
result.footnotes        # str (concatenated footnotes)
result.headers_footers  # str (concatenated headers/footers)
result.annotations      # str (concatenated annotations)
```

#### XlsxContent / XlsContent (.xlsx, .xls)

```python
result.metadata   # XlsxMetadata / XlsMetadata (title, creator, created, modified, ...)
result.sheets     # List[XlsxSheet / XlsSheet]

# Each sheet:
sheet.name   # str (sheet name)
sheet.data   # List[Dict[str, Any]] (rows as dictionaries)
sheet.text   # str (text representation)
```

#### PptxContent (.pptx)

```python
result.metadata   # PptxMetadata (title, author, created, modified, ...)
result.slides     # List[PPTXSlide]

# Each slide:
slide.slide_number          # int (1-indexed)
slide.title                 # str
slide.footer                # str
slide.content_placeholders  # List[str] (body content)
slide.other_textboxes       # List[str] (free-form text)
slide.images                # List[PPTXImage] (filename, content_type, size_bytes, blob)
slide.text                  # str (pre-computed combined text)
```

#### PptContent (.ppt)

```python
result.metadata   # PptMetadata (title, author, num_slides, created, modified, ...)
result.slides     # List[PptSlideContent]
result.all_text   # List[str] (flat list of all text)

# Each slide:
slide.slide_number   # int (1-indexed)
slide.title          # str | None
slide.body_text      # List[str]
slide.other_text     # List[str]
slide.notes          # List[str] (speaker notes)
slide.text_combined  # str (property: title + body + other)
slide.all_text       # List[PptTextBlock] (with text_type info)
```

#### PdfContent (.pdf)

```python
result.metadata    # PdfMetadata (total_pages)
result.pages       # Dict[int, PdfPage] (1-indexed)

# Each page:
page.text    # str
page.images  # List[PdfImage] (index, name, width, height, data, format)
```

#### PlainTextContent (.txt, .csv, .json, .tsv)

```python
result.content   # str (full file content)
result.metadata  # FileMetadataInterface (filename, file_extension, file_path, folder_path)
```

#### EmailContent (.eml, .msg, .mbox)

```python
result.from_email    # EmailAddress (name, address)
result.to_emails     # List[EmailAddress]
result.to_cc         # List[EmailAddress]
result.to_bcc        # List[EmailAddress]
result.reply_to      # List[EmailAddress]
result.subject       # str
result.in_reply_to   # str (message ID of parent email)
result.body_plain    # str (plain text body)
result.body_html     # str (HTML body)
result.metadata      # EmailMetadata (date, message_id, plus file metadata)

# EmailAddress structure:
email.name     # str (display name)
email.address  # str (email address)
```

## Examples

### Bulk Processing

```python
import sharepoint2text
from pathlib import Path

def extract_all_documents(folder: Path) -> dict[str, list[str]]:
    """Extract text from all supported files in a folder."""
    results = {}

    for file_path in folder.rglob("*"):
        if sharepoint2text.is_supported_file(str(file_path)):
            try:
                # Collect all content from the generator (handles mbox with multiple emails)
                texts = [result.get_full_text() for result in sharepoint2text.read_file(file_path)]
                results[str(file_path)] = texts
            except Exception as e:
                print(f"Failed to extract {file_path}: {e}")

    return results
```

### Extract Images

```python
import sharepoint2text

# From PDF
result = next(sharepoint2text.read_file("document.pdf"))
for page_num, page in result.pages.items():
    for img in page.images:
        with open(f"page{page_num}_{img.name}.{img.format}", "wb") as out:
            out.write(img.data)

# From PowerPoint
result = next(sharepoint2text.read_file("slides.pptx"))
for slide in result.slides:
    for img in slide.images:
        with open(img.filename, "wb") as out:
            out.write(img.blob)

# From Word
result = next(sharepoint2text.read_file("document.docx"))
for img in result.images:
    if img.data:
        with open(img.filename, "wb") as out:
            out.write(img.data.getvalue())
```

### Email Processing

```python
import sharepoint2text

# Process a single email file (.eml or .msg)
email = next(sharepoint2text.read_file("message.eml"))
print(f"From: {email.from_email.name} <{email.from_email.address}>")
print(f"Subject: {email.subject}")
print(f"Date: {email.metadata.date}")
print(f"Body:\n{email.body_plain}")

# Process a mailbox with multiple emails (.mbox)
for i, email in enumerate(sharepoint2text.read_file("archive.mbox")):
    print(f"\n--- Email {i + 1} ---")
    print(f"From: {email.from_email.address}")
    print(f"To: {', '.join(e.address for e in email.to_emails)}")
    print(f"Subject: {email.subject}")
    if email.to_cc:
        print(f"CC: {', '.join(e.address for e in email.to_cc)}")
```

### RAG Pipeline Integration

```python
import sharepoint2text

def prepare_for_rag(file_path: str) -> list[dict]:
    """Prepare document chunks for RAG ingestion."""
    chunks = []

    # Handle all content items from the generator
    for result in sharepoint2text.read_file(file_path):
        meta = result.get_metadata()

        for i, unit in enumerate(result.iterator()):
            if unit.strip():  # Skip empty units
                chunks.append({
                    "text": unit,
                    "metadata": {
                        "source": file_path,
                        "chunk_index": i,
                        "author": getattr(meta, "author", None),
                        "title": getattr(meta, "title", None),
                    }
                })
    return chunks
```

## Requirements

- Python >= 3.10
- olefile >= 0.47
- openpyxl >= 3.1.5
- pandas >= 2.3.3
- pypdf >= 6.5.0
- python-docx >= 1.2.0
- python-pptx >= 1.0.2
- python-calamine >= 0.6.1
- mail-parser >= 3.15.0
- msg-parser >= 1.2.0

## License

Apache 2.0 - see [LICENSE](LICENSE) for details.
