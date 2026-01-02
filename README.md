# sharepoint-to-text

A **pure Python** library for extracting text, metadata, and structured elements from Microsoft Office files—both modern (`.docx`, `.xlsx`, `.pptx`) and legacy (`.doc`, `.xls`, `.ppt`) formats—plus PDF, email formats, and plain text.

**Install:** `pip install sharepoint-to-text`
**Python import:** `import sharepoint2text`
**CLI (text):** `sharepoint2text /path/to/file.docx > extraction.txt`
**CLI (JSON):** `sharepoint2text --json /path/to/file.docx > extraction.json` (no binary by default; add `--binary` to include)

## What You Get

- **Unified API**: `sharepoint2text.read_file(path)` yields one or more typed extraction results.
- **Typed results**: each format returns a specific dataclass (e.g. `DocxContent`, `PdfContent`) that also supports the common interface.
- **Text**: `get_full_text()` or `iterate_text()` (pages / slides / sheets depending on format).
- **Structured content**: tables and images where the format supports it.
- **Metadata**: file metadata (plus format-specific metadata where available).
- **Serialization**: `result.to_json()` returns a JSON-serializable dict.

## Why This Library?

### Pure Python, No System Dependencies

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

### Legacy Microsoft Office

| Format             | Extension | Description                    |
|--------------------|-----------|--------------------------------|
| Word 97-2003       | `.doc`    | Word 97-2003 documents         |
| Excel 97-2003      | `.xls`    | Excel 97-2003 spreadsheets     |
| PowerPoint 97-2003 | `.ppt`    | PowerPoint 97-2003 presentations |

### Modern Microsoft Office

| Format          | Extension | Description                    |
|-----------------|-----------|--------------------------------|
| Word 2007+      | `.docx`   | Word 2007+ documents           |
| Excel 2007+     | `.xlsx`   | Excel 2007+ spreadsheets       |
| PowerPoint 2007+| `.pptx`   | PowerPoint 2007+ presentations |

### OpenDocument

| Format       | Extension | Description               |
|--------------|-----------|---------------------------|
| Text         | `.odt`    | OpenDocument Text         |
| Presentation | `.odp`    | OpenDocument Presentation |
| Spreadsheet  | `.ods`    | OpenDocument Spreadsheet  |

### Email

| Format | Extension | Description                           |
|--------|-----------|---------------------------------------|
| EML    | `.eml`    | RFC 822 email format                  |
| MSG    | `.msg`    | Microsoft Outlook email format        |
| MBOX   | `.mbox`   | Unix mailbox format (multiple emails) |

### Plain Text

| Format     | Extension | Description              |
|------------|-----------|--------------------------|
| Plain Text | `.txt`    | Plain text files         |
| Markdown   | `.md`     | Markdown                 |
| RTF        | `.rtf`    | Rich Text Format         |
| CSV        | `.csv`    | Comma-separated values   |
| TSV        | `.tsv`    | Tab-separated values     |
| JSON       | `.json`   | JSON files               |

### PDF

| Format | Extension | Description    |
|--------|-----------|----------------|
| PDF    | `.pdf`    | PDF documents  |

### HTML

| Format | Extension | Description         |
|--------|-----------|---------------------|
| HTML   | `.html`   | HTML documents      |
| HTML   | `.htm`    | HTML documents      |

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

## Libraries

### Core Libraries (runtime)

These are required for normal use of the library:

- `defusedxml`: Hardened XML parsing for OOXML/ODF formats
- `mail-parser`: RFC 822 email parsing (`.eml`)
- `msg-parser`: Outlook `.msg` extraction
- `olefile`: OLE2 container parsing for legacy Office formats
- `openpyxl`: `.xlsx` parsing
- `pypdf`: `.pdf` parsing
- `xlrd`: `.xls` parsing

### Development Libraries

These are only needed for development workflows:

- `pytest`: test runner
- `pre-commit`: linting/format hooks
- `black`: code formatter

## Quick Start

### The Unified Interface

`sharepoint2text.read_file(...)` returns a **generator** of extraction results implementing a common interface. Most formats yield a single item, but some (notably `.mbox`) can yield multiple items.

```python
import sharepoint2text

# Works identically for ANY supported format
# Most formats yield a single item, so use next() for convenience
for result in sharepoint2text.read_file("document.docx"):  # or .doc, .pdf, .pptx, etc.
    # Methods available on ALL content types:
    text = result.get_full_text()  # Complete text as a single string
    metadata = result.get_metadata()  # File metadata (author, dates, etc.)

    # Iterate over logical units (varies by format - see below)
    for unit in result.iterate_text():
        print(unit)

    # Iterate over extracted images
    for image in result.iterate_images():
        print(image)

    # Iterate over extracted tables
    for table in result.iterate_tables():
        print(table)

# For single-item formats, you can use next() directly:
result = next(sharepoint2text.read_file("document.docx"))
print(result.get_full_text())
```

Notes: `ImageInterface` provides `get_bytes()`, `get_content_type()`, `get_caption()`, `get_description()`, and `get_metadata()` (unit index, image index, content type, width, height). `TableInterface` provides `get_table()` (rows as lists) and `get_dim()` (rows, columns).

Most results also expose **format-specific structured fields** (e.g. `PdfContent.pages`, `PptxContent.slides`, `XlsxContent.sheets`) in addition to the common interface—see **Return Types** below.

### JSON Output (`to_json()`)

All extraction results support `to_json()` for a JSON-serializable representation of the extracted data (including nested dataclasses).

```python
import json
import sharepoint2text

result = next(sharepoint2text.read_file("document.docx"))
print(json.dumps(result.to_json()))
```

To restore objects from JSON, use `ExtractionInterface.from_json(...)`.

```python
from sharepoint2text.extractors.data_types import ExtractionInterface

restored = ExtractionInterface.from_json(result.to_json())
```

### Understanding `iterate_text()` Output by Format

Different file formats have different natural structural units:

| Format | `iterate_text()` yields | Notes |
|--------|-------------------------|-------|
| `.docx`, `.doc`, `.odt` | 1 item (full text) | Word/text documents have no page structure in the file format |
| `.xlsx`, `.xls`, `.ods` | 1 item per **sheet** | Each yield contains sheet content |
| `.pptx`, `.ppt`, `.odp` | 1 item per **slide** | Each yield contains slide text |
| `.pdf` | 1 item per **page** | Each yield contains page text |
| `.eml`, `.msg` | 1 item (email body) | Plain text or HTML body |
| `.mbox` | 1 item per **email** | Mailboxes can contain multiple emails |
| `.txt`, `.csv`, `.json`, `.tsv` | 1 item (full content) | Single unit |

**Note on Word documents:** The `.doc` and `.docx` file formats do not store page boundaries—pages are a rendering artifact determined by fonts, margins, and printer settings. The library returns the full document as a single text unit.

**Note on generators:** All extractors return generators. Most formats yield a single content object, but `.mbox` files can yield multiple `EmailContent` objects (one per email in the mailbox). Use `next()` for single-item formats or iterate with `for` to handle all cases.

### Choosing Between `get_full_text()` and `iterate_text()`

The interface provides two methods for accessing text content, and **you must decide which is appropriate for your use case**:

| Method | Returns | Best for |
|--------|---------|----------|
| `get_full_text()` | All text as a single string | Simple extraction, full-text search, when structure doesn't matter |
| `iterate_text()` | Yields logical units (pages, slides, sheets) | RAG pipelines, per-unit indexing, preserving document structure |

**For RAG and vector storage:** Consider whether storing pages/slides/sheets as separate chunks with metadata (e.g., page numbers) benefits your retrieval strategy. This allows more precise source attribution when users query your system.

```python
# Option 1: Store entire document as one chunk
result = next(sharepoint2text.read_file("report.pdf"))
store_in_vectordb(text=result.get_full_text(), metadata={"source": "report.pdf"})

# Option 2: Store each page separately with page numbers
result = next(sharepoint2text.read_file("report.pdf"))
for page_num, page_text in enumerate(result.iterate_text(), start=1):
    store_in_vectordb(
        text=page_text,
        metadata={"source": "report.pdf", "page": page_num}
    )
```

**Trade-offs to consider:**
- **Per-unit storage** enables citing specific pages/slides in responses, but creates more chunks
- **Full-text storage** is simpler and may work better for small documents
- **Word documents** (`.doc`, `.docx`) only yield one unit from `iterate_text()` since they lack page structure—for these formats, both methods are equivalent

### Format-Specific Notes on `get_full_text()`

`get_full_text()` is intended as a convenient “best default” for each format. In a few formats it intentionally differs from a plain `"\n".join(iterate_text())`, or it omits optional content unless you opt in:

| Format | `get_full_text()` default behavior | Not included by default / where to find it |
|--------|------------------------------------|--------------------------------------------|
| `.doc` | Prepends `metadata.title` (if present) and returns main document body | `footnotes`, `headers_footers`, `annotations` are separate fields (`DocContent`) |
| `.docx` | Returns `base_full_text` (formulas omitted by default) | Pass `include_formulas=True` / `include_comments=True` to `DocxContent.get_full_text(...)` |
| `.ppt` | Per-slide `title + body + other` concatenated | Speaker notes live in `slide.notes` (`PptSlideContent`) |
| `.pptx` | Per-slide `base_text` concatenated | Pass `include_formulas/include_comments/include_image_captions` to `PptxContent.get_full_text(...)` |
| `.odp` | Per-slide `text_combined` concatenated | Pass `include_notes/include_annotations` to `OdpContent.get_full_text(...)` |
| `.xls` | Concatenation of sheet `text` blocks (no sheet names) | Sheet names are available as `sheet.name` (`XlsSheet`) |
| `.xlsx`, `.ods` | Includes sheet name + sheet text for each sheet | Images are available via `iterate_images()` / sheet image lists |
| `.pdf` | Concatenation of extracted page text | Tables/images are available via `iterate_tables()` / `iterate_images()` (`PdfContent.pages`) |
| `.eml`, `.msg`, `.mbox` | Returns `body_plain` when present, else `body_html` | Attachments are in `EmailContent.attachments` and can be extracted via `iterate_supported_attachments()` |
| `.txt`, `.csv`, `.tsv`, `.json`, `.md`, `.html` | Returns stripped content (leading/trailing whitespace removed) | Use the raw fields (`.content`) if you need untrimmed text |
| `.rtf` | Returns the extractor’s `full_text` when available | `iterate_text()` yields per-page text when explicit `\\page` breaks exist |

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
for page_num, page in enumerate(result.pages, start=1):
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

## CLI

After installation, a `sharepoint2text` command is available. It accepts a single file path and prints the extracted full text to stdout by default.

```bash
sharepoint2text /path/to/file.pdf > extraction.txt
```

To emit structured output, use `--json` (prints `result.to_json()` to stdout).

```bash
sharepoint2text --json /path/to/file.pdf > extraction.json
```

Some formats include binary payloads (e.g., embedded images in Office/PDF files, email attachments). The CLI omits binary payloads in JSON by default (emits `null` for binary fields). Use `--binary` to include base64 blobs:

```bash
sharepoint2text --json /path/to/file.pdf > extraction.json

# include binary payloads
sharepoint2text --json --binary /path/to/file.pdf > extraction.with-binary.json
```

- Without `--json`, multiple items (e.g. `.mbox`) are separated by a blank line.
- With `--json`, stdout is a single JSON object (one item) or a JSON array (multiple items).

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
sharepoint2text.read_odt(file: io.BytesIO, path: str | None = None) -> Generator[OdtContent, Any, None]
sharepoint2text.read_odp(file: io.BytesIO, path: str | None = None) -> Generator[OdpContent, Any, None]
sharepoint2text.read_ods(file: io.BytesIO, path: str | None = None) -> Generator[OdsContent, Any, None]
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
    def iterate_text() -> Iterator[str]          # Iterate over logical units
    def iterate_images() -> Generator[ImageInterface, None, None]
    def iterate_tables() -> Generator[TableInterface, None, None]
    def get_full_text() -> str                   # Complete text as string
    def get_metadata() -> FileMetadataInterface  # Metadata with to_dict()
    def to_json() -> dict                        # JSON-serializable representation
    @classmethod
    def from_json(data: dict) -> "ExtractionInterface"
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

#### OdpContent (.odp)

```python
result.metadata   # OdpMetadata (title, creator, creation_date, generator, ...)
result.slides     # List[OdpSlide]

# Each slide:
slide.slide_number   # int (1-indexed)
slide.name           # str (slide name)
slide.title          # str
slide.body_text      # List[str]
slide.other_text     # List[str]
slide.tables         # List[List[List[str]]] (tables on slide)
slide.annotations    # List[OdpAnnotation] (comments)
slide.images         # List[OdpImage] (embedded images with href, name, data, size_bytes)
slide.notes          # List[str] (speaker notes)
slide.text_combined  # str (property: title + body + other)
```

#### OdsContent (.ods)

```python
result.metadata   # OdsMetadata (title, creator, creation_date, generator, ...)
result.sheets     # List[OdsSheet]

# Each sheet:
sheet.name         # str (sheet name)
sheet.data         # List[Dict[str, Any]] (row data with column keys A, B, C, ...)
sheet.text         # str (tab-separated cell values, newline-separated rows)
sheet.annotations  # List[OdsAnnotation] (cell comments)
sheet.images       # List[OdsImage] (embedded images)
```

#### PdfContent (.pdf)

```python
result.metadata    # PdfMetadata (total_pages)
result.pages       # List[PdfPage]

# Each page:
page.text    # str
page.images  # List[PdfImage] (index, name, width, height, data, format)
page.tables  # List[List[List[str]]]
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

#### HtmlContent (.html, .htm)

```python
result.content   # str (plain text content)
result.tables    # List[List[List[str]]] (table cell values)
result.headings  # List[Dict[str, str]] (level/text)
result.links     # List[Dict[str, str]] (text/href)
result.metadata  # HtmlMetadata (title, language, charset, ...)
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
for page_num, page in enumerate(result.pages, start=1):
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

        for i, unit in enumerate(result.iterate_text()):
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

## Exceptions

- `ExtractionFileFormatNotSupportedError`: Raised when no extractor exists for a given file type (e.g., unsupported extension/MIME mapping in the router).
- `ExtractionFileEncryptedError`: Raised when an extractor detects encryption or password protection (e.g., encrypted PDF, OOXML/ODF password-protected files, legacy Office with FILEPASS/encryption flags).
- `LegacyMicrosoftParsingError`: Raised when legacy Office parsing fails for non-encryption reasons (corrupt OLE streams, invalid headers, or unsupported legacy variations).

## License

Apache 2.0 - see [LICENSE](LICENSE) for details.
