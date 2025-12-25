# sharepoint-to-text

A Python library for extracting plain text content from files typically found in SharePoint repositories. Supports both modern Office Open XML formats and legacy binary formats (Word 97-2003, Excel 97-2003, PowerPoint 97-2003), plus PDF documents.

## Why this library?

Enterprise SharePoints often contain decades of accumulated documents in various formats. While modern `.docx`, `.xlsx`, and `.pptx` files are well-supported by existing libraries, legacy `.doc`, `.xls`, and `.ppt` files remain common and are harder to process. This library provides a unified interface for extracting text from all these formats, making it ideal for:

- Building RAG (Retrieval-Augmented Generation) pipelines over SharePoint content
- Document indexing and search systems
- Content migration projects
- Automated document processing workflows

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
| JSON              | `.json`   | JSON                             |
| Text              | `.txt`    | Plain text                       |
| CSV               | `.csv`    | CSV                              |
| TSV               | `.tsv`    | TSV                              |

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

### Using read_file (Recommended)

The simplest way to extract content from any supported file:

```python
import sharepoint2text

# Extract content from any supported file
result = sharepoint2text.read_file("quarterly_report.docx")
print(result.get_full_text())

# Works with any supported format
result = sharepoint2text.read_file("budget.xlsx")
for sheet in result.sheets:
    print(f"Sheet: {sheet.name}, Rows: {len(sheet.data)}")
```

### Check if a File is Supported

```python
import sharepoint2text

if sharepoint2text.is_supported_file("document.docx"):
    result = sharepoint2text.read_file("document.docx")
```

### Using Format-Specific Extractors

For more control, use the format-specific extractors directly. These take a `BytesIO` object:

```python
import sharepoint2text
import io

# Extract from a Word document
with open("document.docx", "rb") as f:
    result = sharepoint2text.read_docx(io.BytesIO(f.read()))

print(f"Author: {result.metadata.author}")
print(f"Paragraphs: {len(result.paragraphs)}")
print(f"Tables: {len(result.tables)}")

# Extract from a PDF
with open("report.pdf", "rb") as f:
    result = sharepoint2text.read_pdf(io.BytesIO(f.read()))

for page_num, page in result.pages.items():
    print(f"Page {page_num}: {page.text[:100]}...")
```

### Working with Bytes from APIs

Useful when receiving files from APIs or network requests:

```python
import sharepoint2text
import io

def extract_from_sharepoint_response(filename: str, content: bytes):
    extractor = sharepoint2text.get_extractor(filename)
    return extractor(io.BytesIO(content))

# Example usage
result = extract_from_sharepoint_response("budget.xlsx", file_bytes)
for sheet in result.sheets:
    print(f"Sheet: {sheet.name}, Rows: {len(sheet.data)}")
```

## API Reference

### Functions

```python
import sharepoint2text

# Read any supported file (recommended)
result = sharepoint2text.read_file(path: str | Path)

# Check if a file type is supported
supported = sharepoint2text.is_supported_file(path: str) -> bool

# Get an extractor function for a file type
extractor = sharepoint2text.get_extractor(path: str) -> Callable[[io.BytesIO], ...]

# Format-specific extractors (take io.BytesIO, return typed dataclass)
sharepoint2text.read_docx(file: io.BytesIO) -> MicrosoftDocxContent
sharepoint2text.read_doc(file: io.BytesIO) -> MicrosoftDocContent
sharepoint2text.read_xlsx(file: io.BytesIO) -> MicrosoftXlsxContent
sharepoint2text.read_xls(file: io.BytesIO) -> MicrosoftXlsContent
sharepoint2text.read_pptx(file: io.BytesIO) -> MicrosoftPptxContent
sharepoint2text.read_ppt(file: io.BytesIO) -> PPTContent
sharepoint2text.read_pdf(file: io.BytesIO) -> PdfContent
sharepoint2text.read_plain_text(file: io.BytesIO) -> PlainTextContent
```

### Return Types

All extractors return dataclasses that implement a common interface:

```python
# Common interface for all content types
result.get_full_text() -> str      # Full text as a single string
result.iterator() -> Iterator[str]  # Iterate over text units (pages, slides, sheets)
```

#### MicrosoftDocxContent (.docx)

```python
result.metadata      # MicrosoftDocxMetadata (title, author, created, modified, ...)
result.paragraphs    # List[MicrosoftDocxParagraph]
result.tables        # List[List[List[str]]]
result.images        # List[MicrosoftDocxImage]
result.headers       # List[MicrosoftDocxHeaderFooter]
result.footers       # List[MicrosoftDocxHeaderFooter]
result.hyperlinks    # List[MicrosoftDocxHyperlink]
result.footnotes     # List[MicrosoftDocxNote]
result.comments      # List[MicrosoftDocxComment]
```

#### MicrosoftDocContent (.doc)

```python
result.metadata         # MicrosoftDocMetadata (title, author, num_pages, num_words, ...)
result.main_text        # str
result.footnotes        # str
result.headers_footers  # str
result.annotations      # str
```

#### MicrosoftXlsxContent / MicrosoftXlsContent (.xlsx, .xls)

```python
result.metadata   # MicrosoftXlsx/XlsMetadata (title, creator, created, modified, ...)
result.sheets     # List[MicrosoftXlsx/XlsSheet]

# Each sheet has:
sheet.name   # str
sheet.data   # List[Dict[str, Any]] - rows as dictionaries
sheet.text   # str - text representation
```

#### MicrosoftPptxContent / PPTContent (.pptx, .ppt)

```python
result.metadata  # MicrosoftPptxMetadata / dict
result.slides    # List[MicrosoftPptxSlide] / List[SlideContent]

# Each slide has:
slide.slide_number          # int
slide.title                 # str
slide.content_placeholders  # List[str] (.pptx)
slide.body_text             # List[str] (.ppt)
slide.images                # List[MicrosoftPptxImage] (.pptx)
```

#### PdfContent (.pdf)

```python
result.metadata     # PdfMetadata (total_pages)
result.pages        # Dict[int, PdfPage]

# Each page has:
page.text    # str
page.images  # List[PdfImage] (index, name, width, height, data, format)
```

For full type definitions, see the [source code](https://github.com/Horsmann/sharepoint-to-text/tree/master/sharepoint2text/extractors).

## Examples

### Extract All Text from a PowerPoint

```python
import sharepoint2text

def get_presentation_text(filepath: str) -> str:
    result = sharepoint2text.read_file(filepath)
    # Use the common interface - works for both .ppt and .pptx
    return result.get_full_text()

print(get_presentation_text("presentation.pptx"))
```

### Process Multiple Files

```python
import sharepoint2text
from pathlib import Path

def extract_all_documents(folder: Path) -> dict[str, str]:
    results = {}

    for file_path in folder.rglob("*"):
        if sharepoint2text.is_supported_file(str(file_path)):
            try:
                result = sharepoint2text.read_file(file_path)
                results[str(file_path)] = result.get_full_text()
            except Exception as e:
                print(f"Failed to extract {file_path}: {e}")

    return results

documents = extract_all_documents(Path("./sharepoint_export"))
```

### Extract Images from Documents

```python
import sharepoint2text
import io

# Extract images from PDF
with open("document.pdf", "rb") as f:
    result = sharepoint2text.read_pdf(io.BytesIO(f.read()))

for page_num, page in result.pages.items():
    for img in page.images:
        with open(f"page{page_num}_{img.name}.{img.format}", "wb") as out:
            out.write(img.data)

# Extract images from PowerPoint
with open("slides.pptx", "rb") as f:
    result = sharepoint2text.read_pptx(io.BytesIO(f.read()))

for slide in result.slides:
    for img in slide.images:
        with open(img.filename, "wb") as out:
            out.write(img.blob)
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

## License

Apache 2.0 - see [LICENSE](LICENSE) for details.
