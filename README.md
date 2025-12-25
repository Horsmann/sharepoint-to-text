# sharepoint-to-text

A Python library for extracting plain text content from files typically found in SharePoint repositories. Supports both modern Office Open XML formats and legacy binary formats (Word 97-2003, Excel 97-2003, PowerPoint 97-2003), plus PDF documents.

## Why this library?

Enterprise SharePoints often contain decades of accumulated documents in various formats. While modern `.docx`, `.xlsx`, and `.pptx` files are well-supported by existing libraries, legacy `.doc`, `.xls`, and `.ppt` files remain common and are harder to process. This library provides a unified interface for extracting text from all these formats, making it ideal for:

- Building RAG (Retrieval-Augmented Generation) pipelines over SharePoint content
- Document indexing and search systems
- Content migration projects
- Automated document processing workflows

## Supported Formats

| Format | Extension | Description |
|--------|-----------|-------------|
| Modern Word | `.docx` | Word 2007+ documents |
| Legacy Word | `.doc` | Word 97-2003 documents |
| Modern Excel | `.xlsx` | Excel 2007+ spreadsheets |
| Legacy Excel | `.xls` | Excel 97-2003 spreadsheets |
| Modern PowerPoint | `.pptx` | PowerPoint 2007+ presentations |
| Legacy PowerPoint | `.ppt` | PowerPoint 97-2003 presentations |
| PDF | `.pdf` | PDF documents |

## Installation

```bash
pip install sharepoint-to-text
```

Or install from source:

```bash
git clone https://github.com/your-org/sharepoint-to-text.git
cd sharepoint-to-text
pip install -e .
```

## Quick Start

### Using the Router (Recommended)

The router automatically detects the file type and returns the appropriate extractor:

```python
from sharepoint2text.router import get_extractor
import io

# Get the extractor based on filename
extractor = get_extractor("quarterly_report.docx")

# Extract content from a file
with open("quarterly_report.docx", "rb") as f:
    result = extractor(io.BytesIO(f.read()))

# Access extracted text
print(result["full_text"])
```

### Working with Bytes Directly

Useful when receiving files from APIs or network requests:

```python
from sharepoint2text.router import get_extractor
import io

def extract_from_sharepoint_response(filename: str, content: bytes) -> dict:
    extractor = get_extractor(filename)
    return extractor(io.BytesIO(content))

# Example usage
result = extract_from_sharepoint_response("budget.xlsx", file_bytes)
for sheet_name, records in result["content"].items():
    print(f"Sheet: {sheet_name}, Rows: {len(records)}")
```

### Using Extractors Directly

You can also import and use specific extractors:

```python
from sharepoint2text.extractors.docx_extractor import read_docx
from sharepoint2text.extractors.pdf_extractor import read_pdf
import io

# Extract from a Word document
with open("document.docx", "rb") as f:
    result = read_docx(io.BytesIO(f.read()))

print(f"Author: {result['metadata']['author']}")
print(f"Paragraphs: {len(result['paragraphs'])}")
print(f"Tables: {len(result['tables'])}")

# Extract from a PDF
with open("report.pdf", "rb") as f:
    result = read_pdf(io.BytesIO(f.read()))

for page_num, page_data in result["pages"].items():
    print(f"Page {page_num}: {page_data['text'][:100]}...")
```

## API Reference

### Router

```python
from sharepoint2text.router import get_extractor

extractor = get_extractor(path: str) -> Callable[[io.BytesIO], dict]
```

Returns an extractor function based on the file extension. The file does not need to exist; only the path/filename is used for detection.

**Raises:** `RuntimeError` if the file type is not supported.

### Return Structures

#### Word Documents (.docx, .doc)

```python
{
    "metadata": {
        "title": str,
        "author": str,
        "created": datetime,
        "modified": datetime,
        ...
    },
    "paragraphs": [...],      # .docx only
    "tables": [...],          # .docx only
    "images": [...],          # .docx only
    "full_text": str,         # .docx: concatenated text
    "text": str,              # .doc: main document text
}
```

#### Excel Spreadsheets (.xlsx, .xls)

```python
{
    "metadata": {
        "title": str,
        "creator": str,
        ...
    },
    "content": {              # .xlsx
        "Sheet1": [{"col1": val, "col2": val}, ...],
        "Sheet2": [...],
    },
    "sheets": {               # .xls
        "Sheet1": [{"col1": val, "col2": val}, ...],
    }
}
```

#### PowerPoint Presentations (.pptx, .ppt)

```python
{
    "metadata": {
        "title": str,
        "author": str,
        ...
    },
    "slides": [
        {
            "slide_number": int,
            "title": str | None,
            "body_text": [...],           # .ppt
            "content_placeholders": [...], # .pptx
            "images": [...],              # .pptx
        },
        ...
    ],
    "slide_count": int,       # .ppt only
}
```

#### PDF Documents (.pdf)

```python
{
    "metadata": {
        "total_pages": int,
    },
    "pages": {
        1: {
            "text": str,
            "images": [
                {
                    "name": str,
                    "width": int,
                    "height": int,
                    "data": bytes,
                    "format": str,
                },
                ...
            ],
        },
        ...
    },
}
```

## Examples

### Extract All Text from a PowerPoint

```python
from sharepoint2text.router import get_extractor
import io

def get_presentation_text(filepath: str) -> str:
    extractor = get_extractor(filepath)
    with open(filepath, "rb") as f:
        result = extractor(io.BytesIO(f.read()))

    texts = []
    for slide in result["slides"]:
        if slide.get("title"):
            texts.append(slide["title"])
        # Handle both .ppt and .pptx formats
        for text in slide.get("body_text", []) + slide.get("content_placeholders", []):
            texts.append(text)

    return "\n".join(texts)

print(get_presentation_text("presentation.pptx"))
```

### Process Multiple Files

```python
from sharepoint2text.router import get_extractor
from pathlib import Path
import io

def extract_all_documents(folder: Path) -> dict[str, dict]:
    results = {}
    supported = {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf"}

    for file_path in folder.rglob("*"):
        if file_path.suffix.lower() in supported:
            try:
                extractor = get_extractor(str(file_path))
                with open(file_path, "rb") as f:
                    results[str(file_path)] = extractor(io.BytesIO(f.read()))
            except Exception as e:
                print(f"Failed to extract {file_path}: {e}")

    return results

documents = extract_all_documents(Path("./sharepoint_export"))
```

### Extract Images from Documents

```python
from sharepoint2text.extractors.pdf_extractor import read_pdf
from sharepoint2text.extractors.pptx_extractor import read_pptx
import io

# Extract images from PDF
with open("document.pdf", "rb") as f:
    result = read_pdf(io.BytesIO(f.read()))

for page_num, page_data in result["pages"].items():
    for img in page_data["images"]:
        with open(f"page{page_num}_{img['name']}.{img['format']}", "wb") as out:
            out.write(img["data"])

# Extract images from PowerPoint
with open("slides.pptx", "rb") as f:
    result = read_pptx(io.BytesIO(f.read()))

for slide in result["slides"]:
    for img in slide.get("images", []):
        with open(img["filename"], "wb") as out:
            out.write(img["blob"])
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

MIT
