# Command Line Interface

The `sharepoint2text` CLI provides a command-line interface for extracting text and structured content from files and folders.

## Installation

The CLI is automatically available after installing the package:

```bash
uv add sharepoint-to-text
# or
pip install sharepoint-to-text
```

## Quick Start

```bash
# Extract text from a single file
sharepoint2text --file document.docx

# Extract all supported files from a folder
sharepoint2text --folder /path/to/documents

# Extract specific file types from a folder
sharepoint2text --folder /path/to/documents --suffixes .docx,.pdf
```

---

## Input Options

### Single File Extraction

Extract content from a single file:

```bash
sharepoint2text --file /path/to/document.docx
sharepoint2text -f /path/to/document.pdf
```

### Folder Extraction

Extract content from all supported files in a folder:

```bash
# Recursive (default) - includes subdirectories
sharepoint2text --folder /path/to/documents

# Non-recursive - top-level only
sharepoint2text --folder /path/to/documents --no-recursive

# Short form
sharepoint2text -d /path/to/documents
```

### Filtering by File Type

When using `--folder`, filter by specific file extensions:

```bash
# Extract only Word and PDF files
sharepoint2text --folder /path/to/documents --suffixes .docx,.pdf

# Extract only text files (short form)
sharepoint2text -d /path/to/documents -s .txt,.md

# Suffixes work with or without leading dot
sharepoint2text -d /path/to/documents -s docx,pdf,txt
```

---

## Output Options

### Standard Output (Default)

By default, extracted content is written to stdout:

```bash
# Print to terminal
sharepoint2text --file document.docx

# Redirect to file using shell
sharepoint2text --file document.docx > output.txt
```

### Single File Output

Write all extracted content to a single file:

```bash
# Single file extraction to file
sharepoint2text --file document.docx --output result.txt

# Folder extraction combined into single file
sharepoint2text --folder /path/to/docs --output combined.txt

# JSON output to file
sharepoint2text --folder /path/to/docs --json --output results.json
```

### Folder Output (Mirrored Structure)

When extracting from a folder, you can write each file separately to an output folder, preserving the directory structure:

```bash
# Write each file separately to output folder
sharepoint2text --folder /input/docs --output /output/extracted/

# The output structure mirrors the input:
# /input/docs/report.docx      -> /output/extracted/report.txt
# /input/docs/sub/data.xlsx    -> /output/extracted/sub/data.txt
# /input/docs/sub/notes.pdf    -> /output/extracted/sub/notes.txt
```

**How it works:**

- If `--output` is an existing directory, files are written separately
- If `--output` is a new path without extension or ending with `/`, it's created as a directory
- If `--output` has a file extension (e.g., `.txt`, `.json`), all results are combined into that file

```bash
# Existing folder -> separate files
sharepoint2text -d ./docs -o ./output/

# New path without extension -> creates folder, separate files
sharepoint2text -d ./docs -o ./extracted

# Path with extension -> single combined file
sharepoint2text -d ./docs -o ./results.txt
```

---

## Output Formats

### Plain Text (Default)

Extracts full text content:

```bash
sharepoint2text --file document.docx
```

### JSON Format

Emit structured JSON with full extraction objects:

```bash
sharepoint2text --file document.docx --json
sharepoint2text -f document.docx -j
```

Output structure:
```json
[
    {
        "_type": "DocxContent",
        "paragraphs": [...],
        "tables": [...],
        "metadata": {...}
    }
]
```

### JSON Unit Format

Emit JSON with per-unit extraction (pages, slides, sheets, etc.):

```bash
sharepoint2text --file presentation.pptx --json-unit
sharepoint2text -f spreadsheet.xlsx -u
```

Output structure:
```json
[
    {
        "_type": "PptxUnit",
        "slide_number": 1,
        "title": "Introduction",
        "content": "..."
    },
    {
        "_type": "PptxUnit",
        "slide_number": 2,
        ...
    }
]
```

---

## Additional Options

### Include Images

Include base64-encoded image data in JSON output:

```bash
sharepoint2text --file document.docx --json --include-images
sharepoint2text -f document.docx -j -i
```

> Note: `--include-images` requires `--json` or `--json-unit`

### Exclude Email Attachments

Skip extracting supported attachments from email files:

```bash
sharepoint2text --file message.eml --no-attachments
sharepoint2text -f message.msg -n
```

### File Size Limit

Control maximum file size (default: 100 MiB):

```bash
# Set limit to 50 MiB
sharepoint2text --file large.pdf --max-file-size-mb 50

# Disable size limit
sharepoint2text --file huge.pdf --max-file-size-mb 0

# Short form
sharepoint2text -f large.pdf -m 200
```

### Version

Display CLI version:

```bash
sharepoint2text --version
sharepoint2text -v
```

---

## Complete Option Reference

| Option | Short | Description |
|--------|-------|-------------|
| `--file FILE` | `-f` | Path to a single file to extract |
| `--folder FOLDER` | `-d` | Path to a folder to extract from (recursive by default) |
| `--suffixes LIST` | `-s` | Comma-separated suffixes to filter (e.g., `.docx,.pdf`) |
| `--no-recursive` | | Only extract top-level files (no subdirectories) |
| `--output PATH` | `-o` | Output path: file (combined) or folder (separate files) |
| `--json` | `-j` | Emit structured JSON output |
| `--json-unit` | `-u` | Emit per-unit JSON output |
| `--include-images` | `-i` | Include base64 image data in JSON |
| `--no-attachments` | `-n` | Skip email attachment extraction |
| `--max-file-size-mb N` | `-m` | Maximum file size in MiB (default: 100, 0 to disable) |
| `--version` | `-v` | Show version and exit |
| `--help` | `-h` | Show help message and exit |

---

## Examples

### Basic Extraction

```bash
# Extract text from Word document
sharepoint2text -f report.docx

# Extract text from PDF
sharepoint2text -f paper.pdf

# Extract from Excel spreadsheet
sharepoint2text -f data.xlsx
```

### Batch Processing

```bash
# Extract all supported files from a project
sharepoint2text -d ./project --output ./extracted/

# Extract only Office documents as JSON
sharepoint2text -d ./documents -s .docx,.xlsx,.pptx -j -o results.json

# Extract PDFs with folder structure preserved
sharepoint2text -d ./papers -s .pdf -o ./text_versions/
```

### Email Processing

```bash
# Extract email with attachments
sharepoint2text -f message.eml

# Extract email without attachments
sharepoint2text -f message.eml -n

# Extract mailbox to JSON
sharepoint2text -f inbox.mbox -j -o emails.json
```

### JSON Workflows

```bash
# Extract and pipe to jq for processing
sharepoint2text -f document.docx -j | jq '.[] | .metadata'

# Extract units for chunking
sharepoint2text -f report.pdf -u | jq '.[] | .text'

# Batch extract to JSON with images
sharepoint2text -d ./docs -j -i -o ./output.json
```

### Integration with Other Tools

```bash
# Count words in extracted text
sharepoint2text -f document.docx | wc -w

# Search within extracted content
sharepoint2text -d ./docs | grep -i "keyword"

# Extract and index
sharepoint2text -d ./documents -j | ./index_to_elasticsearch.sh
```

---

## Exit Codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | Error (invalid arguments, extraction failure, file not found, etc.) |

---

## Error Handling

The CLI provides clear error messages for common issues:

```bash
# File not found
$ sharepoint2text -f nonexistent.docx
sharepoint2text: File not found: nonexistent.docx

# Unsupported format
$ sharepoint2text -f image.png
sharepoint2text: File format not supported: .png

# File too large
$ sharepoint2text -f huge.pdf -m 10
sharepoint2text: File size 52428800 bytes exceeds CLI maximum of 10485760 bytes

# Invalid option combination
$ sharepoint2text -f doc.docx -s .pdf
sharepoint2text: --suffixes can only be used with --folder
```

---

## Tips and Best Practices

1. **Use folder output for batch processing**: When extracting many files, use folder output (`-o /output/folder/`) to preserve organization.

2. **Filter by type for faster processing**: Use `--suffixes` to only process relevant file types.

3. **Disable images for speed**: Images are disabled by default. Only use `--include-images` when needed.

4. **Use JSON for programmatic access**: The JSON output is structured and machine-readable.

5. **Use JSON-unit for chunking**: For RAG/indexing workflows, `--json-unit` provides pre-chunked content.

6. **Check file sizes first**: For unknown folders, consider using `--max-file-size-mb` to skip very large files.

---

## Related Documentation

- [README.md](../README.md) - Library overview and Python API
- [format-matrix.md](format-matrix.md) - Supported formats and behavior
- [direct-extractors.md](direct-extractors.md) - Direct Python API usage
