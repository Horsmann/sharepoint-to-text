# Command Line Interface

The `sharepoint2text` command extracts one file or a folder. Plain text is the
default; both JSON modes emit the stable version-2 schema.

## Quick Start

```bash
sharepoint2text --file document.docx
sharepoint2text --file document.docx --json
sharepoint2text --file report.pdf --json-unit
sharepoint2text --folder ./docs --suffixes .docx,.pdf
```

## Input

Exactly one input source is required:

```bash
# Single file
sharepoint2text --file /path/to/document.pdf
sharepoint2text -f /path/to/document.pdf

# Recursive folder extraction
sharepoint2text --folder /path/to/documents

# Top-level folder only
sharepoint2text --folder /path/to/documents --no-recursive

# Filter folder input
sharepoint2text --folder /path/to/documents --suffixes .docx,.pdf,.txt
```

Suffixes can be written with or without a leading dot.

## Output

### Plain text

```bash
sharepoint2text --file document.docx
sharepoint2text --folder ./docs --output combined.txt
```

Multiple documents are separated by a blank line.

### Document JSON

```bash
sharepoint2text --file document.docx --json
sharepoint2text -f document.docx -j
```

The output is always an array. Each item is a version-2 envelope:

```json
[
  {
    "schema": "sharepoint2text.extraction",
    "version": 2,
    "document": {
      "format": "docx",
      "source": {
        "filename": "document.docx",
        "extension": ".docx",
        "path": "/path/to/document.docx"
      },
      "metadata": {"keywords": [], "properties": {}},
      "units": [],
      "document_images": [],
      "document_tables": [],
      "document_annotations": [],
      "attachments": [],
      "properties": {}
    }
  }
]
```

### Unit JSON

```bash
sharepoint2text --file presentation.pptx --json-unit
sharepoint2text -f spreadsheet.xlsx -u
```

This is also an array of complete version-2 envelopes, but each envelope
contains exactly one item in `document.units`. Source and document metadata stay
available in every item, making the output self-contained for streaming and
indexing. Document-level images, tables, annotations, and attachments are
retained in every envelope.

### Binary payloads

Binary data is omitted by default. Request image extraction and base64 encoding
of image and attachment payloads explicitly:

```bash
sharepoint2text --file report.pdf --json --include-binary
sharepoint2text -f report.pdf -u -i
```

`--include-binary` requires `--json` or `--json-unit`. The former
`--include-images` spelling remains available as a compatibility alias.

### File and folder destinations

```bash
# One output file
sharepoint2text --file document.docx --output result.txt
sharepoint2text --folder ./docs --json --output results.json

# Mirror input paths below an output folder
sharepoint2text --folder /input/docs --output /output/extracted/
```

For folder input:

- an existing output directory produces one file per input;
- a new extensionless output path is created as a directory;
- an output path with an extension combines all results in one file.

Per-file output uses `.txt` for plain text and `.json` for structured output.
If one input yields multiple documents, such as an `.mbox` mailbox, all of its
documents are written together rather than overwriting one another.

## Attachments

Email attachment records are included by default. Omit them with:

```bash
sharepoint2text --file message.eml --no-attachments
sharepoint2text -f message.msg -n
```

## Size Limits

The default maximum input size is 100 MiB:

```bash
sharepoint2text --file large.pdf --max-file-size-mb 50
sharepoint2text -f large.pdf -m 200
sharepoint2text --file trusted.pdf --max-file-size-mb 0
```

Zero disables the file-size check. Archive decompression protections remain
separate.

## ZIP-Bomb Limits

For a trusted ZIP, Office, or OpenDocument file that exceeds the default
ZIP-bomb thresholds, multiply every default threshold by the same whole number:

```bash
sharepoint2text --file trusted-export.zip --zip-bomb-limit-multiplier 2
```

The multiplier must be an integer from `2` through `10`. To disable ZIP-bomb
checks entirely for trusted input, use the literal value `none`:

```bash
sharepoint2text --file trusted-export.zip --zip-bomb-limit-multiplier none
```

Omitting the option preserves the default protections. Disabling them can allow
a malicious archive to exhaust memory or disk resources.

## Option Reference

| Option | Short | Description |
|---|---|---|
| `--file FILE` | `-f` | Extract one file |
| `--folder FOLDER` | `-d` | Extract a folder recursively |
| `--suffixes LIST` | `-s` | Filter folder input with comma-separated suffixes |
| `--no-recursive` | | Inspect only the folder's top level |
| `--output PATH` | `-o` | Write combined output or mirror into a directory |
| `--json` | `-j` | Emit version-2 document envelopes |
| `--json-unit` | `-u` | Emit one version-2 envelope per unit |
| `--include-binary` | `-i` | Extract images and encode image and attachment payloads as base64 |
| `--no-attachments` | `-n` | Omit email attachment records and payloads |
| `--max-file-size-mb N` | `-m` | Maximum input size; default 100, zero disables |
| `--zip-bomb-limit-multiplier N` | `--zblm` | Multiply all ZIP-bomb limits by 2..10; `none` disables |
| `--version` | `-v` | Print the installed version |
| `--help` | `-h` | Show command help |

`--suffixes` and `--no-recursive` require folder input. `--json` and
`--json-unit` are mutually exclusive. `--include-images` is a compatibility
alias for `--include-binary`.

## Exit Codes

| Code | Meaning |
|---:|---|
| `0` | Extraction succeeded |
| `1` | Arguments, validation, I/O, extraction, or serialization failed |
| `2` | Command syntax could not be parsed by `argparse` |

Errors are written to stderr. Folder extraction logs individual skipped files
and continues when possible. File extraction errors are reported without a
Python traceback.

## Related Documentation

- [README.md](../README.md) — normalized Python API
- [format-matrix.md](format-matrix.md) — format behavior and caveats
