# Contributing to sharepoint-to-text

Thank you for your interest in contributing to sharepoint-to-text. This document
provides guidelines and instructions for contributing.

## Getting Started

### Prerequisites

- Python 3.10 or higher
- [uv](https://github.com/astral-sh/uv) (required for development)

### Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/Horsmann/sharepoint-to-text.git
   cd sharepoint-to-text
   ```

2. Create a virtual environment and install dependencies:
   ```bash
   # Uses uv.lock and installs all dependency groups (including dev)
   uv sync --all-groups

   # Optional: activate the created virtual environment
   source .venv/bin/activate  # On Windows: .venv\\Scripts\\activate
   ```

3. Install pre-commit hooks:
   ```bash
   uv run pre-commit install
   ```

## Development Workflow

### Running Tests

```bash
uv run pytest
```

### Type Checking

```bash
uv run mypy .
```

### Code Formatting

This project uses [Black](https://github.com/psf/black) for code formatting:

```bash
uv run black sharepoint2text
```

### Pre-commit Hooks

Pre-commit hooks run automatically on `git commit`. To run them manually:

```bash
uv run pre-commit run --all-files
```

## Making Changes

### Branching Strategy

1. Create a new branch for your feature or bugfix:
   ```bash
   git checkout -b feature/your-feature-name
   # or
   git checkout -b fix/your-bugfix-name
   ```

2. Make your changes in small, focused commits.

3. Write or update tests as needed.

4. Ensure `uv run pytest` and `uv run mypy .` both pass before submitting.

### Commit Messages

Write clear, descriptive commit messages:

- Use the present tense ("Add feature" not "Added feature")
- Use the imperative mood ("Move cursor to..." not "Moves cursor to...")
- Limit the first line to 72 characters or less
- Reference issues and pull requests liberally after the first line

### Pull Requests

1. Update the CHANGELOG.md with your changes under the `[Unreleased]` section.

2. Ensure your code passes all tests and linting checks.

3. Submit a pull request with a clear title and description.

4. Link any related issues in the PR description.

## Adding Support for New File Formats

If you want to add support for a new file format:

1. Create a new extractor module in `sharepoint2text/parsing/extractors/`:
   - Follow the naming convention: `{format}_extractor.py`
   - Implement a fully typed
     `read_{format}(file_like: BinaryIO, path: str | None = None)` generator.
   - Yield one or more internal parser records compatible with
     `ExtractionRecord` from `sharepoint2text/parsing/extractors/_records.py`.
   - Populate source metadata from `path` when it is available.
   - Keep parser-specific records internal; public entry points normalize them
     to `ExtractedDocument`.
   - Keep behavior consistent with existing extractors:
     - Single-document formats yield exactly one item (e.g., `.pdf`, `.docx`)
     - Multi-item formats yield multiple items (notably `.mbox`, one per email)

2. Update `sharepoint2text/parsing/router.py`:
   - Register the extension and internal extractor key in `_EXTRACTOR_REGISTRY`.
   - Add a literal import case to `_load_registered_extractor`; this is the extractor allowlist.
   - Add aliases to `_EXTENSION_ALIASES` or `_COMPOUND_EXTENSIONS` when needed.
   - Add MIME routing in `sharepoint2text/parsing/mime_types.py` when a stable media type exists.

3. Update normalization when the new record needs explicit handling:
   - Add its fallback format mapping or unit kind in `sharepoint2text/parsing/_normalization.py`.
   - Map useful scalar details to namespaced `properties` keys.
   - Preserve canonical ownership: assets assigned to a unit must not also
     appear in document-level collections.

4. Add tests in `sharepoint2text/tests/`:
   - Create test fixtures in `sharepoint2text/tests/resources/`
   - Add extractor coverage to the relevant module under `sharepoint2text/tests/extractors/`.
   - Add extension, alias, and MIME routing coverage to `test_router.py`.
   - Add a public-boundary check to `test_integration.py` when appropriate,
     confirming `read_file` or `read_bytes` yields `ExtractedDocument`.
   - Add codec coverage to `test_models.py` only when the normalized model or
     version-2 wire schema changes.

5. Update documentation:
   - Add the format to the README.md supported formats table
   - Add its cardinality, unit shape, and caveats to `doc/format-matrix.md`.
   - Document any format-specific namespaced properties consumers may rely on.

## Code Style Guidelines

- Follow PEP 8 guidelines.
- Add complete type hints to every function signature.
- Write docstrings for all public functions and classes.
- Keep functions focused and reasonably sized.
- Use explicit exception handling; do not add bare `except:` blocks.

## Design Notes

- **Keep the public model normalized**: all public extraction entry points yield
  `ExtractedDocument`. Format-specific parser records belong in `_records.py`
  and remain internal.
- **Use the centralized codec**: serialize normalized documents with
  `document_to_dict`, `document_to_json`, `document_from_dict`, and
  `document_from_json` from `sharepoint2text.parsing.models`.
- **Keep routing deterministic**: extension-based routing should work regardless
  of platform MIME databases; MIME routing is a secondary path.
- **Namespace format-specific properties**: use keys such as `xlsx.hidden`
  rather than adding format-specific public fields.
- **Use library exceptions** from `sharepoint2text/parsing/exceptions.py` for
  user-facing failure modes:
  - `ExtractionFileFormatNotSupportedError` for unsupported formats
  - `ExtractionFileEncryptedError` for password-protected/encrypted content
  - `ExtractionLegacyMicrosoftParsingError` for legacy Office parsing failures
  - `ExtractionZipBombError` for unsafe ZIP-based input
  - `ExtractionFailedError` for unexpected extraction failures (usually wrapped by `read_file`)

## Notes on uv

This repository uses `uv.lock` and dependency groups. For development tests and
tooling, use `uv sync --all-groups`.

## Reporting Issues

When reporting issues, please include:

- Python version
- Operating system
- Minimal reproducible example
- Full error traceback (if applicable)
- Sample file (if possible and not containing sensitive data)

## License

By contributing to sharepoint-to-text, you agree that your contributions will be licensed under the Apache 2.0 License.
