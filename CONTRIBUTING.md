# Contributing to sharepoint-to-text

Thank you for your interest in contributing to sharepoint-to-text! This document provides guidelines and instructions for contributing.

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

4. Ensure all tests pass before submitting.

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

1. Create a new extractor module in `sharepoint2text/extractors/`:
   - Follow the naming convention: `{format}_extractor.py`
   - Implement a `read_{format}(file_like: io.BytesIO, path: str | None = None)` function
   - Return a **generator** yielding one or more typed dataclasses that implement `ExtractionInterface`
   - Populate file metadata via `metadata.populate_from_path(path)` when `path` is provided
   - Keep behavior consistent with existing extractors:
     - Single-document formats yield exactly one item (e.g., `.pdf`, `.docx`)
     - Multi-item formats yield multiple items (notably `.mbox`, one per email)

2. Update `sharepoint2text/router.py`:
   - Add the extension → extractor registration in `_EXTRACTOR_REGISTRY`
   - If needed, add a MIME type entry in `sharepoint2text/mime_types.py` so MIME-based detection can route too
   - If your extension has common aliases (e.g., `.htm` vs `.html`), add an alias in the router so detection is OS-independent

3. Add tests in `sharepoint2text/tests/`:
   - Create test fixtures in `sharepoint2text/tests/resources/`
   - Add extraction tests in `test_extractions.py`
   - Add router tests in `test_router.py`
   - If your output supports `to_json()`, add a round-trip check (serialize → deserialize) in `test_serialization.py`

4. Update documentation:
   - Add the format to the README.md supported formats table
   - Document the return type and what `iterate_units()` yields (unit text + unit metadata) for the format

## Code Style Guidelines

- Follow PEP 8 guidelines
- Use type hints for function parameters and return values
- Write docstrings for public functions and classes
- Keep functions focused and reasonably sized

## Design Notes

- **Prefer dataclasses** for extraction outputs. They work with `to_json()`/`from_json()` via the shared serialization layer in `sharepoint2text/extractors/serialization.py`.
- **Keep routing deterministic**: extension-based routing should work regardless of platform MIME databases; MIME routing is a helpful secondary path.
- **Use library exceptions** from `sharepoint2text/exceptions.py` for user-facing failure modes:
  - `ExtractionFileFormatNotSupportedError` for unsupported formats
  - `ExtractionFileEncryptedError` for password-protected/encrypted content
  - `LegacyMicrosoftParsingError` for legacy Office parsing failures
  - `ExtractionFailedError` for unexpected extraction failures (usually wrapped by `read_file`)

## Notes on pip

This repository uses `uv.lock` and dependency groups. For development (tests/linting), use `uv sync --all-groups`.

## Reporting Issues

When reporting issues, please include:

- Python version
- Operating system
- Minimal reproducible example
- Full error traceback (if applicable)
- Sample file (if possible and not containing sensitive data)

## License

By contributing to sharepoint-to-text, you agree that your contributions will be licensed under the Apache 2.0 License.
