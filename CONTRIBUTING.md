# Contributing to sharepoint-to-text

Thank you for your interest in contributing to sharepoint-to-text! This document provides guidelines and instructions for contributing.

## Getting Started

### Prerequisites

- Python 3.10 or higher
- [uv](https://github.com/astral-sh/uv) (recommended) or pip

### Development Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/your-org/sharepoint-to-text.git
   cd sharepoint-to-text
   ```

2. Create a virtual environment and install dependencies:
   ```bash
   # Using uv (recommended)
   uv venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   uv pip install -e ".[dev]"

   # Or using pip
   python -m venv .venv
   source .venv/bin/activate
   pip install -e ".[dev]"
   ```

3. Install pre-commit hooks:
   ```bash
   pre-commit install
   ```

## Development Workflow

### Running Tests

```bash
pytest
```

### Code Formatting

This project uses [Black](https://github.com/psf/black) for code formatting:

```bash
black sharepoint2text
```

### Pre-commit Hooks

Pre-commit hooks run automatically on `git commit`. To run them manually:

```bash
pre-commit run --all-files
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
   - Implement a `read_{format}(file: io.BytesIO) -> dict` function
   - Return a dictionary with consistent structure (metadata, content, etc.)

2. Update `sharepoint2text/router.py`:
   - Add the MIME type mapping
   - Add the extractor mapping

3. Add tests in `sharepoint2text/tests/`:
   - Create test fixtures in `sharepoint2text/tests/resources/`
   - Add extraction tests in `test_extractions.py`
   - Add router tests in `test_router.py`

4. Update documentation:
   - Add the format to the README.md supported formats table
   - Document the return structure

## Code Style Guidelines

- Follow PEP 8 guidelines
- Use type hints for function parameters and return values
- Write docstrings for public functions and classes
- Keep functions focused and reasonably sized

## Reporting Issues

When reporting issues, please include:

- Python version
- Operating system
- Minimal reproducible example
- Full error traceback (if applicable)
- Sample file (if possible and not containing sensitive data)

## License

By contributing to sharepoint-to-text, you agree that your contributions will be licensed under the Apache 2.0 License.
