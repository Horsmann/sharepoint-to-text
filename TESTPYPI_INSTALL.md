# Installing from TestPyPI

This guide explains how to test-install a pre-release of `sharepoint-to-text` from [TestPyPI](https://test.pypi.org/) using [uv](https://docs.astral.sh/uv/).

## What is TestPyPI?

TestPyPI is a separate instance of the Python Package Index that allows package maintainers to test their release process without affecting the production PyPI. It's useful for:

- Verifying the package builds and installs correctly
- Testing pre-release versions before publishing to production PyPI
- Catching packaging issues early

## Prerequisites

- Python >= 3.10
- uv installed ([installation guide](https://docs.astral.sh/uv/getting-started/installation/))

## Installation Steps

### 1. Create a Project Environment (Recommended)

Always test in a clean project directory to avoid conflicts:

```bash
# Create a clean test project
mkdir testpypi-env
cd testpypi-env

# Initialize a minimal project
uv init --lib

# Create a new virtual environment
uv venv

# Activate it
# On macOS/Linux:
source .venv/bin/activate

# On Windows:
.venv\Scripts\activate
```

### 2. Install from TestPyPI

Use `--default-index` to install from TestPyPI, combined with `--index` to pull dependencies from the regular PyPI (since TestPyPI may not have all dependencies):

```bash
uv add \
  --index-strategy unsafe-best-match \
  --default-index https://test.pypi.org/simple/ \
  --index https://pypi.org/simple/ \
  sharepoint-to-text
```
or
```bash
uv pip \
   install -i https://test.pypi.org/simple/ \
   --index-strategy unsafe-best-match \
   sharepoint-to-text==<version>
```

**Note:** TestPyPI normalizes version strings. A version like `0.2.0.rc01` in `pyproject.toml` becomes `0.2.0rc1` on TestPyPI.

Optional: faster AES handling for encrypted PDFs (avoids the slow fallback crypto and large-PDF image skips):

```bash
uv add \
  --index-strategy unsafe-best-match \
  --default-index https://test.pypi.org/simple/ \
  --index https://pypi.org/simple/ \
  "sharepoint-to-text[pdf-crypto]"
```

### 3. Test Pre-releases

1) Create a project environment (see step 1).
2) source .venv/bin/activate
3) Run below command to install the latest pre-release in this environment
```bash
uv add --prerelease allow \
  --index-strategy unsafe-best-match \
  --default-index https://test.pypi.org/simple/ \
  --index https://pypi.org/simple/ \
  sharepoint-to-text
```
4) Run `uv run python script.py` (or `python script.py` if the venv is activated)

## Verifying the Installation

After installation, verify that the correct version is installed:

```bash
uv run python -c "import importlib.metadata as m; print(m.version('sharepoint-to-text'))"
```

Test a basic import:

```bash
uv run python -c "import sharepoint2text; print('Success!')"
```

Run a quick functionality test:

```python
import sharepoint2text

# Check supported formats
print(sharepoint2text.is_supported_file("document.docx"))  # Should print: True

# List available functions
print(dir(sharepoint2text))
```

## Troubleshooting

### Package Not Found

If you get a "package not found" error:

1. Check that the package was successfully deployed to TestPyPI by visiting:
   https://test.pypi.org/project/sharepoint-to-text/

2. Verify you're using the correct package name (`sharepoint-to-text`, not `sharepoint2text`)

3. Try with `--prerelease=allow` flag to include pre-release versions

### Dependency Issues

If dependencies fail to install:

1. Ensure you're using `--index https://pypi.org/simple/` to fetch dependencies from the regular PyPI

2. Install dependencies manually first:
   ```bash
   uv add olefile openpyxl pandas pypdf python-docx python-pptx python-calamine mail-parser msg-parser
   ```

3. Then install from TestPyPI using the same indexes:
   ```bash
   uv add --default-index https://test.pypi.org/simple/ --index https://pypi.org/simple/ sharepoint-to-text
   ```

### Version Conflicts

If you have an existing installation causing conflicts:

```bash
uv remove sharepoint-to-text
uv add --default-index https://test.pypi.org/simple/ --index https://pypi.org/simple/ sharepoint-to-text
```

## Quick Reference

| Task | Command |
|------|---------|
| Install latest from TestPyPI | `uv add --default-index https://test.pypi.org/simple/ --index https://pypi.org/simple/ sharepoint-to-text` |
| Install specific version | `uv add --default-index https://test.pypi.org/simple/ --index https://pypi.org/simple/ sharepoint-to-text==X.Y.Z` |
| Install with pre-releases | `uv add --prerelease allow --default-index https://test.pypi.org/simple/ --index https://pypi.org/simple/ sharepoint-to-text` |
| Check installed version | `uv run python -c "import importlib.metadata as m; print(m.version('sharepoint-to-text'))"` |
| Uninstall | `uv remove sharepoint-to-text` |

## Links

- [TestPyPI Project Page](https://test.pypi.org/project/sharepoint-to-text/)
- [Production PyPI Project Page](https://pypi.org/project/sharepoint-to-text/)
- [uv Documentation](https://docs.astral.sh/uv/)
- [TestPyPI Documentation](https://packaging.python.org/en/latest/guides/using-testpypi/)
