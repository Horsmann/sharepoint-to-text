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

### 1. Create a Virtual Environment (Recommended)

Always test in a fresh virtual environment to avoid conflicts:

```bash
# Create a new virtual environment
uv venv testpypi-env

# Activate it
# On macOS/Linux:
source testpypi-env/bin/activate

# On Windows:
testpypi-env\Scripts\activate
```

### 2. Install from TestPyPI

Use the `--index-url` flag to install from TestPyPI, combined with `--extra-index-url` to pull dependencies from the regular PyPI (since TestPyPI may not have all dependencies):


```bash
uv pip install --upgrade \
  --pre \
  --index-strategy unsafe-best-match \
  -i https://test.pypi.org/simple/ \
  --extra-index-url https://pypi.org/simple \
  sharepoint-to-text
```

**Note:** TestPyPI normalizes version strings. A version like `0.2.0.rc01` in `pyproject.toml` becomes `0.2.0rc1` on TestPyPI.

### 4. Test Pre-releases

1) Create a virtual env.
2) source .venv/bin/activate
3) Run below command to install latest pre-release in this environment
```bash
uv pip install --prerelease=allow --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ sharepoint-to-text
```
4) python script.py (note: not `uv run python script.py` - we want to use the pre-release we just installed!)

## Verifying the Installation

After installation, verify that the correct version is installed:

```bash
uv pip show sharepoint-to-text
```

Test a basic import:

```bash
python -c "import sharepoint2text; print('Success!')"
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

1. Ensure you're using `--extra-index-url https://pypi.org/simple/` to fetch dependencies from the regular PyPI

2. Install dependencies manually first:
   ```bash
   uv pip install olefile openpyxl pandas pypdf python-docx python-pptx python-calamine mail-parser msg-parser
   ```

3. Then install from TestPyPI without dependencies:
   ```bash
   uv pip install --index-url https://test.pypi.org/simple/ --no-deps sharepoint-to-text
   ```

### Version Conflicts

If you have an existing installation causing conflicts:

```bash
uv pip uninstall sharepoint-to-text
uv pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ sharepoint-to-text
```

## Quick Reference

| Task | Command |
|------|---------|
| Install latest from TestPyPI | `uv pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ sharepoint-to-text` |
| Install specific version | `uv pip install --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ sharepoint-to-text==X.Y.Z` |
| Install with pre-releases | `uv pip install --prerelease=allow --index-url https://test.pypi.org/simple/ --extra-index-url https://pypi.org/simple/ sharepoint-to-text` |
| Check installed version | `uv pip show sharepoint-to-text` |
| Uninstall | `uv pip uninstall sharepoint-to-text` |

## Links

- [TestPyPI Project Page](https://test.pypi.org/project/sharepoint-to-text/)
- [Production PyPI Project Page](https://pypi.org/project/sharepoint-to-text/)
- [uv Documentation](https://docs.astral.sh/uv/)
- [TestPyPI Documentation](https://packaging.python.org/en/latest/guides/using-testpypi/)
