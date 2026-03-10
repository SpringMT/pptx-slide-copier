# Releasing to PyPI

## Prerequisites

```bash
pip install build twine
```

PyPI account and API token are required. Create a token at https://pypi.org/manage/account/token/

## Release Steps

### 1. Update version

Update the version in **both** files:

- `pyproject.toml` (`version = "X.Y.Z"`)
- `pptx_slide_copier/__init__.py` (`__version__ = "X.Y.Z"`)

### 2. Commit and tag

```bash
git add pyproject.toml pptx_slide_copier/__init__.py
git commit -m "bump version to X.Y.Z"
git tag vX.Y.Z
git push origin main --tags
```

### 3. Clean previous builds

```bash
rm -rf dist/ build/ *.egg-info
```

### 4. Build

```bash
python -m build
```

This creates `dist/pptx_slide_copier-X.Y.Z.tar.gz` and `dist/pptx_slide_copier-X.Y.Z-py3-none-any.whl`.

### 5. Verify the package (optional)

```bash
twine check dist/*
```

### 6. Upload to TestPyPI (optional)

```bash
twine upload --repository testpypi dist/*
```

Verify at https://test.pypi.org/project/pptx-slide-copier/

### 7. Upload to PyPI

```bash
twine upload dist/*
```

You will be prompted for credentials. Use `__token__` as the username and the API token as the password.

To avoid entering credentials each time, create `~/.pypirc`:

```ini
[pypi]
username = __token__
password = pypi-XXXXXXXXXXXXXXXX
```

### 8. Verify

```bash
pip install --upgrade pptx-slide-copier
python -c "import pptx_slide_copier; print(pptx_slide_copier.__version__)"
```
