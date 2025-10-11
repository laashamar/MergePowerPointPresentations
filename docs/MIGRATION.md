# Migration Guide: Python Package Refactoring

## Overview

This document explains the refactoring of the PowerPoint Merger from a flat Python script structure to a modern, installable Python package following best practices.

## What Changed?

### Project Structure

**Before (Flat Structure):**

```

MergePowerPointPresentations/
├── main.py
├── app.py
├── app_logger.py
├── gui.py
├── powerpoint_core.py
├── run_with_logging.py
├── requirements.txt
└── tests/

```

**After (src Layout):**

```

MergePowerPointPresentations/
├── src/
│   └── merge_powerpoint/          # Main package
│       ├── __init__.py             # Package exports
│       ├── __main__.py             # CLI entry point
│       ├── app.py                  # Refactored
│       ├── app_logger.py           # Refactored
│       ├── gui.py                  # Refactored
│       └── powerpoint_core.py      # Refactored
├── main.py                         # Compatibility shim
├── app.py                          # Compatibility shim
├── app_logger.py                   # Compatibility shim
├── gui.py                          # Compatibility shim
├── powerpoint_core.py              # Compatibility shim
├── run_with_logging.py             # Updated wrapper
├── pyproject.toml                  # Modern config (NEW)
├── requirements.txt                # Still supported
└── tests/                          # Unchanged

```

## Key Improvements

### 1. Modern Package Configuration (`pyproject.toml`)

Replaced `setup.py` approach with modern `pyproject.toml` (PEP 518, PEP 621):

- ✅ Single configuration file for everything
- ✅ Standardized metadata format
- ✅ Automatic CLI script registration
- ✅ Development dependencies management
- ✅ Tool configurations (Black, Ruff, pytest)

### 2. Code Quality Standards

All code now follows strict quality standards:

- **Black Formatted**: 100% PEP 8 compliant formatting
- **Ruff Linted**: Zero linting violations
- **Comprehensive Docstrings**: PEP 257 compliant documentation
- **Type-hint Ready**: Structure supports future type annotations

### 3. Professional Package Structure

The src layout provides:

- Import isolation during development
- Clear separation of concerns
- Better testability
- Industry-standard organization

### 4. CLI Entry Point

New command-line interface after installation:

```bash

# After: pip install .
merge-powerpoint

# Still works:
python main.py
python -m merge_powerpoint

```

## For Users

### Installation Changes

**Before:**

```bash

pip install -r requirements.txt
python main.py

```

**After:**

```bash

pip install .
merge-powerpoint  # New CLI command!

```

### Running the Application

**Multiple options now available:**

```bash

# Option 1: CLI command (recommended after installation)
merge-powerpoint

# Option 2: Python module
python -m merge_powerpoint

# Option 3: Legacy scripts (still work)
python main.py
python run_with_logging.py

```

### Imports (for programmatic use)

**Before:**

```python

from powerpoint_core import PowerPointMerger
from gui import MainWindow

```

**After (both work):**

```python

# New way (recommended)
from merge_powerpoint.powerpoint_core import PowerPointMerger
from merge_powerpoint.gui import MainUI

# Old way (still works via compatibility shims)
from powerpoint_core import PowerPointMerger
from gui import MainUI

```

## For Developers

### Setting Up Development Environment

**New approach:**

```bash

# Clone repository
git clone https://github.com/laashamar/MergePowerPointPresentations.git
cd MergePowerPointPresentations

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install in editable mode with dev tools
pip install -e ".[dev]"

```

This installs:

- The package in editable mode
- All dependencies
- Development tools (pytest, black, ruff, etc.)

### Code Quality Workflow

**New tools and commands:**

```bash

# Format code
black src/merge_powerpoint/

# Check formatting
black --check src/merge_powerpoint/

# Lint code
ruff check src/merge_powerpoint/

# Fix auto-fixable issues
ruff check --fix src/merge_powerpoint/

# Run tests
pytest tests/

# Run tests with coverage
pytest --cov=src/merge_powerpoint tests/

```

### Adding New Features

When adding new code:

1. Add modules to `src/merge_powerpoint/`
2. Format with Black
3. Check with Ruff
4. Add comprehensive docstrings
5. Update `__init__.py` exports if needed
6. Run tests

## Backward Compatibility

### Compatibility Shims

The root-level `.py` files are now "shims" that import from the new package:

```python

# app.py (compatibility shim)
import sys
from pathlib import Path

src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from merge_powerpoint.app import AppController  # noqa: E402, F401

```

This ensures:

- ✅ Existing code continues to work
- ✅ Tests don't need modification
- ✅ Legacy scripts still function
- ✅ Gradual migration path

### Tests

**No changes required!** The tests in `tests/` directory work unchanged because:

1. Compatibility shims provide old import paths
2. Test configuration already supports the new structure
3. conftest.py adds root to path

## Benefits Summary

### For Users

- ✅ Professional CLI command (`merge-powerpoint`)
- ✅ Easy installation (`pip install .`)
- ✅ Multiple ways to run (CLI, module, script)
- ✅ All existing usage patterns still work

### For Developers

- ✅ Modern package structure (src layout)
- ✅ Comprehensive development tools
- ✅ Automated code quality checks
- ✅ Better import isolation
- ✅ Follows Python best practices

### For the Project

- ✅ Professional, maintainable codebase
- ✅ Ready for PyPI publication
- ✅ Easier onboarding for contributors
- ✅ Better documentation
- ✅ Industry-standard structure

## Troubleshooting

### Import Errors

**Problem:** `ModuleNotFoundError: No module named 'merge_powerpoint'`

**Solution:**

```bash

# Install the package
pip install -e .

```

### CLI Command Not Found

**Problem:** `merge-powerpoint: command not found`

**Solution:**

```bash

# Ensure package is installed
pip install .

# Or use alternative methods
python -m merge_powerpoint
python main.py

```

### Tests Failing

**Problem:** Tests can't import modules

**Solution:** Tests should work without modification. If issues persist:

```bash

# Ensure you're in the project root
cd /path/to/MergePowerPointPresentations

# Run tests from root
pytest tests/

```

## Questions?

For questions about the refactoring:

- Check the [ARCHITECTURE.md](ARCHITECTURE.md) for technical details
- See [README.md](../README.md) for usage instructions
- Open an issue on GitHub for help

## Timeline

- **Before Refactoring**: Flat script structure
- **After Refactoring**: Modern package with src layout
- **Compatibility**: Maintained indefinitely through shims
- **Recommendation**: Adopt new patterns for new code
