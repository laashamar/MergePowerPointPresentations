# Examples

This directory contains example scripts demonstrating how to use the PowerPoint Merger application.

## Available Examples

### example_refactored_gui.py

Demonstrates the refactored PySide6 GUI with all features:
- Two-column layout with drag-and-drop
- Signal-based architecture
- Custom event handlers
- Proper application configuration

**Run:**
```bash
python examples/example_refactored_gui.py
```

**Features demonstrated:**
- Application initialization with QApplication
- Setting application metadata (name, organization, version)
- Creating and injecting PowerPointMerger backend
- Connecting to all UI signals (files_added, file_removed, etc.)
- Custom event logging
- Error handling
- Welcome dialog

## Requirements

All examples require the package to be installed:

```bash
pip install -e .
```

Or for development:

```bash
pip install -e ".[dev]"
```

## Platform Requirements

- **Windows**: Full functionality with COM automation
- **Linux/macOS**: GUI works but merge requires mocking or Windows VM

For testing on non-Windows platforms:

```bash
export QT_QPA_PLATFORM=offscreen  # For headless testing
python examples/example_refactored_gui.py
```

## Learn More

- See [GUI_GUIDE.md](../docs/GUI_GUIDE.md) for API reference
- See [ARCHITECTURE.md](../docs/ARCHITECTURE.md) for design details
- See [tests/test_gui_refactored.py](../tests/test_gui_refactored.py) for testing examples
