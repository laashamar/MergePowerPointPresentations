# GUI Migration Guide: Transitioning to the Refactored PySide6 UI

This document provides guidance for using the new refactored PySide6 GUI interface.

## Overview

The refactored GUI (`gui_refactored.py`) implements a modern, two-column interface following PySide6 best practices with comprehensive features including drag-and-drop, threading, and signal-based architecture.

## Quick Start

### Basic Usage

```python
from merge_powerpoint.gui_refactored import MainUI
from merge_powerpoint.powerpoint_core import PowerPointMerger
from PySide6.QtWidgets import QApplication
import sys

app = QApplication(sys.argv)
app.setApplicationName("PowerPoint Merger")
app.setOrganizationName("MergePowerPoint")

# Create backend and inject into UI
merger = PowerPointMerger()
window = MainUI(merger=merger)
window.setWindowTitle("PowerPoint Presentation Merger")
window.resize(1000, 600)
window.show()

sys.exit(app.exec())
```

## Key Features

### 1. Two-Column Layout (3:1 Ratio)

**Left Column**: Main interaction area
- Empty state with drop zone when no files
- Active state with file list when files are added
- Drag-and-drop support for .pptx files

**Right Column**: Configuration and actions
- Clear list button
- Output file configuration
- Merge button

### 2. Signal-Based Architecture

The UI emits signals for all major events:

```python
# Available signals
window.files_added.connect(handler)      # List[str] of file paths
window.file_removed.connect(handler)     # str file path
window.order_changed.connect(handler)    # List[str] new order
window.clear_requested.connect(handler)  # No parameters
window.merge_requested.connect(handler)  # str output path
```

### 3. Threading

Merge operations run in a background QThread, keeping the UI responsive:

```python
# The worker is automatically created and managed
# Progress updates come through signals
worker.progress.connect(on_progress)     # (int current, int total)
worker.finished.connect(on_finished)     # (bool success, str path, str error)
```

### 4. Settings Persistence

The UI remembers the last save location:

```python
# Automatic via QSettings
# Uses application name and organization from QApplication
```

## API Reference

### MainUI Class

```python
class MainUI(QWidget):
    """Main user interface widget.
    
    Signals:
        files_added(list): File paths added
        file_removed(str): File path removed  
        order_changed(list): New file order
        clear_requested(): List cleared
        merge_requested(str): Merge started with output path
    """
    
    def __init__(self, merger: Optional[PowerPointMerger] = None, parent=None):
        """Initialize UI.
        
        Args:
            merger: PowerPointMerger instance (dependency injection)
            parent: Optional parent widget
        """
```

### FileListModel Class

```python
class FileListModel(QStandardItemModel):
    """Model for file list management."""
    
    def add_files(self, paths: List[str]) -> List[str]:
        """Add files, return rejected paths."""
    
    def remove_file(self, path: str) -> bool:
        """Remove file, return success."""
    
    def clear_all(self):
        """Remove all files."""
    
    def get_file_paths(self) -> List[str]:
        """Get ordered file paths."""
    
    def reorder_files(self, new_order: List[str]):
        """Update file order."""
```

## Testing with pytest-qt

```python
import pytest
from PySide6.QtTest import QSignalSpy
from merge_powerpoint.gui_refactored import MainUI
from merge_powerpoint.powerpoint_core import PowerPointMerger

@pytest.fixture
def main_ui(qtbot):
    merger = PowerPointMerger()
    ui = MainUI(merger=merger)
    qtbot.addWidget(ui)
    ui.show()
    qtbot.waitExposed(ui)
    return ui

def test_files_added_signal(main_ui, qtbot, mocker):
    # Mock file system
    mocker.patch('os.path.exists', return_value=True)
    mocker.patch('os.path.isfile', return_value=True)
    
    # Create spy
    spy = QSignalSpy(main_ui.files_added)
    
    # Trigger action
    main_ui._on_files_dropped(["/path/to/file.pptx"])
    
    # Verify
    assert spy.count() == 1
```

## UI States

### Empty State
- Drop zone visible
- File list hidden
- Clear button disabled
- Merge button disabled

### Active State (1 file)
- Drop zone hidden
- File list visible
- Clear button enabled
- Merge button disabled (need 2+ files)

### Active State (2+ files)
- Drop zone hidden
- File list visible
- Clear button enabled
- Merge button enabled

### Merging State
- All controls disabled
- Progress bar visible
- UI remains responsive

## Internationalization

All UI strings are centralized in `UI_STRINGS` dictionary:

```python
UI_STRINGS = {
    "window_title": "PowerPoint Presentation Merger",
    "drop_zone_text": "Drag and drop PowerPoint files here",
    "browse_button": "Browse for Files...",
    "clear_list_button": "Clear List",
    "merge_button": "Merge Presentations",
    # ... etc
}
```

To add a new language:
1. Copy `UI_STRINGS` dict
2. Translate values
3. Load based on locale
4. Pass to MainUI or modify at module level

## Icons and Resources

Icons are managed through Qt resource system:

```bash
# Compile resources
cd resources
pyside6-rcc icons.qrc -o ../src/merge_powerpoint/icons_rc.py
```

Icons are SVG format in `resources/icons/`:
- `plus.svg` - Drop zone icon
- `trash.svg` - Clear button
- `close.svg` - Remove file
- `powerpoint.svg` - File icon
- `folder.svg` - Save location

## Accessibility

The UI implements accessibility features:
- All buttons have text labels
- Keyboard navigation with logical tab order
- Tooltips on all interactive elements
- Screen reader compatible

## Best Practices

### 1. Dependency Injection
Always inject the PowerPointMerger:
```python
merger = PowerPointMerger()
ui = MainUI(merger=merger)
```

### 2. Connect Before Show
Connect signals before showing the window:
```python
ui = MainUI(merger=merger)
ui.files_added.connect(my_handler)
ui.show()
```

### 3. Set Application Identity
Always set app name for QSettings:
```python
app = QApplication(sys.argv)
app.setApplicationName("PowerPoint Merger")
app.setOrganizationName("MergePowerPoint")
```

### 4. Handle Errors
Connect to finished signal to handle errors:
```python
def on_merge_finished(success, path, error):
    if not success:
        print(f"Error: {error}")
```

## Troubleshooting

### Icons Not Showing
Compile the resource file:
```bash
pyside6-rcc resources/icons.qrc -o src/merge_powerpoint/icons_rc.py
```

### Settings Not Saving
Set application identity before creating MainUI.

### Tests Failing
Set Qt platform for headless testing:
```bash
export QT_QPA_PLATFORM=offscreen
pytest tests/test_gui_refactored.py
```

### UI Freezing
The refactored UI uses threading - if it freezes, there may be an issue with the worker thread. Check logs for errors.

## Migration from Original GUI

The original `gui.py` remains available but new projects should use `gui_refactored.py`:

**Advantages**:
- Modern Qt patterns
- Better testability  
- Non-blocking operations
- Signal-based architecture
- Comprehensive test coverage
- Improved user experience

See the full test suite in `tests/test_gui_refactored.py` for complete usage examples.
