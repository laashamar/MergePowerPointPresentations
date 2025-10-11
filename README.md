# PowerPoint Presentation Merger

A modern Python utility to merge multiple PowerPoint (.pptx) files into a
single presentation. This tool uses COM automation to ensure that all
formatting, animations, and embedded content are preserved with perfect fidelity
during the merge process.

## Features

- **Merge multiple .pptx files** in a specified order
- **Preserves original formatting**, transitions, and animations
- **Modern PySide6 GUI** with drag-and-drop support
- **Two-column interface** for intuitive file management
- **Real-time progress tracking** during merge operations
- **Responsive UI** with threaded merge operations (UI never freezes)
- **File reordering** via drag-and-drop
- **Settings persistence** - remembers last save location
- **Comprehensive testing** with pytest-qt
- **Modern package structure** for easy installation and development
- **Command-line interface** for easy execution

## Modern GUI (PySide6)

The application features a modern GUI built with PySide6, offering an intuitive two-column layout:

### Two-Column Layout

- **Left Column (3:1 ratio)**: Main interaction area with smart state management
  - **Empty State**: Shows a drop zone with icon and "Browse for Files" button
  - **Active State**: Displays file list with drag-and-drop reordering
- **Right Column (1:1 ratio)**: Configuration and actions
  - Clear list button with icon
  - Output file configuration
  - Save location selector
  - Large, prominent "Merge Presentations" button

### Key Features

- **Drag-and-Drop**: Drop .pptx files directly onto the interface
- **File Validation**: Only accepts .pptx files, prevents duplicates
- **Signal-Based Architecture**: Follows Qt best practices with proper signal/slot connections
- **Threaded Merge**: Background worker thread prevents UI freezing
- **Progress Feedback**: Indeterminate progress bar during merge operations
- **Keyboard Navigation**: Full keyboard accessibility with logical tab order
- **Settings Persistence**: Remembers last save location between sessions
- **Internationalization Ready**: All UI strings centralized for easy translation

### Using the GUI Programmatically

The GUI can be embedded in your own applications. Usage example:

```python

from merge_powerpoint.gui import MainUI
from merge_powerpoint.powerpoint_core import PowerPointMerger
from PySide6.QtWidgets import QApplication, QMainWindow
import sys

app = QApplication(sys.argv)
merger = PowerPointMerger()

# MainUI is a QWidget, so embed it in a QMainWindow
main_window = QMainWindow()
ui = MainUI(merger=merger)
main_window.setCentralWidget(ui)
main_window.setWindowTitle("PowerPoint Presentation Merger")
main_window.resize(1000, 600)
main_window.show()

sys.exit(app.exec())

```

## Installation

### Option 1: Install from Source (Recommended)

1. **Clone this repository:**

   ```bash

   git clone https://github.com/laashamar/MergePowerPointPresentations.git
   cd MergePowerPointPresentations

   ```

2. **Create a virtual environment (recommended):**

   ```bash

   python -m venv venv
   source venv/bin/activate  # On Windows, use `venv\Scripts\activate`

   ```

3. **Install the package:**

   ```bash

   pip install .

   ```

   Or for development with additional tools:

   ```bash

   pip install -e ".[dev]"

   ```

### Option 2: Pre-built Executable (For End Users)

1. Download the latest release from the [Releases page](https://github.com/laashamar/MergePowerPointPresentations/releases)
2. Extract the archive to your desired location
3. Run `MergePowerPoint.exe` directly

## Usage

### Command Line

After installation, you can run the application using the CLI command:

```bash

merge-powerpoint

```

Or run it as a Python module:

```bash

python -m merge_powerpoint

```

### Step-by-Step Workflow

1. **Add Files**: Click "Add Files" to select PowerPoint presentations (.pptx) to merge
2. **Reorder**: Use "Move Up" and "Move Down" buttons to arrange files in the desired order
3. **Merge**: Click "Merge Files" to combine all presentations into a single file
4. **Save**: Choose a location and filename for the merged presentation

### Features in Detail

- **Add Files**: Select one or multiple PowerPoint files using the file dialog
- **Remove Selected**: Remove specific files from the merge list
- **Clear All**: Remove all files from the list at once
- **Move Up/Down**: Reorder files to control the sequence in the merged presentation
- **Progress Tracking**: Visual progress bar shows merge progress

## Development

### Code Quality Tools

This project uses modern Python development tools:

- **Black**: Code formatting (PEP 8 compliant)
- **Ruff**: Fast Python linter
- **pytest**: Testing framework with pytest-qt for GUI testing
- **mypy**: Static type checking

### Running Tests

The project includes comprehensive test coverage:

```bash

# Run all tests
pytest tests/

# Run with coverage report
pytest --cov=src tests/

# Run only GUI tests
pytest tests/test_gui.py -v

# Run with verbose output
pytest -v tests/

```

Test coverage includes:

- **Model Tests**: FileListModel with file management logic (9 tests)
- **Widget Tests**: DropZoneWidget and UI components (3 tests)
- **Integration Tests**: MainUI with signal emissions and state management (14 tests)
- **Worker Tests**: MergeWorker threading behavior (2 tests)
- **Utility Tests**: UI strings and configuration (1 test)

### Architecture

The GUI follows modern PySide6 best practices:

1. **Dependency Injection**: Backend logic (PowerPointMerger) is injected into the UI
2. **Model-View Pattern**: Uses QStandardItemModel for file list management
3. **Signal-Slot Architecture**: UI events emit signals that can be connected to any backend
4. **Threading**: Merge operations run in QThread to prevent UI freezing
5. **Settings Persistence**: Uses QSettings for cross-session configuration
6. **Resource Management**: Icons managed through Qt resource system (.qrc)

### Code Formatting

```bash

# Auto-format code
ruff check --fix src/ tests/

# Check formatting without changes
ruff check src/ tests/

```

### Linting

```bash

ruff check src/

```

## Python Version Compatibility

**Important**: This application is developed and tested using Python 3.8 to 3.12.

- The `pywin32` library supports Python 3.13
- `PySide6` does not yet have official pre-compiled wheels for Python 3.13

**Recommended**: Use Python 3.12 for guaranteed stability.

## Platform Requirements

- **Operating System**: Windows (COM automation requirement)
- **Python**: 3.8 or higher (3.12 recommended)
- **Microsoft PowerPoint**: Must be installed and licensed

## Documentation

- 🏗️ [**ARCHITECTURE.md**](docs/ARCHITECTURE.md) - Technical architecture and design patterns
- 📝 [**CHANGELOG.md**](docs/CHANGELOG.md) - Version history and release notes
- 🚀 [**PLANNED_FEATURE_ENHANCEMENTS.md**](docs/PLANNED_FEATURE_ENHANCEMENTS.md) -
  Planned features and roadmap
- 🤝 [**CONTRIBUTING.md**](docs/CONTRIBUTING.md) - How to contribute to the project
- 📜 [**CODE_OF_CONDUCT.md**](docs/CODE_OF_CONDUCT.md) - Community guidelines

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
