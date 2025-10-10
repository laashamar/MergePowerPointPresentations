# Architecture Documentation

## PowerPoint Presentation Merger - System Architecture

### Overview

The PowerPoint Presentation Merger is a Python desktop application built with a **modern package structure** following the **src layout** pattern. It provides an intuitive GUI workflow for merging multiple PowerPoint presentations using COM automation for reliable slide copying, with comprehensive logging capabilities.

### Design Principles

- **Modern Package Structure**: Follows PEP 518/621 with src layout
- **Separation of Concerns**: Each module has a single, well-defined responsibility
- **Modular Structure**: Clear boundaries between GUI, business logic, and infrastructure
- **COM Integration**: Native PowerPoint automation for perfect fidelity
- **Comprehensive Logging**: Full observability for debugging and monitoring
- **User-Centric Design**: Intuitive GUI with clear validation
- **Backward Compatibility**: Compatibility shims for existing code

## Module Architecture

### Package Structure (src layout)

```text
src/merge_powerpoint/           # Main package
├── __init__.py                 # Package initialization and exports
├── __main__.py                 # CLI entry point (python -m merge_powerpoint)
├── app.py                      # Application controller
├── app_logger.py               # Logging configuration
├── gui.py                      # GUI components (PySide6)
└── powerpoint_core.py          # PowerPoint COM automation

Root Level (Compatibility):
├── main.py                     # Standard entry point (uses src package)
├── run_with_logging.py         # Entry point with logging (uses src package)
├── app.py                      # Compatibility shim → src/merge_powerpoint/app.py
├── app_logger.py               # Compatibility shim → src/merge_powerpoint/app_logger.py
├── gui.py                      # Compatibility shim → src/merge_powerpoint/gui.py
└── powerpoint_core.py          # Compatibility shim → src/merge_powerpoint/powerpoint_core.py

Configuration:
├── pyproject.toml              # Modern Python project configuration (PEP 518/621)
├── pytest.ini                  # pytest configuration
├── .flake8                     # Flake8 linting configuration
└── .pylintrc                   # Pylint configuration
```

### Module Dependencies

```text
main.py
└── merge_powerpoint package
    ├── app.py (AppController)
    │   └── powerpoint_core.py (PowerPointMerger)
    ├── gui.py (MainWindow)
    │   ├── PySide6.QtWidgets
    │   └── powerpoint_core.py (PowerPointMerger)
    └── app_logger.py (setup_logging)

CLI: merge-powerpoint command
└── merge_powerpoint.__main__.main()
    └── Same structure as main.py

run_with_logging.py
├── merge_powerpoint.app_logger (setup_logging)
└── main.main()
```

## Detailed Module Specifications

### 1. Entry Points

#### `merge-powerpoint` CLI Command

- **Purpose**: Primary CLI entry point (installed via pip)
- **Implementation**: Defined in `pyproject.toml` as console script
- **Module**: `merge_powerpoint.__main__:main`
- **Usage**: `merge-powerpoint` (after installation)
- **Features**:
  - Simple command-line invocation
  - No need to specify Python explicitly
  - Works from any directory after installation

#### `python -m merge_powerpoint`

- **Purpose**: Module execution entry point
- **Module**: `src/merge_powerpoint/__main__.py`
- **Usage**: `python -m merge_powerpoint`
- **Features**:
  - Works without installation (with PYTHONPATH set)
  - Direct module execution

#### `main.py`

- **Purpose**: Traditional script entry point
- **Responsibilities**:
  - Import from refactored package
  - Launch the main application
  - Provides backward compatibility
- **Usage**: `python main.py`
- **Implementation**: Wrapper that imports from `merge_powerpoint` package

#### `run_with_logging.py`

- **Purpose**: Entry point with exception logging
- **Responsibilities**:
  - Configure logging
  - Wrap main() with exception handling
  - Log critical errors
- **Usage**: `python run_with_logging.py`
- **Features**:
  - Enhanced error reporting
  - Catches unhandled exceptions

### 2. Core Package (`merge_powerpoint`)

#### `__init__.py` - Package Initialization

- **Purpose**: Define package exports and version
- **Exports**:
  - `AppController`
  - `PowerPointMerger`
  - `PowerPointError`
  - `__version__`

#### `app.py` - Application Controller

- **Purpose**: High-level application controller
- **Key Class**: `AppController`
- **Base Class**: Inherits from `PowerPointMerger`
- **Responsibilities**:
  - Provide application-specific functionality
  - Can be extended without modifying core logic
- **Usage**: Used by GUI to manage merge operations

#### `app_logger.py` - Logging Configuration

- **Purpose**: Centralized logging setup
- **Key Function**: `setup_logging()`
- **Features**:
  - Creates logs directory automatically
  - Configures file and console handlers
  - INFO level and above
  - Structured log format
- **Returns**: Configured root logger

### 3. Presentation Layer

#### `gui.py` - User Interface

- **Purpose**: PySide6-based graphical user interface
- **Framework**: PySide6 (Qt for Python)
- **Key Class**: `MainWindow`
- **Features**:
  - File list management
  - Add/Remove/Clear operations
  - File reordering (Move Up/Down)
  - Merge with progress tracking
  - Input validation
- **Components**:
  - QListWidget for file display
  - QPushButtons for actions
  - QProgressBar for merge progress
  - QFileDialog for file selection
  - File dialog integration
  - Move Up/Down file reordering
  - Keyboard shortcuts (Enter key support)
  - Visual feedback and styling

### 4. Business Logic Layer

#### `powerpoint_core.py` - PowerPoint Operations

- **Purpose**: Core PowerPoint merging functionality
- **Key Classes**:

  **PowerPointError**
  - Custom exception for PowerPoint-related errors
  - Used throughout the module for error handling

  **PowerPointCore**
  - Low-level COM automation for PowerPoint
  - Handles PowerPoint instance management
  - Methods:
    - `__init__()`: Initialize COM automation
    - `merge_presentations()`: Merge files using COM
  - Platform: Windows only (COM requirement)
  - COM Operations:
    - PowerPoint application connection/creation
    - Presentation manipulation
    - Slide insertion from files
    - File saving and cleanup

  **PowerPointMerger**
  - High-level file management and merging
  - Methods:
    - `add_files()`: Add files to merge list
    - `remove_file()`, `remove_files()`: Remove files
    - `clear_files()`: Clear all files
    - `move_file_up()`, `move_file_down()`: Reorder files
    - `get_files()`: Get current file list
    - `merge()`: High-level merge with progress callback
  - Used by GUI and application controller

## Application Workflow

### GUI Workflow

```text
1. Application Startup
   ├── Launch via CLI (merge-powerpoint) or script (python main.py)
   ├── Initialize logging
   ├── Create QApplication
   └── Show MainWindow

2. File Management
   ├── User clicks "Add Files"
   ├── QFileDialog shows file selection
   ├── Files added to PowerPointMerger
   └── GUI updates file list

3. File Reordering (Optional)
   ├── User selects file in list
   ├── Clicks "Move Up" or "Move Down"
   ├── PowerPointMerger reorders files
   └── GUI refreshes display

4. Merge Operation
   ├── User clicks "Merge Files"
   ├── Validation: At least 2 files required
   ├── QFileDialog for output path
   ├── PowerPointMerger.merge() called
   ├── Progress bar updates via callback
   └── Success/Error message shown

5. Application Exit
   └── User closes main window

```

### Package Installation Workflow

```text
1. Installation
   ├── pip install . (or pip install -e . for development)
   ├── setuptools builds package from pyproject.toml
   ├── Dependencies installed (PySide6, pywin32, comtypes)
   └── CLI entry point registered: merge-powerpoint

2. CLI Execution
   ├── User runs: merge-powerpoint
   ├── Python executes: merge_powerpoint.__main__:main()
   ├── Application launches
   └── GUI appears

3. Module Execution
   ├── User runs: python -m merge_powerpoint
   ├── Python executes: src/merge_powerpoint/__main__.py
   └── Same as CLI execution
```

## Technical Implementation Details

### COM Automation Architecture

```python
# PowerPoint COM Integration Pattern (using comtypes)
import comtypes.client

# Initialize COM
comtypes.CoInitialize()
powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
powerpoint.Visible = True

# Create destination presentation
base_presentation = powerpoint.Presentations.Add()

# Insert slides from each file
for file_path in file_paths:
    abs_path = os.path.abspath(file_path)
    slide_count = base_presentation.Slides.Count
    # InsertFromFile inserts slides after the specified index
    base_presentation.Slides.InsertFromFile(abs_path, slide_count)

# Save the merged presentation
base_presentation.SaveAs(output_path)
base_presentation.Close()

# Cleanup
comtypes.CoUninitialize()
```

### Logging Architecture

```text
Root Logger (configured by app_logger.setup_logging())
├── FileHandler → logs/app.log (persistent storage)
└── StreamHandler → Console (real-time display)
```

### Error Handling Strategy

- **Input Validation**: File existence and type checking
- **COM Exception Handling**: Platform-specific error handling
- **Resource Management**: Proper COM cleanup with __del__
- **User Feedback**: Clear error messages via QMessageBox
- **Logging Integration**: All errors logged with stack traces
- **Custom Exceptions**: PowerPointError for domain-specific errors

## External Dependencies

### Required Dependencies

| Package | Version | Purpose | Usage |
|---------|---------|---------|-------|
| `PySide6` | >=6.7 | GUI framework | All user interface components |
| `pywin32` | >=311 | COM automation (optional) | Alternative to comtypes |
| `comtypes` | >=1.2.0 | COM automation | PowerPoint integration |

### Optional Dependencies (Development)

| Package | Version | Purpose |
|---------|---------|---------|
| `pytest` | >=8.0.0 | Testing framework |
| `pytest-qt` | >=4.2.0 | Qt/PySide6 testing support |
| `pytest-cov` | >=4.1.0 | Code coverage |
| `pytest-mock` | >=3.12.0 | Mocking support |
| `black` | >=24.0.0 | Code formatting |
| `ruff` | >=0.1.0 | Fast linting |
| `flake8` | >=7.0.0 | Legacy linting |
| `pylint` | >=3.0.0 | Static analysis |
| `mypy` | >=1.8.0 | Type checking |

### Standard Library Dependencies

| Module | Purpose |
|--------|---------|
| `logging` | Application logging |
| `os` | File system operations |
| `sys` | System integration and path manipulation |
| `pathlib` | Modern path handling |

## Platform Requirements

### System Requirements

- **Operating System**: Windows (COM automation requirement)
- **Python Version**: 3.8 or higher (3.12 recommended)
- **Microsoft PowerPoint**: Must be installed and licensed
- **Memory**: Minimal (GUI-based application)
- **Storage**: Minimal footprint

### Architecture Limitations

1. **Windows-Only**: COM automation restricts to Windows platform
2. **PowerPoint Dependency**: Requires installed PowerPoint application
3. **Single-Threading**: GUI operations are single-threaded
4. **Memory Usage**: Large presentations may require significant memory
5. **No Undo**: Merge operations cannot be reversed

## Design Decisions and Rationale

### 1. COM Automation vs. Library-Based Approach

**Decision**: Use COM automation instead of `python-pptx`

**Rationale**:

- Perfect slide fidelity preservation
- Native PowerPoint operations
- Animation and embedded content support
- Reliable slide copying mechanism

### 2. Modal GUI Workflow

**Decision**: Sequential modal windows instead of single complex interface

**Rationale**:

- Clear step-by-step progression
- Simplified state management
- Reduced user cognitive load
- Easy validation at each step

### 3. Move Up/Down vs. Drag-and-Drop

**Decision**: Button-based reordering instead of drag-and-drop

**Rationale**:

- Eliminated unreliable `tkinterdnd2` dependency
- More accessible interface
- Simpler implementation and maintenance
- Better cross-platform compatibility

### 4. Dual Entry Points

**Decision**: Separate entry points for standard and logging modes

**Rationale**:

- Clean separation of concerns
- Optional debugging capabilities
- User choice in interface complexity
- Development vs. production usage patterns

## Future Enhancement Opportunities

### Short-term Improvements

- Progress indicators during merge operations
- Slide preview capabilities
- Batch processing support
- Custom slide selection (partial merges)

### Long-term Enhancements

- Cross-platform support (alternative to COM)
- Undo/redo functionality
- Template and theme preservation options
- Network file support
- Plugin architecture for extensibility

## Security Considerations

### COM Security

- PowerPoint COM objects have full application access
- File system access through PowerPoint application
- Potential for macro execution in source files

### Mitigation Strategies

- Read-only source file access
- Input validation and sanitization
- Error boundary isolation
- User permission requirements

## Performance Characteristics

### Typical Performance

- **Small presentations** (1-10 slides): < 5 seconds
- **Medium presentations** (11-50 slides): 5-30 seconds
- **Large presentations** (50+ slides): 30+ seconds

### Performance Factors

- Number and complexity of slides
- Embedded media and animations
- Available system memory
- PowerPoint application responsiveness

## Troubleshooting Guide

### Common Issues

1. **COM Errors**: Ensure PowerPoint is properly installed and licensed
2. **Permission Errors**: Run with appropriate file system permissions
3. **Memory Issues**: Close other applications for large merges
4. **File Lock Errors**: Ensure source files are not open in PowerPoint

### Diagnostic Tools

- Live logging window (`run_with_logging.py`)
- Error summary generation
- Detailed exception logging
- File validation reporting

## Modern Package Structure Benefits

### Why src Layout?

The refactored codebase uses the **src layout** pattern, which provides several advantages:

1. **Import Isolation**: Prevents accidentally importing from the source directory during development
2. **Clear Separation**: Distinguishes source code from tests, docs, and configuration
3. **Installation Testing**: Forces testing against installed package, not source directory
4. **Best Practice**: Follows modern Python packaging standards (PEP 518, PEP 621)

### Backward Compatibility

The refactoring maintains backward compatibility through:

- **Compatibility Shims**: Root-level modules import from `src/merge_powerpoint/`
- **Test Compatibility**: Existing tests work without modification
- **Legacy Entry Points**: `main.py` and `run_with_logging.py` continue to work
- **Gradual Migration**: Old import patterns continue to function

### Code Quality Improvements

- **Black Formatting**: All code formatted to PEP 8 standards (100 char line length)
- **Ruff Linting**: Fast, comprehensive linting with zero violations
- **Comprehensive Docstrings**: PEP 257 compliant documentation for all modules/classes/functions
- **Type Hints Ready**: Structure supports future type annotation addition

### Installation Methods

**Development Installation:**
```bash
pip install -e .
```

**Production Installation:**
```bash
pip install .
```

**Development with Tools:**
```bash
pip install -e ".[dev]"
```

### CLI Access

After installation, the package provides multiple entry points:

```bash
# Primary CLI command (recommended)
merge-powerpoint

# Module execution
python -m merge_powerpoint

# Legacy script execution
python main.py
python run_with_logging.py
```
