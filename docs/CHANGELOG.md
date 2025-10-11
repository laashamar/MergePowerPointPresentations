# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed - Major Refactoring

* **Modern PySide6 GUI**
    * Migrated from tkinter to PySide6 (Qt for Python) for a modern, professional interface
    * Single-window application with intuitive file management
    * File list with selection and reordering capabilities
    * Built-in progress bar for visual feedback during merge operations
    * Responsive UI that remains interactive during merge operations

* **Refactored PySide6 GUI** (NEW - `gui_refactored.py`)
    * Modern two-column layout with 3:1 stretch ratio
    * Smart UI state management with empty state (drop zone) and active state (file list)
    * Full drag-and-drop support for .pptx files
    * Custom QStandardItemModel for scalable file list management
    * QListView with FileItemDelegate for card-style items
    * Signal-based architecture following Qt best practices
    * QThread-based MergeWorker for truly non-blocking operations
    * QSettings integration for persistent configuration (remembers last save location)
    * Comprehensive i18n support via UI_STRINGS dictionary
    * Full keyboard navigation and accessibility features
    * Dependency injection pattern for improved testability

* **Modern Package Structure**
    * Reorganized codebase using src layout pattern (PEP 518/621)
    * All code moved to `src/merge_powerpoint/` package
    * Created compatibility shims in root for backward compatibility
    * Added proper package initialization and exports
    * Implemented CLI entry point: `merge-powerpoint` command

* **Enhanced Code Quality**
    * All code formatted with Black (100 char line length)
    * Comprehensive docstrings following PEP 257 and Google style
    * Type hints ready structure
    * Zero linting violations with Ruff

### Added

* **Qt Resources System**
    * SVG icons for modern, scalable UI elements (plus, trash, close, powerpoint, folder)
    * Compiled .qrc resources for efficient loading and high-DPI support
    * Icon resource file at `src/merge_powerpoint/icons_rc.py`

* **Comprehensive Testing**
    * 29 new pytest-qt tests specifically for refactored GUI
    * Model tests for FileListModel (9 tests)
    * Widget tests for DropZoneWidget (3 tests)
    * Integration tests for MainUI (14 tests)
    * Worker tests for MergeWorker (2 tests)
    * UI strings validation (1 test)
    * Total test suite: 51 tests, 100% pass rate

* **Documentation**
    * New GUI_GUIDE.md with complete API reference and usage examples
    * Updated README with refactored GUI features and architecture section
    * Standalone example script (`examples/example_refactored_gui.py`)
    * Examples directory with comprehensive documentation
    * Migration guide for transitioning to refactored GUI

* **Progress Tracking**
    * Visual progress bar shows merge progress in real-time
    * Progress callback system for detailed operation tracking
    * Non-blocking merge operations with worker threads

* **Modern Development Tools**
    * Added pytest for testing with pytest-qt for GUI tests
    * Added Black for code formatting
    * Added Ruff for fast linting
    * Added mypy for type checking
    * Development installation mode with `pip install -e ".[dev]"`

### Technical Architecture

* **Design Patterns**
    * Dependency Injection: PowerPointMerger injected into MainUI
    * Model-View: QStandardItemModel with QListView
    * Signal-Slot: Type-safe Qt signal connections
    * Worker Thread: QThread for background operations
    * Settings Persistence: QSettings for cross-session state

### Removed

* Removed tkinter dependencies
* Removed tkinterdnd2 dependency
* Removed modal window workflow in favor of single-window design
* Removed legacy GUI implementation

## [1.0.0] - 2025-10-06

### Added

* **Initial Release**
    * Intuitive 4-step GUI for merging PowerPoint presentations.
    * Uses COM automation for perfect fidelity copying of slides, including animations and formatting.
    * Features file reordering, automatic slideshow launch, and robust error handling.
    * Includes standard and debug (`run_with_logging.py`) entry points with live logging.
    * Comprehensive documentation for users and developers.