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

* **Modern Package Structure**
    * Reorganized codebase using src layout pattern (PEP 518/621)
    * All code moved to `src/merge_powerpoint/` package
    * Created compatibility shims in root for backward compatibility
    * Added proper package initialization and exports
    * Implemented CLI entry point: `merge-powerpoint` command

* **Enhanced Code Quality**
    * All code formatted with Black (100 char line length)
    * Comprehensive docstrings following PEP 257
    * Type hints ready structure
    * Zero linting violations with Ruff

### Added

* **Progress Tracking**
    * Visual progress bar shows merge progress in real-time
    * Progress callback system for detailed operation tracking
    * Non-blocking merge operations

* **Modern Development Tools**
    * Added pytest for testing with pytest-qt for GUI tests
    * Added Black for code formatting
    * Added Ruff for fast linting
    * Added mypy for type checking
    * Development installation mode with `pip install -e ".[dev]"`

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