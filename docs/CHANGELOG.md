# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added - Version 2.0 Features

* **Complete GUI Refactor with PySide6**
  * Migrated from tkinter to PySide6 for modern, cross-platform GUI
  * Enhanced user interface with better responsiveness and styling
  * Improved drag-and-drop functionality using Qt's native DnD system
  * Better file management with visual feedback and progress indicators

* **Modern Python Packaging**
  * Implemented pyproject.toml for standardized project configuration
  * Added proper package structure with entry points
  * Installable package with `pip install -e .` for development
  * Centralized configuration for all development tools

* **Streamlined Development Tools**
  * Replaced multiple linting tools (black, flake8, pylint, isort) with ruff
  * Unified code formatting and linting in single fast tool
  * Updated CI/CD pipeline for modern packaging standards
  * Improved test configuration and coverage reporting

* **Enhanced Documentation**
  * Updated README with modern installation instructions
  * Added comprehensive development guidelines
  * Created planned feature enhancements documentation
  * Improved code documentation and type hints

### Changed

* **GUI Framework**: Complete migration from tkinter to PySide6
* **Project Structure**: Modernized to follow current Python packaging standards
* **Dependencies**: Streamlined to essential packages only
* **Build System**: Uses pyproject.toml instead of setup.py and requirements.txt
* **Code Quality**: Unified linting and formatting with ruff

### Removed

* **Legacy Dependencies**: Removed tkinter-specific packages (tkinterdnd2, customtkinter)
* **Redundant Tools**: Eliminated multiple linting tools in favor of ruff
* **Old Configuration**: Removed setup.py, requirements.txt, pytest.ini in favor of pyproject.toml

### Technical

* Updated to PySide6>=6.7 for modern Qt GUI framework
* Implemented ruff>=0.1.0 for fast Python linting and formatting
* Enhanced test suite with pytest-qt for GUI testing
* Improved CI/CD with Windows-focused testing pipeline

## [1.1.0] - 2025-10-06

### Added - Phase 3 Features (Legacy tkinter version)

* **Drag-and-Drop File Addition**
  * Users can drag .pptx files directly onto the application window
  * Only valid .pptx files are accepted; other file types are silently ignored
  * Integrated tkinterdnd2 library for cross-platform drag-and-drop support

* **Drag-and-Drop List Reordering**
  * File order in the list can be changed by clicking and dragging file labels
  * Real-time visual feedback with numbered file list
  * Changes to order are immediately reflected in the internal file list

* **Dynamic Status Feedback During Merge**
  * Merge process runs in separate thread to keep GUI responsive
  * Real-time progress updates showing current file and slide being processed
  * Status messages: "Merging [filename] (slide X of Y)...", "Merge Complete!", or error details
  * Thread-safe GUI updates using self.after() method

* **Post-Merge Actions**
  * Two new buttons appear after successful merge:
    - "Open Presentation": Opens the merged file in the default application
    - "Show in Explorer": Opens file explorer and highlights the merged file
  * Cross-platform support for Windows, macOS, and Linux
  * Buttons are hidden by default and only shown after successful merge

### Changed

* Updated `powerpoint_core.merge_presentations()` to accept optional progress callback
* Enhanced logging throughout the merge process for better debugging
* All new code follows PEP8 standards with proper docstrings (PEP257)

### Technical

* Added tkinterdnd2>=0.3.0 to requirements.txt
* Implemented threading module for non-blocking merge operations
* Added subprocess module for cross-platform file operations

## [1.0.0] - 2025-10-06

### Added

* **Initial Release**
    * Intuitive 4-step GUI for merging PowerPoint presentations.
    * Uses COM automation for perfect fidelity copying of slides, including animations and formatting.
    * Features file reordering, automatic slideshow launch, and robust error handling.
    * Includes standard and debug (`run_with_logging.py`) entry points with live logging.
    * Comprehensive documentation for users and developers.