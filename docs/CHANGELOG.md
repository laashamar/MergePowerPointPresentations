# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added - Phase 3 Features

* **Drag-and-Drop File Addition**
  * Users can now drag .pptx files directly onto the application window to add them to the merge list
    * Only valid .pptx files are accepted; other file types are silently ignored
    * Integrated tkinterdnd2 library for cross-platform drag-and-drop support

* **Drag-and-Drop List Reordering**
    * File order in the list can be changed by clicking and dragging file labels
    * Real-time visual feedback with numbered file list
    * Changes to order are immediately reflected in the internal file list

* **Dynamic Status Feedback During Merge**
    * Merge process now runs in a separate thread to keep the GUI responsive
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