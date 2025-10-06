# Changelog

All notable changes to the PowerPoint Presentation Merger project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Updated

- All Norwegian text translated to English throughout the application
- Updated ARCHITECTURE.md with current module structure and comprehensive documentation
- Rewritten README.md with professional formatting and complete feature documentation
- Improved Markdown formatting across all documentation files

## [2.1.0] - 2025-10-05

### Added in 2.1.0

- **Live Logging System**: Comprehensive logging infrastructure with GUI display
- **Dual Entry Points**:
  - `main.py` for standard usage
  - `run_with_logging.py` for debugging with live log window
- **Error Collection**: Automatic error summarization and reporting
- **Log File Output**: Persistent logging to `merge_powerpoint.log` in Downloads folder
- **Threading Support**: Non-blocking GUI operations during merge process
- **Enhanced Error Handling**: Improved exception management with detailed logging

### Changed in 2.1.0

- **Module Renaming**: `core.py` renamed to `powerpoint_core.py` for better clarity
- **Logging Integration**: All modules now include comprehensive logging
- **User Messages**: Enhanced error messages and user feedback
- **GUI Improvements**: Better visual feedback and keyboard shortcuts
- **Architecture**: Modular structure with clear separation of concerns

### Fixed in 2.1.0

- **Resource Management**: Improved COM object cleanup and disposal
- **Error Recovery**: Better handling of PowerPoint application errors
- **Input Validation**: Enhanced file validation and user input checking
- **Memory Management**: Optimized memory usage during merge operations

## [2.0.0] - 2025-10-05

### Added

- **Modular Architecture**: Complete restructuring into separate modules
  - `app.py`: Application orchestration and state management
  - `gui.py`: All GUI components and user interactions
  - `powerpoint_core.py`: PowerPoint COM automation logic
  - `main.py`: Application entry point
- **COM Automation**: Native PowerPoint integration for perfect slide copying
- **Move Up/Down Controls**: Button-based file reordering system
- **Input Validation**: Comprehensive validation at each workflow step
- **Error Handling**: Robust error management with user-friendly messages
- **PEP 8 Compliance**: Code formatted according to Python style guidelines

### Changed

- **Dependency Management**: Replaced unreliable libraries
  - Removed `python-pptx` (unreliable merging)
  - Removed `tkinterdnd2` (unreliable drag-and-drop)
  - Added `pywin32` for COM automation
- **GUI Framework**: Simplified to standard `tkinter` widgets
- **File Selection**: Replaced drag-and-drop with file dialog
- **Merge Algorithm**: Complete rewrite using COM automation
- **Slideshow Launch**: Native PowerPoint slideshow instead of subprocess

### Removed

- **Drag-and-Drop**: Removed unreliable drag-and-drop functionality
- **External GUI Dependencies**: Eliminated `tkinterdnd2` dependency
- **Python-pptx**: Removed library-based merging approach

### Fixed

- **Slide Copying**: Perfect fidelity preservation of all content
- **Animation Support**: Proper copying of animations and transitions
- **Embedded Content**: Reliable handling of embedded media and objects
- **Error Recovery**: Improved handling of merge failures

## [1.0.0] - 2025-10-04

### Initial Release

- **Basic PowerPoint Merger**: Core functionality for merging presentations
- **GUI Interface**: Simple tkinter-based user interface
- **File Selection**: Basic file selection functionality
- **Drag-and-Drop**: Initial drag-and-drop support (later removed)
- **Python-pptx Integration**: Library-based merging (later replaced)
- **Basic Error Handling**: Simple error management
- **Documentation**: Initial project documentation

### Features

- Merge multiple PowerPoint presentations
- Basic GUI workflow
- Simple file selection
- Slideshow launch capability

## [0.1.0] - 2025-10-04

### Project Setup

- **Repository Initialization**: Project setup and initial commit
- **Basic Structure**: Initial project structure and planning
- **Documentation Setup**: Basic README and project files
- **Development Environment**: Initial development setup

---

## Version History Summary

| Version | Release Date | Key Features |
|---------|--------------|--------------|
| **2.1.0** | 2025-10-05 | Live logging, dual entry points, enhanced error handling |
| **2.0.0** | 2025-10-05 | Modular architecture, COM automation, reliable merging |
| **1.0.0** | 2025-10-04 | Initial GUI application with basic merging |
| **0.1.0** | 2025-10-04 | Project initialization and setup |

## Migration Guide

### From v1.x to v2.x

- **Dependencies**: Install `pywin32` and remove `python-pptx`, `tkinterdnd2`
- **Usage**: No changes to user workflow, but improved reliability
- **Requirements**: Ensure Microsoft PowerPoint is installed and licensed

### From v2.0 to v2.1

- **New Features**: Optional live logging mode available via `run_with_logging.py`
- **Logging**: Log files now saved to Downloads folder
- **Compatibility**: Fully backward compatible with v2.0 usage patterns

## Development Notes

### Build Requirements

- **Python**: 3.6 or higher
- **Platform**: Windows (COM automation requirement)
- **Dependencies**: See `requirements.txt` for complete list

### Testing

- Manual testing on Windows 10/11
- PowerPoint 2016, 2019, and Microsoft 365 compatibility
- Various presentation sizes and complexity levels

### Known Issues

- **Platform Limitation**: Windows-only due to COM automation
- **PowerPoint Dependency**: Requires installed PowerPoint application
- **Memory Usage**: Large presentations may require significant memory

### Future Roadmap

- **Cross-Platform Support**: Investigate alternatives to COM automation
- **Progress Indicators**: Visual progress during merge operations
- **Batch Processing**: Multiple merge operations in sequence
- **Slide Selection**: Choose specific slides to merge
- **Template Preservation**: Better handling of presentation templates

---

*For detailed technical information, see [ARCHITECTURE.md](ARCHITECTURE.md)*

*For usage instructions, see [README.md](README.md)*
