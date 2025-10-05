# PowerPoint Presentation Merger

A powerful Python GUI application for merging multiple PowerPoint (.pptx) files into a single presentation with perfect fidelity. Uses COM automation to ensure all formatting, animations, and embedded content are preserved during the merge process.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.6+](https://img.shields.io/badge/python-3.6+-blue.svg)](https://www.python.org/downloads/)
[![Windows](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

## Features

- **Step-by-Step GUI Workflow**: Intuitive 4-step process for merging presentations
- **Perfect Slide Copying**: COM automation preserves all formatting, animations, and embedded content
- **File Management**: Easy file selection with validation and reordering capabilities
- **Automatic Slideshow**: Launches merged presentation immediately after creation
- **Comprehensive Logging**: Optional live logging for debugging and troubleshooting
- **Error Handling**: Robust error management with clear user feedback
- **Dual Entry Points**: Choose between standard mode or debug mode with live logging

## System Requirements

### Operating System

- **Windows 7/8/10/11** (COM automation requires Windows)

### Software Dependencies

- **Python 3.6 or higher** (for source code execution)
- **Microsoft PowerPoint** (must be installed and licensed)

### Hardware Requirements

- **Memory**: 4GB RAM minimum (8GB recommended for large presentations)
- **Storage**: 100MB available space
- **Display**: 1024x768 minimum resolution

## Installation

### Option 1: Pre-built Executable (Recommended for End Users)

1. Download the latest release from the [Releases page](https://github.com/laashamar/MergePowerPointPresentations/releases)
2. Extract the archive to your desired location
3. Run `MergePowerPoint.exe` directly

### Option 2: From Source Code (For Developers)

1. **Clone the repository**:

   ```bash
   git clone https://github.com/laashamar/MergePowerPointPresentations.git
   cd MergePowerPointPresentations
   ```

2. **Install dependencies**:

   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:

   ```bash
   # Standard mode
   python main.py
   
   # Debug mode with live logging
   python run_with_logging.py
   ```

## Usage Guide

### Quick Start

1. **Launch the application** using one of the methods above
2. **Follow the 4-step workflow** described below
3. **Enjoy your merged presentation** with automatic slideshow launch

### Detailed Workflow

#### Step 1: Number of Files

- Enter the number of PowerPoint files you want to merge
- Must be a positive integer
- Press **Enter** or click **Next** to continue

#### Step 2: File Selection

- Click **"Add Files from Disk"** to browse for `.pptx` files
- Select exactly the number of files specified in Step 1
- Only PowerPoint (.pptx) files are accepted
- Duplicate selections are automatically prevented

#### Step 3: Output Filename

- Enter a name for your merged presentation
- The `.pptx` extension is added automatically if omitted
- Press **Enter** or click **Next** to continue

#### Step 4: File Ordering

- Use **"Move Up"** and **"Move Down"** buttons to arrange files
- Files will be merged in the order shown (top to bottom)
- Select a file in the list before using move buttons
- Click **"Create New File"** to start the merge process

#### Step 5: Merge and Launch

- The application automatically merges your presentations
- Progress is shown in real-time (if using debug mode)
- Merged presentation is saved to your specified location
- Slideshow launches automatically upon completion

## Application Modes

### Standard Mode (`main.py`)

- **Best for**: Regular usage and end users
- **Features**: Clean GUI workflow without logging overhead
- **Command**: `python main.py`

### Debug Mode (`run_with_logging.py`)

- **Best for**: Troubleshooting and development
- **Features**:
  - Live logging window with real-time status updates
  - Detailed error reporting and diagnostics
  - Log file saved to `Downloads/merge_powerpoint.log`
  - Threading for non-blocking operation
- **Command**: `python run_with_logging.py`

## Architecture Overview

The application follows a modular architecture with clear separation of concerns:

```text
Entry Points:
├── main.py                     # Standard entry point
└── run_with_logging.py        # Debug entry point with logging

Core Modules:
├── app.py                     # Application orchestration
├── gui.py                     # User interface components
├── powerpoint_core.py         # COM automation logic
└── logger.py                  # Logging infrastructure
```

### Key Components

- **`app.py`**: Central workflow coordination and state management
- **`gui.py`**: All GUI windows with validation and user interactions
- **`powerpoint_core.py`**: PowerPoint COM automation for merging and slideshow
- **`logger.py`**: Comprehensive logging system with multiple output targets
- **`main.py`**: Simple entry point for standard usage
- **`run_with_logging.py`**: Advanced entry point with live debugging

## Technical Details

### COM Automation

The application uses Microsoft's COM (Component Object Model) automation to:

- Launch PowerPoint application instances
- Create and manipulate presentation objects
- Copy slides with perfect fidelity
- Save merged presentations
- Launch slideshows programmatically

### Error Handling

- **Input Validation**: Real-time validation at each step
- **File Verification**: Existence and format checking
- **COM Exception Management**: Graceful handling of PowerPoint errors
- **User Feedback**: Clear error messages with actionable guidance
- **Resource Cleanup**: Proper disposal of COM objects

### Performance Characteristics

| Presentation Size | Typical Merge Time |
|------------------|-------------------|
| Small (1-10 slides) | < 5 seconds |
| Medium (11-50 slides) | 5-30 seconds |
| Large (50+ slides) | 30+ seconds |

*Performance depends on slide complexity, embedded media, and system specifications.*

## Troubleshooting

### Common Issues

#### "PowerPoint application not found"

- **Cause**: PowerPoint is not installed or not properly registered
- **Solution**: Install Microsoft PowerPoint and ensure it's licensed

#### "Permission denied" errors

- **Cause**: Insufficient file system permissions
- **Solution**: Run as administrator or check file/folder permissions

#### "File is locked" errors

- **Cause**: Source files are open in PowerPoint
- **Solution**: Close all PowerPoint windows before merging

#### Memory issues with large presentations

- **Cause**: Insufficient system memory
- **Solution**: Close other applications and try merging fewer files at once

### Getting Help

1. **Enable Debug Mode**: Use `run_with_logging.py` for detailed error information
2. **Check Log Files**: Review `merge_powerpoint.log` in your Downloads folder
3. **Verify System Requirements**: Ensure all dependencies are met
4. **Report Issues**: Create an issue on the [GitHub repository](https://github.com/laashamar/MergePowerPointPresentations/issues)

## Development

### Setting Up Development Environment

1. **Clone and install dependencies** (see Installation section)
2. **Run tests** (if available):

   ```bash
   python -m pytest
   ```

3. **Follow code style guidelines**:
   - PEP 8 for Python code formatting
   - Comprehensive logging for debugging
   - Modular architecture principles

### Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Dependencies

### Required Python Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `pywin32` | Latest | COM automation for PowerPoint |
| `tkinter` | Built-in | GUI framework |

### Standard Library Dependencies

- `logging` - Application logging
- `os` - File system operations
- `threading` - Background processing
- `sys` - System integration

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- **Microsoft PowerPoint** for COM automation capabilities
- **Python Community** for excellent libraries and documentation
- **Contributors** who have helped improve this project

## Changelog

### Version History

See [CHANGELOG.md](CHANGELOG.md) for detailed version history and release notes.

---

Made with ❤️ for PowerPoint users who need reliable presentation merging
