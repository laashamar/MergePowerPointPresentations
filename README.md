# PowerPoint Presentation Merger

A modern Python utility to merge multiple PowerPoint (.pptx) files into a single presentation. This tool uses COM automation to ensure that all formatting, animations, and embedded content are preserved with perfect fidelity during the merge process.

## Features

- **Merge multiple .pptx files** in a specified order
- **Preserves original formatting**, transitions, and animations
- **Simple graphical user interface (GUI)** to manage files
- **Reorder files** before merging
- **Responsive UI** that does not freeze during the merge process
- **Modern package structure** with `src` layout
- **Command-line interface** for easy execution

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
- **pytest**: Testing framework
- **mypy**: Static type checking

### Running Tests

```bash
pytest tests/
```

### Code Formatting

```bash
black src/
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
- 🚀 [**PLANNED_FEATURE_ENHANCEMENTS.md**](docs/PLANNED_FEATURE_ENHANCEMENTS.md) - Planned features and roadmap
- 🤝 [**CONTRIBUTING.md**](docs/CONTRIBUTING.md) - How to contribute to the project
- 📜 [**CODE_OF_CONDUCT.md**](docs/CODE_OF_CONDUCT.md) - Community guidelines

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
