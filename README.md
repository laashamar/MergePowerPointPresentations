# PowerPoint Presentation Merger

A Python GUI application for merging multiple PowerPoint (`.pptx`) files into a single presentation using COM automation for reliable and accurate merging.

## Features

- User-friendly GUI with step-by-step workflow
- File dialog for easy file selection
- Move Up/Down buttons for reordering presentations before merging
- COM automation for perfect slide copying (preserves all formatting, animations, and content)
- Automatic slideshow launch after merging
- Comprehensive error handling

## Requirements

- Python 3.6 or higher
- Windows OS (required for COM automation)
- Microsoft PowerPoint installed on the system

## Installation

1. Clone this repository:
```bash
git clone https://github.com/laashamar/MergePowerPointPresentations.git
cd MergePowerPointPresentations
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the application:
```bash
python main.py
```

The application will guide you through a 4-step process:

### Step 1: Number of Files
Enter the number of PowerPoint files you want to merge.

### Step 2: Select Files
Click "Add Files from Disk" to browse and select the `.pptx` files you want to merge.

### Step 3: New Filename
Enter a name for the merged presentation (`.pptx` extension is added automatically).

### Step 4: Set Merge Order
Use the "Move Up" and "Move Down" buttons to reorder the files. The files will be merged in the order shown.

Click "Create New File" to merge the presentations and launch the slideshow.

## Architecture

The application is structured into modular components:

- **main.py**: Entry point for the application
- **app.py**: Application orchestration and state management
- **gui.py**: All GUI windows and components
- **core.py**: PowerPoint merging logic using COM automation

## Notes

- The application uses COM automation to ensure perfect copying of all slide content, formatting, and animations
- The application automatically opens the merged presentation in slideshow mode
- All slides from each presentation are copied in order
- The GUI closes automatically after launching the slideshow

## License

MIT License - See LICENSE file for details
