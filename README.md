# PowerPoint Presentation Merger

A Python GUI application for merging multiple PowerPoint (`.pptx`) files into a single presentation.

## Features

- User-friendly GUI with step-by-step workflow
- Drag-and-drop support for file selection
- File dialog for traditional file selection
- Drag-and-drop reordering of presentations before merging
- Automatic slideshow launch after merging
- Comprehensive error handling

## Requirements

- Python 3.6 or higher
- Windows OS (for PowerPoint slideshow launch)
- PowerPoint installed on the system

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
python merge_presentations.py
```

The application will guide you through a 4-step process:

### Step 1: Number of Files
Enter the number of PowerPoint files you want to merge.

### Step 2: Select Files
Add files using either:
- **Drag and Drop**: Drag `.pptx` files from File Explorer into the listbox
- **File Dialog**: Click "Add Files from Disk" to browse and select files

### Step 3: New Filename
Enter a name for the merged presentation (`.pptx` extension is added automatically).

### Step 4: Set Merge Order
Reorder the files by dragging them up or down in the list. The files will be merged in the order shown.

Click "Create New File" to merge the presentations and launch the slideshow.

## Notes

- The application automatically opens the merged presentation in slideshow mode
- All slides from each presentation are copied in order
- The GUI closes automatically after launching the slideshow

## License

MIT License - See LICENSE file for details
