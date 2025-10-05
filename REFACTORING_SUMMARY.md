# Refactoring Summary

## What Was Changed

### 1. Dependencies
**BEFORE:**
- `python-pptx` (unreliable for merging)
- `tkinterdnd2` (unreliable drag-and-drop)

**AFTER:**
- `pywin32` (reliable COM automation)
- Standard `tkinter` only (no external GUI extensions)

### 2. Architecture
**BEFORE:**
- Single monolithic file (`merge_presentations.py`)
- All logic in one `PowerPointMerger` class

**AFTER:**
- Modular structure with clear separation of concerns:
  - `main.py` - Entry point
  - `app.py` - Application orchestration
  - `gui.py` - GUI components
  - `core.py` - Business logic

### 3. GUI Changes
**File Selection Window:**
- **Removed**: Drag-and-drop functionality
- **Kept**: "Add Files from Disk" button
- **Uses**: Standard `tk.Tk()` instead of `TkinterDnD.Tk()`

**File Reordering Window:**
- **Removed**: Drag-and-drop listbox reordering
- **Added**: "Move Up" and "Move Down" buttons
- **Uses**: Standard `tk.Listbox` with button controls

### 4. Merging Logic
**BEFORE (python-pptx):**
```python
# Created blank slides and attempted to copy shapes
merged_prs = Presentation()
for slide in source.slides:
    new_slide = merged_prs.slides.add_slide(layout)
    for shape in slide.shapes:
        copy_shape(shape, new_slide)  # Often failed
```

**AFTER (COM automation):**
```python
# Uses PowerPoint's native copy/paste
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
destination = powerpoint.Presentations.Add()
for file in files:
    source = powerpoint.Presentations.Open(file)
    for slide in source.Slides:
        slide.Copy()
        destination.Slides.Paste()  # Perfect fidelity
```

### 5. Slideshow Launch
**BEFORE:**
```python
subprocess.Popen(['powerpnt.exe', '/s', path])
```

**AFTER:**
```python
presentation.SlideShowSettings.Run()
```

## Benefits

1. **Perfect Slide Copying**: COM automation preserves all formatting, animations, and embedded content
2. **More Reliable**: No dependency on unreliable third-party libraries
3. **Better Maintainability**: Modular structure makes code easier to understand and modify
4. **PEP 8 Compliant**: All code follows Python style guidelines
5. **Clearer Architecture**: Each module has a single, well-defined responsibility

## Module Responsibilities

| Module | Responsibility |
|--------|---------------|
| `main.py` | Application entry point |
| `app.py` | State management and workflow orchestration |
| `gui.py` | All GUI windows and user interactions |
| `core.py` | PowerPoint operations using COM automation |

## How to Use

```bash
# Install dependencies
pip install -r requirements.txt

# Run the application
python main.py
```

The application guides users through:
1. Enter number of files
2. Select files via file dialog
3. Enter output filename
4. Reorder files with Move Up/Down buttons
5. Merge and launch slideshow
