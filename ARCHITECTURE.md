# ARCHITECTURE.md

## PowerPoint Presentation Merger - Architecture Overview

### Design Pattern
The application uses a **modular architecture** with clear separation of concerns:
- **GUI Layer** (`gui.py`): Handles all user interface components
- **Application Layer** (`app.py`): Manages application state and workflow orchestration
- **Core Layer** (`core.py`): Implements PowerPoint merging logic using COM automation
- **Entry Point** (`main.py`): Application startup

### Module Structure

```
main.py                      # Entry point
├── app.py                   # Application orchestration
    ├── gui.py               # GUI windows
    │   ├── show_number_of_files_window()
    │   ├── show_file_selection_window()
    │   ├── show_filename_window()
    │   └── show_reorder_window()
    └── core.py              # Backend logic
        ├── merge_presentations()
        └── launch_slideshow()
```

### State Management
The `PowerPointMergerApp` class in `app.py` maintains four key state variables:
- `num_files`: Expected number of files to merge
- `selected_files`: List of file paths selected by user
- `output_filename`: Name for the merged presentation
- `file_order`: Final order of files after user reordering

### GUI Framework
- **Primary**: `tkinter` for all GUI components
- **File Selection**: Standard `tkinter.filedialog` for file selection
- **Reordering**: Move Up/Down buttons using standard `tkinter` widgets

### Workflow Progression
The application follows a strict sequential flow:
1. User inputs number of files → stored in `num_files`
2. User selects files → stored in `selected_files`
3. User enters filename → stored in `output_filename`
4. User reorders files → updates `file_order`
5. System merges and launches → uses all stored state

Each window is **modal** - the next window only opens after the previous one completes successfully and the user clicks the corresponding action button.

### Merge Algorithm

```python
# Using COM automation with win32com.client
PowerPoint = Dispatch("PowerPoint.Application")
destination_prs = PowerPoint.Presentations.Add()

for each source_file in file_order:
    source_prs = PowerPoint.Presentations.Open(source_file)
    for each slide in source_prs.Slides:
        slide.Copy()
        destination_prs.Slides.Paste()
    source_prs.Close()

destination_prs.SaveAs(output_path)
destination_prs.SlideShowSettings.Run()
```

### Error Handling Strategy
- **Validation**: Input validation at each step before proceeding
- **User Feedback**: Clear error messages via `messagebox.showerror()`
- **Graceful Degradation**: If PowerPoint launch fails, presentation is still saved
- **Try-Except Blocks**: Wrap file operations and merge logic
- **Resource Cleanup**: Properly close COM objects to prevent resource leaks

### External Dependencies
- **pywin32**: COM automation for PowerPoint interaction
  - `win32com.client.Dispatch("PowerPoint.Application")`
  - Direct PowerPoint COM automation
- **tkinter**: Standard Python GUI library (included with Python)
  - All standard widgets (no external extensions)

### Platform Considerations
- **Windows-specific**: COM automation requires Windows OS
- **PowerPoint Required**: Microsoft PowerPoint must be installed
- **No Cross-Platform Support**: COM automation is Windows-only

### Limitations & Design Decisions
1. **COM-Based Approach**: Chosen for reliable and accurate slide copying
   - Preserves all formatting, animations, and embedded content
   - Native PowerPoint operations ensure perfect fidelity
   
2. **No Drag-and-Drop**: Removed unreliable `tkinterdnd2` library
   - Replaced with Move Up/Down buttons
   - More reliable and simpler implementation
   
3. **No Undo**: Once merge is initiated, it cannot be cancelled
   - Simplifies state management
   - Acceptable for single-purpose tool

### Future Enhancement Opportunities
- Add progress indicators during merge operations
- Support for batch processing multiple merge operations
- Undo/redo functionality in file reordering
- Preview of presentations before merging
- Custom slide selection (choose specific slides to merge)
