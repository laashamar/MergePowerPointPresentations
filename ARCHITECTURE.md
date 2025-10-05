# ARCHITECTURE.md

## PowerPoint Presentation Merger - Architecture Overview

### Design Pattern
The application uses a **class-based architecture** with a single main class `PowerPointMerger` that manages the entire workflow state and GUI windows.

### Class Structure

```
PowerPointMerger
├── __init__()              # Initialize application state
├── run()                   # Entry point - start first window
├── show_number_of_files_window()
├── show_file_selection_window()
├── show_filename_window()
├── show_reorder_window()
├── merge_and_launch()      # Backend merge logic
└── _copy_shape()           # Helper method for shape copying
```

### State Management
The class maintains four key state variables:
- `num_files`: Expected number of files to merge
- `selected_files`: List of file paths selected by user
- `output_filename`: Name for the merged presentation
- `file_order`: Final order of files after user reordering

### GUI Framework
- **Primary**: `tkinter` for standard GUI components
- **Extended**: `tkinterdnd2` for drag-and-drop functionality
  - File selection window uses `TkinterDnD.Tk()`
  - Reordering window uses `TkinterDnD.Tk()`

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
for each source_file in file_order:
    for each slide in source_file.slides:
        1. Add new blank slide to destination
        2. For each shape in source slide:
            - Copy shape element
            - Add to destination slide
```

### Error Handling Strategy
- **Validation**: Input validation at each step before proceeding
- **User Feedback**: Clear error messages via `messagebox.showerror()`
- **Graceful Degradation**: If PowerPoint launch fails, presentation is still saved
- **Try-Except Blocks**: Wrap file operations and merge logic

### External Dependencies
1. **python-pptx**: PowerPoint file manipulation
   - `Presentation()`: Load/create presentations
   - `slides`: Access slide collection
   - `shapes`: Access shape collection
   
2. **tkinterdnd2**: Drag-and-drop support
   - `DND_FILES`: File drop target type
   - `TkinterDnD.Tk()`: Enhanced root window

3. **subprocess**: Launch PowerPoint
   - `Popen(['powerpnt.exe', '/s', path])`

### Platform Considerations
- **Windows-specific**: PowerPoint launch uses `powerpnt.exe`
- **Cross-platform GUI**: tkinter components work on all platforms
- **Cross-platform merge**: python-pptx works on all platforms

### Limitations & Design Decisions
1. **Shape Copying**: Uses element-level copying due to python-pptx limitations
   - No direct slide.clone() method available
   - Element copying preserves most properties
   
2. **Layout Handling**: Uses first available layout for new slides
   - Simplifies implementation
   - Works for most use cases
   
3. **No Undo**: Once merge is initiated, it cannot be cancelled
   - Simplifies state management
   - Acceptable for single-purpose tool

### Future Enhancement Opportunities
- Preview of selected slides before merging
- Progress bar during merge operation
- Ability to select specific slides from each presentation
- Support for custom slide layouts
- Cross-platform slideshow launch
- Batch merge operations
