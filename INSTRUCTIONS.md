

**Objective:**

Create a Python script that provides a Graphical User Interface (GUI) for merging multiple PowerPoint (`.pptx`) files into a single, new `.pptx` file. The script should guide the user through selecting, ordering, and naming the files, and then automatically open the final merged presentation in slideshow mode. The script should be designed to run on Windows.

**Required Libraries:**

* `tkinter` for the GUI components (including `tkinter.filedialog`).
* `python-pptx` for reading and writing PowerPoint files.
* `os` and `subprocess` for opening the final file.

**GUI Workflow (Step-by-Step):**

1.  **Initial Window: Number of Files**
    * Display a window with the title "Step 1: Number of Files".
    * Include a label: "How many files should be merged?"
    * Provide a text entry box for the user to input a number.
    * Include a "Next" button. When clicked, this window closes and the next one opens.

2.  **Second Window: File Selection**
    * Display a window with the title "Step 2: Select Files".
    * This window must support two ways of adding `.pptx` files:
        * **Drag and Drop:** Implement a listbox where the user can drag and drop their `.pptx` files from File Explorer.
        * **File Dialog Button:** Include a button labeled "Add Files from Disk" that opens a standard file selection dialog, allowing the user to select multiple `.pptx` files.
    * Display the full paths of the selected files in the listbox.
    * Include an "OK" button. The script should validate that the number of files selected matches the number provided in Step 1 before proceeding. If it doesn't match, show an error message.

3.  **Third Window: New Filename**
    * After the user clicks "OK," open a new dialog with a label: "New file name".
    * Provide a text entry box.
    * The file extension `.pptx` should be hardcoded. The user only types the base name (e.g., "MergedPresentation"), and the script will save it as "MergedPresentation.pptx". Do not allow the user to change the extension.
    * Include a "Next" button.

4.  **Fourth Window: Reorder Files**
    * Display a window with the title "Step 3: Set Merge Order".
    * Include a label: "In what order should the presentations be merged? (Drag to reorder)".
    * Display the selected file names in a listbox.
    * Implement drag-and-drop functionality so the user can reorder the items in the list to define the final merge sequence.
    * Include a "Create New File" button.

**Backend Logic (Executed upon clicking "Create New File"):**

1.  **Merge Presentations by Slide:**
    * Create a new, blank `pptx.Presentation` object.
    * Iterate through the list of source `.pptx` files **in the exact order specified by the user in the reordering window**.
    * For each source presentation, iterate through all of its slides. The slides from the second presentation should be appended after the last slide of the first, and so on.
    * To copy a slide, add a new slide to the destination presentation using the same slide layout. Then, iterate through all shapes on the source slide and copy each shape to the newly created slide in the destination presentation. *Note: `python-pptx` does not support direct copying of slides between presentations, so you must replicate them by copying their layouts and shapes.*

2.  **Save the File:**
    * Save the destination presentation using the filename provided by the user in Step 3, ensuring the `.pptx` extension is appended.

3.  **Launch Slideshow:**
    * After successfully saving the file, automatically open the newly created `.pptx` file directly in slideshow mode starting from the first slide.
    * Use the `subprocess` module to call PowerPoint from the command line on Windows: `powerpnt.exe /s "C:\\path\\to\\newfile.pptx"`.

4.  **Script Termination:**
    * After launching the slideshow, the GUI application should close, and the script should terminate gracefully in the background.

**Requirements & Best Practices:**

* Follow **PEP 8** styling guidelines for clean, readable code.
* Include **comments** to explain complex parts of the code, especially the GUI layout and the slide-copying logic.
* Implement basic **error handling**. For example, handle cases where the user enters non-numeric input for the number of files, or if a selected file is not a valid `.pptx` presentation.