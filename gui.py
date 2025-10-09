"""
GUI module for PowerPoint Merger application.

This module contains all the GUI components and windows for the
step-by-step PowerPoint merging workflow.
"""
import logging
import os
import tkinter as tk
from tkinter import filedialog, messagebox


def show_number_of_files_window(callback):
    """
    Display Step 1: Window for entering number of files to merge.

    Args:
        callback: Function to call with the number of files when Next is clicked
    """
    logging.info("Showing 'Number of files' window (Step 1).")
    window = tk.Tk()
    window.title("Step 1: Number of Files")
    window.geometry("400x150")

    # Label
    label = tk.Label(
        window,
        text="How many files should be merged?",
        font=("Arial", 12)
    )
    label.pack(pady=20)

    # Entry box
    entry = tk.Entry(window, font=("Arial", 12), width=10)
    entry.pack(pady=10)
    entry.focus_set()

    def on_next():
        """Handle Next button click."""
        num_str = entry.get()
        logging.info(f"User clicked 'Next' with input: '{num_str}'.")
        try:
            # Validate input is a positive integer
            num = int(num_str)
            if num <= 0:
                raise ValueError("Number must be positive")
            logging.info(f"Input validated. Number of files: {num}.")
            window.destroy()
            callback(num)
        except ValueError:
            error_msg = "Please enter a valid positive integer."
            logging.error(
                f"Invalid input for number of files: '{num_str}'. Showing error message.")
            messagebox.showerror(
                "Invalid Input",
                error_msg
            )

    # Next button
    next_btn = tk.Button(
        window,
        text="Next",
        command=on_next,
        font=("Arial", 12),
        width=10
    )
    next_btn.pack(pady=10)
    window.bind('<Return>', lambda event=None: next_btn.invoke())

    window.mainloop()


def show_file_selection_window(num_files, callback):
    """
    Display Step 2: Window for selecting files via file dialog.

    Args:
        num_files: Expected number of files to select
        callback: Function to call with the selected files list when OK is clicked
    """
    logging.info(
        f"Showing 'Select files' window (Step 2). Expecting {num_files} files.")
    window = tk.Tk()
    window.title("Step 2: Select Files")
    window.geometry("600x400")

    selected_files = []

    # Label
    label = tk.Label(
        window,
        text=f"Select {num_files} PowerPoint file(s):",
        font=("Arial", 12)
    )
    label.pack(pady=10)

    # Listbox for displaying files
    listbox_frame = tk.Frame(window)
    listbox_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(listbox_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(
        listbox_frame,
        font=("Arial", 10),
        yscrollcommand=scrollbar.set
    )
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)

    def add_files_from_disk():
        """Handle Add Files button click."""
        logging.info("Opening file dialog to select files.")
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if not files:
            logging.info("No files were selected in the file dialog.")
            return

        logging.info(f"{len(files)} file(s) selected: {files}")
        for file in files:
            if file not in selected_files:
                selected_files.append(file)
                listbox.insert(tk.END, file)
                logging.info(f"Added file to list: {file}")
            else:
                logging.warning(
                    f"File is already in the list and was ignored: {file}")

    # Add Files button
    add_btn = tk.Button(
        window,
        text="Add Files from Disk",
        command=add_files_from_disk,
        font=("Arial", 12)
    )
    add_btn.pack(pady=5)

    def on_ok():
        """Handle OK button click."""
        logging.info("User clicked 'OK' in file selection window.")
        if len(selected_files) != num_files:
            error_msg = (f"Please select exactly {num_files} file(s). "
                         f"You have selected {len(selected_files)} file(s).")
            logging.error(f"Wrong number of files selected. {error_msg}")
            messagebox.showerror("Invalid Selection", error_msg)
        else:
            logging.info(
                "Correct number of files selected. Validating file paths.")
            for file in selected_files:
                if not os.path.exists(file):
                    error_msg = f"File does not exist: {file}"
                    logging.error(error_msg)
                    messagebox.showerror("File Not Found", error_msg)
                    return
                if not file.lower().endswith('.pptx'):
                    error_msg = f"File is not a .pptx file: {file}"
                    logging.error(error_msg)
                    messagebox.showerror("Invalid File", error_msg)
                    return
            logging.info(
                "All file paths are validated. Continuing to next step.")
            window.destroy()
            callback(selected_files)

    # OK button
    ok_btn = tk.Button(
        window,
        text="OK",
        command=on_ok,
        font=("Arial", 12),
        width=10
    )
    ok_btn.pack(pady=10)

    window.mainloop()


def show_filename_window(callback):
    """
    Display Step 3: Window for entering the output filename.

    Args:
        callback: Function to call with the filename when Next is clicked
    """
    logging.info("Showing 'New filename' window (Step 3).")
    window = tk.Tk()
    window.title("New Filename")
    window.geometry("400x150")

    # Label
    label = tk.Label(
        window,
        text="New file name:",
        font=("Arial", 12)
    )
    label.pack(pady=20)

    # Entry box
    entry = tk.Entry(window, font=("Arial", 12), width=30)
    entry.pack(pady=10)
    entry.focus_set()

    def on_next():
        """Handle Next button click."""
        filename = entry.get().strip()
        logging.info(f"User clicked 'Next' with filename: '{filename}'.")
        if not filename:
            error_msg = "Please enter a filename."
            logging.error("Filename input was empty.")
            messagebox.showerror("Invalid Input", error_msg)
            return

        # Ensure .pptx extension
        if not filename.lower().endswith('.pptx'):
            original_filename = filename
            filename += '.pptx'
            logging.info(
                f"Added '.pptx' to filename. From '{original_filename}' to '{filename}'.")

        logging.info(
            f"Filename validated: '{filename}'. Continuing to next step.")
        window.destroy()
        callback(filename)

    # Next button
    next_btn = tk.Button(
        window,
        text="Next",
        command=on_next,
        font=("Arial", 12),
        width=10
    )
    next_btn.pack(pady=10)
    window.bind('<Return>', lambda event=None: next_btn.invoke())

    window.mainloop()


def show_reorder_window(selected_files, callback):
    """
    Display Step 4: Window for reordering files using Move Up/Down buttons.

    Args:
        selected_files: List of file paths to reorder
        callback: Function to call with the reordered files list when Create is clicked
    """
    logging.info("Showing 'Change order' window (Step 4).")
    window = tk.Tk()
    window.title("Step 4: Set Merge Order")
    window.geometry("600x450")

    # Label
    label = tk.Label(
        window,
        text="In what order should the presentations be merged?",
        font=("Arial", 12),
        wraplength=550
    )
    label.pack(pady=10)

    # Initialize file order
    file_order = selected_files.copy()
    logging.info(
        f"Initial file order: {[os.path.basename(f) for f in file_order]}")

    # Listbox for displaying files
    listbox_frame = tk.Frame(window)
    listbox_frame.pack(pady=10, padx=20, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(listbox_frame)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    listbox = tk.Listbox(
        listbox_frame,
        font=("Arial", 10),
        yscrollcommand=scrollbar.set
    )
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Populate listbox with filenames (not full paths for readability)
    for file in file_order:
        listbox.insert(tk.END, os.path.basename(file))

    if listbox.size() > 0:
        listbox.selection_set(0)  # Select first item by default

    def move_up():
        """Move selected item up in the list."""
        selection = listbox.curselection()
        if not selection:
            logging.warning(
                "Move Up button clicked without an element being selected.")
            messagebox.showinfo(
                "No Selection",
                "Please select an item to move.")
            return

        index = selection[0]
        if index == 0:
            logging.info(
                "Ignoring 'Move Up' as element is already at the top.")
            return

        logging.info(
            f"Moving '{
                os.path.basename(
                    file_order[index])}' up from position {index}.")
        # Swap items in file_order list
        file_order[index], file_order[index -
                                      1] = file_order[index - 1], file_order[index]

        # Update listbox
        listbox.delete(0, tk.END)
        for file in file_order:
            listbox.insert(tk.END, os.path.basename(file))

        # Reselect the moved item
        listbox.selection_set(index - 1)
        logging.info(f"New order: {[os.path.basename(f) for f in file_order]}")

    def move_down():
        """Move selected item down in the list."""
        selection = listbox.curselection()
        if not selection:
            logging.warning(
                "Move Down button clicked without an element being selected.")
            messagebox.showinfo(
                "No Selection",
                "Please select an item to move.")
            return

        index = selection[0]
        if index == len(file_order) - 1:
            logging.info(
                "Ignoring 'Move Down' as element is already at the bottom.")
            return

        logging.info(
            f"Moving '{
                os.path.basename(
                    file_order[index])}' down from position {index}.")
        # Swap items in file_order list
        file_order[index], file_order[index +
                                      1] = file_order[index + 1], file_order[index]

        # Update listbox
        listbox.delete(0, tk.END)
        for file in file_order:
            listbox.insert(tk.END, os.path.basename(file))

        # Reselect the moved item
        listbox.selection_set(index + 1)
        logging.info(f"New order: {[os.path.basename(f) for f in file_order]}")

    # Button frame for Move Up and Move Down buttons
    button_frame = tk.Frame(window)
    button_frame.pack(pady=5)

    move_up_btn = tk.Button(
        button_frame,
        text="Move Up",
        command=move_up,
        font=("Arial", 12),
        width=12
    )
    move_up_btn.pack(side=tk.LEFT, padx=5)

    move_down_btn = tk.Button(
        button_frame,
        text="Move Down",
        command=move_down,
        font=("Arial", 12),
        width=12
    )
    move_down_btn.pack(side=tk.LEFT, padx=5)

    def on_create():
        """Handle Create New File button click."""
        logging.info("User clicked 'Create New File'.")
        logging.info(f"Final file order for merging: {file_order}")
        window.destroy()
        callback(file_order)

    # Create New File button
    create_btn = tk.Button(
        window,
        text="Create New File",
        command=on_create,
        font=("Arial", 12),
        width=15,
        bg="#4CAF50",
        fg="white"
    )
    create_btn.pack(pady=10)

    window.mainloop()
