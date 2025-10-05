"""
GUI module for PowerPoint Merger application.

This module contains all the GUI components and windows for the
step-by-step PowerPoint merging workflow.
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox


def show_number_of_files_window(callback):
    """
    Display Step 1: Window for entering number of files to merge.

    Args:
        callback: Function to call with the number of files when Next is clicked
    """
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

    def on_next():
        """Handle Next button click."""
        try:
            # Validate input is a positive integer
            num = int(entry.get())
            if num <= 0:
                raise ValueError("Number must be positive")
            window.destroy()
            callback(num)
        except ValueError:
            messagebox.showerror(
                "Invalid Input",
                "Please enter a valid positive number."
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

    window.mainloop()


def show_file_selection_window(num_files, callback):
    """
    Display Step 2: Window for selecting files via file dialog.

    Args:
        num_files: Expected number of files to select
        callback: Function to call with the selected files list when OK is clicked
    """
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
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )
        for file in files:
            if file not in selected_files:
                selected_files.append(file)
                listbox.insert(tk.END, file)

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
        if len(selected_files) != num_files:
            messagebox.showerror(
                "Invalid Selection",
                f"Please select exactly {num_files} file(s). "
                f"You have selected {len(selected_files)} file(s)."
            )
        else:
            # Validate all files are valid .pptx files
            for file in selected_files:
                if not os.path.exists(file):
                    messagebox.showerror(
                        "File Not Found",
                        f"File does not exist: {file}"
                    )
                    return
                if not file.endswith('.pptx'):
                    messagebox.showerror(
                        "Invalid File",
                        f"File is not a .pptx file: {file}"
                    )
                    return
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

    def on_next():
        """Handle Next button click."""
        filename = entry.get().strip()
        if not filename:
            messagebox.showerror(
                "Invalid Input",
                "Please enter a filename."
            )
            return

        # Ensure .pptx extension
        if not filename.endswith('.pptx'):
            filename += '.pptx'

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

    window.mainloop()


def show_reorder_window(selected_files, callback):
    """
    Display Step 4: Window for reordering files using Move Up/Down buttons.

    Args:
        selected_files: List of file paths to reorder
        callback: Function to call with the reordered files list when Create is clicked
    """
    window = tk.Tk()
    window.title("Step 3: Set Merge Order")
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

    # Populate listbox with filenames (not full paths for readability)
    for file in file_order:
        listbox.insert(tk.END, os.path.basename(file))

    def move_up():
        """Move selected item up in the list."""
        selection = listbox.curselection()
        if not selection:
            messagebox.showinfo(
                "No Selection",
                "Please select an item to move."
            )
            return

        index = selection[0]
        if index == 0:
            # Already at the top
            return

        # Swap items in file_order list
        file_order[index], file_order[index - 1] = \
            file_order[index - 1], file_order[index]

        # Update listbox
        listbox.delete(0, tk.END)
        for file in file_order:
            listbox.insert(tk.END, os.path.basename(file))

        # Reselect the moved item
        listbox.selection_set(index - 1)

    def move_down():
        """Move selected item down in the list."""
        selection = listbox.curselection()
        if not selection:
            messagebox.showinfo(
                "No Selection",
                "Please select an item to move."
            )
            return

        index = selection[0]
        if index == len(file_order) - 1:
            # Already at the bottom
            return

        # Swap items in file_order list
        file_order[index], file_order[index + 1] = \
            file_order[index + 1], file_order[index]

        # Update listbox
        listbox.delete(0, tk.END)
        for file in file_order:
            listbox.insert(tk.END, os.path.basename(file))

        # Reselect the moved item
        listbox.selection_set(index + 1)

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
        window.destroy()
        callback(file_order)

    # Create New File button
    create_btn = tk.Button(
        window,
        text="Create New File",
        command=on_create,
        font=("Arial", 12),
        width=15
    )
    create_btn.pack(pady=10)

    window.mainloop()
