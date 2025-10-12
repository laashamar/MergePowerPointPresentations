"""
GUI module for PowerPoint Merger application.

This module contains the modern two-column GUI with drag-and-drop
support for PowerPoint file merging.
"""
import logging
import os
import tkinter as tk
from tkinter import filedialog, messagebox
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    logging.warning("tkinterdnd2 not available. Drag-and-drop will be disabled.")


class PowerPointMergerGUI:
    """Modern GUI for PowerPoint Merger with two-column layout."""

    def __init__(self, merge_callback):
        """
        Initialize the GUI.

        Args:
            merge_callback: Function to call when merge is requested
                           Should accept (file_list, output_path) parameters
        """
        self.merge_callback = merge_callback
        self.file_list = []  # List of file paths in merge order

        # Create main window
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("PowerPoint Merger")
        self.root.geometry("900x600")

        # Set application icon
        icon_path = os.path.join(
            os.path.dirname(__file__),
            "resources",
            "MergePowerPoint.ico"
        )
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception as e:
                logging.warning(f"Could not set application icon: {e}")

        self._create_widgets()
        self._update_merge_queue_display()

    def _create_widgets(self):
        """Create and layout all GUI widgets."""
        # Main container with two columns
        main_container = tk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Column 1: Merge Queue (left side, wider)
        self.queue_frame = tk.Frame(main_container, relief=tk.RIDGE, borderwidth=2)
        self.queue_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        queue_label = tk.Label(
            self.queue_frame,
            text="Merge Queue",
            font=("Arial", 14, "bold")
        )
        queue_label.pack(pady=(10, 5))

        # Container for drop zone or file list
        self.content_frame = tk.Frame(self.queue_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Column 2: Configuration & Actions (right side)
        config_frame = tk.Frame(main_container, relief=tk.RIDGE, borderwidth=2)
        config_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(5, 0))
        config_frame.config(width=300)

        config_label = tk.Label(
            config_frame,
            text="Configuration",
            font=("Arial", 14, "bold")
        )
        config_label.pack(pady=(10, 5))

        # Output folder selector
        folder_frame = tk.LabelFrame(
            config_frame,
            text="Output Location",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        folder_frame.pack(fill=tk.X, padx=10, pady=10)

        self.output_folder_var = tk.StringVar(value=os.path.expanduser("~\\Desktop"))

        folder_entry = tk.Entry(
            folder_frame,
            textvariable=self.output_folder_var,
            font=("Arial", 9),
            state="readonly"
        )
        folder_entry.pack(fill=tk.X, pady=(0, 5))

        browse_folder_btn = tk.Button(
            folder_frame,
            text="Browse",
            command=self._browse_output_folder,
            font=("Arial", 10)
        )
        browse_folder_btn.pack(fill=tk.X)

        # Output filename
        filename_frame = tk.LabelFrame(
            config_frame,
            text="Output Filename",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=10
        )
        filename_frame.pack(fill=tk.X, padx=10, pady=10)

        self.output_filename_var = tk.StringVar(value="merged_presentation.pptx")

        filename_entry = tk.Entry(
            filename_frame,
            textvariable=self.output_filename_var,
            font=("Arial", 9)
        )
        filename_entry.pack(fill=tk.X)

        # Action buttons
        button_frame = tk.Frame(config_frame)
        button_frame.pack(fill=tk.X, padx=10, pady=20)

        self.merge_btn = tk.Button(
            button_frame,
            text="Merge Presentations",
            command=self._on_merge,
            font=("Arial", 11, "bold"),
            bg="#4CAF50",
            fg="white",
            height=2
        )
        self.merge_btn.pack(fill=tk.X, pady=(0, 10))

        clear_btn = tk.Button(
            button_frame,
            text="Clear Queue",
            command=self._clear_queue,
            font=("Arial", 10)
        )
        clear_btn.pack(fill=tk.X)

        # Status label at bottom
        self.status_var = tk.StringVar(value="Ready")
        status_label = tk.Label(
            config_frame,
            textvariable=self.status_var,
            font=("Arial", 9),
            fg="blue",
            wraplength=280,
            justify=tk.LEFT
        )
        status_label.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

    def _create_drop_zone(self):
        """Create the initial drop zone interface."""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        drop_container = tk.Frame(self.content_frame, bg="#f0f0f0")
        drop_container.pack(fill=tk.BOTH, expand=True)

        # Large plus sign icon (using label with large font)
        plus_label = tk.Label(
            drop_container,
            text="+",
            font=("Arial", 72, "bold"),
            bg="#f0f0f0",
            fg="#888888"
        )
        plus_label.pack(expand=True, pady=(50, 10))

        # Instructional text
        instruction_label = tk.Label(
            drop_container,
            text="Drag and drop PowerPoint files here",
            font=("Arial", 12),
            bg="#f0f0f0",
            fg="#555555"
        )
        instruction_label.pack()

        # Browse button as alternative
        browse_btn = tk.Button(
            drop_container,
            text="Browse for Files",
            command=self._browse_files,
            font=("Arial", 11),
            width=20,
            height=2
        )
        browse_btn.pack(pady=(20, 50))

        # Enable drag-and-drop if available
        if HAS_DND:
            drop_container.drop_target_register(DND_FILES)
            drop_container.dnd_bind('<<Drop>>', self._on_drop)

    def _create_file_list(self):
        """Create the file list interface with reordering capability."""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        # Scrollable canvas for file cards
        canvas = tk.Canvas(self.content_frame)
        scrollbar = tk.Scrollbar(self.content_frame, orient="vertical", command=canvas.yview)
        self.file_list_frame = tk.Frame(canvas)

        self.file_list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.file_list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Create file cards
        for i, file_path in enumerate(self.file_list):
            self._create_file_card(i, file_path)

    def _create_file_card(self, index, file_path):
        """Create a card widget for a file in the queue."""
        card = tk.Frame(self.file_list_frame, relief=tk.RAISED, borderwidth=1, bg="white")
        card.pack(fill=tk.X, padx=5, pady=2)

        # File info frame
        info_frame = tk.Frame(card, bg="white")
        info_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)

        # PowerPoint icon (using emoji/text)
        icon_label = tk.Label(
            info_frame,
            text="📊",
            font=("Arial", 16),
            bg="white"
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 10))

        # Filename
        filename = os.path.basename(file_path)
        name_label = tk.Label(
            info_frame,
            text=filename,
            font=("Arial", 10),
            bg="white",
            anchor="w"
        )
        name_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Create tooltip showing full path
        self._create_tooltip(name_label, file_path)

        # Reorder buttons
        button_frame = tk.Frame(card, bg="white")
        button_frame.pack(side=tk.LEFT, padx=5)

        up_btn = tk.Button(
            button_frame,
            text="↑",
            command=lambda idx=index: self._move_file_up(idx),
            font=("Arial", 10),
            width=2
        )
        up_btn.pack(side=tk.LEFT, padx=2)

        down_btn = tk.Button(
            button_frame,
            text="↓",
            command=lambda idx=index: self._move_file_down(idx),
            font=("Arial", 10),
            width=2
        )
        down_btn.pack(side=tk.LEFT, padx=2)

        # Remove button
        remove_btn = tk.Button(
            card,
            text="✕",
            command=lambda idx=index: self._remove_file(idx),
            font=("Arial", 10),
            fg="red",
            bg="white",
            borderwidth=0,
            width=2
        )
        remove_btn.pack(side=tk.RIGHT, padx=5)

    def _create_tooltip(self, widget, text):
        """Create a tooltip for a widget."""
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")

            label = tk.Label(
                tooltip,
                text=text,
                background="#ffffe0",
                relief=tk.SOLID,
                borderwidth=1,
                font=("Arial", 9)
            )
            label.pack()

            widget.tooltip = tooltip

        def hide_tooltip(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                delattr(widget, 'tooltip')

        widget.bind('<Enter>', show_tooltip)
        widget.bind('<Leave>', hide_tooltip)

    def _update_merge_queue_display(self):
        """Update the merge queue display based on file list state."""
        if not self.file_list:
            self._create_drop_zone()
            self.merge_btn.config(state=tk.DISABLED)
        else:
            self._create_file_list()
            self.merge_btn.config(state=tk.NORMAL)

    def _browse_files(self):
        """Open file dialog to browse for PowerPoint files."""
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[
                ("PowerPoint Files", "*.pptx *.ppsx"),
                ("All Files", "*.*")
            ]
        )

        if files:
            self._add_files(list(files))

    def _add_files(self, file_paths):
        """
        Add files to the merge queue with validation.

        Args:
            file_paths: List of file paths to add
        """
        added_count = 0

        for file_path in file_paths:
            # Validate file type
            if not file_path.lower().endswith(('.pptx', '.ppsx')):
                messagebox.showwarning(
                    "Invalid File Type",
                    f"Only PowerPoint files (.pptx and .ppsx) are supported.\n\n"
                    f"Invalid file: {os.path.basename(file_path)}"
                )
                continue

            # Check if file already exists in queue
            if file_path in self.file_list:
                messagebox.showinfo(
                    "Duplicate File",
                    f"This file has already been added.\n\n{os.path.basename(file_path)}"
                )
                continue

            # Check file access
            if not os.path.exists(file_path):
                messagebox.showerror(
                    "File Not Found",
                    f"The specified file does not exist.\n\n{file_path}"
                )
                continue

            try:
                # Try to open the file to check permissions
                with open(file_path, 'rb'):
                    pass

                # File is valid and accessible, add to queue
                self.file_list.append(file_path)
                added_count += 1
                logging.info(f"Added file to queue: {file_path}")

            except PermissionError:
                messagebox.showerror(
                    "Access Denied",
                    f"Access denied. Unable to open the file.\n\n"
                    f"{os.path.basename(file_path)}\n\n"
                    f"The file may be in use by another application or you may "
                    f"lack the necessary permissions."
                )
            except Exception as e:
                messagebox.showerror(
                    "File Access Error",
                    f"Could not access file: {os.path.basename(file_path)}\n\n"
                    f"Error: {str(e)}"
                )

        if added_count > 0:
            self._update_merge_queue_display()
            self.status_var.set(f"Added {added_count} file(s) to queue")

    def _on_drop(self, event):
        """Handle drag-and-drop event."""
        # Parse dropped files
        files = self.root.tk.splitlist(event.data)
        self._add_files(files)

    def _move_file_up(self, index):
        """Move file up in the queue."""
        if index > 0:
            self.file_list[index], self.file_list[index-1] = \
                self.file_list[index-1], self.file_list[index]
            self._update_merge_queue_display()
            logging.info(f"Moved file up: {os.path.basename(self.file_list[index])}")

    def _move_file_down(self, index):
        """Move file down in the queue."""
        if index < len(self.file_list) - 1:
            self.file_list[index], self.file_list[index+1] = \
                self.file_list[index+1], self.file_list[index]
            self._update_merge_queue_display()
            logging.info(f"Moved file down: {os.path.basename(self.file_list[index])}")

    def _remove_file(self, index):
        """Remove file from the queue."""
        removed_file = self.file_list.pop(index)
        logging.info(f"Removed file from queue: {removed_file}")
        self._update_merge_queue_display()
        self.status_var.set(f"Removed: {os.path.basename(removed_file)}")

    def _clear_queue(self):
        """Clear all files from the queue."""
        if self.file_list:
            if messagebox.askyesno(
                "Clear Queue",
                "Are you sure you want to remove all files from the queue?"
            ):
                self.file_list.clear()
                self._update_merge_queue_display()
                self.status_var.set("Queue cleared")
                logging.info("Merge queue cleared")

    def _browse_output_folder(self):
        """Open folder dialog to select output location."""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_folder_var.get()
        )

        if folder:
            self.output_folder_var.set(folder)
            logging.info(f"Output folder set to: {folder}")

    def _on_merge(self):
        """Handle merge button click."""
        if not self.file_list:
            messagebox.showwarning(
                "No Files",
                "Please add at least one PowerPoint file to the merge queue."
            )
            return

        # Validate output filename
        filename = self.output_filename_var.get().strip()
        if not filename:
            messagebox.showerror(
                "Invalid Filename",
                "Please enter a valid output filename."
            )
            return

        # Ensure .pptx extension
        if not filename.lower().endswith('.pptx'):
            filename += '.pptx'
            self.output_filename_var.set(filename)

        # Build full output path
        output_path = os.path.join(self.output_folder_var.get(), filename)

        # Check if output file already exists
        if os.path.exists(output_path):
            if not messagebox.askyesno(
                "File Exists",
                f"The file '{filename}' already exists.\n\n"
                f"Do you want to overwrite it?"
            ):
                return

        # Update status
        self.status_var.set(f"Merging {len(self.file_list)} presentations...")
        self.merge_btn.config(state=tk.DISABLED)
        self.root.update()

        # Call merge callback
        logging.info(f"Starting merge of {len(self.file_list)} files to {output_path}")
        self.merge_callback(self.file_list.copy(), output_path)

    def update_status(self, message):
        """Update the status label."""
        self.status_var.set(message)
        self.root.update()

    def enable_merge_button(self):
        """Re-enable the merge button."""
        if self.file_list:
            self.merge_btn.config(state=tk.NORMAL)

    def run(self):
        """Start the GUI main loop."""
        logging.info("Starting PowerPoint Merger GUI")
        self.root.mainloop()


def show_modern_gui(merge_callback):
    """
    Display the modern PowerPoint Merger GUI.

    Args:
        merge_callback: Function to call when merge is requested.
                       Should accept (file_list, output_path) parameters.
    """
    gui = PowerPointMergerGUI(merge_callback)
    gui.run()
