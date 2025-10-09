"""Main GUI module for PowerPoint Merger application using CustomTkinter."""

import logging
import os
import subprocess
import sys
import threading
from tkinter import filedialog

import customtkinter
from tkinterdnd2 import DND_FILES, TkinterDnD

# --- Path adjustment to allow importing from the parent directory ---
# This is the crucial part that allows the new GUI to find the core logic.
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# --- Application Core Logic Imports ---
import powerpoint_core  # noqa: E402, F401 - Will be used in future phases


# --- Main Application Class ---
class App(customtkinter.CTk):
    """Main application window for PowerPoint merger."""

    def __init__(self):
        """Initialize the main application window."""
        super().__init__()

        # Set window title and geometry
        self.title("Merge PowerPoint Presentations")
        self.geometry("800x600")

        # Initialize file list
        self.file_list = []

        # Initialize selected file index for reordering
        self.selected_file_index = None

        # Configure logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        # Configure grid layout (3 rows, 3 columns)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)

        # 2.1: File Display List
        self.file_list_frame = customtkinter.CTkScrollableFrame(
            self,
            label_text="Selected Presentations"
        )
        self.file_list_frame.grid(
            row=0,
            column=1,
            columnspan=2,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        # Configure drag-and-drop for file addition
        self._setup_drag_and_drop()

        # 2.2: File Management Buttons
        file_buttons_frame = customtkinter.CTkFrame(self)
        file_buttons_frame.grid(
            row=0,
            column=0,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        self.add_button = customtkinter.CTkButton(
            file_buttons_frame,
            text="Add Presentation(s)",
            command=self.add_files
        )
        self.add_button.pack(padx=10, pady=10)

        self.remove_button = customtkinter.CTkButton(
            file_buttons_frame,
            text="Remove Selected",
            command=self.remove_selected_file
        )
        self.remove_button.pack(padx=10, pady=10)

        self.clear_button = customtkinter.CTkButton(
            file_buttons_frame,
            text="Clear All",
            command=self.clear_file_list
        )
        self.clear_button.pack(padx=10, pady=10)

        # 2.3: Output File Configuration
        output_frame = customtkinter.CTkFrame(self)
        output_frame.grid(
            row=1,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="ew"
        )

        output_label = customtkinter.CTkLabel(
            output_frame,
            text="Output File:"
        )
        output_label.pack(side="left", padx=10, pady=10)

        self.output_path_entry = customtkinter.CTkEntry(
            output_frame,
            placeholder_text="Select output path..."
        )
        self.output_path_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=10,
            pady=10
        )

        self.save_as_button = customtkinter.CTkButton(
            output_frame,
            text="Save As...",
            command=self.select_output_path
        )
        self.save_as_button.pack(side="left", padx=10, pady=10)

        # 2.4: Primary Action and Status
        action_frame = customtkinter.CTkFrame(self)
        action_frame.grid(
            row=2,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="ew"
        )

        self.merge_button = customtkinter.CTkButton(
            action_frame,
            text="Merge Presentations",
            command=self.merge_presentations,
            state="disabled"
        )
        self.merge_button.pack(pady=10)

        self.status_label = customtkinter.CTkLabel(
            action_frame,
            text="Ready"
        )
        self.status_label.pack(pady=5)

        # Post-merge action buttons (hidden by default)
        self.open_file_button = customtkinter.CTkButton(
            action_frame,
            text="Open Presentation",
            command=self.open_merged_file
        )

        self.show_in_explorer_button = customtkinter.CTkButton(
            action_frame,
            text="Show in Explorer",
            command=self.show_in_file_explorer
        )

        # Store the output path for post-merge actions
        self.last_merged_file_path = None

    def _setup_drag_and_drop(self):
        """Configure drag-and-drop functionality for file addition."""
        try:
            # Register the scrollable frame as a drop target
            self.file_list_frame.drop_target_register(DND_FILES)
            self.file_list_frame.dnd_bind('<<Drop>>', self._on_drop)
            logging.info("Drag-and-drop configured successfully")
        except Exception as e:
            logging.warning(f"Could not configure drag-and-drop: {e}")

    def _on_drop(self, event):
        """
        Handle file drop events.

        Args:
            event: Drop event containing file paths
        """
        try:
            # Parse dropped files (tkinterdnd2 returns them as a string)
            files = self._parse_drop_files(event.data)

            added_count = 0
            for file in files:
                # Validate file extension
                if os.path.splitext(file)[1].lower() == '.pptx':
                    if file not in self.file_list:
                        self.file_list.append(file)
                        added_count += 1
                        logging.info(f"Added file via drag-and-drop: {file}")

            if added_count > 0:
                self.update_file_list_display()
                self.update_merge_button_status()
                logging.info(f"Total files added: {added_count}")

        except Exception as e:
            logging.error(f"Error handling dropped files: {e}")

    def _parse_drop_files(self, data):
        """
        Parse file paths from drop event data.

        Args:
            data: Raw drop event data

        Returns:
            list: List of file paths
        """
        # Handle different formats that tkinterdnd2 might return
        files = []
        if isinstance(data, str):
            # Split by spaces, but handle paths with spaces (enclosed in {})
            parts = []
            current = ""
            in_braces = False

            for char in data:
                if char == '{':
                    in_braces = True
                elif char == '}':
                    in_braces = False
                    if current:
                        parts.append(current)
                        current = ""
                elif char == ' ' and not in_braces:
                    if current:
                        parts.append(current)
                        current = ""
                else:
                    current += char

            if current:
                parts.append(current)

            files = parts

        return files

    def add_files(self):
        """Open file dialog to add PowerPoint files."""
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )

        for file in files:
            if file not in self.file_list:
                self.file_list.append(file)

        self.update_file_list_display()
        self.update_merge_button_status()

    def update_file_list_display(self):
        """Update the file list display in the scrollable frame."""
        # Clear current contents
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()

        # Display each file with drag-and-drop reordering support
        for idx, file_path in enumerate(self.file_list):
            # Create a frame for each file item
            file_frame = customtkinter.CTkFrame(self.file_list_frame)
            file_frame.pack(fill="x", padx=5, pady=2)

            label = customtkinter.CTkLabel(
                file_frame,
                text=f"{idx + 1}. {os.path.basename(file_path)}",
                anchor="w"
            )
            label.pack(side="left", fill="x", expand=True, padx=5)

            # Bind mouse events for drag-and-drop reordering
            label.bind("<Button-1>", lambda e, i=idx: self._on_label_click(i))
            label.bind("<B1-Motion>", self._on_label_drag)
            label.bind("<ButtonRelease-1>", self._on_label_release)

    def _on_label_click(self, index):
        """
        Handle mouse click on file label for drag-and-drop reordering.

        Args:
            index: Index of the clicked file in the file_list
        """
        self.selected_file_index = index
        logging.info(f"Selected file for reordering: {os.path.basename(self.file_list[index])}")

    def _on_label_drag(self, event):
        """
        Handle mouse drag motion for reordering.

        Args:
            event: Motion event
        """
        # Visual feedback could be added here if needed
        pass

    def _on_label_release(self, event):
        """
        Handle mouse release to complete reordering.

        Args:
            event: Button release event
        """
        if self.selected_file_index is None:
            return

        # Find which label the mouse is over
        widget = event.widget.winfo_containing(event.x_root, event.y_root)

        if widget is None:
            self.selected_file_index = None
            return

        # Find the target index by checking all frames
        target_index = None
        for idx, child in enumerate(self.file_list_frame.winfo_children()):
            if child == widget or child == widget.master:
                target_index = idx
                break

        if target_index is not None and target_index != self.selected_file_index:
            # Perform the reordering
            file_to_move = self.file_list.pop(self.selected_file_index)
            self.file_list.insert(target_index, file_to_move)

            logging.info(
                f"Reordered: moved '{os.path.basename(file_to_move)}' "
                f"from position {self.selected_file_index + 1} to {target_index + 1}"
            )

            self.update_file_list_display()

        self.selected_file_index = None

    def remove_selected_file(self):
        """Remove the last file from the list (placeholder implementation)."""
        if self.file_list:
            self.file_list.pop()
            self.update_file_list_display()
            self.update_merge_button_status()

    def clear_file_list(self):
        """Clear all files from the list."""
        self.file_list.clear()
        self.update_file_list_display()
        self.update_merge_button_status()

    def select_output_path(self):
        """Open save-as dialog to select output file path."""
        file_path = filedialog.asksaveasfilename(
            title="Save Merged Presentation As",
            defaultextension=".pptx",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )

        if file_path:
            self.output_path_entry.delete(0, "end")
            self.output_path_entry.insert(0, file_path)
            self.update_merge_button_status()

    def update_merge_button_status(self):
        """Enable or disable merge button based on conditions."""
        # Check if at least 2 files are selected
        has_files = len(self.file_list) >= 2

        # Check if output path is specified
        has_output = bool(self.output_path_entry.get().strip())

        # Enable button if both conditions are met
        if has_files and has_output:
            self.merge_button.configure(state="normal")
        else:
            self.merge_button.configure(state="disabled")

    def merge_presentations(self):
        """Execute the merge process in a separate thread with progress updates."""
        output_path = self.output_path_entry.get()

        if not self.file_list or not output_path:
            logging.error("Cannot merge: missing files or output path")
            return

        # Disable merge button during processing
        self.merge_button.configure(state="disabled")
        self.status_label.configure(text="Starting merge...")

        # Hide post-merge buttons
        self.open_file_button.pack_forget()
        self.show_in_explorer_button.pack_forget()

        logging.info(f"Starting merge of {len(self.file_list)} files")

        # Run merge in separate thread
        merge_thread = threading.Thread(
            target=self._perform_merge_thread,
            args=(self.file_list.copy(), output_path),
            daemon=True
        )
        merge_thread.start()

    def _perform_merge_thread(self, file_list, output_path):
        """
        Perform the actual merge operation in a separate thread.

        Args:
            file_list: List of files to merge
            output_path: Path for the output file
        """
        try:
            success, result_path, error_msg = powerpoint_core.merge_presentations(
                file_list,
                output_path,
                progress_callback=self._merge_progress_callback
            )

            if success:
                self.last_merged_file_path = result_path
                self._update_status_safe("Merge Complete!")
                self._show_post_merge_buttons()
                logging.info(f"Merge completed successfully: {result_path}")
            else:
                self._update_status_safe(f"Error: {error_msg}")
                logging.error(f"Merge failed: {error_msg}")

        except Exception as e:
            error_msg = str(e)
            self._update_status_safe(f"Error: {error_msg}")
            logging.error(f"Merge exception: {error_msg}", exc_info=True)

        finally:
            # Re-enable merge button
            self.after(0, lambda: self.merge_button.configure(state="normal"))

    def _merge_progress_callback(self, filename, current_slide, total_slides):
        """
        Thread-safe callback for merge progress updates.

        Args:
            filename: Name of the file currently being processed
            current_slide: Current slide number being processed
            total_slides: Total number of slides in the current file
        """
        status_text = f"Merging \"{filename}\" (slide {current_slide} of {total_slides})..."
        self._update_status_safe(status_text)
        logging.info(f"Progress: {filename} - slide {current_slide}/{total_slides}")

    def _update_status_safe(self, text):
        """
        Update status label in a thread-safe manner using self.after().

        Args:
            text: Status text to display
        """
        self.after(0, lambda: self.status_label.configure(text=text))

    def _show_post_merge_buttons(self):
        """Show post-merge action buttons in a thread-safe manner."""
        def show_buttons():
            self.open_file_button.pack(pady=5)
            self.show_in_explorer_button.pack(pady=5)

        self.after(0, show_buttons)

    def open_merged_file(self):
        """Open the merged presentation file using the system default application."""
        if not self.last_merged_file_path:
            logging.warning("No merged file available to open")
            return

        try:
            if sys.platform == "win32":
                os.startfile(self.last_merged_file_path)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", self.last_merged_file_path])
            else:  # Linux and others
                subprocess.run(["xdg-open", self.last_merged_file_path])

            logging.info(f"Opened file: {self.last_merged_file_path}")

        except Exception as e:
            logging.error(f"Could not open file: {e}")
            self.status_label.configure(text=f"Error opening file: {e}")

    def show_in_file_explorer(self):
        """Open file explorer and highlight the merged file."""
        if not self.last_merged_file_path:
            logging.warning("No merged file available to show")
            return

        try:
            if sys.platform == "win32":
                subprocess.run(['explorer', '/select,', self.last_merged_file_path])
            elif sys.platform == "darwin":  # macOS
                subprocess.run(['open', '-R', self.last_merged_file_path])
            else:  # Linux
                # Most Linux file managers don't support selecting files
                # So we open the containing directory
                directory = os.path.dirname(self.last_merged_file_path)
                subprocess.run(['xdg-open', directory])

            logging.info(f"Showed file in explorer: {self.last_merged_file_path}")

        except Exception as e:
            logging.error(f"Could not show file in explorer: {e}")
            self.status_label.configure(text=f"Error showing file: {e}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
