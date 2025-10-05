"""
Application module for PowerPoint Merger.

This module orchestrates the workflow, manages application state,
and connects GUI events with the backend merging logic.
"""

from tkinter import messagebox
import gui
import core


class PowerPointMergerApp:
    """Main application class that manages state and workflow."""

    def __init__(self):
        """Initialize the application state."""
        self.num_files = 0
        self.selected_files = []
        self.output_filename = ""
        self.file_order = []

    def run(self):
        """Start the application with the first window."""
        gui.show_number_of_files_window(self._on_number_of_files_entered)

    def _on_number_of_files_entered(self, num_files):
        """Handle number of files input from Step 1."""
        self.num_files = num_files
        gui.show_file_selection_window(
            self.num_files,
            self._on_files_selected
        )

    def _on_files_selected(self, selected_files):
        """Handle file selection from Step 2."""
        self.selected_files = selected_files
        gui.show_filename_window(self._on_filename_entered)

    def _on_filename_entered(self, filename):
        """Handle filename input from Step 3."""
        self.output_filename = filename
        gui.show_reorder_window(
            self.selected_files,
            self._on_files_reordered
        )

    def _on_files_reordered(self, file_order):
        """Handle file reordering from Step 4."""
        self.file_order = file_order
        self._merge_and_launch()

    def _merge_and_launch(self):
        """Merge presentations and launch the slideshow."""
        # Merge presentations using COM automation
        success, output_path, error_msg = core.merge_presentations(
            self.file_order,
            self.output_filename
        )

        if not success:
            messagebox.showerror(
                "Merge Failed",
                f"Failed to merge presentations: {error_msg}"
            )
            return

        messagebox.showinfo(
            "Success",
            f"Presentation merged successfully!\n\nSaved to: {output_path}"
        )

        # Launch slideshow using COM automation
        success, error_msg = core.launch_slideshow(output_path)

        if not success:
            messagebox.showwarning(
                "Launch Failed",
                f"Presentation saved successfully to {output_path}, "
                f"but failed to launch slideshow: {error_msg}"
            )


def start_app():
    """Start the PowerPoint Merger application."""
    app = PowerPointMergerApp()
    app.run()
