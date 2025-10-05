"""
Application module for PowerPoint Merger.

This module orchestrates the workflow, manages application state,
and connects GUI events with the backend merging logic.
"""
import logging
from tkinter import messagebox
import gui
import powerpoint_core  # Changed from 'core'


class PowerPointMergerApp:
    """Main application class that manages state and workflow."""

    def __init__(self):
        """Initialize the application state."""
        self.num_files = 0
        self.selected_files = []
        self.output_filename = ""
        self.file_order = []
        logging.info("PowerPointMergerApp instance initialized.")

    def run(self):
        """Start the application with the first window."""
        logging.info("Application workflow starting.")
        gui.show_number_of_files_window(self._on_number_of_files_entered)

    def _on_number_of_files_entered(self, num_files):
        """Handle number of files input from Step 1."""
        logging.info(f"Step 1 completed. Number of files set to: {num_files}.")
        self.num_files = num_files
        gui.show_file_selection_window(
            self.num_files,
            self._on_files_selected
        )

    def _on_files_selected(self, selected_files):
        """Handle file selection from Step 2."""
        logging.info(f"Step 2 completed. {len(selected_files)} files selected.")
        logging.debug(f"Selected files: {selected_files}")
        self.selected_files = selected_files
        gui.show_filename_window(self._on_filename_entered)

    def _on_filename_entered(self, filename):
        """Handle filename input from Step 3."""
        logging.info(f"Step 3 completed. Output filename set to: '{filename}'.")
        self.output_filename = filename
        gui.show_reorder_window(
            self.selected_files,
            self._on_files_reordered
        )

    def _on_files_reordered(self, file_order):
        """Handle file reordering from Step 4."""
        logging.info("Step 4 completed. File order has been set.")
        logging.debug(f"Final file order: {file_order}")
        self.file_order = file_order
        self._merge_and_launch()

    def _merge_and_launch(self):
        """Merge presentations and launch the slideshow."""
        logging.info("Starting merge and slideshow of presentations.")
        
        # Merge presentations using COM automation
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            self.file_order,
            self.output_filename
        )

        if not success:
            # Error already logged in core module
            messagebox.showerror(
                "Merge Failed",
                f"Failed to merge presentations: {error_msg}"
            )
            return

        logging.info(f"Merge was successful. File saved to: {output_path}")
        messagebox.showinfo(
            "Success",
            f"Presentation was merged successfully!\n\nSaved to: {output_path}"
        )

        # Launch slideshow using COM automation
        logging.info(f"Starting slideshow for: {output_path}")
        success, error_msg = powerpoint_core.launch_slideshow(output_path)

        if not success:
            # Error already logged in core module
            messagebox.showwarning(
                "Launch Failed",
                f"Presentation was saved to {output_path}, "
                f"but could not be started: {error_msg}"
            )


def start_app():
    """Start the PowerPoint Merger application."""
    logging.info("Creating and running a new instance of PowerPointMergerApp.")
    app = PowerPointMergerApp()
    app.run()

