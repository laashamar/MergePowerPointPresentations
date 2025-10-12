"""
Application module for PowerPoint Merger.

This module orchestrates the workflow, manages application state,
and connects GUI events with the backend merging logic.
"""
import logging
from tkinter import messagebox
import gui
import powerpoint_core


class PowerPointMergerApp:
    """Main application class that manages state and workflow."""

    def __init__(self):
        """Initialize the application state."""
        self.gui_instance = None
        logging.info("PowerPointMergerApp instance initialized.")

    def run(self):
        """Start the application with the modern GUI."""
        logging.info("Application workflow starting.")
        gui.show_modern_gui(self._on_merge_requested)

    def _on_merge_requested(self, file_list, output_path):
        """
        Handle merge request from GUI.

        Args:
            file_list: List of file paths in merge order
            output_path: Full path for the output file
        """
        logging.info(f"Merge requested for {len(file_list)} files")
        logging.info(f"Output path: {output_path}")

        # Merge presentations using COM automation
        success, saved_path, error_msg = powerpoint_core.merge_presentations(
            file_list,
            output_path
        )

        if not success:
            # Error already logged in core module
            messagebox.showerror(
                "Merge Failed",
                f"Failed to merge presentations: {error_msg}"
            )
            # Re-enable merge button if GUI instance is available
            return

        logging.info(f"Merge was successful. File saved to: {saved_path}")
        messagebox.showinfo(
            "Success",
            f"Presentation was merged successfully!\n\nSaved to: {saved_path}"
        )

        # Launch slideshow using COM automation
        logging.info(f"Starting slideshow for: {saved_path}")
        success, error_msg = powerpoint_core.launch_slideshow(saved_path)

        if not success:
            # Error already logged in core module
            messagebox.showwarning(
                "Launch Failed",
                f"Presentation was saved to {saved_path}, "
                f"but could not be started: {error_msg}"
            )


def start_app():
    """Start the PowerPoint Merger application."""
    logging.info("Creating and running a new instance of PowerPointMergerApp.")
    app = PowerPointMergerApp()
    app.run()
