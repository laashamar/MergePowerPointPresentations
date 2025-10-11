#!/usr/bin/env python3
"""
Example script demonstrating the refactored PySide6 GUI.

This standalone example shows how to:
1. Initialize the application with proper configuration
2. Create and inject the backend merger
3. Connect to UI signals for custom behavior
4. Launch the modern two-column interface

Run this file directly:
    python examples/example_refactored_gui.py
"""

import logging
import sys
from pathlib import Path

# Add src to path for development
sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from PySide6.QtWidgets import QApplication, QMessageBox

from merge_powerpoint.gui_refactored import UI_STRINGS, MainUI
from merge_powerpoint.powerpoint_core import PowerPointMerger

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def on_files_added(file_paths):
    """Custom handler for when files are added."""
    logger.info(f"Files added event: {len(file_paths)} files")
    for path in file_paths:
        logger.info(f"  - {Path(path).name}")


def on_file_removed(file_path):
    """Custom handler for when a file is removed."""
    logger.info(f"File removed event: {Path(file_path).name}")


def on_order_changed(new_order):
    """Custom handler for when the file order changes."""
    logger.info(f"Order changed event: {len(new_order)} files")
    for i, path in enumerate(new_order, 1):
        logger.info(f"  {i}. {Path(path).name}")


def on_clear_requested():
    """Custom handler for when the list is cleared."""
    logger.info("Clear requested event")


def on_merge_requested(output_path):
    """Custom handler for when merge is requested."""
    logger.info(f"Merge requested event: {output_path}")


def main():
    """Main entry point for the example application."""
    logger.info("Starting PowerPoint Merger with refactored GUI")

    # Create the Qt application
    app = QApplication(sys.argv)

    # Set application metadata (required for QSettings)
    app.setApplicationName("PowerPoint Merger")
    app.setOrganizationName("MergePowerPoint")
    app.setApplicationVersion("2.0.0")

    logger.info("Application metadata configured")

    # Create the backend merger
    try:
        merger = PowerPointMerger()
        logger.info("PowerPointMerger instance created")
    except Exception as e:
        logger.error(f"Failed to create PowerPointMerger: {e}")
        QMessageBox.critical(
            None,
            "Initialization Error",
            f"Failed to initialize PowerPoint merger:\n{e}"
        )
        return 1

    # Create the UI with dependency injection
    window = MainUI(merger=merger)
    window.setWindowTitle(UI_STRINGS["window_title"])
    window.resize(1000, 600)

    logger.info("MainUI instance created")

    # Connect to signals for custom behavior
    window.files_added.connect(on_files_added)
    window.file_removed.connect(on_file_removed)
    window.order_changed.connect(on_order_changed)
    window.clear_requested.connect(on_clear_requested)
    window.merge_requested.connect(on_merge_requested)

    logger.info("Signal handlers connected")

    # Show the window
    window.show()
    logger.info("Window displayed")

    # Display welcome message
    QMessageBox.information(
        window,
        "Welcome",
        "Welcome to PowerPoint Merger!\n\n"
        "Features:\n"
        "• Drag and drop .pptx files\n"
        "• Reorder files by dragging\n"
        "• Click 'Browse for Files...' to select files\n"
        "• Click 'Merge Presentations' when ready\n\n"
        "Check the console for event logging."
    )

    # Run the event loop
    logger.info("Entering Qt event loop")
    exit_code = app.exec()

    logger.info(f"Application exiting with code: {exit_code}")
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
