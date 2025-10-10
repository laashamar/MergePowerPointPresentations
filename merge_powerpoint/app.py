import logging
from PySide6.QtCore import QObject, Slot
from PySide6.QtWidgets import QFileDialog

# MODIFIED: Use relative imports to refer to other modules in the package.
from .gui import MainWindow
from .powerpoint_core import PowerPointMerger

class AppController(QObject):
    """
    The main controller for the application.
    Connects the GUI (View) with the business logic (Model).
    """
    def __init__(self):
        super().__init__()
        self.files_to_merge = []
        self.merger = PowerPointMerger()  # Use PowerPointMerger instance
        self.main_window = MainWindow(self)  # Pass self (the controller) to the window

    def show_main_window(self):
        """Makes the main application window visible."""
        self.main_window.show()
        logging.info("Main window shown.")

    @Slot()
    def add_files(self):
        """Opens a file dialog to select PowerPoint files to add to the list."""
        logging.info("Add Files button clicked.")
        files, _ = QFileDialog.getOpenFileNames(
            self.main_window,
            "Select PowerPoint files to merge",
            "",
            "PowerPoint Files (*.pptx)"
        )
        if files:
            self.files_to_merge.extend(files)
            self.main_window.update_file_list(self.files_to_merge)
            logging.info(f"Added {len(files)} files. Total: {len(self.files_to_merge)}.")

    @Slot()
    def remove_selected_file(self):
        """Removes the currently selected file from the list."""
        selected_items = self.main_window.file_list.selectedItems()
        if not selected_items:
            logging.warning("Remove file clicked, but no file was selected.")
            return

        selected_file = selected_items[0].text()
        self.files_to_merge.remove(selected_file)
        self.main_window.update_file_list(self.files_to_merge)
        logging.info(f"Removed file: {selected_file}. Total: {len(self.files_to_merge)}.")

    @Slot()
    def merge_files(self):
        """
_x000D_
        Handles the logic for merging the selected PowerPoint files.
        """
        logging.info("Merge Files button clicked.")
        if len(self.files_to_merge) < 2:
            self.main_window.show_message("Error", "Please select at least two files to merge.")
            logging.error("Merge failed: Less than two files selected.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self.main_window,
            "Save Merged File As",
            "merged_presentation.pptx",
            "PowerPoint Files (*.pptx)"
        )

        if output_path:
            self.main_window.progress_bar.setVisible(True)
            self.main_window.progress_bar.setValue(0)
            try:
                # Use the PowerPointMerger instance
                self.merger.add_files(self.files_to_merge)
                self.merger.merge(output_path, self.main_window.update_progress)
                self.main_window.show_message("Success", f"Files merged successfully to:\n{output_path}")
                logging.info(f"Merge successful. Output: {output_path}")
            except Exception as e:
                self.main_window.show_message("Error", f"An error occurred during merge: {e}")
                logging.error(f"An exception occurred during merge: {e}", exc_info=True)
            finally:
                self.main_window.progress_bar.setVisible(False)

