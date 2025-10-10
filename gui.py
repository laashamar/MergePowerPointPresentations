"""
This module defines the main graphical user interface (GUI) for the PowerPoint Presentation Merger application.
It uses the PySide6 framework to create a window where users can add, manage, and merge PowerPoint files.

The MainWindow class is the central component of the GUI, providing the following features:
- A list view to display the PowerPoint files selected for merging.
- Buttons to add, remove, clear, and reorder the files.
- A "Merge Files" button to initiate the merging process.
- A progress bar to provide feedback during the merge operation.
- Integration with the PowerPointMerger class from powerpoint_core.py to handle the backend logic.

The GUI is designed to be intuitive and user-friendly, guiding the user through the process of merging presentations.
"""

import sys
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, QWidget,
    QPushButton, QListWidget, QFileDialog, QMessageBox, QProgressBar,
    QListWidgetItem
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from powerpoint_core import PowerPointMerger
from app_logger import setup_logging  # Corrected import

# Set up logging
logger = setup_logging()  # Corrected function call

class MainWindow(QMainWindow):
    """Main window for the PowerPoint Merger application."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PowerPoint Presentation Merger")
        self.setGeometry(100, 100, 800, 600)
        
        # Set the application icon
        # This path is relative to the execution directory
        icon_path = Path("resources/MergePowerPoint.ico")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.merger = PowerPointMerger()
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self._setup_ui()

    def _setup_ui(self):
        """Set up the UI components."""
        # File list display
        self.file_list_widget = QListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.ExtendedSelection)
        self.layout.addWidget(self.file_list_widget)

        # Button layout
        button_layout = QHBoxLayout()

        self.add_button = QPushButton("&Add Files")
        self.add_button.clicked.connect(self.add_files)
        button_layout.addWidget(self.add_button)

        self.remove_button = QPushButton("&Remove Selected")
        self.remove_button.clicked.connect(self.remove_selected_files)
        button_layout.addWidget(self.remove_button)
        
        self.clear_button = QPushButton("&Clear All")
        self.clear_button.clicked.connect(self.clear_all_files)
        button_layout.addWidget(self.clear_button)

        # Spacer to push reordering buttons to the right
        button_layout.addStretch()
        
        self.move_up_button = QPushButton("Move &Up")
        self.move_up_button.clicked.connect(self.move_file_up)
        button_layout.addWidget(self.move_up_button)

        self.move_down_button = QPushButton("Move &Down")
        self.move_down_button.clicked.connect(self.move_file_down)
        button_layout.addWidget(self.move_down_button)

        self.layout.addLayout(button_layout)

        # Merge button and progress bar
        self.merge_button = QPushButton("&Merge Files")
        self.merge_button.clicked.connect(self.merge_files)
        self.layout.addWidget(self.merge_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.layout.addWidget(self.progress_bar)
        
        self.update_button_states()

    def add_files(self):
        """Open a dialog to add PowerPoint files."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PowerPoint Files", "", "PowerPoint Presentations (*.pptx)"
        )
        if files:
            self.merger.add_files(files)
            self.update_file_list()
            logger.info("Added files: %s", files)

    def remove_selected_files(self):
        """Remove the selected files from the list."""
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items:
            return
        
        files_to_remove = [item.text() for item in selected_items]
        self.merger.remove_files(files_to_remove)
        self.update_file_list()
        logger.info("Removed files: %s", files_to_remove)

    def clear_all_files(self):
        """Clear all files from the list."""
        self.merger.clear_files()
        self.update_file_list()
        logger.info("Cleared all files.")

    def update_file_list(self):
        """Update the file list widget with the current list of files."""
        self.file_list_widget.clear()
        for file in self.merger.get_files():
            self.file_list_widget.addItem(QListWidgetItem(file))
        self.update_button_states()

    def merge_files(self):
        """Initiate the file merging process."""
        if len(self.merger.get_files()) < 2:
            QMessageBox.warning(self, "Not enough files", "Please add at least two files to merge.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged File", "", "PowerPoint Presentation (*.pptx)"
        )
        if output_path:
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            try:
                self.merger.merge(output_path, self.update_progress)
                QMessageBox.information(self, "Success", f"Successfully merged files to {output_path}")
                logger.info("Successfully merged files to %s", output_path)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"An error occurred during merging: {e}")
                logger.error("Merging failed: %s", e, exc_info=True)
            finally:
                self.progress_bar.setVisible(False)
                
    def move_file_up(self):
        """Move the selected file up in the list."""
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items or len(selected_items) > 1:
            return # Only move one item at a time
        
        current_index = self.file_list_widget.row(selected_items[0])
        if self.merger.move_file_up(current_index):
            self.update_file_list()
            self.file_list_widget.setCurrentRow(current_index - 1)
            logger.info("Moved file up: %s", selected_items[0].text())
            
    def move_file_down(self):
        """Move the selected file down in the list."""
        selected_items = self.file_list_widget.selectedItems()
        if not selected_items or len(selected_items) > 1:
            return
            
        current_index = self.file_list_widget.row(selected_items[0])
        if self.merger.move_file_down(current_index):
            self.update_file_list()
            self.file_list_widget.setCurrentRow(current_index + 1)
            logger.info("Moved file down: %s", selected_items[0].text())

    def update_progress(self, value):
        """Update the progress bar."""
        self.progress_bar.setValue(value)
        QApplication.processEvents()
        
    def update_button_states(self):
        """Enable or disable buttons based on the application state."""
        has_files = len(self.merger.get_files()) > 0
        has_selection = len(self.file_list_widget.selectedItems()) > 0
        can_merge = len(self.merger.get_files()) >= 2
        
        self.remove_button.setEnabled(has_files and has_selection)
        self.clear_button.setEnabled(has_files)
        self.merge_button.setEnabled(can_merge)
        
        # Logic for move up/down buttons
        can_move_up = False
        can_move_down = False
        if has_selection and len(self.file_list_widget.selectedItems()) == 1:
            selected_index = self.file_list_widget.currentRow()
            if selected_index > 0:
                can_move_up = True
            if selected_index < self.file_list_widget.count() - 1:
                can_move_down = True
        
        self.move_up_button.setEnabled(can_move_up)
        self.move_down_button.setEnabled(can_move_down)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

