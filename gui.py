"""
GUI for the PowerPoint Merger application using PySide6.
"""
import sys
from PySide6.QtCore import QSize, Qt, QUrl
from PySide6.QtGui import QIcon, QDesktopServices
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
    QListWidget, QHBoxLayout, QFileDialog, QMessageBox, QLabel,
    QLineEdit, QAbstractItemView
)

from app import AppController


class MainWindow(QMainWindow):
    """Main application window for the PowerPoint Merger."""

    def __init__(self, controller: AppController):
        """
        Initialize the main window.

        :param controller: The application controller instance.
        """
        super().__init__()
        self.merger = controller
        self.setWindowTitle("PowerPoint Presentation Merger")
        self.setMinimumSize(QSize(400, 300))
        self.setWindowIcon(QIcon("resources/MergePowerPoint.ico"))

        # --- Widgets ---
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)

        self.add_button = QPushButton("&Add Files")
        self.remove_button = QPushButton("Remove")
        self.move_up_button = QPushButton("Move Up")
        self.move_down_button = QPushButton("Move Down")
        self.merge_button = QPushButton("Merge")
        self.help_button = QPushButton("Help")

        self.output_path_line_edit = QLineEdit()
        self.output_path_line_edit.setPlaceholderText("Select output file path...")
        self.output_button = QPushButton("Browse")

        # --- Layouts ---
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)

        file_buttons_layout = QHBoxLayout()
        file_buttons_layout.addWidget(self.add_button)
        file_buttons_layout.addWidget(self.remove_button)
        file_buttons_layout.addStretch()
        file_buttons_layout.addWidget(self.move_up_button)
        file_buttons_layout.addWidget(self.move_down_button)

        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Output File:"))
        output_layout.addWidget(self.output_path_line_edit)
        output_layout.addWidget(self.output_button)

        action_buttons_layout = QHBoxLayout()
        action_buttons_layout.addWidget(self.help_button)
        action_buttons_layout.addStretch()
        action_buttons_layout.addWidget(self.merge_button)

        main_layout.addLayout(file_buttons_layout)
        main_layout.addWidget(self.list_widget)
        main_layout.addLayout(output_layout)
        main_layout.addLayout(action_buttons_layout)

        self.setCentralWidget(central_widget)

        # --- Connections ---
        self.add_button.clicked.connect(self.add_files)
        self.remove_button.clicked.connect(self.remove_files)
        self.move_up_button.clicked.connect(self.move_up)
        self.move_down_button.clicked.connect(self.move_down)
        self.merge_button.clicked.connect(self.merge_files)
        self.output_button.clicked.connect(self.select_output_file)
        self.help_button.clicked.connect(self.show_help)

    def update_file_list(self):
        """Updates the list widget with the current list of files."""
        self.list_widget.clear()
        self.list_widget.addItems(self.merger.get_files())

    def update_progress(self, current_step, total_steps):
        """
        Placeholder for updating a progress bar.
        This will be implemented with the status bar feature.
        """
        # TODO: Implement progress bar logic here
        pass

    def select_output_file(self):
        """Opens a dialog to select the output file path."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged File", "", "PowerPoint Presentations (*.pptx)"
        )
        if file_path:
            self.output_path_line_edit.setText(file_path)

    def add_files(self):
        """Opens a dialog to add files and updates the list."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Add PowerPoint Files", "", "PowerPoint Presentations (*.pptx *.ppt)"
        )
        if files:
            self.merger.add_files(files)
            self.update_file_list()

    def remove_files(self):
        """Removes the selected file from the list."""
        selected_items = self.list_widget.selectedItems()
        if not selected_items:
            return
        for item in selected_items:
            self.merger.remove_file(item.text())
        self.update_file_list()

    def move_up(self):
        """Moves the selected file up in the list."""
        current_row = self.list_widget.currentRow()
        if current_row > 0:
            self.merger.move_file_up(current_row)
            self.update_file_list()
            self.list_widget.setCurrentRow(current_row - 1)

    def move_down(self):
        """Moves the selected file down in the list."""
        current_row = self.list_widget.currentRow()
        if 0 <= current_row < self.list_widget.count() - 1:
            self.merger.move_file_down(current_row)
            self.update_file_list()
            self.list_widget.setCurrentRow(current_row + 1)

    def merge_files(self):
        """Initiates the file merging process."""
        output_path = self.output_path_line_edit.text()
        if not output_path:
            QMessageBox.warning(self, "Output Path Missing", "Please specify an output file path.")
            return

        if not self.merger.get_files():
            QMessageBox.warning(self, "No files", "Please add files to merge.")
            return

        try:
            self.merger.merge(output_path, self.update_progress)
            QMessageBox.information(self, "Success", "Presentations merged successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred during merge: {e}")

    def show_help(self):
        """Opens the project's README file in a web browser."""
        url = QUrl("https://github.com/laashamar/MergePowerPointPresentations/blob/main/README.md")
        QDesktopServices.openUrl(url)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    controller = AppController()
    window = MainWindow(controller)
    window.show()
    sys.exit(app.exec())

