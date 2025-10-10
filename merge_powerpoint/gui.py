import logging
from PySide6.QtCore import Slot
from PySide6.QtWidgets import (
    QMainWindow, QVBoxLayout, QPushButton, QListWidget, QWidget,
    QLabel, QProgressBar, QMessageBox, QHBoxLayout
)

# MODIFIED: This class is no longer needed here as the logic is in powerpoint_core
# and the UI updates are handled by the MainWindow itself.
# from .powerpoint_core import PowerPointMerger

class MainWindow(QMainWindow):
    """
    The main window for the application (the View).
    It is responsible for displaying the UI and emitting signals to the controller.
    """
    def __init__(self, controller):
        super().__init__()
        self.controller = controller  # Store a reference to the controller

        self.setWindowTitle("PowerPoint Merger")
        self.setGeometry(100, 100, 500, 400)

        # Main layout and central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # --- UI Elements ---
        self.label = QLabel("Select PowerPoint files to merge:")
        self.file_list = QListWidget()
        self.file_list.setAlternatingRowColors(True)

        self.add_button = QPushButton("Add Files")
        self.remove_button = QPushButton("Remove Selected")
        self.merge_button = QPushButton("Merge Files")
        self.merge_button.setStyleSheet("font-weight: bold;")

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)

        # --- Layouts ---
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)

        main_layout.addWidget(self.label)
        main_layout.addWidget(self.file_list)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(self.merge_button)
        main_layout.addWidget(self.progress_bar)

        # --- Connections ---
        # Connect UI element signals to the controller's slots
        self.add_button.clicked.connect(self.controller.add_files)
        self.remove_button.clicked.connect(self.controller.remove_selected_file)
        self.merge_button.clicked.connect(self.controller.merge_files)

    @Slot(list)
    def update_file_list(self, files):
        """Clears and repopulates the file list widget."""
        self.file_list.clear()
        self.file_list.addItems(files)
        logging.debug("File list UI updated.")

    @Slot(int)
    def update_progress(self, value):
        """Updates the progress bar's value."""
        self.progress_bar.setValue(value)

    def show_message(self, title, message):
        """Displays a message box to the user."""
        QMessageBox.information(self, title, message)

