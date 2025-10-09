"""
GUI for the PowerPoint Merger application, built with PySide6.

This module defines the main window, UI components, and event handling,
including running the merge process in a background thread to keep the
UI responsive and handling drag-and-drop for files.
"""

import logging
import os

from PySide6.QtCore import QObject, QThread, Signal, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from app import PowerPointMerger


class DroppableListWidget(QListWidget):
    """
    A QListWidget subclass that handles drag-and-drop for .pptx files.
    """

    # Signal to emit when files are dropped
    filesDropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        """Accepts the drag event if it contains file URLs."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        """Accepts the move event if it contains file URLs."""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event):
        """Handles the drop event, filtering for .pptx files and emitting a signal."""
        logging.info("Drop event detected.")
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            files_to_add = []
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(".pptx"):
                    files_to_add.append(file_path)
                else:
                    logging.warning("Skipped non-pptx file: %s", file_path)

            if files_to_add:
                logging.info("Emitting filesDropped signal with %d files.", len(files_to_add))
                self.filesDropped.emit(files_to_add)
        else:
            super().dropEvent(event)


class Worker(QObject):
    """
    A worker object that performs the merging task in a separate thread.
    """

    finished = Signal()
    error = Signal(str)

    def __init__(self, files_to_merge, output_path):
        """
        Initializes the worker with the necessary data.

        Args:
            files_to_merge (list): A list of source file paths.
            output_path (str): The path for the output file.
        """
        super().__init__()
        self.files_to_merge = files_to_merge
        self.output_path = output_path

    def run(self):
        """
        Executes the merge operation in this thread. It creates its own
        instance of PowerPointMerger to ensure COM objects are thread-safe.
        """
        try:
            # Create a new merger instance within this thread
            thread_merger = PowerPointMerger()
            thread_merger.file_paths = self.files_to_merge
            thread_merger.merge(self.output_path)
            thread_merger.close()
            self.finished.emit()
        except Exception as e:
            logging.error("Worker thread error: %s", e, exc_info=True)
            self.error.emit(str(e))


class MainWindow(QWidget):
    """The main application window."""

    def __init__(self, merger):
        """
        Initializes the main window.

        Args:
            merger (PowerPointMerger): The application logic controller for UI state.
        """
        super().__init__()
        self.merger = merger
        self.thread = None
        self.worker = None

        self.setup_ui()

    def setup_ui(self):
        """Sets up the user interface of the main window."""
        self.setWindowTitle("PowerPoint Presentation Merger")
        icon_path = os.path.join(
            os.path.dirname(__file__), "resources", "MergePowerPoint.ico"
        )
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.resize(600, 400)

        # --- Widgets ---
        self.file_list_widget = DroppableListWidget()
        self.file_list_widget.setSelectionMode(QListWidget.SingleSelection)
        # REMOVED: self.file_list_widget.setDragDropMode(QListWidget.InternalMove)
        # This was preventing external drops. Internal reordering can be
        # re-enabled later if needed, but this is the fix for external drops.

        self.add_button = QPushButton("Add Files")
        self.remove_button = QPushButton("Remove Selected")
        self.up_button = QPushButton("Move Up")
        self.down_button = QPushButton("Move Down")

        self.output_path_edit = QLineEdit()
        self.output_path_edit.setPlaceholderText("Path for merged file...")
        self.browse_button = QPushButton("Browse...")

        self.merge_button = QPushButton("Merge Presentations")
        self.status_bar = QStatusBar()
        self.status_bar.showMessage("Ready")

        # --- Layouts ---
        main_layout = QVBoxLayout(self)
        file_buttons_layout = QHBoxLayout()
        output_layout = QHBoxLayout()

        file_buttons_layout.addWidget(self.add_button)
        file_buttons_layout.addWidget(self.remove_button)
        file_buttons_layout.addStretch()
        file_buttons_layout.addWidget(self.up_button)
        file_buttons_layout.addWidget(self.down_button)

        output_layout.addWidget(QLabel("Output File:"))
        output_layout.addWidget(self.output_path_edit)
        output_layout.addWidget(self.browse_button)

        main_layout.addWidget(QLabel("Presentations to Merge (in order):"))
        main_layout.addWidget(self.file_list_widget)
        main_layout.addLayout(file_buttons_layout)
        main_layout.addLayout(output_layout)
        main_layout.addWidget(self.merge_button)
        main_layout.addWidget(self.status_bar)

        # --- Connections ---
        self.add_button.clicked.connect(self.open_add_files_dialog)
        self.file_list_widget.filesDropped.connect(self.handle_added_files)
        self.remove_button.clicked.connect(self.remove_file)
        self.up_button.clicked.connect(self.move_file_up)
        self.down_button.clicked.connect(self.move_file_down)
        self.browse_button.clicked.connect(self.browse_output_file)
        self.merge_button.clicked.connect(self.merge_files)

    def open_add_files_dialog(self):
        """Opens a dialog to add presentation files."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PowerPoint Presentations",
            "",
            "PowerPoint Files (*.pptx);;All Files (*)",
        )
        if files:
            self.handle_added_files(files)

    def handle_added_files(self, files):
        """Adds a list of files to the merger and updates the UI."""
        self.merger.add_files(files)
        self.update_file_list()

    def remove_file(self):
        """Removes the selected file from the list."""
        current_item = self.file_list_widget.currentItem()
        if not current_item:
            return
        current_row = self.file_list_widget.row(current_item)
        self.merger.remove_file(current_row)
        self.update_file_list()

        # Restore selection to a valid item
        if self.file_list_widget.count() > 0:
            new_index = min(current_row, self.file_list_widget.count() - 1)
            self.file_list_widget.setCurrentRow(new_index)

    def move_file_up(self):
        """Moves the selected file up in the list."""
        current_row = self.file_list_widget.currentRow()
        if current_row > 0:
            self.merger.move_file_up(current_row)
            self.update_file_list()
            self.file_list_widget.setCurrentRow(current_row - 1)

    def move_file_down(self):
        """Moves the selected file down in the list."""
        current_row = self.file_list_widget.currentRow()
        if 0 <= current_row < self.file_list_widget.count() - 1:
            self.merger.move_file_down(current_row)
            self.update_file_list()
            self.file_list_widget.setCurrentRow(current_row + 1)

    def update_file_list(self):
        """Refreshes the file list widget from the merger's file list."""
        self.file_list_widget.clear()
        for file_path in self.merger.file_paths:
            self.file_list_widget.addItem(os.path.basename(file_path))

    def browse_output_file(self):
        """Opens a dialog to select the output file path."""
        output_path, _ = QFileDialog.getSaveFileName(
            self, "Save Merged Presentation As...", "", "PowerPoint Files (*.pptx)"
        )
        if output_path:
            self.output_path_edit.setText(output_path)

    def merge_files(self):
        """Starts the merge process in a separate thread."""
        output_path = self.output_path_edit.text()
        if not output_path:
            QMessageBox.warning(
                self, "Output Path Missing", "Please specify an output file path."
            )
            return

        if not self.merger.file_paths:
            QMessageBox.warning(
                self, "Not Enough Files", "Please add at least one presentation to merge."
            )
            return

        self.set_controls_enabled(False)
        self.status_bar.showMessage("Merging... please wait.")

        # Pass the data (not the object) to the worker
        self.thread = QThread()
        self.worker = Worker(self.merger.file_paths, output_path)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_merge_finished)
        self.worker.error.connect(self.on_merge_error)

        # Clean up the thread and worker when done
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(lambda: setattr(self, 'thread', None))

        self.thread.start()

    def on_merge_finished(self):
        """Handles the successful completion of the merge process."""
        self.status_bar.showMessage("Merge successful!", 5000)
        QMessageBox.information(self, "Success", "Presentations merged successfully!")
        self.set_controls_enabled(True)

    def on_merge_error(self, error_message):
        """Handles errors that occur during the merge process."""
        self.status_bar.showMessage("Error occurred during merge.", 5000)
        QMessageBox.critical(
            self, "Merge Error", f"An error occurred:\n{error_message}"
        )
        self.set_controls_enabled(True)

    def set_controls_enabled(self, enabled):
        """
        Enables or disables UI controls to prevent user interaction
        during a process.

        Args:
            enabled (bool): True to enable controls, False to disable.
        """
        self.file_list_widget.setEnabled(enabled)
        self.add_button.setEnabled(enabled)
        self.remove_button.setEnabled(enabled)
        self.up_button.setEnabled(enabled)
        self.down_button.setEnabled(enabled)
        self.output_path_edit.setEnabled(enabled)
        self.browse_button.setEnabled(enabled)
        self.merge_button.setEnabled(enabled)

    def closeEvent(self, event):
        """
        Handles the window close event, ensuring the worker thread is
        finished before closing.
        """
        logging.info("Main window close event triggered.")
        if self.thread and self.thread.isRunning():
            reply = QMessageBox.question(self, 'Confirm Close',
                                       "A merge process is still running. "
                                       "Are you sure you want to quit?",
                                       QMessageBox.Yes | QMessageBox.No,
                                       QMessageBox.No)

            if reply == QMessageBox.No:
                event.ignore()
                return

        logging.info("Closing application.")
        event.accept()

