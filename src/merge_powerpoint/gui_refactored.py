"""Refactored modern GUI for PowerPoint Merger using PySide6.

This module implements a modern, two-column interface with drag-and-drop
support, custom item delegates, and signal-based architecture following
best practices for PySide6 applications.
"""

import logging
import os
from typing import List, Optional

from PySide6.QtCore import (
    QModelIndex,
    QSettings,
    Qt,
    QThread,
    Signal,
    Slot,
)
from PySide6.QtGui import (
    QDragEnterEvent,
    QDropEvent,
    QIcon,
    QPainter,
    QStandardItem,
    QStandardItemModel,
)
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStyledItemDelegate,
    QVBoxLayout,
    QWidget,
)

# Import compiled resources
try:
    from merge_powerpoint import icons_rc  # noqa: F401
except ImportError:
    icons_rc = None  # Icons won't be available if resources not compiled

from merge_powerpoint.powerpoint_core import PowerPointError, PowerPointMerger

logger = logging.getLogger(__name__)


# UI strings for internationalization
UI_STRINGS = {
    "window_title": "PowerPoint Presentation Merger",
    "drop_zone_text": "Drag and drop PowerPoint files here",
    "browse_button": "Browse for Files...",
    "clear_list_button": "Clear List",
    "output_group_title": "Output File",
    "output_filename_default": "merged-presentation.pptx",
    "save_to_button": "Save to...",
    "merge_button": "Merge Presentations",
    "remove_tooltip": "Remove this file",
    "not_enough_files_title": "Not Enough Files",
    "not_enough_files_message": "Please add at least two PowerPoint files to merge.",
    "invalid_file_title": "Invalid File",
    "invalid_file_message": "Only .pptx files are accepted.",
    "merge_success_title": "Success",
    "merge_success_message": "Presentation saved to {path}",
    "merge_error_title": "Merge Error",
    "file_not_readable_title": "File Error",
    "file_not_readable_message": "Cannot read file: {path}",
}


class FileListModel(QStandardItemModel):
    """Model for managing the list of PowerPoint files.

    This model provides a data structure for file paths with support for
    drag-and-drop reordering and duplicate prevention.
    """

    def __init__(self, parent=None):
        """Initialize the file list model.

        Args:
            parent: Optional parent QObject.
        """
        super().__init__(parent)
        self.file_paths: List[str] = []

    def add_files(self, paths: List[str]) -> List[str]:
        """Add files to the model, preventing duplicates.

        Args:
            paths: List of file paths to add.

        Returns:
            List of paths that were rejected (duplicates or invalid).
        """
        rejected = []
        for path in paths:
            if path in self.file_paths:
                rejected.append(path)
                logger.debug(f"Rejected duplicate file: {path}")
            elif not path.lower().endswith('.pptx'):
                rejected.append(path)
                logger.debug(f"Rejected non-pptx file: {path}")
            else:
                self.file_paths.append(path)
                item = QStandardItem(path)
                item.setData(path, Qt.UserRole)
                self.appendRow(item)
                logger.info(f"Added file: {path}")
        return rejected

    def remove_file(self, path: str) -> bool:
        """Remove a file from the model.

        Args:
            path: File path to remove.

        Returns:
            True if file was removed, False if not found.
        """
        if path in self.file_paths:
            idx = self.file_paths.index(path)
            self.removeRow(idx)
            self.file_paths.remove(path)
            logger.info(f"Removed file: {path}")
            return True
        return False

    def clear_all(self):
        """Remove all files from the model."""
        self.clear()
        self.file_paths = []
        logger.info("Cleared all files")

    def get_file_paths(self) -> List[str]:
        """Get the current ordered list of file paths.

        Returns:
            List of file paths in current order.
        """
        return self.file_paths.copy()

    def reorder_files(self, new_order: List[str]):
        """Update the file order based on a new list.

        Args:
            new_order: New ordered list of file paths.
        """
        self.file_paths = new_order.copy()
        self.clear()
        for path in new_order:
            item = QStandardItem(path)
            item.setData(path, Qt.UserRole)
            self.appendRow(item)
        logger.info(f"Reordered files: {len(new_order)} items")

    def supportedDragActions(self):
        """Support move operations for drag and drop."""
        return Qt.MoveAction

    def supportedDropActions(self):
        """Support move operations for drag and drop."""
        return Qt.MoveAction


class FileItemDelegate(QStyledItemDelegate):
    """Custom delegate for rendering file items as cards.

    Each file item is rendered with an icon, filename, and remove button.
    """

    remove_clicked = Signal(str)  # Emits the file path

    def __init__(self, parent=None):
        """Initialize the delegate.

        Args:
            parent: Optional parent QObject.
        """
        super().__init__(parent)
        self._icon = QIcon(":/icons/powerpoint.svg")

    def paint(self, painter: QPainter, option, index: QModelIndex):
        """Paint the file item as a card.

        Args:
            painter: QPainter to use for drawing.
            option: Style options for the item.
            index: Model index of the item.
        """
        super().paint(painter, option, index)
        # The base implementation will handle the text
        # In a full implementation, we would custom-draw the entire card here

    def sizeHint(self, option, index: QModelIndex):
        """Return the size hint for the item.

        Args:
            option: Style options for the item.
            index: Model index of the item.

        Returns:
            QSize for the item.
        """
        size = super().sizeHint(option, index)
        size.setHeight(max(size.height(), 48))  # Minimum height for card
        return size


class DropZoneWidget(QFrame):
    """Widget displayed when no files are in the list.

    Shows a centered icon and text encouraging users to drag and drop files.
    """

    files_dropped = Signal(list)  # Emits list of file paths

    def __init__(self, parent=None):
        """Initialize the drop zone widget.

        Args:
            parent: Optional parent widget.
        """
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFrameShape(QFrame.StyledPanel)
        self.setStyleSheet("""
            DropZoneWidget {
                background-color: #f5f5f5;
                border: 2px dashed #cccccc;
                border-radius: 8px;
            }
            DropZoneWidget:hover {
                border-color: #999999;
                background-color: #eeeeee;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        # Plus icon
        icon_label = QLabel()
        icon_label.setPixmap(
            QIcon(":/icons/plus.svg").pixmap(64, 64)
        )
        icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(icon_label)

        # Text
        text_label = QLabel(UI_STRINGS["drop_zone_text"])
        text_label.setAlignment(Qt.AlignCenter)
        font = text_label.font()
        font.setPointSize(14)
        text_label.setFont(font)
        text_label.setStyleSheet("color: #666666;")
        layout.addWidget(text_label)

        # Browse button
        browse_button = QPushButton(UI_STRINGS["browse_button"])
        browse_button.clicked.connect(self._browse_files)
        browse_button.setMinimumWidth(150)
        layout.addWidget(browse_button, alignment=Qt.AlignCenter)

    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter events.

        Args:
            event: Drag enter event.
        """
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        """Handle drop events.

        Args:
            event: Drop event containing file URLs.
        """
        if event.mimeData().hasUrls():
            file_paths = []
            for url in event.mimeData().urls():
                if url.isLocalFile():
                    path = url.toLocalFile()
                    if path.lower().endswith('.pptx'):
                        file_paths.append(path)
            if file_paths:
                self.files_dropped.emit(file_paths)
            event.acceptProposedAction()
        else:
            event.ignore()

    def _browse_files(self):
        """Open file dialog to browse for PowerPoint files."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select PowerPoint Files",
            "",
            "PowerPoint Presentations (*.pptx)",
        )
        if files:
            self.files_dropped.emit(files)


class MergeWorker(QThread):
    """Worker thread for performing merge operations.

    Signals:
        progress: Emitted with (current, total) during merge.
        finished: Emitted with (success, output_path, error_message).
    """

    progress = Signal(int, int)
    finished = Signal(bool, str, str)

    def __init__(self, file_paths: List[str], output_path: str, merger: PowerPointMerger):
        """Initialize the merge worker.

        Args:
            file_paths: List of file paths to merge.
            output_path: Output file path.
            merger: PowerPointMerger instance to use.
        """
        super().__init__()
        self.file_paths = file_paths
        self.output_path = output_path
        self.merger = merger

    def run(self):
        """Execute the merge operation in the worker thread."""
        try:
            logger.info(f"Starting merge of {len(self.file_paths)} files")

            def progress_callback(current, total):
                self.progress.emit(current, total)

            # Clear and add files to merger
            self.merger.clear_files()
            self.merger.add_files(self.file_paths)

            # Perform merge
            self.merger.merge(self.output_path, progress_callback)

            self.finished.emit(True, self.output_path, "")
            logger.info(f"Merge completed successfully: {self.output_path}")
        except PowerPointError as e:
            logger.error(f"Merge failed: {e}", exc_info=True)
            self.finished.emit(False, "", str(e))
        except Exception as e:
            logger.error(f"Unexpected error during merge: {e}", exc_info=True)
            self.finished.emit(False, "", str(e))


class MainUI(QWidget):
    """Main user interface widget for PowerPoint merger.

    This widget implements a two-column layout with file management on the left
    and configuration/actions on the right.

    Signals:
        files_added: Emitted when files are added (list[str] of paths).
        file_removed: Emitted when a file is removed (str path).
        order_changed: Emitted when file order changes (list[str] of paths).
        clear_requested: Emitted when clear all is requested.
        merge_requested: Emitted when merge is requested (str output path).
    """

    files_added = Signal(list)
    file_removed = Signal(str)
    order_changed = Signal(list)
    clear_requested = Signal()
    merge_requested = Signal(str)

    def __init__(self, merger: Optional[PowerPointMerger] = None, parent=None):
        """Initialize the main UI.

        Args:
            merger: Optional PowerPointMerger instance (injected dependency).
            parent: Optional parent widget.
        """
        super().__init__(parent)
        self.merger = merger or PowerPointMerger()
        self.file_model = FileListModel(self)
        self.settings = QSettings("MergePowerPoint", "Merger")
        self.merge_worker: Optional[MergeWorker] = None

        self._setup_ui()
        self._connect_signals()
        self._restore_settings()

    def _setup_ui(self):
        """Set up the user interface components."""
        main_layout = QHBoxLayout(self)
        main_layout.setSpacing(16)
        main_layout.setContentsMargins(16, 16, 16, 16)

        # Left column: File list (3 parts in stretch)
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)

        # Drop zone (shown when empty)
        self.drop_zone = DropZoneWidget()
        self.drop_zone.files_dropped.connect(self._on_files_dropped)
        left_layout.addWidget(self.drop_zone)

        # File list view (shown when not empty)
        self.file_list_view = QListView()
        self.file_list_view.setModel(self.file_model)
        self.file_list_view.setItemDelegate(FileItemDelegate())
        self.file_list_view.setDragEnabled(True)
        self.file_list_view.setAcceptDrops(True)
        self.file_list_view.setDropIndicatorShown(True)
        self.file_list_view.setDragDropMode(QListView.InternalMove)
        self.file_list_view.setSelectionMode(QListView.SingleSelection)
        self.file_list_view.setVisible(False)
        left_layout.addWidget(self.file_list_view)

        main_layout.addWidget(left_widget, 3)

        # Right column: Configuration and actions (1 part in stretch)
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)

        # Clear list button
        self.clear_button = QPushButton(UI_STRINGS["clear_list_button"])
        self.clear_button.setIcon(QIcon(":/icons/trash.svg"))
        self.clear_button.clicked.connect(self._on_clear_clicked)
        self.clear_button.setEnabled(False)
        right_layout.addWidget(self.clear_button)

        right_layout.addSpacing(16)

        # Output file configuration
        output_group = QGroupBox(UI_STRINGS["output_group_title"])
        output_layout = QVBoxLayout(output_group)

        self.output_filename_edit = QLineEdit()
        self.output_filename_edit.setText(UI_STRINGS["output_filename_default"])
        self.output_filename_edit.setPlaceholderText("merged-presentation.pptx")
        output_layout.addWidget(self.output_filename_edit)

        self.save_to_button = QPushButton(UI_STRINGS["save_to_button"])
        self.save_to_button.setIcon(QIcon(":/icons/folder.svg"))
        self.save_to_button.clicked.connect(self._on_save_to_clicked)
        output_layout.addWidget(self.save_to_button)

        right_layout.addWidget(output_group)

        right_layout.addStretch()

        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        right_layout.addWidget(self.progress_bar)

        # Merge button
        self.merge_button = QPushButton(UI_STRINGS["merge_button"])
        self.merge_button.setMinimumHeight(48)
        self.merge_button.clicked.connect(self._on_merge_clicked)
        self.merge_button.setEnabled(False)
        font = self.merge_button.font()
        font.setPointSize(12)
        font.setBold(True)
        self.merge_button.setFont(font)
        right_layout.addWidget(self.merge_button)

        main_layout.addWidget(right_widget, 1)

        # Set initial state
        self._update_ui_state()

    def _connect_signals(self):
        """Connect internal signals to slots."""
        self.file_model.rowsInserted.connect(self._on_model_changed)
        self.file_model.rowsRemoved.connect(self._on_model_changed)
        self.file_model.modelReset.connect(self._on_model_changed)

    def _restore_settings(self):
        """Restore saved settings from previous session."""
        last_save_dir = self.settings.value("last_save_dir", "")
        if last_save_dir and os.path.isdir(last_save_dir):
            self.last_save_dir = last_save_dir
        else:
            self.last_save_dir = os.path.expanduser("~")

    def _save_settings(self):
        """Save current settings for next session."""
        if hasattr(self, 'last_save_dir'):
            self.settings.setValue("last_save_dir", self.last_save_dir)

    def _update_ui_state(self):
        """Update UI state based on file list."""
        has_files = len(self.file_model.file_paths) > 0
        can_merge = len(self.file_model.file_paths) >= 2

        # Toggle drop zone vs file list
        self.drop_zone.setVisible(not has_files)
        self.file_list_view.setVisible(has_files)

        # Update button states
        self.clear_button.setEnabled(has_files)
        self.merge_button.setEnabled(can_merge)

    def _on_files_dropped(self, file_paths: List[str]):
        """Handle files dropped onto the drop zone.

        Args:
            file_paths: List of file paths that were dropped.
        """
        # Validate files exist and are readable
        valid_files = []
        for path in file_paths:
            if not os.path.exists(path):
                logger.warning(f"File does not exist: {path}")
                continue
            if not os.path.isfile(path):
                logger.warning(f"Not a file: {path}")
                continue
            # Skip the read check for now to avoid blocking
            valid_files.append(path)

        if valid_files:
            rejected = self.file_model.add_files(valid_files)
            if rejected:
                # Show brief status about rejected files
                logger.info(f"Rejected {len(rejected)} files (duplicates or invalid)")
            self.files_added.emit(valid_files)

    @Slot()
    def _on_clear_clicked(self):
        """Handle clear button click."""
        self.file_model.clear_all()
        self.clear_requested.emit()

    @Slot()
    def _on_save_to_clicked(self):
        """Handle save to button click."""
        initial_dir = self.last_save_dir
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Merged Presentation",
            os.path.join(initial_dir, self.output_filename_edit.text()),
            "PowerPoint Presentations (*.pptx)",
        )
        if file_path:
            # Update the filename field and save directory
            self.output_filename_edit.setText(os.path.basename(file_path))
            self.last_save_dir = os.path.dirname(file_path)
            self._save_settings()

    @Slot()
    def _on_merge_clicked(self):
        """Handle merge button click."""
        file_paths = self.file_model.get_file_paths()
        if len(file_paths) < 2:
            QMessageBox.warning(
                self,
                UI_STRINGS["not_enough_files_title"],
                UI_STRINGS["not_enough_files_message"]
            )
            return

        # Get output path
        filename = self.output_filename_edit.text().strip()
        if not filename:
            filename = UI_STRINGS["output_filename_default"]

        # Ensure .pptx extension
        if not filename.lower().endswith('.pptx'):
            filename += '.pptx'

        output_path = os.path.join(self.last_save_dir, filename)

        # Confirm overwrite if file exists
        if os.path.exists(output_path):
            reply = QMessageBox.question(
                self,
                "Confirm Overwrite",
                f"File already exists:\n{output_path}\n\nOverwrite?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        # Start merge in worker thread
        self._start_merge(file_paths, output_path)

    def _start_merge(self, file_paths: List[str], output_path: str):
        """Start the merge operation in a worker thread.

        Args:
            file_paths: List of file paths to merge.
            output_path: Path where merged file will be saved.
        """
        # Disable UI during merge
        self._set_ui_enabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        # Create and start worker
        self.merge_worker = MergeWorker(file_paths, output_path, self.merger)
        self.merge_worker.progress.connect(self._on_merge_progress)
        self.merge_worker.finished.connect(self._on_merge_finished)
        self.merge_worker.start()

        self.merge_requested.emit(output_path)

    @Slot(int, int)
    def _on_merge_progress(self, current: int, total: int):
        """Handle merge progress updates.

        Args:
            current: Current progress value.
            total: Total number of items.
        """
        if total > 0:
            percentage = int((current / total) * 100)
            self.progress_bar.setValue(percentage)

    @Slot(bool, str, str)
    def _on_merge_finished(self, success: bool, output_path: str, error_message: str):
        """Handle merge completion.

        Args:
            success: Whether merge was successful.
            output_path: Path to output file.
            error_message: Error message if unsuccessful.
        """
        # Re-enable UI
        self._set_ui_enabled(True)
        self.progress_bar.setVisible(False)

        if success:
            # Show success message with link to open folder
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle(UI_STRINGS["merge_success_title"])
            msg.setText(UI_STRINGS["merge_success_message"].format(path=output_path))
            msg.setStandardButtons(QMessageBox.Ok)

            # Add button to open folder
            open_folder_button = msg.addButton("Open Folder", QMessageBox.ActionRole)
            msg.exec()

            if msg.clickedButton() == open_folder_button:
                self._open_folder(output_path)
        else:
            QMessageBox.critical(
                self,
                UI_STRINGS["merge_error_title"],
                f"Merge failed:\n{error_message}"
            )

        self.merge_worker = None

    def _open_folder(self, file_path: str):
        """Open the folder containing the specified file.

        Args:
            file_path: Path to the file.
        """
        import platform
        import subprocess

        folder = os.path.dirname(os.path.abspath(file_path))

        try:
            if platform.system() == "Windows":
                subprocess.run(["explorer", "/select,", file_path])
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", "-R", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder])
        except Exception as e:
            logger.warning(f"Failed to open folder: {e}")

    def _set_ui_enabled(self, enabled: bool):
        """Enable or disable UI controls.

        Args:
            enabled: Whether to enable controls.
        """
        self.clear_button.setEnabled(enabled and len(self.file_model.file_paths) > 0)
        self.merge_button.setEnabled(enabled and len(self.file_model.file_paths) >= 2)
        self.save_to_button.setEnabled(enabled)
        self.output_filename_edit.setEnabled(enabled)
        self.file_list_view.setEnabled(enabled)
        self.drop_zone.setEnabled(enabled)

    @Slot()
    def _on_model_changed(self):
        """Handle changes to the file model."""
        self._update_ui_state()
        if self.file_model.rowCount() > 0:
            self.order_changed.emit(self.file_model.get_file_paths())


if __name__ == "__main__":
    import sys

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    app = QApplication(sys.argv)
    app.setApplicationName("PowerPoint Merger")
    app.setOrganizationName("MergePowerPoint")

    window = MainUI()
    window.setWindowTitle(UI_STRINGS["window_title"])
    window.resize(1000, 600)
    window.show()

    sys.exit(app.exec())
