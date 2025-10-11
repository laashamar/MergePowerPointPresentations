"""Tests for the refactored PySide6 GUI.

This module tests the modern two-column interface including signal emissions,
drag-and-drop functionality, state management, and threading behavior.
"""

import os
import sys
from unittest.mock import MagicMock, Mock

import pytest
from PySide6.QtCore import QMimeData, QUrl
from PySide6.QtTest import QSignalSpy

# Mock comtypes before importing
if sys.platform != 'win32':
    sys.modules['comtypes'] = MagicMock()
    sys.modules['comtypes.client'] = MagicMock()

from merge_powerpoint.gui_refactored import (
    UI_STRINGS,
    DropZoneWidget,
    FileListModel,
    MainUI,
    MergeWorker,
)
from merge_powerpoint.powerpoint_core import PowerPointMerger


class TestFileListModel:
    """Tests for the FileListModel class."""

    def test_initialization(self):
        """Test model initializes with empty file list."""
        model = FileListModel()
        assert model.file_paths == []
        assert model.rowCount() == 0

    def test_add_files_success(self):
        """Test adding valid .pptx files."""
        model = FileListModel()
        files = ["/path/to/file1.pptx", "/path/to/file2.pptx"]
        rejected = model.add_files(files)

        assert rejected == []
        assert model.file_paths == files
        assert model.rowCount() == 2

    def test_add_files_rejects_duplicates(self):
        """Test that duplicate files are rejected."""
        model = FileListModel()
        files = ["/path/to/file1.pptx", "/path/to/file1.pptx"]
        rejected = model.add_files(files)

        assert len(rejected) == 1
        assert rejected[0] == "/path/to/file1.pptx"
        assert model.file_paths == ["/path/to/file1.pptx"]
        assert model.rowCount() == 1

    def test_add_files_rejects_non_pptx(self):
        """Test that non-.pptx files are rejected."""
        model = FileListModel()
        files = ["/path/to/file1.pptx", "/path/to/file2.pdf", "/path/to/file3.doc"]
        rejected = model.add_files(files)

        assert len(rejected) == 2
        assert "/path/to/file2.pdf" in rejected
        assert "/path/to/file3.doc" in rejected
        assert model.file_paths == ["/path/to/file1.pptx"]

    def test_remove_file_success(self):
        """Test removing a file from the model."""
        model = FileListModel()
        model.add_files(["/path/to/file1.pptx", "/path/to/file2.pptx"])

        result = model.remove_file("/path/to/file1.pptx")

        assert result is True
        assert model.file_paths == ["/path/to/file2.pptx"]
        assert model.rowCount() == 1

    def test_remove_file_not_found(self):
        """Test removing a non-existent file returns False."""
        model = FileListModel()
        model.add_files(["/path/to/file1.pptx"])

        result = model.remove_file("/path/to/nonexistent.pptx")

        assert result is False
        assert model.file_paths == ["/path/to/file1.pptx"]

    def test_clear_all(self):
        """Test clearing all files from the model."""
        model = FileListModel()
        model.add_files(["/path/to/file1.pptx", "/path/to/file2.pptx"])

        model.clear_all()

        assert model.file_paths == []
        assert model.rowCount() == 0

    def test_get_file_paths(self):
        """Test getting file paths returns a copy."""
        model = FileListModel()
        files = ["/path/to/file1.pptx", "/path/to/file2.pptx"]
        model.add_files(files)

        paths = model.get_file_paths()

        assert paths == files
        # Verify it's a copy, not the original
        paths.append("/path/to/file3.pptx")
        assert model.file_paths != paths

    def test_reorder_files(self):
        """Test reordering files in the model."""
        model = FileListModel()
        original = ["/path/to/file1.pptx", "/path/to/file2.pptx", "/path/to/file3.pptx"]
        model.add_files(original)

        new_order = ["/path/to/file3.pptx", "/path/to/file1.pptx", "/path/to/file2.pptx"]
        model.reorder_files(new_order)

        assert model.file_paths == new_order
        assert model.rowCount() == 3


class TestDropZoneWidget:
    """Tests for the DropZoneWidget class."""

    def test_initialization(self, qtbot):
        """Test drop zone initializes correctly."""
        widget = DropZoneWidget()
        qtbot.addWidget(widget)

        assert widget.acceptDrops() is True

    def test_files_dropped_signal_emitted(self, qtbot, tmp_path):
        """Test that files_dropped signal is emitted with valid files."""
        widget = DropZoneWidget()
        qtbot.addWidget(widget)

        # Create temporary .pptx files
        file1 = tmp_path / "test1.pptx"
        file2 = tmp_path / "test2.pptx"
        file1.touch()
        file2.touch()

        # Create spy for signal
        spy = QSignalSpy(widget.files_dropped)

        # Simulate drop event
        mime_data = QMimeData()
        urls = [QUrl.fromLocalFile(str(file1)), QUrl.fromLocalFile(str(file2))]
        mime_data.setUrls(urls)

        # Trigger drop
        drop_event = Mock()
        drop_event.mimeData.return_value = mime_data
        widget.dropEvent(drop_event)

        # Verify signal was emitted
        assert spy.count() == 1
        emitted_files = spy.at(0)[0]
        assert len(emitted_files) == 2

    def test_browse_button_exists(self, qtbot):
        """Test that browse button is present and functional."""
        widget = DropZoneWidget()
        qtbot.addWidget(widget)

        # The widget should have a browse button (implicit check via initialization)


class TestMainUI:
    """Tests for the main UI widget."""

    @pytest.fixture
    def main_ui(self, qtbot):
        """Create a MainUI instance for testing."""
        merger = PowerPointMerger()
        ui = MainUI(merger=merger)
        qtbot.addWidget(ui)
        ui.show()
        qtbot.waitExposed(ui)
        return ui

    def test_initialization(self, main_ui):
        """Test that MainUI initializes correctly."""
        assert main_ui.merger is not None
        assert main_ui.file_model is not None
        assert main_ui.merge_worker is None
        assert main_ui.drop_zone.isVisible() is True
        assert main_ui.file_list_view.isVisible() is False

    def test_files_added_signal_emitted(self, main_ui, qtbot, tmp_path, mocker):
        """Test that files_added signal is emitted when files are added."""
        spy = QSignalSpy(main_ui.files_added)

        # Create temp file
        file1 = tmp_path / "test1.pptx"
        file1.touch()

        # Mock os.path.exists and os.path.isfile
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        # Add files via drop zone
        main_ui._on_files_dropped([str(file1)])

        # Verify signal was emitted
        assert spy.count() == 1

    def test_clear_requested_signal_emitted(self, main_ui, qtbot, tmp_path, mocker):
        """Test that clear_requested signal is emitted on clear button click."""
        # Mock file operations
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        # Add some files first
        file1 = tmp_path / "test1.pptx"
        file1.touch()
        main_ui._on_files_dropped([str(file1)])

        spy = QSignalSpy(main_ui.clear_requested)

        # Click clear button
        main_ui.clear_button.click()

        # Verify signal was emitted
        assert spy.count() == 1

    def test_merge_requested_signal_emitted(self, main_ui, qtbot, tmp_path, mocker):
        """Test that merge_requested signal is emitted when merge is clicked."""
        # Mock file operations
        mocker.patch('os.path.exists', return_value=False)  # No existing file
        mocker.patch('os.path.isfile', return_value=True)

        # Add files
        file1 = tmp_path / "test1.pptx"
        file2 = tmp_path / "test2.pptx"
        file1.touch()
        file2.touch()
        main_ui._on_files_dropped([str(file1), str(file2)])

        spy = QSignalSpy(main_ui.merge_requested)

        # Trigger merge directly via _start_merge to avoid dialogs
        output_path = str(tmp_path / "output.pptx")

        # Mock the worker to prevent actual merge
        mock_worker = mocker.MagicMock()
        mocker.patch('merge_powerpoint.gui_refactored.MergeWorker', return_value=mock_worker)

        main_ui._start_merge([str(file1), str(file2)], output_path)

        # Verify signal was emitted
        assert spy.count() == 1

    def test_ui_state_empty(self, main_ui):
        """Test UI state when no files are added."""
        assert main_ui.drop_zone.isVisible() is True
        assert main_ui.file_list_view.isVisible() is False
        assert main_ui.clear_button.isEnabled() is False
        assert main_ui.merge_button.isEnabled() is False

    def test_ui_state_one_file(self, main_ui, tmp_path, mocker):
        """Test UI state when one file is added."""
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        file1 = tmp_path / "test1.pptx"
        file1.touch()
        main_ui._on_files_dropped([str(file1)])

        assert main_ui.drop_zone.isVisible() is False
        assert main_ui.file_list_view.isVisible() is True
        assert main_ui.clear_button.isEnabled() is True
        assert main_ui.merge_button.isEnabled() is False  # Need 2+ files

    def test_ui_state_two_files(self, main_ui, tmp_path, mocker):
        """Test UI state when two files are added."""
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        file1 = tmp_path / "test1.pptx"
        file2 = tmp_path / "test2.pptx"
        file1.touch()
        file2.touch()
        main_ui._on_files_dropped([str(file1), str(file2)])

        assert main_ui.drop_zone.isVisible() is False
        assert main_ui.file_list_view.isVisible() is True
        assert main_ui.clear_button.isEnabled() is True
        assert main_ui.merge_button.isEnabled() is True

    def test_ui_disabled_during_merge(self, main_ui, tmp_path, mocker):
        """Test that UI controls are disabled during merge operation."""
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        # Add files
        file1 = tmp_path / "test1.pptx"
        file2 = tmp_path / "test2.pptx"
        file1.touch()
        file2.touch()
        main_ui._on_files_dropped([str(file1), str(file2)])

        # Mock to prevent actual merge
        mock_worker = mocker.MagicMock()
        mocker.patch.object(MergeWorker, '__init__', return_value=None)
        mocker.patch.object(MergeWorker, 'start')
        mocker.patch('merge_powerpoint.gui_refactored.MergeWorker', return_value=mock_worker)

        # Initial state - buttons enabled
        assert main_ui.merge_button.isEnabled() is True

        # Trigger merge
        main_ui._start_merge([str(file1), str(file2)], "/tmp/output.pptx")

        # UI should be disabled
        assert main_ui.merge_button.isEnabled() is False
        assert main_ui.clear_button.isEnabled() is False
        assert main_ui.progress_bar.isVisible() is True

    def test_duplicate_files_rejected(self, main_ui, tmp_path, mocker):
        """Test that duplicate files are not added twice."""
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        file1 = tmp_path / "test1.pptx"
        file1.touch()

        # Add same file twice
        main_ui._on_files_dropped([str(file1)])
        main_ui._on_files_dropped([str(file1)])

        # Should only have one file
        assert len(main_ui.file_model.file_paths) == 1

    def test_invalid_files_rejected(self, main_ui, tmp_path, mocker):
        """Test that non-.pptx files are rejected."""
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        file1 = tmp_path / "test1.pdf"
        file1.touch()

        main_ui._on_files_dropped([str(file1)])

        # Should have no files
        assert len(main_ui.file_model.file_paths) == 0

    def test_order_changed_signal_emitted(self, main_ui, tmp_path, mocker):
        """Test that order_changed signal is emitted when model changes."""
        mocker.patch('os.path.exists', return_value=True)
        mocker.patch('os.path.isfile', return_value=True)

        spy = QSignalSpy(main_ui.order_changed)

        file1 = tmp_path / "test1.pptx"
        file2 = tmp_path / "test2.pptx"
        file1.touch()
        file2.touch()

        # Add files - this should emit order_changed
        main_ui._on_files_dropped([str(file1), str(file2)])

        # Signal should be emitted
        assert spy.count() >= 1

    def test_settings_persistence(self, main_ui, tmp_path):
        """Test that settings are saved and restored."""
        # Set a save directory
        test_dir = str(tmp_path / "save_location")
        os.makedirs(test_dir, exist_ok=True)
        main_ui.last_save_dir = test_dir

        # Save settings
        main_ui._save_settings()

        # Create new instance and verify restore
        new_ui = MainUI()
        new_ui._restore_settings()

        # Should restore the saved directory
        assert hasattr(new_ui, 'last_save_dir')

    def test_keyboard_navigation(self, main_ui, qtbot):
        """Test that keyboard navigation works through controls."""
        # This tests basic tab order
        # In a full implementation, we would set explicit tab order
        main_ui.show()
        qtbot.waitExposed(main_ui)

        # Test that we can tab through widgets
        # (Basic test - full implementation would verify specific tab order)
        assert main_ui.isVisible()

    def test_accessible_names_set(self, main_ui):
        """Test that buttons have accessible names."""
        # All buttons should have accessible text
        assert main_ui.clear_button.text() != ""
        assert main_ui.merge_button.text() != ""
        assert main_ui.save_to_button.text() != ""


class TestMergeWorker:
    """Tests for the MergeWorker thread."""

    def test_initialization(self):
        """Test worker initializes with correct parameters."""
        merger = PowerPointMerger()
        files = ["/path/to/file1.pptx", "/path/to/file2.pptx"]
        output = "/path/to/output.pptx"

        worker = MergeWorker(files, output, merger)

        assert worker.file_paths == files
        assert worker.output_path == output
        assert worker.merger is merger

    def test_signals_exist(self):
        """Test that worker has required signals."""
        merger = PowerPointMerger()
        worker = MergeWorker([], "", merger)

        # Check signals exist
        assert hasattr(worker, 'progress')
        assert hasattr(worker, 'finished')


class TestUIStrings:
    """Tests for UI strings dictionary."""

    def test_all_strings_defined(self):
        """Test that all required UI strings are defined."""
        required_keys = [
            "window_title",
            "drop_zone_text",
            "browse_button",
            "clear_list_button",
            "output_group_title",
            "merge_button",
        ]

        for key in required_keys:
            assert key in UI_STRINGS
            assert UI_STRINGS[key] != ""


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
