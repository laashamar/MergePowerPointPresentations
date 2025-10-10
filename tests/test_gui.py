"""
Tests for the GUI components of the PowerPoint Presentation Merger application.
"""
import pytest
from unittest.mock import MagicMock, patch
from PySide6.QtCore import Qt

# conftest.py provides the main_app and mock_file_dialog fixtures


def test_main_window_initialization(main_app):
    """Test that the main window initializes with the correct title and widgets."""
    assert main_app.windowTitle() == "PowerPoint Presentation Merger"
    # The `&` is a mnemonic for shortcuts and is part of the actual text
    assert main_app.add_button.text() == "&Add Files"
    assert main_app.remove_button.text() == "&Remove Selected"
    assert main_app.clear_button.text() == "&Clear All"
    assert main_app.merge_button.text() == "&Merge Presentations"
    assert main_app.move_up_button.text() == "Move &Up"
    assert main_app.move_down_button.text() == "Move &Down"


def test_add_files_button(mock_file_dialog, main_app):
    """Test the 'Add Files' functionality."""
    main_app.add_files_button.click()
    mock_file_dialog.assert_called_once()
    assert main_app.file_list_widget.count() == 2
    assert main_app.file_list_widget.item(0).text() == "file1.pptx"
    assert main_app.file_list_widget.item(1).text() == "file2.pptx"


def test_remove_selected_files_button(main_app):
    """Test removing a selected file from the list."""
    main_app.merger.add_files(["a.pptx", "b.pptx", "c.pptx"])
    main_app.update_file_list()
    main_app.file_list_widget.setCurrentRow(1)  # Select "b.pptx"
    main_app.remove_files_button.click()
    assert main_app.file_list_widget.count() == 2
    assert main_app.file_list_widget.item(0).text() == "a.pptx"
    assert main_app.file_list_widget.item(1).text() == "c.pptx"


def test_clear_all_files_button(main_app):
    """Test clearing all files from the list."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    assert main_app.file_list_widget.count() == 2
    main_app.clear_files_button.click()
    assert main_app.file_list_widget.count() == 0


@patch('gui.QMessageBox')
@patch('gui.QFileDialog.getSaveFileName', return_value=('merged.pptx', 'PowerPoint Presentation (*.pptx)'))
def test_merge_button_success(mock_save_dialog, mock_message_box, main_app, qtbot):
    """Test a successful merge operation."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.merger.merge = MagicMock()

    # Ensure the button is enabled before clicking, as it requires at least 2 files.
    assert main_app.merge_button.isEnabled()
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)

    main_app.merger.merge.assert_called_once_with('merged.pptx', main_app.update_progress)
    mock_save_dialog.assert_called_once()
    mock_message_box.information.assert_called_once_with(
        main_app, "Success", "Presentations merged successfully into merged.pptx"
    )


def test_merge_button_no_files(main_app, qtbot):
    """Test that the merge button is disabled when no files are present."""
    assert not main_app.merge_button.isEnabled()
    main_app.merger.merge = MagicMock()
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)
    main_app.merger.merge.assert_not_called()


@patch('gui.QMessageBox')
@patch('gui.QFileDialog.getSaveFileName', return_value=('', None))
def test_merge_button_cancel(mock_save_dialog, mock_message_box, main_app, qtbot):
    """Test cancelling the merge operation's save dialog."""
    # Add two files to ensure the merge button is enabled
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.merger.merge = MagicMock()

    assert main_app.merge_button.isEnabled()
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)

    mock_save_dialog.assert_called_once()
    main_app.merger.merge.assert_not_called()
    mock_message_box.information.assert_not_called()


def test_move_file_up_button(main_app):
    """Test moving a selected file up in the list."""
    main_app.merger.add_files(["a.pptx", "b.pptx", "c.pptx"])
    main_app.update_file_list()
    main_app.file_list_widget.setCurrentRow(1)  # Select "b.pptx"
    main_app.move_up_button.click()
    assert main_app.merger.get_files_list() == ["b.pptx", "a.pptx", "c.pptx"]
    assert main_app.file_list_widget.currentRow() == 0


def test_move_file_down_button(main_app):
    """Test moving a selected file down in the list."""
    main_app.merger.add_files(["a.pptx", "b.pptx", "c.pptx"])
    main_app.update_file_list()
    main_app.file_list_widget.setCurrentRow(1)  # Select "b.pptx"
    main_app.move_down_button.click()
    assert main_app.merger.get_files_list() == ["a.pptx", "c.pptx", "b.pptx"]
    assert main_app.file_list_widget.currentRow() == 2


def test_update_progress(main_app):
    """Test that the progress bar and status label are updated correctly."""
    main_app.update_progress(5, 10, "file.pptx")
    assert main_app.progress_bar.value() == 50
    assert main_app.status_label.text() == "Processing slide 5 of 10 in file.pptx"

