import pytest
from unittest.mock import MagicMock, patch

# MODIFIED: Import from the package
from merge_powerpoint.app import AppController
from merge_powerpoint.gui import MainWindow


@pytest.fixture
def app_controller():
    """Returns a clean AppController instance for each test."""
    return AppController()


def test_app_controller_initialization(qapp):
    """
    Test that the AppController initializes correctly.
    """
    controller = AppController()
    assert controller.files_to_merge == []
    assert isinstance(controller.main_window, MainWindow)
    # Test that the main_window's controller is the one we just created
    assert controller.main_window.controller == controller

@patch('PySide6.QtWidgets.QFileDialog.getOpenFileNames')
def test_add_files_selected(mock_get_open_files, app_controller):
    """
    Test adding files when the user selects files in the dialog.
    """
    mock_get_open_files.return_value = (['file1.pptx', 'file2.pptx'], None)
    app_controller.main_window.update_file_list = MagicMock()

    app_controller.add_files()

    assert app_controller.files_to_merge == ['file1.pptx', 'file2.pptx']
    app_controller.main_window.update_file_list.assert_called_once_with(['file1.pptx', 'file2.pptx'])

@patch('PySide6.QtWidgets.QFileDialog.getOpenFileNames')
def test_add_files_cancelled(mock_get_open_files, app_controller):
    """
    Test adding files when the user cancels the dialog.
    """
    mock_get_open_files.return_value = ([], None)
    app_controller.main_window.update_file_list = MagicMock()

    app_controller.add_files()

    assert app_controller.files_to_merge == []
    app_controller.main_window.update_file_list.assert_not_called()

def test_remove_selected_file_with_selection(app_controller):
    """
    Test removing a file when one is selected in the list.
    """
    app_controller.files_to_merge = ['file1.pptx', 'file2.pptx']
    # Mock the list widget and its selectedItems method
    mock_item = MagicMock()
    mock_item.text.return_value = 'file1.pptx'
    app_controller.main_window.file_list.selectedItems = MagicMock(return_value=[mock_item])
    app_controller.main_window.update_file_list = MagicMock()

    app_controller.remove_selected_file()

    assert app_controller.files_to_merge == ['file2.pptx']
    app_controller.main_window.update_file_list.assert_called_once_with(['file2.pptx'])

def test_remove_selected_file_without_selection(app_controller):
    """
    Test removing a file when nothing is selected.
    """
    app_controller.files_to_merge = ['file1.pptx']
    app_controller.main_window.file_list.selectedItems = MagicMock(return_value=[])
    app_controller.main_window.update_file_list = MagicMock()

    app_controller.remove_selected_file()

    assert app_controller.files_to_merge == ['file1.pptx']
    app_controller.main_window.update_file_list.assert_not_called()

# MODIFIED: Patch the PowerPointMerger.merge method instead of merge_presentations function
@patch('PySide6.QtWidgets.QFileDialog.getSaveFileName')
@patch('merge_powerpoint.app.PowerPointMerger.merge')
def test_merge_files_success(mock_merge, mock_save_file, app_controller):
    """
    Test the successful merge process.
    """
    app_controller.files_to_merge = ['file1.pptx', 'file2.pptx']
    mock_save_file.return_value = ('output.pptx', None)
    app_controller.main_window.show_message = MagicMock()
    app_controller.main_window.progress_bar.setVisible = MagicMock()
    app_controller.main_window.progress_bar.setValue = MagicMock()

    app_controller.merge_files()

    mock_merge.assert_called_once()
    app_controller.main_window.show_message.assert_called_with("Success", "Files merged successfully to:\noutput.pptx")
    # Called once to show, once to hide
    assert app_controller.main_window.progress_bar.setVisible.call_count == 2

@patch('PySide6.QtWidgets.QFileDialog.getSaveFileName')
def test_merge_files_insufficient_files(mock_save_file, app_controller):
    """
    Test merging when fewer than two files are selected.
    """
    app_controller.files_to_merge = ['file1.pptx']
    app_controller.main_window.show_message = MagicMock()

    app_controller.merge_files()

    app_controller.main_window.show_message.assert_called_with("Error", "Please select at least two files to merge.")
    mock_save_file.assert_not_called()

# MODIFIED: Patch the PowerPointMerger.merge method instead of merge_presentations function
@patch('PySide6.QtWidgets.QFileDialog.getSaveFileName')
@patch('merge_powerpoint.app.PowerPointMerger.merge', side_effect=Exception("Test error"))
def test_merge_files_exception(mock_merge, mock_save_file, app_controller):
    """
    Test the merge process when an exception occurs.
    """
    app_controller.files_to_merge = ['file1.pptx', 'file2.pptx']
    mock_save_file.return_value = ('output.pptx', None)
    app_controller.main_window.show_message = MagicMock()
    app_controller.main_window.progress_bar.setVisible = MagicMock()

    app_controller.merge_files()

    mock_merge.assert_called_once()
    app_controller.main_window.show_message.assert_called_with("Error", "An error occurred during merge: Test error")
    assert app_controller.main_window.progress_bar.setVisible.call_count == 2

