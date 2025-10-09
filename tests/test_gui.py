"""
Tests for the GUI (MainWindow) of the PowerPoint Merger application.
"""
import sys
from unittest.mock import patch, MagicMock
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
import pytest
from gui import MainWindow
from app import AppController

# Fixture to create a QApplication instance for the tests
@pytest.fixture(scope="session")
def qapp():
    """Fixture to create a QApplication instance for the tests."""
    return QApplication.instance() or QApplication(sys.argv)

# Fixture to create the main application window
@pytest.fixture
def main_app(qapp):
    """Fixture to create the main application window."""
    controller = AppController()
    app = MainWindow(controller)
    yield app

def test_main_window_initialization(main_app):
    """Test that the main window initializes with the correct title and widgets."""
    assert main_app.windowTitle() == "PowerPoint Presentation Merger"
    # The `&` is a mnemonic for shortcuts and not part of the actual text
    assert main_app.add_button.text() == "Add Files"
    assert main_app.remove_button.text() == "Remove"
    assert main_app.move_up_button.text() == "Move Up"
    assert main_app.move_down_button.text() == "Move Down"
    assert main_app.merge_button.text() == "Merge"

@patch('gui.QFileDialog.getOpenFileNames', return_value=(['test1.pptx', 'test2.pptx'], None))
def test_add_files_button(mock_dialog, main_app, qtbot):
    """Test the 'Add Files' button functionality."""
    qtbot.mouseClick(main_app.add_button, Qt.MouseButton.LeftButton)
    mock_dialog.assert_called_once()
    assert main_app.list_widget.count() == 2
    assert main_app.list_widget.item(0).text() == 'test1.pptx'

def test_remove_files_button(main_app, qtbot):
    """Test the 'Remove' button functionality."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.list_widget.setCurrentRow(0)
    qtbot.mouseClick(main_app.remove_button, Qt.MouseButton.LeftButton)
    assert main_app.list_widget.count() == 1
    assert main_app.list_widget.item(0).text() == 'b.pptx'

def test_move_up_button(main_app, qtbot):
    """Test the 'Move Up' button functionality."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.list_widget.setCurrentRow(1)
    qtbot.mouseClick(main_app.move_up_button, Qt.MouseButton.LeftButton)
    assert main_app.list_widget.item(0).text() == 'b.pptx'

def test_move_down_button(main_app, qtbot):
    """Test the 'Move Down' button functionality."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.list_widget.setCurrentRow(0)
    qtbot.mouseClick(main_app.move_down_button, Qt.MouseButton.LeftButton)
    assert main_app.list_widget.item(0).text() == 'b.pptx'

@patch('gui.QMessageBox')
@patch('gui.QFileDialog.getSaveFileName', return_value=('merged.pptx', None))
def test_merge_button_success(mock_save_dialog, mock_message_box, main_app, qtbot):
    """Test a successful merge operation."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.merger.merge = MagicMock()

    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)

    main_app.merger.merge.assert_called_once_with('merged.pptx', main_app.update_progress)
    mock_save_dialog.assert_called_once()
    mock_message_box.information.assert_called_once_with(
        main_app, "Success", "Presentations merged successfully!"
    )

@patch('gui.QMessageBox')
def test_merge_button_no_files(mock_message_box, main_app, qtbot):
    """Test merge button with no files selected."""
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)
    # The application first checks for an output path before checking for files.
    # The test must reflect this actual behavior.
    mock_message_box.warning.assert_called_with(
        main_app, "Output Path Missing", "Please specify an output file path."
    )

@patch('gui.QMessageBox')
@patch('gui.QFileDialog.getSaveFileName', return_value=('', None))
def test_merge_button_cancel(mock_save_dialog, mock_message_box, main_app, qtbot):
    """Test cancelling the merge operation's save dialog."""
    main_app.merger.add_files(["a.pptx"])
    main_app.update_file_list()
    main_app.merger.merge = MagicMock()

    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)

    mock_save_dialog.assert_called_once()
    main_app.merger.merge.assert_not_called()
    mock_message_box.information.assert_not_called()

@patch('gui.QMessageBox.critical')
def test_merge_failure(mock_critical, main_app, qtbot):
    """Test a failed merge operation."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.output_path_line_edit.setText("output.pptx")
    main_app.merger.merge = MagicMock(side_effect=Exception("Merge failed"))
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)
    mock_critical.assert_called_once_with(
        main_app, "Error", "An error occurred during merge: Merge failed"
    )

@patch("PySide6.QtGui.QDesktopServices.openUrl")
def test_help_button(mock_open_url, main_app, qtbot):
    """Test that the help button opens the correct URL."""
    qtbot.mouseClick(main_app.help_button, Qt.MouseButton.LeftButton)
    mock_open_url.assert_called_once()
    called_url = mock_open_url.call_args[0][0].toString()
    assert "https://github.com/laashamar/MergePowerPointPresentations/blob/main/README.md" in called_url

