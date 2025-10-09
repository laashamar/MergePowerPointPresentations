# tests/test_gui.py

"""
Tests for the GUI of the PowerPoint Merger application.
"""

import os
from unittest.mock import patch, MagicMock
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication, QFileDialog
import pytest
from gui import MainWindow

# Fixture to create the application and main window instances
@pytest.fixture
def main_app(qtbot):
    """
    Creates the main application window for testing.
    """
    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    qtbot.addWidget(window)
    yield window
    window.close()

def test_main_window_initialization(main_app):
    """Test that the main window initializes with the correct title and widgets."""
    assert main_app.windowTitle() == "PowerPoint Presentation Merger"
    assert main_app.add_button.text() == "&Add Files"
    assert main_app.remove_button.text() == "&Remove"
    assert main_app.up_button.text() == "Move &Up"
    assert main_app.down_button.text() == "Move &Down"
    assert main_app.merge_button.text() == "&Merge"
    assert main_app.output_label.text().startswith("Output File:")

def test_add_files_button(main_app, qtbot, monkeypatch):
    """Test the 'Add Files' button functionality."""
    test_files = ["test1.pptx", "test2.pptx"]

    # Mock QFileDialog.getOpenFileNames to return our test files
    monkeypatch.setattr(QFileDialog, 'getOpenFileNames', lambda *args, **kwargs: (test_files, None))

    qtbot.mouseClick(main_app.add_button, Qt.MouseButton.LeftButton)
    assert main_app.merger.file_paths == test_files
    assert main_app.file_list_widget.count() == 2
    assert main_app.file_list_widget.item(0).text() == os.path.basename(test_files[0])
    assert main_app.file_list_widget.item(1).text() == os.path.basename(test_files[1])

def test_remove_button(main_app, qtbot):
    """Test the 'Remove' button functionality."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.file_list_widget.setCurrentRow(0)
    qtbot.mouseClick(main_app.remove_button, Qt.MouseButton.LeftButton)
    assert main_app.merger.file_paths == ["b.pptx"]

def test_move_buttons(main_app, qtbot):
    """Test the 'Move Up' and 'Move Down' buttons."""
    main_app.merger.add_files(["a.pptx", "b.pptx", "c.pptx"])
    main_app.update_file_list()

    # Test move down
    main_app.file_list_widget.setCurrentRow(0)  # Select 'a.pptx'
    qtbot.mouseClick(main_app.down_button, Qt.MouseButton.LeftButton)
    assert main_app.merger.get_files() == ["b.pptx", "a.pptx", "c.pptx"]

    # Test move up
    main_app.file_list_widget.setCurrentRow(1)  # Select 'a.pptx'
    qtbot.mouseClick(main_app.up_button, Qt.MouseButton.LeftButton)
    assert main_app.merger.get_files() == ["a.pptx", "b.pptx", "c.pptx"]

@patch('gui.QMessageBox')
@patch('gui.QFileDialog.getSaveFileName', return_value=('merged.pptx', None))
def test_merge_button_success(mock_save_dialog, mock_message_box, main_app, qtbot):
    """Test a successful merge operation."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    main_app.merger.merge = MagicMock()

    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)

    main_app.merger.merge.assert_called_once_with('merged.pptx', main_app.update_progress)
    mock_message_box.information.assert_called_once()

@patch('gui.QMessageBox')
def test_merge_button_no_files(mock_message_box, main_app, qtbot):
    """Test merge button with no files selected."""
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)
    mock_message_box.warning.assert_called_with(main_app, "No files", "Please add files to merge.")
