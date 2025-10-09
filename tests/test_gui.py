"""
Tests for the PySide6 GUI.
"""

from unittest.mock import patch
from PySide6.QtCore import Qt
import pytest

# It's important to import the classes we need to test and mock
from gui import MainWindow
from app import PowerPointMerger


@pytest.fixture
def merger_app():
    """A fixture to create a PowerPointMerger instance for tests."""
    # We can mock the core dependency to avoid actual COM calls
    with patch('app.PowerPointCore'):
        merger = PowerPointMerger()
        yield merger


@pytest.fixture
def main_app(qtbot, merger_app):
    """
    Create and return the main application window, now correctly initialized
    with a PowerPointMerger instance.
    """
    app_window = MainWindow(merger_app)
    qtbot.addWidget(app_window)
    return app_window


def test_main_window_initialization(main_app):
    """Test that the main window initializes with the correct title and widgets."""
    assert main_app.windowTitle() == "PowerPoint Presentation Merger"
    assert main_app.add_button.text() == "&Add Files"
    assert main_app.merge_button.text() == "&Merge Presentations"
    assert main_app.file_list_widget.count() == 0


def test_add_files_button(main_app, qtbot, monkeypatch):
    """Test that the 'Add Files' button updates the list."""
    # Mock the file dialog to return a predefined list of files
    monkeypatch.setattr(
        "PySide6.QtWidgets.QFileDialog.getOpenFileNames",
        lambda *args, **kwargs: (["test1.pptx", "test2.pptx"], "")
    )
    qtbot.mouseClick(main_app.add_button, Qt.MouseButton.LeftButton)
    assert main_app.file_list_widget.count() == 2
    assert main_app.file_list_widget.item(0).text() == "test1.pptx"


def test_remove_file_button(main_app, qtbot):
    """Test that the 'Remove Selected' button works correctly."""
    main_app.merger.add_files(["a.pptx", "b.pptx", "c.pptx"])
    main_app.update_file_list()

    main_app.file_list_widget.setCurrentRow(1)  # Select 'b.pptx'
    qtbot.mouseClick(main_app.remove_button, Qt.MouseButton.LeftButton)

    assert main_app.file_list_widget.count() == 2
    assert main_app.file_list_widget.item(0).text() == "a.pptx"
    assert main_app.file_list_widget.item(1).text() == "c.pptx"


def test_move_buttons(main_app, qtbot):
    """Test the 'Move Up' and 'Move Down' buttons."""
    main_app.merger.add_files(["a.pptx", "b.pptx", "c.pptx"])
    main_app.update_file_list()

    # Test move down
    main_app.file_list_widget.setCurrentRow(0)  # Select 'a.pptx'
    qtbot.mouseClick(main_app.down_button, Qt.MouseButton.LeftButton)
    assert main_app.merger.get_files() == ["b.pptx", "a.pptx", "c.pptx"]

    # Test move up
    main_app.file_list_widget.setCurrentRow(1)  # Select 'a.pptx' again
    qtbot.mouseClick(main_app.up_button, Qt.MouseButton.LeftButton)
    assert main_app.merger.get_files() == ["a.pptx", "b.pptx", "c.pptx"]


@patch('gui.QMessageBox')
def test_merge_button_no_files(mock_msg_box, main_app, qtbot):
    """Test that the merge button shows a warning if no files are added."""
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)
    mock_msg_box.warning.assert_called_once()


@patch('gui.QMessageBox')
def test_merge_button_no_output_path(mock_msg_box, main_app, qtbot):
    """Test that the merge button shows a warning if no output path is set."""
    main_app.merger.add_files(["a.pptx", "b.pptx"])
    main_app.update_file_list()
    qtbot.mouseClick(main_app.merge_button, Qt.MouseButton.LeftButton)
    mock_msg_box.warning.assert_called_once()

