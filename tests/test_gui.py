import pytest
from PySide6.QtCore import Qt

# MODIFIED: Imports are no longer needed here because the fixtures
# in conftest.py already provide the necessary application objects.

def test_main_window_initial_state(main_window):
    """
    Test the initial state of the MainWindow.
    """
    assert main_window.windowTitle() == "PowerPoint Merger"
    assert main_window.file_list.count() == 0
    assert main_window.merge_button.isEnabled()
    assert not main_window.progress_bar.isVisible()

def test_update_file_list(main_window):
    """
    Test that the file list widget updates correctly.
    """
    files = ["C:/path/one.pptx", "C:/path/two.pptx"]
    main_window.update_file_list(files)
    
    assert main_window.file_list.count() == 2
    assert main_window.file_list.item(0).text() == "C:/path/one.pptx"
    assert main_window.file_list.item(1).text() == "C:/path/two.pptx"

def test_update_progress(main_window):
    """
    Test that the progress bar updates correctly.
    """
    main_window.update_progress(50)
    assert main_window.progress_bar.value() == 50

def test_add_button_click_triggers_controller(main_window, qtbot, mocker):
    """
    Test that clicking the 'Add Files' button calls the controller's method.
    """
    # Mock the controller's method to check if it's called
    mocker.patch.object(main_window.controller, 'add_files')
    qtbot.mouseClick(main_window.add_button, Qt.LeftButton)
    main_window.controller.add_files.assert_called_once()

def test_remove_button_click_triggers_controller(main_window, qtbot, mocker):
    """
    Test that clicking the 'Remove Selected' button calls the controller's method.
    """
    mocker.patch.object(main_window.controller, 'remove_selected_file')
    qtbot.mouseClick(main_window.remove_button, Qt.LeftButton)
    main_window.controller.remove_selected_file.assert_called_once()

def test_merge_button_click_triggers_controller(main_window, qtbot, mocker):
    """
    Test that clicking the 'Merge Files' button calls the controller's method.
    """
    mocker.patch.object(main_window.controller, 'merge_files')
    qtbot.mouseClick(main_window.merge_button, Qt.LeftButton)
    main_window.controller.merge_files.assert_called_once()

