import pytest
from PySide6.QtCore import Qt


def test_main_window_initial_state(main_window):
    """
    Test the initial state of the MainWindow.
    """
    # Check the correct window title
    assert main_window.windowTitle() == "PowerPoint Presentation Merger"
    # Check that file list is empty
    assert main_window.file_list_widget.count() == 0
    # Merge button should be disabled when there are no files
    assert not main_window.merge_button.isEnabled()
    # Progress bar should be hidden initially
    assert not main_window.progress_bar.isVisible()


def test_update_file_list(main_window):
    """
    Test that the file list widget updates correctly.
    To update the file list, we manipulate the merger and call update_file_list().
    """
    files = ["C:/path/one.pptx", "C:/path/two.pptx"]
    # Add files to the merger
    main_window.merger.add_files(files)
    # Update the GUI to reflect the changes
    main_window.update_file_list()
    
    assert main_window.file_list_widget.count() == 2
    assert main_window.file_list_widget.item(0).text() == "C:/path/one.pptx"
    assert main_window.file_list_widget.item(1).text() == "C:/path/two.pptx"


def test_update_progress(main_window):
    """
    Test that the progress bar updates correctly.
    The update_progress method takes (value, total) parameters.
    """
    main_window.update_progress(1, 2)
    assert main_window.progress_bar.value() == 50
    
    main_window.update_progress(2, 4)
    assert main_window.progress_bar.value() == 50


def test_add_button_click(main_window, qtbot, mocker):
    """
    Test that clicking the 'Add Files' button calls the add_files method.
    """
    # Mock the add_files method to verify it's called
    mocker.patch.object(main_window, 'add_files')
    qtbot.mouseClick(main_window.add_button, Qt.LeftButton)
    main_window.add_files.assert_called_once()


def test_remove_button_click(main_window, qtbot, mocker):
    """
    Test that clicking the 'Remove Selected' button calls the remove_selected_files method.
    """
    # Add files and select one to enable the remove button
    main_window.merger.add_files(['file1.pptx', 'file2.pptx'])
    main_window.update_file_list()
    # Select an item (this triggers selection changed signal)
    item = main_window.file_list_widget.item(0)
    item.setSelected(True)
    main_window.file_list_widget.setCurrentItem(item)
    # Force update button states
    main_window.update_button_states()
    
    # Mock the remove_selected_files method
    mocker.patch.object(main_window, 'remove_selected_files')
    qtbot.mouseClick(main_window.remove_button, Qt.LeftButton)
    main_window.remove_selected_files.assert_called_once()


def test_merge_button_click(main_window, qtbot, mocker):
    """
    Test that clicking the 'Merge Files' button calls the merge_files method.
    """
    # Add at least 2 files to enable the merge button
    main_window.merger.add_files(['file1.pptx', 'file2.pptx'])
    main_window.update_file_list()
    
    # Mock the merge_files method
    mocker.patch.object(main_window, 'merge_files')
    qtbot.mouseClick(main_window.merge_button, Qt.LeftButton)
    main_window.merge_files.assert_called_once()
