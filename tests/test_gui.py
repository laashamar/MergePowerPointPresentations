"""Tests for the refactored MainUI widget.

This module tests the modern two-column interface including basic
functionality, state management, and user interactions.
"""
from PySide6.QtCore import Qt


def test_main_ui_initial_state(main_window):
    """Test the initial state of the MainUI widget."""
    # Check that file model is empty
    assert len(main_window.file_model.file_paths) == 0
    assert main_window.file_model.rowCount() == 0
    # Merge button should be disabled when there are no files
    assert not main_window.merge_button.isEnabled()
    # Clear button should be disabled
    assert not main_window.clear_button.isEnabled()
    # Progress bar should be hidden initially
    assert not main_window.progress_bar.isVisible()
    # Drop zone should be visible, file list hidden
    assert main_window.drop_zone.isVisible()
    assert not main_window.file_list_view.isVisible()


def test_add_files(main_window):
    """Test adding files to the file model."""
    files = ["/path/to/file1.pptx", "/path/to/file2.pptx"]
    rejected = main_window.file_model.add_files(files)

    assert rejected == []
    assert main_window.file_model.rowCount() == 2
    assert main_window.file_model.file_paths == files
    # Merge button should be enabled with 2+ files
    assert main_window.merge_button.isEnabled()
    # Clear button should be enabled
    assert main_window.clear_button.isEnabled()
    # File list should be visible, drop zone hidden
    assert main_window.file_list_view.isVisible()
    assert not main_window.drop_zone.isVisible()


def test_clear_files(main_window):
    """Test clearing all files."""
    # Add some files first
    files = ["/path/to/file1.pptx", "/path/to/file2.pptx"]
    main_window.file_model.add_files(files)

    # Clear them
    main_window.file_model.clear_all()

    assert main_window.file_model.rowCount() == 0
    assert len(main_window.file_model.file_paths) == 0
    # UI should reset to initial state
    assert not main_window.merge_button.isEnabled()
    assert not main_window.clear_button.isEnabled()
    assert main_window.drop_zone.isVisible()
    assert not main_window.file_list_view.isVisible()


def test_merge_progress_update(main_window):
    """Test that the progress bar updates correctly."""
    # Make progress bar visible first
    main_window.progress_bar.setVisible(True)

    # Simulate progress updates
    main_window._on_merge_progress(1, 2)
    assert main_window.progress_bar.value() == 50

    main_window._on_merge_progress(2, 4)
    assert main_window.progress_bar.value() == 50

    main_window._on_merge_progress(4, 4)
    assert main_window.progress_bar.value() == 100


def test_clear_button_click(main_window, qtbot, mocker):
    """Test that clicking the clear button clears files."""
    # Add some files
    files = ["/path/to/file1.pptx", "/path/to/file2.pptx"]
    main_window.file_model.add_files(files)

    # Mock the signal to verify it's emitted
    mocker.patch.object(main_window, 'clear_requested')

    # Click the clear button
    qtbot.mouseClick(main_window.clear_button, Qt.LeftButton)

    # Verify files are cleared
    assert main_window.file_model.rowCount() == 0
    main_window.clear_requested.emit.assert_called_once()


def test_merge_button_enabled_state(main_window):
    """Test that merge button is enabled only with 2+ files."""
    # Initially disabled
    assert not main_window.merge_button.isEnabled()

    # Add one file - should still be disabled
    main_window.file_model.add_files(["/path/to/file1.pptx"])
    assert not main_window.merge_button.isEnabled()

    # Add second file - should be enabled
    main_window.file_model.add_files(["/path/to/file2.pptx"])
    assert main_window.merge_button.isEnabled()


def test_reject_duplicate_files(main_window):
    """Test that duplicate files are rejected."""
    file = "/path/to/file1.pptx"

    # Add file first time - should succeed
    rejected = main_window.file_model.add_files([file])
    assert rejected == []
    assert main_window.file_model.rowCount() == 1

    # Try to add same file again - should be rejected
    rejected = main_window.file_model.add_files([file])
    assert len(rejected) == 1
    assert rejected[0] == file
    assert main_window.file_model.rowCount() == 1  # Count unchanged


def test_reject_non_pptx_files(main_window):
    """Test that non-.pptx files are rejected."""
    invalid_files = ["/path/to/file.txt", "/path/to/file.pdf"]
    rejected = main_window.file_model.add_files(invalid_files)

    assert len(rejected) == 2
    assert main_window.file_model.rowCount() == 0

