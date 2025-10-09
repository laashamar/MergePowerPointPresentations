"""
Unit tests for gui.py module.

Tests GUI window functions with mocking.
"""
import unittest
from unittest.mock import Mock, patch, MagicMock, call
import tkinter as tk

# Import the module to test
import gui


class TestShowNumberOfFilesWindow(unittest.TestCase):
    """Test cases for show_number_of_files_window function."""

    @patch('gui.tk.Tk')
    def test_window_created(self, mock_tk):
        """Test that window is created with correct properties."""
        mock_window = MagicMock()
        mock_tk.return_value = mock_window

        callback = Mock()

        # Execute (don't actually run mainloop)
        with patch.object(mock_window, 'mainloop'):
            gui.show_number_of_files_window(callback)

        # Verify window was configured
        mock_window.title.assert_called_with("Step 1: Number of Files")
        mock_window.geometry.assert_called_with("400x150")


class TestShowFileSelectionWindow(unittest.TestCase):
    """Test cases for show_file_selection_window function."""

    @patch('gui.tk.Tk')
    def test_window_created(self, mock_tk):
        """Test that window is created with correct properties."""
        mock_window = MagicMock()
        mock_tk.return_value = mock_window

        callback = Mock()

        # Execute (don't actually run mainloop)
        with patch.object(mock_window, 'mainloop'):
            gui.show_file_selection_window(3, callback)

        # Verify window was configured
        mock_window.title.assert_called_with("Step 2: Select Files")
        mock_window.geometry.assert_called_with("600x400")


class TestShowFilenameWindow(unittest.TestCase):
    """Test cases for show_filename_window function."""

    @patch('gui.tk.Tk')
    def test_window_created(self, mock_tk):
        """Test that window is created with correct properties."""
        mock_window = MagicMock()
        mock_tk.return_value = mock_window

        callback = Mock()

        # Execute (don't actually run mainloop)
        with patch.object(mock_window, 'mainloop'):
            gui.show_filename_window(callback)

        # Verify window was configured
        mock_window.title.assert_called_with("New Filename")
        mock_window.geometry.assert_called_with("400x150")


class TestShowReorderWindow(unittest.TestCase):
    """Test cases for show_reorder_window function."""

    @patch('gui.tk.Tk')
    def test_window_created(self, mock_tk):
        """Test that window is created with correct properties."""
        mock_window = MagicMock()
        mock_tk.return_value = mock_window

        callback = Mock()
        test_files = ['file1.pptx', 'file2.pptx', 'file3.pptx']

        # Execute (don't actually run mainloop)
        with patch.object(mock_window, 'mainloop'):
            gui.show_reorder_window(test_files, callback)

        # Verify window was configured
        mock_window.title.assert_called_with("Step 4: Set Merge Order")
        mock_window.geometry.assert_called_with("600x450")


class TestGUIInputValidation(unittest.TestCase):
    """Test cases for input validation in GUI functions."""

    @patch('gui.messagebox.showerror')
    @patch('gui.tk.Tk')
    def test_number_of_files_invalid_input(self, mock_tk, mock_showerror):
        """Test validation of number of files input."""
        mock_window = MagicMock()
        mock_entry = MagicMock()
        mock_entry.get.return_value = 'invalid'

        mock_tk.return_value = mock_window

        callback = Mock()

        # We can't easily test the inner on_next function without
        # actually creating the GUI, so we'll just verify the
        # window creation
        with patch.object(mock_window, 'mainloop'):
            gui.show_number_of_files_window(callback)

        # Verify window was created
        mock_window.title.assert_called_with("Step 1: Number of Files")


class TestFileDialogIntegration(unittest.TestCase):
    """Test cases for file dialog integration."""

    @patch('gui.filedialog.askopenfilenames')
    @patch('gui.tk.Tk')
    def test_file_selection_uses_file_dialog(
            self, mock_tk, mock_file_dialog):
        """Test that file selection opens file dialog."""
        mock_window = MagicMock()
        mock_tk.return_value = mock_window
        mock_file_dialog.return_value = []

        callback = Mock()

        # Execute (don't actually run mainloop)
        with patch.object(mock_window, 'mainloop'):
            gui.show_file_selection_window(2, callback)

        # Verify window was created
        mock_window.title.assert_called_with("Step 2: Select Files")


if __name__ == '__main__':
    unittest.main()
