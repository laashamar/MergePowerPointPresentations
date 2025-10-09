"""
Unit tests for app.py module.

Tests the PowerPointMergerApp class and its workflow orchestration.
"""
import unittest
from unittest.mock import Mock, patch, MagicMock
import logging

# Import the module to test
import app


class TestPowerPointMergerApp(unittest.TestCase):
    """Test cases for PowerPointMergerApp class."""

    def setUp(self):
        """Set up test fixtures."""
        self.app_instance = app.PowerPointMergerApp()

    def test_initialization(self):
        """Test that PowerPointMergerApp initializes with correct state."""
        self.assertEqual(self.app_instance.num_files, 0)
        self.assertEqual(self.app_instance.selected_files, [])
        self.assertEqual(self.app_instance.output_filename, "")
        self.assertEqual(self.app_instance.file_order, [])

    @patch('app.gui.show_number_of_files_window')
    def test_run_starts_workflow(self, mock_show_window):
        """Test that run() starts the application workflow."""
        self.app_instance.run()
        mock_show_window.assert_called_once()

    def test_on_number_of_files_entered(self):
        """Test that number of files is stored correctly."""
        with patch('app.gui.show_file_selection_window') as mock_show:
            self.app_instance._on_number_of_files_entered(3)
            self.assertEqual(self.app_instance.num_files, 3)
            mock_show.assert_called_once_with(
                3, self.app_instance._on_files_selected)

    def test_on_files_selected(self):
        """Test that selected files are stored correctly."""
        test_files = ['file1.pptx', 'file2.pptx', 'file3.pptx']
        with patch('app.gui.show_filename_window') as mock_show:
            self.app_instance._on_files_selected(test_files)
            self.assertEqual(self.app_instance.selected_files, test_files)
            mock_show.assert_called_once()

    def test_on_filename_entered(self):
        """Test that output filename is stored correctly."""
        test_filename = 'output.pptx'
        self.app_instance.selected_files = ['file1.pptx', 'file2.pptx']
        with patch('app.gui.show_reorder_window') as mock_show:
            self.app_instance._on_filename_entered(test_filename)
            self.assertEqual(self.app_instance.output_filename, test_filename)
            mock_show.assert_called_once()

    def test_on_files_reordered(self):
        """Test that file order is stored and merge is triggered."""
        test_order = ['file2.pptx', 'file1.pptx']
        with patch.object(
                self.app_instance, '_merge_and_launch') as mock_merge:
            self.app_instance._on_files_reordered(test_order)
            self.assertEqual(self.app_instance.file_order, test_order)
            mock_merge.assert_called_once()

    @patch('app.powerpoint_core.merge_presentations')
    @patch('app.powerpoint_core.launch_slideshow')
    @patch('app.messagebox.showinfo')
    def test_merge_and_launch_success(
            self,
            mock_showinfo,
            mock_launch,
            mock_merge):
        """Test successful merge and launch workflow."""
        # Setup
        self.app_instance.file_order = ['file1.pptx', 'file2.pptx']
        self.app_instance.output_filename = 'output.pptx'
        mock_merge.return_value = (True, '/path/to/output.pptx', None)
        mock_launch.return_value = (True, None)

        # Execute
        self.app_instance._merge_and_launch()

        # Verify
        mock_merge.assert_called_once_with(
            ['file1.pptx', 'file2.pptx'], 'output.pptx')
        mock_launch.assert_called_once_with('/path/to/output.pptx')
        mock_showinfo.assert_called_once()

    @patch('app.powerpoint_core.merge_presentations')
    @patch('app.messagebox.showerror')
    def test_merge_and_launch_merge_failure(self, mock_showerror, mock_merge):
        """Test merge failure handling."""
        # Setup
        self.app_instance.file_order = ['file1.pptx', 'file2.pptx']
        self.app_instance.output_filename = 'output.pptx'
        mock_merge.return_value = (False, '', 'Merge error')

        # Execute
        self.app_instance._merge_and_launch()

        # Verify
        mock_showerror.assert_called_once()

    @patch('app.powerpoint_core.merge_presentations')
    @patch('app.powerpoint_core.launch_slideshow')
    @patch('app.messagebox.showinfo')
    @patch('app.messagebox.showwarning')
    def test_merge_and_launch_slideshow_failure(
            self, mock_warning, mock_showinfo, mock_launch, mock_merge):
        """Test slideshow launch failure handling."""
        # Setup
        self.app_instance.file_order = ['file1.pptx', 'file2.pptx']
        self.app_instance.output_filename = 'output.pptx'
        mock_merge.return_value = (True, '/path/to/output.pptx', None)
        mock_launch.return_value = (False, 'Launch error')

        # Execute
        self.app_instance._merge_and_launch()

        # Verify
        mock_merge.assert_called_once()
        mock_launch.assert_called_once()
        mock_showinfo.assert_called_once()
        mock_warning.assert_called_once()

    @patch('app.PowerPointMergerApp')
    def test_start_app(self, mock_app_class):
        """Test start_app function creates and runs app instance."""
        mock_instance = Mock()
        mock_app_class.return_value = mock_instance

        app.start_app()

        mock_app_class.assert_called_once()
        mock_instance.run.assert_called_once()


if __name__ == '__main__':
    unittest.main()
