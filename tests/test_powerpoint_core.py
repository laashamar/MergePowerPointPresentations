"""
Unit tests for powerpoint_core.py module.

Tests the PowerPoint COM automation functions with mocking.
"""
import unittest
from unittest.mock import Mock, patch, MagicMock
import os

# Import the module to test
import powerpoint_core


class TestMergePresentations(unittest.TestCase):
    """Test cases for merge_presentations function."""

    @patch('powerpoint_core.win32com.client.Dispatch')
    @patch('os.path.abspath')
    def test_merge_presentations_success(self, mock_abspath, mock_dispatch):
        """Test successful merging of presentations."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint

        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination

        mock_source = MagicMock()
        mock_source.Slides.Count = 3
        mock_powerpoint.Presentations.Open.return_value = mock_source

        mock_abspath.side_effect = lambda x: f'/abs/{x}'

        file_order = ['file1.pptx', 'file2.pptx']
        output_filename = 'output.pptx'

        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            file_order, output_filename
        )

        # Verify
        self.assertTrue(success)
        self.assertEqual(output_path, '/abs/output.pptx')
        self.assertIsNone(error_msg)
        mock_destination.SaveAs.assert_called_once_with('/abs/output.pptx')

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_presentations_handles_error(self, mock_dispatch):
        """Test error handling in merge_presentations."""
        # Setup mock to raise exception
        mock_dispatch.side_effect = Exception('COM error')

        file_order = ['file1.pptx', 'file2.pptx']
        output_filename = 'output.pptx'

        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            file_order, output_filename
        )

        # Verify
        self.assertFalse(success)
        self.assertEqual(output_path, '')
        self.assertIn('COM error', error_msg)

    @patch('powerpoint_core.win32com.client.Dispatch')
    @patch('os.path.abspath')
    def test_merge_presentations_removes_default_slide(
            self, mock_abspath, mock_dispatch):
        """Test that default blank slide is removed."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint

        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination

        mock_source = MagicMock()
        mock_source.Slides.Count = 2
        mock_powerpoint.Presentations.Open.return_value = mock_source

        mock_abspath.side_effect = lambda x: f'/abs/{x}'

        file_order = ['file1.pptx']
        output_filename = 'output.pptx'

        # Execute
        powerpoint_core.merge_presentations(file_order, output_filename)

        # Verify default slide was deleted
        mock_destination.Slides.assert_called()

    @patch('powerpoint_core.win32com.client.Dispatch')
    @patch('os.path.abspath')
    def test_merge_presentations_with_progress_callback(
            self, mock_abspath, mock_dispatch):
        """Test merge_presentations with progress callback."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint

        mock_destination = MagicMock()
        mock_destination.Slides.Count = 0
        mock_powerpoint.Presentations.Add.return_value = mock_destination

        mock_source = MagicMock()
        mock_source.Slides.Count = 3
        mock_powerpoint.Presentations.Open.return_value = mock_source

        mock_abspath.side_effect = lambda x: f'/abs/{x}'

        mock_callback = Mock()
        file_order = ['file1.pptx']
        output_filename = 'output.pptx'

        # Execute
        powerpoint_core.merge_presentations(
            file_order, output_filename, progress_callback=mock_callback
        )

        # Verify callback was called
        self.assertEqual(mock_callback.call_count, 3)


class TestLaunchSlideshow(unittest.TestCase):
    """Test cases for launch_slideshow function."""

    @patch('powerpoint_core.win32com.client.Dispatch')
    @patch('os.path.abspath')
    def test_launch_slideshow_success(self, mock_abspath, mock_dispatch):
        """Test successful slideshow launch."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint

        mock_presentation = MagicMock()
        mock_powerpoint.Presentations.Open.return_value = mock_presentation

        mock_abspath.return_value = '/abs/path/to/file.pptx'

        # Execute
        success, error_msg = powerpoint_core.launch_slideshow('file.pptx')

        # Verify
        self.assertTrue(success)
        self.assertIsNone(error_msg)
        mock_presentation.SlideShowSettings.Run.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_launch_slideshow_handles_error(self, mock_dispatch):
        """Test error handling in launch_slideshow."""
        # Setup mock to raise exception
        mock_dispatch.side_effect = Exception('COM error')

        # Execute
        success, error_msg = powerpoint_core.launch_slideshow('file.pptx')

        # Verify
        self.assertFalse(success)
        self.assertIn('COM error', error_msg)

    @patch('powerpoint_core.win32com.client.Dispatch')
    @patch('os.path.abspath')
    def test_launch_slideshow_makes_powerpoint_visible(
            self, mock_abspath, mock_dispatch):
        """Test that PowerPoint is made visible."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint

        mock_presentation = MagicMock()
        mock_powerpoint.Presentations.Open.return_value = mock_presentation

        mock_abspath.return_value = '/abs/path/to/file.pptx'

        # Execute
        powerpoint_core.launch_slideshow('file.pptx')

        # Verify
        self.assertTrue(mock_powerpoint.Visible)


if __name__ == '__main__':
    unittest.main()
