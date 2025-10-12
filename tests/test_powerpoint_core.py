"""
Unit tests for powerpoint_core module.

This module tests the core PowerPoint merging functionality using mocking
to avoid requiring actual PowerPoint installation during testing.
"""
import os
import sys
import pytest
from unittest.mock import Mock, MagicMock, patch, call

# Mock win32com before importing powerpoint_core
sys.modules['win32com'] = MagicMock()
sys.modules['win32com.client'] = MagicMock()

import powerpoint_core


class TestMergePresentations:
    """Tests for the merge_presentations function."""

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_successful_merge(self, mock_dispatch):
        """Test successful merge of multiple PowerPoint files."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        # Create mock destination presentation
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1  # Start with default blank slide
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        # Create mock source presentations with slides
        mock_source1 = MagicMock()
        mock_source1.Slides.Count = 3
        mock_source1.Slides.Range.return_value = MagicMock()
        
        mock_source2 = MagicMock()
        mock_source2.Slides.Count = 2
        mock_source2.Slides.Range.return_value = MagicMock()
        
        # Configure Open to return different presentations
        mock_powerpoint.Presentations.Open.side_effect = [mock_source1, mock_source2]
        
        # Test files
        test_files = ['file1.pptx', 'file2.pptx']
        output_file = 'output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is True
        assert os.path.abspath(output_file) == output_path
        assert error_msg is None
        
        # Verify PowerPoint was initialized
        mock_dispatch.assert_called_once_with("PowerPoint.Application")
        assert mock_powerpoint.Visible is True
        
        # Verify new presentation was created
        mock_powerpoint.Presentations.Add.assert_called_once()
        
        # Verify default blank slide was removed
        mock_destination.Slides.assert_called()
        
        # Verify both source files were opened
        assert mock_powerpoint.Presentations.Open.call_count == 2
        
        # Verify slides were copied and pasted
        assert mock_source1.Slides.Range.call_count == 1
        assert mock_source2.Slides.Range.call_count == 1
        assert mock_destination.Slides.Paste.call_count == 2
        
        # Verify source presentations were closed
        mock_source1.Close.assert_called_once()
        mock_source2.Close.assert_called_once()
        
        # Verify destination was saved
        mock_destination.SaveAs.assert_called_once_with(os.path.abspath(output_file))

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_single_file(self, mock_dispatch):
        """Test merging a single file (edge case)."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        mock_source = MagicMock()
        mock_source.Slides.Count = 5
        mock_source.Slides.Range.return_value = MagicMock()
        mock_powerpoint.Presentations.Open.return_value = mock_source
        
        # Test with single file
        test_files = ['single_file.pptx']
        output_file = 'output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is True
        assert error_msg is None
        assert mock_powerpoint.Presentations.Open.call_count == 1
        mock_destination.SaveAs.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_with_special_characters_in_filename(self, mock_dispatch):
        """Test merging files with special characters and spaces in names."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        mock_source = MagicMock()
        mock_source.Slides.Count = 2
        mock_source.Slides.Range.return_value = MagicMock()
        mock_powerpoint.Presentations.Open.return_value = mock_source
        
        # Test with special characters and spaces
        test_files = ['file with spaces & special chars.pptx']
        output_file = 'output (version 1).pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is True
        assert error_msg is None
        mock_destination.SaveAs.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_empty_file_list(self, mock_dispatch):
        """Test merging with empty file list."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        # Test with empty file list
        test_files = []
        output_file = 'output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify - should succeed but no files opened
        assert success is True
        assert mock_powerpoint.Presentations.Open.call_count == 0
        mock_destination.SaveAs.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_corrupt_file_error(self, mock_dispatch):
        """Test handling of corrupt input files."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        # Simulate corrupt file - Open raises exception
        mock_powerpoint.Presentations.Open.side_effect = Exception("File is corrupted")
        
        # Test files
        test_files = ['corrupt_file.pptx']
        output_file = 'output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is False
        assert output_path == ""
        assert error_msg is not None
        assert "corrupt_file.pptx" in error_msg or "corrupted" in error_msg.lower()
        
        # Verify cleanup was attempted
        mock_destination.Close.assert_called_once()
        mock_powerpoint.Quit.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_powerpoint_initialization_error(self, mock_dispatch):
        """Test handling of PowerPoint initialization failure."""
        # Simulate PowerPoint not available
        mock_dispatch.side_effect = Exception("PowerPoint not installed")
        
        # Test files
        test_files = ['file1.pptx']
        output_file = 'output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is False
        assert output_path == ""
        assert error_msg is not None

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_save_error(self, mock_dispatch):
        """Test handling of save operation failure (e.g., permission denied)."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        mock_source = MagicMock()
        mock_source.Slides.Count = 2
        mock_source.Slides.Range.return_value = MagicMock()
        mock_powerpoint.Presentations.Open.return_value = mock_source
        
        # Simulate save error (permission denied)
        mock_destination.SaveAs.side_effect = Exception("Permission denied")
        
        # Test files
        test_files = ['file1.pptx']
        output_file = 'C:\\readonly\\output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is False
        assert output_path == ""
        assert error_msg is not None
        
        # Verify cleanup was attempted
        mock_destination.Close.assert_called_once()
        mock_source.Close.assert_called_once()
        mock_powerpoint.Quit.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_merge_with_empty_presentation(self, mock_dispatch):
        """Test merging presentations where one has no slides."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_destination = MagicMock()
        mock_destination.Slides.Count = 1
        mock_powerpoint.Presentations.Add.return_value = mock_destination
        
        # First presentation has slides, second is empty
        mock_source1 = MagicMock()
        mock_source1.Slides.Count = 3
        mock_source1.Slides.Range.return_value = MagicMock()
        
        mock_source2 = MagicMock()
        mock_source2.Slides.Count = 0  # Empty presentation
        
        mock_powerpoint.Presentations.Open.side_effect = [mock_source1, mock_source2]
        
        # Test files
        test_files = ['file1.pptx', 'empty_file.pptx']
        output_file = 'output.pptx'
        
        # Execute
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            test_files, output_file
        )
        
        # Verify
        assert success is True
        assert error_msg is None
        
        # Verify first file was processed normally
        assert mock_source1.Slides.Range.call_count == 1
        assert mock_destination.Slides.Paste.call_count == 1
        
        # Verify second file was skipped (no paste call for empty presentation)
        mock_source2.Slides.Range.assert_not_called()


class TestLaunchSlideshow:
    """Tests for the launch_slideshow function."""

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_successful_launch(self, mock_dispatch):
        """Test successful slideshow launch."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        mock_presentation = MagicMock()
        mock_powerpoint.Presentations.Open.return_value = mock_presentation
        
        # Execute
        success, error_msg = powerpoint_core.launch_slideshow('test.pptx')
        
        # Verify
        assert success is True
        assert error_msg is None
        
        # Verify PowerPoint was initialized
        mock_dispatch.assert_called_once_with("PowerPoint.Application")
        assert mock_powerpoint.Visible is True
        
        # Verify presentation was opened
        mock_powerpoint.Presentations.Open.assert_called_once()
        
        # Verify slideshow was started
        mock_presentation.SlideShowSettings.Run.assert_called_once()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_launch_file_not_found_error(self, mock_dispatch):
        """Test handling of missing presentation file."""
        # Setup mocks
        mock_powerpoint = MagicMock()
        mock_dispatch.return_value = mock_powerpoint
        
        # Simulate file not found
        mock_powerpoint.Presentations.Open.side_effect = Exception("File not found")
        
        # Execute
        success, error_msg = powerpoint_core.launch_slideshow('nonexistent.pptx')
        
        # Verify
        assert success is False
        assert error_msg is not None
        assert "nonexistent.pptx" in error_msg or "not found" in error_msg.lower()

    @patch('powerpoint_core.win32com.client.Dispatch')
    def test_launch_powerpoint_not_available(self, mock_dispatch):
        """Test handling when PowerPoint is not available."""
        # Simulate PowerPoint not available
        mock_dispatch.side_effect = Exception("PowerPoint not installed")
        
        # Execute
        success, error_msg = powerpoint_core.launch_slideshow('test.pptx')
        
        # Verify
        assert success is False
        assert error_msg is not None
