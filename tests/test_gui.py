"""
Unit tests for gui module.

This module tests the PowerPointMergerGUI class using mocking to avoid
actually displaying windows during tests.
"""
import os
import sys
import pytest
from unittest.mock import Mock, MagicMock, patch, call

# Mock tkinter and tkinterdnd2 before importing gui
sys.modules['tkinter'] = MagicMock()
sys.modules['tkinter.filedialog'] = MagicMock()
sys.modules['tkinter.messagebox'] = MagicMock()
sys.modules['tkinterdnd2'] = MagicMock()

import tkinter as tk
import gui


class TestPowerPointMergerGUI:
    """Tests for the PowerPointMergerGUI class."""

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)  # Disable drag-and-drop for testing
    def test_initialization(self, mock_tk):
        """Test GUI initialization with callback."""
        mock_callback = Mock()
        
        # Create GUI instance
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Verify initialization
        assert gui_instance.merge_callback == mock_callback
        assert gui_instance.file_list == []
        assert gui_instance.root is not None

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_add_valid_pptx_file(self, mock_tk):
        """Test adding a valid .pptx file to the queue."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create a temporary test file
        test_file = '/tmp/test_file.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Verify file was added
            assert test_file in gui_instance.file_list
            assert len(gui_instance.file_list) == 1
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_add_valid_ppsx_file(self, mock_tk):
        """Test adding a valid .ppsx file to the queue."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create a temporary test file
        test_file = '/tmp/test_file.ppsx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Verify file was added
            assert test_file in gui_instance.file_list
            assert len(gui_instance.file_list) == 1
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.messagebox.showwarning')
    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_reject_invalid_file_type(self, mock_tk, mock_showwarning):
        """Test that invalid file types are rejected."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Try to add an invalid file type
        gui_instance._add_files(['/tmp/invalid.txt'])
        
        # Verify file was not added
        assert len(gui_instance.file_list) == 0
        
        # Verify warning was shown
        mock_showwarning.assert_called_once()
        assert "Invalid File Type" in str(mock_showwarning.call_args)

    @patch('gui.messagebox.showinfo')
    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_reject_duplicate_files(self, mock_tk, mock_showinfo):
        """Test that duplicate files are rejected."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create a temporary test file
        test_file = '/tmp/test_duplicate.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file twice
            gui_instance._add_files([test_file])
            gui_instance._add_files([test_file])
            
            # Verify file was added only once
            assert len(gui_instance.file_list) == 1
            
            # Verify duplicate message was shown
            mock_showinfo.assert_called_once()
            assert "Duplicate File" in str(mock_showinfo.call_args)
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_remove_file_from_queue(self, mock_tk):
        """Test removing a file from the queue."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        for f in [test_file1, test_file2]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2])
            assert len(gui_instance.file_list) == 2
            
            # Remove first file
            gui_instance._remove_file(0)
            
            # Verify file was removed
            assert len(gui_instance.file_list) == 1
            assert test_file1 not in gui_instance.file_list
            assert test_file2 in gui_instance.file_list
        finally:
            # Cleanup
            for f in [test_file1, test_file2]:
                if os.path.exists(f):
                    os.remove(f)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_move_file_up(self, mock_tk):
        """Test moving a file up in the queue."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        test_file3 = '/tmp/test3.pptx'
        for f in [test_file1, test_file2, test_file3]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2, test_file3])
            
            # Move second file up
            gui_instance._move_file_up(1)
            
            # Verify order changed
            assert gui_instance.file_list[0] == test_file2
            assert gui_instance.file_list[1] == test_file1
            assert gui_instance.file_list[2] == test_file3
        finally:
            # Cleanup
            for f in [test_file1, test_file2, test_file3]:
                if os.path.exists(f):
                    os.remove(f)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_move_file_down(self, mock_tk):
        """Test moving a file down in the queue."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        test_file3 = '/tmp/test3.pptx'
        for f in [test_file1, test_file2, test_file3]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2, test_file3])
            
            # Move first file down
            gui_instance._move_file_down(0)
            
            # Verify order changed
            assert gui_instance.file_list[0] == test_file2
            assert gui_instance.file_list[1] == test_file1
            assert gui_instance.file_list[2] == test_file3
        finally:
            # Cleanup
            for f in [test_file1, test_file2, test_file3]:
                if os.path.exists(f):
                    os.remove(f)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_move_file_up_at_boundary(self, mock_tk):
        """Test moving first file up (should not move)."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        for f in [test_file1, test_file2]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2])
            
            # Try to move first file up (should not move)
            gui_instance._move_file_up(0)
            
            # Verify order unchanged
            assert gui_instance.file_list[0] == test_file1
            assert gui_instance.file_list[1] == test_file2
        finally:
            # Cleanup
            for f in [test_file1, test_file2]:
                if os.path.exists(f):
                    os.remove(f)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_move_file_down_at_boundary(self, mock_tk):
        """Test moving last file down (should not move)."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        for f in [test_file1, test_file2]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2])
            
            # Try to move last file down (should not move)
            gui_instance._move_file_down(1)
            
            # Verify order unchanged
            assert gui_instance.file_list[0] == test_file1
            assert gui_instance.file_list[1] == test_file2
        finally:
            # Cleanup
            for f in [test_file1, test_file2]:
                if os.path.exists(f):
                    os.remove(f)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_output_path_construction(self, mock_tk):
        """Test that output path is constructed correctly."""
        mock_callback = Mock()
        
        # Create mock for StringVar that properly tracks set/get
        output_folder_mock = MagicMock()
        output_folder_mock.get.return_value = '/tmp'
        
        output_filename_mock = MagicMock()
        output_filename_mock.get.return_value = 'merged_output.pptx'
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.output_folder_var = output_folder_mock
        gui_instance.output_filename_var = output_filename_mock
        
        # Create a temporary test file
        test_file = '/tmp/test.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Trigger merge (which constructs the path)
            with patch('gui.messagebox.askyesno', return_value=True):
                gui_instance._on_merge()
            
            # Verify callback was called with correct path
            expected_path = '/tmp/merged_output.pptx'
            mock_callback.assert_called_once()
            args = mock_callback.call_args[0]
            assert args[1] == expected_path
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_output_path_appends_pptx_extension(self, mock_tk):
        """Test that .pptx extension is appended if missing."""
        mock_callback = Mock()
        
        # Create mock for StringVar that properly tracks set/get
        output_folder_mock = MagicMock()
        output_folder_mock.get.return_value = '/tmp'
        
        output_filename_mock = MagicMock()
        output_filename_mock.get.return_value = 'merged_output'
        output_filename_mock.set = MagicMock()
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.output_folder_var = output_folder_mock
        gui_instance.output_filename_var = output_filename_mock
        
        # Create a temporary test file
        test_file = '/tmp/test.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Trigger merge
            with patch('gui.messagebox.askyesno', return_value=True):
                gui_instance._on_merge()
            
            # Verify .pptx was appended via set call
            output_filename_mock.set.assert_called_with('merged_output.pptx')
            
            # Verify callback was called
            mock_callback.assert_called_once()
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_merge_callback_invocation(self, mock_tk):
        """Test that merge callback is called with correct parameters."""
        mock_callback = Mock()
        
        # Create mock for StringVar
        output_folder_mock = MagicMock()
        output_folder_mock.get.return_value = '/tmp'
        
        output_filename_mock = MagicMock()
        output_filename_mock.get.return_value = 'output.pptx'
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.output_folder_var = output_folder_mock
        gui_instance.output_filename_var = output_filename_mock
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        for f in [test_file1, test_file2]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2])
            
            # Trigger merge
            with patch('gui.messagebox.askyesno', return_value=True):
                gui_instance._on_merge()
            
            # Verify callback was called with correct parameters
            mock_callback.assert_called_once()
            file_list, output_path = mock_callback.call_args[0]
            
            assert len(file_list) == 2
            assert test_file1 in file_list
            assert test_file2 in file_list
            assert output_path == '/tmp/output.pptx'
        finally:
            # Cleanup
            for f in [test_file1, test_file2]:
                if os.path.exists(f):
                    os.remove(f)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_merge_button_disabled_when_queue_empty(self, mock_tk):
        """Test that merge button is disabled when queue is empty."""
        mock_callback = Mock()
        
        # Create a mock button with proper config method
        mock_button = MagicMock()
        mock_button.config = MagicMock()
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.merge_btn = mock_button
        
        # Call the update method which should disable the button
        gui_instance._update_merge_queue_display()
        
        # Verify button was configured to be disabled
        mock_button.config.assert_called_with(state=tk.DISABLED)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_merge_button_enabled_when_files_added(self, mock_tk):
        """Test that merge button is enabled when files are added."""
        mock_callback = Mock()
        
        # Create a mock button with proper config method
        mock_button = MagicMock()
        mock_button.config = MagicMock()
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.merge_btn = mock_button
        
        # Create a temporary test file
        test_file = '/tmp/test.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Verify button was configured to be enabled
            mock_button.config.assert_called_with(state=tk.NORMAL)
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_merge_button_disabled_after_removing_all_files(self, mock_tk):
        """Test that merge button is disabled after removing all files."""
        mock_callback = Mock()
        
        # Create a mock button with proper config method
        mock_button = MagicMock()
        mock_button.config = MagicMock()
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.merge_btn = mock_button
        
        # Create a temporary test file
        test_file = '/tmp/test.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Remove file
            gui_instance._remove_file(0)
            
            # Verify merge button is disabled
            # Last call should be to disable
            calls = mock_button.config.call_args_list
            assert any(call[1].get('state') == tk.DISABLED for call in calls)
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.messagebox.showwarning')
    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_merge_with_empty_queue_shows_warning(self, mock_tk, mock_showwarning):
        """Test that merging with empty queue shows warning."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Try to merge with empty queue
        gui_instance._on_merge()
        
        # Verify warning was shown
        mock_showwarning.assert_called_once()
        assert "No Files" in str(mock_showwarning.call_args)
        
        # Verify callback was not called
        mock_callback.assert_not_called()

    @patch('gui.messagebox.showerror')
    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_merge_with_empty_filename_shows_error(self, mock_tk, mock_showerror):
        """Test that merging with empty filename shows error."""
        mock_callback = Mock()
        
        # Create mock for StringVar
        output_filename_mock = MagicMock()
        output_filename_mock.get.return_value = ''
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.output_filename_var = output_filename_mock
        
        # Create a temporary test file
        test_file = '/tmp/test.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            # Add file to queue
            gui_instance._add_files([test_file])
            
            # Try to merge
            gui_instance._on_merge()
            
            # Verify error was shown
            mock_showerror.assert_called_once()
            assert "Invalid Filename" in str(mock_showerror.call_args)
            
            # Verify callback was not called
            mock_callback.assert_not_called()
        finally:
            # Cleanup
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.messagebox.showerror')
    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_reject_nonexistent_file(self, mock_tk, mock_showerror):
        """Test that nonexistent files are rejected."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Try to add a nonexistent file
        gui_instance._add_files(['/tmp/nonexistent_file.pptx'])
        
        # Verify file was not added
        assert len(gui_instance.file_list) == 0
        
        # Verify error was shown
        mock_showerror.assert_called_once()
        assert "File Not Found" in str(mock_showerror.call_args)

    @patch('gui.messagebox.askyesno', return_value=True)
    @patch('gui.tk.Tk')
    @patch('gui.HAS_DND', False)
    def test_clear_queue(self, mock_tk, mock_askyesno):
        """Test clearing the queue."""
        mock_callback = Mock()
        
        # Create a mock button with proper config method
        mock_button = MagicMock()
        mock_button.config = MagicMock()
        
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        gui_instance.merge_btn = mock_button
        
        # Create temporary test files
        test_file1 = '/tmp/test1.pptx'
        test_file2 = '/tmp/test2.pptx'
        for f in [test_file1, test_file2]:
            with open(f, 'w') as file:
                file.write('test')
        
        try:
            # Add files to queue
            gui_instance._add_files([test_file1, test_file2])
            assert len(gui_instance.file_list) == 2
            
            # Clear queue
            gui_instance._clear_queue()
            
            # Verify queue was cleared
            assert len(gui_instance.file_list) == 0
            
            # Verify merge button was disabled (check call list)
            calls = mock_button.config.call_args_list
            assert any(call[1].get('state') == tk.DISABLED for call in calls)
            
            # Verify confirmation was requested
            mock_askyesno.assert_called_once()
        finally:
            # Cleanup
            for f in [test_file1, test_file2]:
                if os.path.exists(f):
                    os.remove(f)
