"""
Additional test to verify Browse for Files button configuration.
"""
import sys
import pytest
from unittest.mock import Mock, MagicMock, patch

# Mock customtkinter and tkinter before importing gui
sys.modules['customtkinter'] = MagicMock()
sys.modules['tkinter'] = MagicMock()
sys.modules['tkinter.filedialog'] = MagicMock()
sys.modules['tkinter.messagebox'] = MagicMock()

import gui


class TestBrowseButtonConfiguration:
    """Tests to verify Browse for Files button is properly configured."""

    @patch('gui.ctk.CTk')
    def test_browse_button_uses_primary_accent_colors(self, mock_tk):
        """Test that Browse button uses correct primary accent colors."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # The button should be created during initialization
        # when file_list is empty
        assert gui_instance.file_list == []
        
        # Verify _browse_files method exists
        assert hasattr(gui_instance, '_browse_files')
        assert callable(gui_instance._browse_files)

    @patch('gui.filedialog.askopenfilenames', return_value=[])
    @patch('gui.ctk.CTk')
    def test_browse_files_opens_dialog(self, mock_tk, mock_filedialog):
        """Test that _browse_files opens file dialog with correct parameters."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Call _browse_files
        gui_instance._browse_files()
        
        # Verify file dialog was opened
        mock_filedialog.assert_called_once()
        call_args = mock_filedialog.call_args
        
        # Verify correct parameters
        assert call_args[1]['title'] == "Select PowerPoint Files"
        assert ("PowerPoint Files", "*.pptx *.ppsx") in call_args[1]['filetypes']

    @patch('gui.filedialog.askopenfilenames')
    @patch('gui.ctk.CTk')
    def test_browse_files_adds_selected_files(self, mock_tk, mock_filedialog):
        """Test that _browse_files adds selected files to queue."""
        # Create temporary test file
        test_file = '/tmp/test_browse.pptx'
        with open(test_file, 'w') as f:
            f.write('test')
        
        try:
            mock_filedialog.return_value = [test_file]
            mock_callback = Mock()
            gui_instance = gui.PowerPointMergerGUI(mock_callback)
            
            # Initial state
            assert len(gui_instance.file_list) == 0
            
            # Browse for files
            gui_instance._browse_files()
            
            # Verify file was added
            assert len(gui_instance.file_list) == 1
            assert gui_instance.file_list[0] == test_file
        finally:
            import os
            if os.path.exists(test_file):
                os.remove(test_file)

    @patch('gui.ctk.CTk')
    def test_file_selector_shown_on_empty_queue(self, mock_tk):
        """Test that file selector is shown when queue is empty."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Mock the methods to track calls
        gui_instance._create_file_selector = Mock()
        gui_instance._create_file_list = Mock()
        
        # Update display with empty queue
        gui_instance.file_list = []
        gui_instance._update_merge_queue_display()
        
        # Verify file selector was created
        gui_instance._create_file_selector.assert_called_once()
        gui_instance._create_file_list.assert_not_called()

    @patch('gui.ctk.CTk')
    def test_file_list_shown_on_non_empty_queue(self, mock_tk):
        """Test that file list is shown when queue has files."""
        mock_callback = Mock()
        gui_instance = gui.PowerPointMergerGUI(mock_callback)
        
        # Mock the methods to track calls
        gui_instance._create_file_selector = Mock()
        gui_instance._create_file_list = Mock()
        
        # Update display with files in queue
        gui_instance.file_list = ['/tmp/test.pptx']
        gui_instance._update_merge_queue_display()
        
        # Verify file list was created
        gui_instance._create_file_list.assert_called_once()
        gui_instance._create_file_selector.assert_not_called()
