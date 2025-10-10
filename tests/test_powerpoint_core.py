from unittest.mock import patch, MagicMock, call
import pytest

# MODIFIED: Import functions from the package, not the old class
from merge_powerpoint.powerpoint_core import merge_presentations, find_layout_and_placeholder

# Mock the Presentation class from the python-pptx library
@patch('pptx.Presentation')
def test_merge_presentations_success(mock_presentation_cls):
    """
    Test the happy path for the merge_presentations function.
    """
    # --- Setup Mocks ---
    mock_source_pres1 = MagicMock()
    mock_source_pres1.slides = [MagicMock(), MagicMock()]  # 2 slides
    mock_source_pres2 = MagicMock()
    mock_source_pres2.slides = [MagicMock()]  # 1 slide
    
    mock_dest_pres = MagicMock()
    mock_slide = mock_dest_pres.slides.add_slide.return_value
    mock_slide.shapes = [] # Simulate an empty slide
    
    # Configure the mocked Presentation class to return our mocks
    # We now mock it once for the initial slide count, then for the actual merge
    mock_presentation_cls.side_effect = [
        mock_source_pres1, mock_source_pres2, # For slide count
        mock_dest_pres, mock_source_pres1, mock_source_pres2 # For merging
    ]
    
    mock_callback = MagicMock()
    
    source_files = ['file1.pptx', 'file2.pptx']
    output_file = 'output.pptx'
    
    # --- Run the function ---
    merge_presentations(source_files, output_file, progress_callback=mock_callback)

    # --- Assertions ---
    assert mock_presentation_cls.call_count == 5
    
    assert mock_dest_pres.slides.add_slide.call_count == 3
    
    # Total slides = 3. Should be called after each slide, plus once for saving.
    assert mock_callback.call_count == 4
    mock_callback.assert_has_calls([
        call(33), call(66), call(100), call(100)
    ])
    
    mock_dest_pres.save.assert_called_once_with(output_file)

def test_find_layout_and_placeholder_found():
    """
    Test finding a valid layout and placeholder.
    """
    mock_placeholder = MagicMock()
    mock_placeholder.name = "Content Placeholder 1"
    
    mock_slide_layout = MagicMock()
    mock_slide_layout.placeholders = [mock_placeholder]
    
    mock_presentation = MagicMock()
    mock_presentation.slide_layouts = [mock_slide_layout]
    
    layout, placeholder_idx = find_layout_and_placeholder(mock_presentation, "Content Placeholder 1")
    
    assert layout is mock_slide_layout
    assert placeholder_idx == 0

def test_find_layout_and_placeholder_not_found():
    """
    Test behavior when the placeholder cannot be found.
    """
    mock_presentation = MagicMock()
    mock_presentation.slide_layouts = []

    layout, placeholder_idx = find_layout_and_placeholder(mock_presentation, "NonExistent")

    assert layout is None
    assert placeholder_idx is None

