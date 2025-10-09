# tests/test_powerpoint_core.py

"""
Tests for the PowerPoint core functionalities.
"""

import os
import pytest
from unittest.mock import MagicMock, patch, ANY
from powerpoint_core import PowerPointCore, PowerPointError
import comtypes

# Test data
TEST_FILE_1 = "test1.pptx"
TEST_FILE_2 = "test2.pptx"
OUTPUT_FILE = "merged.pptx"


@pytest.fixture
def mock_powerpoint_app():
    """Fixture for a mocked PowerPoint application object."""
    app = MagicMock()
    app.Presentations.Add.return_value = MagicMock()
    app.Presentations.Open.return_value = MagicMock()
    return app


@patch('powerpoint_core.comtypes.client')
def test_powerpoint_core_initialization_new_instance(mock_comtypes_client):
    """Test that a new PowerPoint instance is created if none is running."""
    mock_comtypes_client.GetActiveObject.side_effect = OSError
    core = PowerPointCore()
    mock_comtypes_client.CreateObject.assert_called_once_with("PowerPoint.Application")
    assert core.powerpoint is not None


@patch('powerpoint_core.comtypes.client')
def test_powerpoint_core_initialization_existing_instance(mock_comtypes_client):
    """Test connection to an existing PowerPoint instance."""
    mock_app = MagicMock()
    mock_comtypes_client.GetActiveObject.return_value = mock_app
    core = PowerPointCore()
    assert core.powerpoint == mock_app


@patch('powerpoint_core.comtypes.client')
def test_powerpoint_core_initialization_failure(mock_comtypes_client):
    """Test that PowerPointError is raised if PowerPoint cannot be started."""
    mock_comtypes_client.GetActiveObject.side_effect = OSError
    mock_comtypes_client.CreateObject.side_effect = OSError
    with pytest.raises(PowerPointError):
        PowerPointCore()


class TestMergePresentations:
    """Tests for the merge_presentations method."""

    @patch('powerpoint_core.comtypes.client')
    def test_merge_presentations_success(self, mock_comtypes_client, mock_powerpoint_app):
        """Test successful merging of presentations."""
        mock_comtypes_client.GetActiveObject.return_value = mock_powerpoint_app
        core = PowerPointCore()

        with patch('os.path.exists', return_value=True):
            core.merge_presentations([TEST_FILE_1, TEST_FILE_2], OUTPUT_FILE)

        mock_powerpoint_app.Presentations.Add.assert_called_once()
        base_presentation = mock_powerpoint_app.Presentations.Add.return_value
        assert base_presentation.Slides.InsertFromFile.call_count == 2
        base_presentation.SaveAs.assert_called_once_with(OUTPUT_FILE)
        base_presentation.Close.assert_called_once()

    @patch('powerpoint_core.comtypes.client')
    def test_merge_presentations_file_not_found(self, mock_comtypes_client, mock_powerpoint_app):
        """Test that FileNotFoundError is raised for non-existent files."""
        mock_comtypes_client.GetActiveObject.return_value = mock_powerpoint_app
        core = PowerPointCore()
        with pytest.raises(FileNotFoundError):
            core.merge_presentations(["non_existent.pptx"], OUTPUT_FILE)

    @patch('powerpoint_core.comtypes.client')
    def test_merge_presentations_handles_error(self, mock_comtypes_client, mock_powerpoint_app):
        """Test that PowerPointError is raised on COM errors during merge."""
        mock_powerpoint_app.Presentations.Add.return_value.Slides.InsertFromFile.side_effect = comtypes.COMError(
            -1, "Mock COM Error", "Mock description"
        )
        mock_comtypes_client.GetActiveObject.return_value = mock_powerpoint_app

        core = PowerPointCore()
        with patch('os.path.exists', return_value=True):
            with pytest.raises(PowerPointError):
                core.merge_presentations([TEST_FILE_1], OUTPUT_FILE)
