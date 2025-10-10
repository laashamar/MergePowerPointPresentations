# tests/test_app_logger.py

"""
Tests for the logging setup.
"""

import logging
from unittest.mock import patch, MagicMock
import app_logger


@patch('app_logger.logging.FileHandler')
@patch('app_logger.logging.StreamHandler')
def test_setup_logging_creates_handlers(mock_stream_handler, mock_file_handler):
    """
    Test that setup_logging creates FileHandler and StreamHandler.
    """
    app_logger.setup_logging()
    mock_file_handler.assert_called_once_with('app.log', 'w', 'utf-8')
    mock_stream_handler.assert_called_once()


@patch('app_logger.logging.FileHandler')
@patch('app_logger.logging.StreamHandler')
def test_setup_logging_configures_correctly(mock_stream_handler, mock_file_handler):
    """
    Test that logging is configured with the correct levels.
    """
    # Create mock handlers
    mock_file_handler_instance = MagicMock()
    mock_stream_handler_instance = MagicMock()
    mock_file_handler.return_value = mock_file_handler_instance
    mock_stream_handler.return_value = mock_stream_handler_instance
    
    app_logger.setup_logging()
    
    # Verify file handler level is set to INFO
    mock_file_handler_instance.setLevel.assert_called_once_with(logging.INFO)
    # Verify stream handler level is set to DEBUG
    mock_stream_handler_instance.setLevel.assert_called_once_with(logging.DEBUG)

