# tests/test_app_logger.py

"""
Tests for the logging setup.
"""

import logging
from unittest.mock import patch, MagicMock
import app_logger


@patch('app_logger.logging.FileHandler')
@patch('app_logger.logging.StreamHandler')
@patch('app_logger.os.makedirs')
@patch('app_logger.os.path.exists')
@patch('app_logger.logging.basicConfig')
def test_setup_logging_creates_directory(mock_basic_config, mock_exists,
                                         mock_makedirs, mock_stream_handler,
                                         mock_file_handler):
    """
    Test that setup_logging creates the log directory if it doesn't exist.
    """
    mock_exists.return_value = False
    app_logger.setup_logging()
    mock_makedirs.assert_called_once_with("logs")


@patch('app_logger.logging.FileHandler')
@patch('app_logger.logging.StreamHandler')
@patch('app_logger.os.path.exists', return_value=True)
@patch('app_logger.logging.basicConfig')
def test_setup_logging_configures_correctly(mock_basic_config, mock_exists,
                                            mock_stream_handler, mock_file_handler):
    """
    Test that logging.basicConfig is called with the correct parameters.
    """
    app_logger.setup_logging()
    assert mock_basic_config.called
    call_args = mock_basic_config.call_args
    assert call_args[1]['level'] == logging.INFO  # logging.INFO

