# tests/test_app_logger.py

"""
Tests for the logging setup.
"""

import logging
from unittest.mock import patch
import app_logger


@patch('os.makedirs')
@patch('os.path.exists')
@patch('logging.basicConfig')
def test_setup_logging_creates_directory(mock_basic_config, mock_exists,
                                         mock_makedirs):
    """
    Test that setup_logging creates the log directory if it doesn't exist.
    """
    mock_exists.return_value = False
    app_logger.setup_logging()
    mock_makedirs.assert_called_once_with("logs")


@patch('os.path.exists', return_value=True)
@patch('logging.basicConfig')
def test_setup_logging_configures_correctly(mock_basic_config, mock_exists):
    """
    Test that logging.basicConfig is called with the correct parameters.
    """
    app_logger.setup_logging()
    assert mock_basic_config.called
    call_args = mock_basic_config.call_args
    assert call_args[1]['level'] == logging.INFO  # logging.INFO

