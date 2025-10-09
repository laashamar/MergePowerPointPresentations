"""
Tests for the logging configuration.
"""

import logging
from unittest.mock import patch
import logger


@patch('os.makedirs')
@patch('os.path.exists')
@patch('logging.basicConfig')
def test_setup_logging_creates_directory(mock_basic_config, mock_exists,
                                         mock_makedirs):
    """
    Test that setup_logging creates the log directory if it doesn't exist.
    """
    mock_exists.return_value = False
    logger.setup_logging()
    mock_makedirs.assert_called_once_with("logs", exist_ok=True)


@patch('os.path.exists', return_value=True)
@patch('logging.basicConfig')
def test_setup_logging_configures_correctly(mock_basic_config, mock_exists):
    """
    Test that logging.basicConfig is called with the correct parameters.
    """
    logger.setup_logging()
    # Check that basicConfig was called.
    mock_basic_config.assert_called_once()
    # Verify the logging level was set correctly.
    call_kwargs = mock_basic_config.call_args.kwargs
    assert call_kwargs['level'] == logging.INFO
    assert 'logs/app.log' in call_kwargs['filename']

