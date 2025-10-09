"""
Unit tests for logger.py module.

Tests the logging infrastructure and configuration.
"""
import unittest
from unittest.mock import Mock, patch, MagicMock, mock_open
import logging
import tkinter as tk
import os
import tempfile

# Import the module to test
import logger


class TestTkinterLogHandler(unittest.TestCase):
    """Test cases for TkinterLogHandler class."""

    def setUp(self):
        """Set up test fixtures."""
        self.mock_text_widget = Mock()
        self.handler = logger.TkinterLogHandler(self.mock_text_widget)

    def test_initialization(self):
        """Test that handler initializes with text widget."""
        self.assertEqual(self.handler.text_widget, self.mock_text_widget)

    def test_emit_logs_to_text_widget(self):
        """Test that emit() writes to text widget."""
        # Create a log record
        record = logging.LogRecord(
            name='test',
            level=logging.INFO,
            pathname='',
            lineno=0,
            msg='Test message',
            args=(),
            exc_info=None
        )

        # Set up formatter
        formatter = logging.Formatter('%(message)s')
        self.handler.setFormatter(formatter)

        # Call emit
        self.handler.emit(record)

        # Verify text widget was updated
        self.mock_text_widget.configure.assert_called()
        self.mock_text_widget.insert.assert_called_once()
        self.mock_text_widget.see.assert_called_once()


class TestErrorListHandler(unittest.TestCase):
    """Test cases for ErrorListHandler class."""

    def setUp(self):
        """Set up test fixtures."""
        logger.error_list.clear()
        self.handler = logger.ErrorListHandler()
        formatter = logging.Formatter('%(levelname)s: %(message)s')
        self.handler.setFormatter(formatter)

    def tearDown(self):
        """Clean up after tests."""
        logger.error_list.clear()

    def test_emit_adds_error_to_list(self):
        """Test that ERROR level messages are added to error_list."""
        record = logging.LogRecord(
            name='test',
            level=logging.ERROR,
            pathname='',
            lineno=0,
            msg='Error message',
            args=(),
            exc_info=None
        )

        self.handler.emit(record)
        self.assertEqual(len(logger.error_list), 1)
        self.assertIn('ERROR', logger.error_list[0])

    def test_emit_adds_critical_to_list(self):
        """Test that CRITICAL level messages are added to error_list."""
        record = logging.LogRecord(
            name='test',
            level=logging.CRITICAL,
            pathname='',
            lineno=0,
            msg='Critical message',
            args=(),
            exc_info=None
        )

        self.handler.emit(record)
        self.assertEqual(len(logger.error_list), 1)

    def test_emit_ignores_info_level(self):
        """Test that INFO level messages are not added to error_list."""
        record = logging.LogRecord(
            name='test',
            level=logging.INFO,
            pathname='',
            lineno=0,
            msg='Info message',
            args=(),
            exc_info=None
        )

        self.handler.emit(record)
        self.assertEqual(len(logger.error_list), 0)


class TestSetupLogging(unittest.TestCase):
    """Test cases for setup_logging function."""

    @patch('logger.FileHandler')
    @patch('os.path.exists')
    @patch('os.remove')
    def test_setup_logging_configures_handlers(
            self, mock_remove, mock_exists, mock_file_handler):
        """Test that setup_logging configures all handlers correctly."""
        mock_exists.return_value = False
        mock_text_widget = Mock()

        # Clear existing handlers
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)

        logger.setup_logging(mock_text_widget)

        # Verify handlers were added
        self.assertGreaterEqual(len(root_logger.handlers), 3)

    @patch('logger.FileHandler')
    @patch('os.path.exists')
    @patch('os.remove')
    def test_setup_logging_removes_old_log_file(
            self, mock_remove, mock_exists, mock_file_handler):
        """Test that setup_logging removes old log file if it exists."""
        mock_exists.return_value = True
        mock_text_widget = Mock()

        logger.setup_logging(mock_text_widget)

        mock_remove.assert_called_once_with(logger.LOG_FILE_PATH)


class TestWriteLogSummary(unittest.TestCase):
    """Test cases for write_log_summary function."""

    def setUp(self):
        """Set up test fixtures."""
        logger.error_list.clear()
        self.temp_file = tempfile.NamedTemporaryFile(
            mode='w', delete=False, suffix='.log')
        self.temp_file.close()
        self.original_log_path = logger.LOG_FILE_PATH
        logger.LOG_FILE_PATH = self.temp_file.name

    def tearDown(self):
        """Clean up after tests."""
        logger.error_list.clear()
        logger.LOG_FILE_PATH = self.original_log_path
        if os.path.exists(self.temp_file.name):
            os.remove(self.temp_file.name)

    def test_write_log_summary_no_errors(self):
        """Test write_log_summary when there are no errors."""
        logger.write_log_summary()

        with open(self.temp_file.name, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn('ERROR SUMMARY', content)
        self.assertIn('No errors were logged', content)

    def test_write_log_summary_with_errors(self):
        """Test write_log_summary when there are errors."""
        logger.error_list.append('ERROR: Test error 1')
        logger.error_list.append('ERROR: Test error 2')

        logger.write_log_summary()

        with open(self.temp_file.name, 'r', encoding='utf-8') as f:
            content = f.read()

        self.assertIn('ERROR SUMMARY', content)
        self.assertIn('Found 2 errors', content)
        self.assertIn('Test error 1', content)
        self.assertIn('Test error 2', content)


if __name__ == '__main__':
    unittest.main()
