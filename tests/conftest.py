"""
Pytest configuration and fixtures for testing.

This file provides common fixtures and mocks for all tests.
"""
import sys
from unittest.mock import MagicMock

# Mock win32com before any imports
win32com_mock = MagicMock()
sys.modules['win32com'] = win32com_mock
sys.modules['win32com.client'] = win32com_mock.client
