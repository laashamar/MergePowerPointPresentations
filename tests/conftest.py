"""
This file contains shared fixtures and configuration for the test suite.
"""
import sys
from unittest.mock import MagicMock

# Mock comtypes for non-Windows platforms before any test modules are loaded
if sys.platform != 'win32':
    # Create a mock comtypes module
    mock_comtypes = MagicMock()
    mock_comtypes_client = MagicMock()
    
    # Add COMError class to the mock
    class MockCOMError(Exception):
        def __init__(self, hresult, text, details):
            super().__init__(text)
            self.hresult = hresult
            self.text = text
            self.details = details
    
    mock_comtypes.COMError = MockCOMError
    mock_comtypes.CoInitialize = MagicMock()
    mock_comtypes.CoUninitialize = MagicMock()
    
    sys.modules['comtypes'] = mock_comtypes
    sys.modules['comtypes.client'] = mock_comtypes_client

# Register the pytest-qt plugin only if PySide6 is available
try:
    import PySide6
    pytest_plugins = "pytestqt"
except ImportError:
    pass

