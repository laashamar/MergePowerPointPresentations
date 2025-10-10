"""
This file contains shared fixtures and configuration for the test suite.
"""
import sys
from unittest.mock import MagicMock

# Mock comtypes for non-Windows platforms before any imports
if sys.platform != 'win32':
    sys.modules['comtypes'] = MagicMock()
    sys.modules['comtypes.client'] = MagicMock()

import pytest

from merge_powerpoint.app import AppController
from merge_powerpoint.gui import MainWindow

# Register the pytest-qt plugin.
pytest_plugins = "pytestqt"


@pytest.fixture
def app_controller(qapp):
    """
    Returns a clean AppController instance for each test.
    Depends on qapp to ensure QApplication is running.
    """
    return AppController()


@pytest.fixture
def main_window(qtbot):
    """
    Creates and returns an instance of the MainWindow.
    This fixture is used for testing the GUI in isolation.
    """
    # MainWindow no longer takes a controller argument - it creates its own PowerPointMerger
    window = MainWindow()

    # Register the widget with qtbot for interaction and garbage collection
    qtbot.addWidget(window)

    # Show the window before the test runs
    window.show()

    return window
