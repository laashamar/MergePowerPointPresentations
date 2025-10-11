"""
This file contains shared fixtures and configuration for the test suite.
"""
import os
import sys
from unittest.mock import MagicMock

# Set Qt platform for headless testing on Linux/CI
if sys.platform != 'win32' and 'DISPLAY' not in os.environ:
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

# Mock comtypes for non-Windows platforms before any imports
if sys.platform != 'win32':
    sys.modules['comtypes'] = MagicMock()
    sys.modules['comtypes.client'] = MagicMock()

import pytest

from merge_powerpoint.app import AppController
from merge_powerpoint.gui import MainUI

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
    Creates and returns an instance of the MainUI.
    This fixture is used for testing the GUI in isolation.
    """
    # MainUI is a QWidget that creates its own PowerPointMerger
    ui = MainUI()

    # Register the widget with qtbot for interaction and garbage collection
    qtbot.addWidget(ui)

    # Show the widget before the test runs
    ui.show()

    return ui
