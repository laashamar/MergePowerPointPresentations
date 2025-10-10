"""
This file contains shared fixtures and configuration for the test suite.
"""
import pytest
from unittest.mock import MagicMock

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
    Creates and returns an instance of the MainWindow with a mocked controller.
    This fixture is used for testing the GUI in isolation.
    """
    # Create a mock controller to isolate the GUI from the app logic
    mock_controller = MagicMock(spec=AppController)

    # Pass the mock_controller directly into the constructor
    window = MainWindow(controller=mock_controller)

    # Register the widget with qtbot for interaction and garbage collection
    qtbot.addWidget(window)

    # Show the window before the test runs
    window.show()

    return window