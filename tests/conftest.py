"""
This module contains shared fixtures for the pytest test suite.

Fixtures defined here are automatically discovered by pytest and can be used in any test
function within the project's test suite without needing to be imported.

This conftest.py sets up the necessary environment for testing the PySide6 GUI application.
It includes fixtures to:
- Ensure the project's root directory is on the Python path for correct module imports.
- Provide a QApplication instance for GUI tests.
- Create a reusable instance of the application's MainWindow.
- Mock file dialogs to prevent them from appearing during tests and to simulate user input.
"""
import sys
from pathlib import Path
import pytest
from unittest.mock import patch, MagicMock

# Add the project root to the Python path to allow for absolute imports
# This makes it possible to import modules like 'gui' and 'powerpoint_core' directly
root_dir = Path(__file__).resolve().parent.parent
sys.path.append(str(root_dir))

from PySide6.QtWidgets import QApplication
from gui import MainWindow
from powerpoint_core import PowerPointMerger # Corrected import

@pytest.fixture(scope="session")
def qapp():
    """Fixture to create a QApplication instance for the test session."""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    yield app
    app.quit()

@pytest.fixture
def main_app(qapp, mocker):
    """Fixture to create an instance of the MainWindow for each test."""
    # Mock the Merger class within the gui module's namespace
    # This prevents the actual PowerPoint COM logic from being triggered during tests
    mocker.patch('gui.PowerPointMerger', autospec=True) # Corrected mock target
    window = MainWindow()
    yield window
    window.close()


@pytest.fixture
def mock_file_dialog(mocker):
    """Fixture to mock the QFileDialog.getOpenFileNames method."""
    mock_dialog = mocker.patch('gui.QFileDialog.getOpenFileNames')
    # Simulate the user selecting two files
    mock_dialog.return_value = (['a.pptx', 'b.pptx'], 'PowerPoint Presentations (*.pptx)')
    return mock_dialog

