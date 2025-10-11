"""Main entry point for the PowerPoint Merger application.

This is the main entry point that initializes and runs the application.
It now uses the refactored package structure from src/merge_powerpoint.
"""

import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from PySide6.QtWidgets import QApplication, QMainWindow  # noqa: E402

from merge_powerpoint.app import AppController  # noqa: E402
from merge_powerpoint.app_logger import setup_logging  # noqa: E402
from merge_powerpoint.gui import MainUI  # noqa: E402


def main():
    """Initialize the application, set up logging, and show the main window.

    Returns:
        int: Application exit code.
    """
    setup_logging()
    app = QApplication(sys.argv)
    app.setApplicationName("PowerPoint Merger")
    app.setOrganizationName("MergePowerPoint")
    
    # AppController is still available but MainUI creates its own merger
    controller = AppController()
    
    # MainUI is a QWidget, so embed it in a QMainWindow
    main_window = QMainWindow()
    ui = MainUI()
    main_window.setCentralWidget(ui)
    main_window.setWindowTitle("PowerPoint Presentation Merger")
    main_window.resize(1000, 600)
    main_window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
