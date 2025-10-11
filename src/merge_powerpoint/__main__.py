"""Main entry point for the PowerPoint Merger application.

This module serves as the entry point when the package is run as a module
using `python -m merge_powerpoint` or via the CLI command.
"""

import sys

from PySide6.QtWidgets import QApplication, QMainWindow

from merge_powerpoint.app_logger import setup_logging
from merge_powerpoint.gui import MainUI


def main():
    """Initialize and run the PowerPoint Merger application.

    Sets up logging, creates the application instance, and displays
    the main window.

    Returns:
        int: Application exit code.
    """
    setup_logging()
    app = QApplication(sys.argv)
    app.setApplicationName("PowerPoint Merger")
    app.setOrganizationName("MergePowerPoint")

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
