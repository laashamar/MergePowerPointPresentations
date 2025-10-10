"""Main entry point for the PowerPoint Merger application.

This module serves as the entry point when the package is run as a module
using `python -m merge_powerpoint` or via the CLI command.
"""

import sys

from PySide6.QtWidgets import QApplication

from merge_powerpoint.app_logger import setup_logging
from merge_powerpoint.gui import MainWindow


def main():
    """Initialize and run the PowerPoint Merger application.

    Sets up logging, creates the application instance, and displays
    the main window.

    Returns:
        int: Application exit code.
    """
    setup_logging()
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
