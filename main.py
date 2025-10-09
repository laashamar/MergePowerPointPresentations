# main.py

"""
This is the main entry point for the PowerPoint Merger application.
"""

import sys
from PySide6.QtWidgets import QApplication
from gui import MainWindow
from app import AppController
from app_logger import setup_logging


def main():
    """
    Initializes the application, sets up logging, and shows the main window.
    """
    setup_logging()
    app = QApplication(sys.argv)
    controller = AppController()
    window = MainWindow(controller)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
