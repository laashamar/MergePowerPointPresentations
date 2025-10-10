import sys
from PySide6.QtWidgets import QApplication

# CORRECTED: The logger file is named app_logger.
from .app_logger import setup_logging
from .app import AppController

def main():
    """
    Main function to initialize and run the application.
    """
    setup_logging()
    app = QApplication(sys.argv)
    
    controller = AppController()
    controller.show_main_window()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()

