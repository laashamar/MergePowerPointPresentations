"""
Development entry point for testing the PowerPoint Merging Tool GUI.

This script is used for development and testing purposes. It imports the main
application window from the `gui` module and runs it. The key functionality
is modifying `sys.path` to allow imports from the parent directory.
"""
import sys
import os
import logging

# --- FIX FOR ModuleNotFoundError ---
# Get the absolute path of the directory containing this script (new_gui).
script_dir = os.path.dirname(os.path.abspath(__file__))
# Get the path of the parent directory (the project root).
project_root = os.path.dirname(script_dir)
# Add the project root to Python's path, making modules in that directory importable.
sys.path.insert(0, project_root)
# ---------------------------------

# Now that the project root is on the path, these imports will succeed.
from gui import MainApplication
from logger import setup_logging

def main():
    """
    Initializes logging and runs the GUI application for testing.
    """
    setup_logging()
    logging.info("Starting the GUI in development/test mode from main_gui.py.")
    try:
        app = MainApplication()
        app.mainloop()
        logging.info("GUI application closed normally.")
    except Exception as e:
        logging.critical("The GUI application encountered a fatal error.", exc_info=True)
        # In a real-world scenario, you might show an error dialog here.

if __name__ == "__main__":
    main()

