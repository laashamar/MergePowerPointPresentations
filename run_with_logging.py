# run_with_logging.py

"""
This script runs the main application and logs any exceptions that occur.
"""

import logging
from app_logger import setup_logging
from main import main

# Set up logging first
setup_logging()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.critical("An unhandled exception occurred: %s", e, exc_info=True)

