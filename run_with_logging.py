"""Run the main application with exception logging.

This script runs the main application and logs any exceptions that occur.
It now uses the refactored package structure from src/merge_powerpoint.
"""

import logging
import sys
from pathlib import Path

# Add src to path for imports
src_path = Path(__file__).parent / "src"
if str(src_path) not in sys.path:
    sys.path.insert(0, str(src_path))

from main import main  # noqa: E402
from merge_powerpoint.app_logger import setup_logging  # noqa: E402

# Set up logging first
setup_logging()

if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as e:
        logging.critical("An unhandled exception occurred: %s", e, exc_info=True)
        sys.exit(1)
