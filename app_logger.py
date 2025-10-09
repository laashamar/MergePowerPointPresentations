# app_logger.py

"""
Configures logging for the application.
"""

import logging
import os
from logging.handlers import RotatingFileHandler

def setup_logging():
    """
    Sets up logging to file and console.
    """
    log_dir = "logs"
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    log_file = os.path.join(log_dir, "app.log")

    # Configure root logger
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(name)s - %(message)s",
        handlers=[
            # File handler
            RotatingFileHandler(log_file, maxBytes=1048576, backupCount=5),
            # Console handler
            logging.StreamHandler()
        ]
    )

