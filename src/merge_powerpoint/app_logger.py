"""Logging configuration module for the PowerPoint Merger application.

This module provides centralized logging configuration for the application,
supporting both file and console logging with appropriate formatting.
"""

import logging
import os


def setup_logging():
    """Set up logging configuration for the application.

    Configures the logging system to output INFO level and higher messages
    to both a log file and the console. Creates the logs directory if it
    doesn't exist.

    Returns:
        logging.Logger: The configured root logger instance.
    """
    # Create logs directory if it doesn't exist
    if not os.path.exists("logs"):
        os.makedirs("logs")

    # Configure logging with basicConfig
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler("logs/app.log", mode="w", encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,  # Force reconfiguration even if already configured
    )

    return logging.getLogger()


if __name__ == "__main__":
    # Example usage to demonstrate the dual logging
    setup_logging()
    logging.debug("This is a debug message.")
    logging.info("This is an info message.")
    logging.warning("This is a warning message.")
    logging.error("This is an error message.")
    logging.critical("This is a critical message.")
