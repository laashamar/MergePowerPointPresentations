"""
This module configures the logging for the application.
"""
import logging
import sys

def setup_logging():
    """
    Set up the root logger to output to both a file and the console.
    - The file will capture INFO level and higher messages.
    - The console will capture DEBUG level and higher messages for development.
    """
    # Get the root logger
    logger = logging.getLogger()
    # Set the lowest capture level to DEBUG to allow all messages to be processed
    logger.setLevel(logging.DEBUG)

    # --- File Handler ---
    # This handler writes logs to 'app.log'
    file_handler = logging.FileHandler('app.log', 'w', 'utf-8')
    file_handler.setLevel(logging.INFO)  # Log INFO and above to the file

    # --- Console (Stream) Handler ---
    # This new handler writes logs to the console (stderr)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG)  # Log DEBUG and above to the console

    # --- Formatter ---
    # Create a consistent format for all log messages
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # Apply the formatter to both handlers
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # --- Add Handlers to Logger ---
    # Clear any existing handlers to avoid duplication
    if logger.hasHandlers():
        logger.handlers.clear()

    # Add the configured handlers
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

if __name__ == '__main__':
    # Example usage to demonstrate the dual logging
    setup_logging()
    logging.debug("This is a debug message. It will only appear in the console.")
    logging.info("This is an info message. It will appear in both the console and app.log.")
    logging.warning("This is a warning message.")
    logging.error("This is an error message.")
    logging.critical("This is a critical message.")
