"""
This module configures the logging for the application.
"""
import logging
import os

def setup_logging():
    """
    Set up logging configuration for the application.
    - Logs INFO level and higher messages.
    - Outputs to both file and console.
    """
    # Create logs directory if it doesn't exist
    if not os.path.exists("logs"):
        os.makedirs("logs")
    
    # Configure logging with basicConfig
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('logs/app.log', mode='w', encoding='utf-8'),
            logging.StreamHandler()
        ],
        force=True  # Force reconfiguration even if logging has already been configured
    )

if __name__ == '__main__':
    # Example usage to demonstrate the dual logging
    setup_logging()
    logging.debug("This is a debug message. It will only appear in the console.")
    logging.info("This is an info message. It will appear in both the console and app.log.")
    logging.warning("This is a warning message.")
    logging.error("This is an error message.")
    logging.critical("This is a critical message.")
