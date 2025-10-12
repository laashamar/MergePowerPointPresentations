import logging
import tkinter as tk
from logging import FileHandler, Handler
import os
from datetime import datetime

# Define the central log file path
LOG_FILE_PATH = os.path.join(os.path.expanduser("~"), "Downloads", "merge_powerpoint.log")

# Global list for collecting error messages
error_list = []

class TkinterLogHandler(Handler):
    """A custom log handler that sends records to a tkinter Text widget or CustomTkinter CTkTextbox."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        """Writes the log message to the Text widget."""
        msg = self.format(record)
        # Check if this is a CustomTkinter widget or regular tkinter Text widget
        try:
            # Try CustomTkinter method first
            if hasattr(self.text_widget, 'configure'):
                self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            if hasattr(self.text_widget, 'configure'):
                self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END) # Auto-scroll
        except Exception as e:
            # Fallback for any issues
            print(f"Error writing to log widget: {e}")

class ErrorListHandler(Handler):
    """A handler that collects all ERROR and CRITICAL messages in a list."""
    def emit(self, record):
        """Adds formatted error messages to the global list."""
        if record.levelno >= logging.ERROR:
            error_list.append(self.format(record))

def setup_logging(log_widget):
    """
    Configures logging to send output to GUI, file and an error list.
    """
    # Delete old log file if it exists for a clean start each run
    if os.path.exists(LOG_FILE_PATH):
        os.remove(LOG_FILE_PATH)

    log_format = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(module)s.%(funcName)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Handler for GUI window
    gui_handler = TkinterLogHandler(log_widget)
    gui_handler.setFormatter(log_format)
    root_logger.addHandler(gui_handler)

    # Handler for file
    file_handler = FileHandler(LOG_FILE_PATH, mode='w', encoding='utf-8')
    file_handler.setFormatter(log_format)
    root_logger.addHandler(file_handler)

    # Handler for collecting errors
    error_handler = ErrorListHandler()
    error_handler.setFormatter(log_format)
    root_logger.addHandler(error_handler)

    logging.info("Logging is set up.")
    logging.info(f"Log file saved to: {LOG_FILE_PATH}")


def write_log_summary():
    """Writes a summary of errors to the end of the log file."""
    try:
        with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
            summary_header = "\n\n" + "="*80 + "\n"
            summary_header += " ERROR SUMMARY ".center(80, "=") + "\n"
            summary_header += "="*80 + "\n\n"
            f.write(summary_header)

            if error_list:
                f.write(f"Found {len(error_list)} errors during execution:\n\n")
                for error in error_list:
                    f.write(error + "\n\n")
            else:
                f.write("No errors were logged during execution.\n")
            
            f.write("\n" + "="*80 + "\n")
        logging.info("Error summary has been written to the log file.")
    except Exception as e:
        logging.error(f"Could not write error summary to log file: {e}", exc_info=True)
