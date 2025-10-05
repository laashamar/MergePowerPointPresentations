import tkinter as tk
from tkinter import messagebox
import logging
import threading
import sys

# Import the new logger configuration and application starter
from logger import setup_logging, write_log_summary
from app import start_app

def run_main_application():
    """Runs the main application and catches any unknown errors."""
    try:
        logging.info("Starting PowerPoint Merger application workflow...")
        start_app()
        logging.info("Application workflow completed (GUI windows closed).")
    except Exception as e:
        # Catch unexpected errors during startup or execution
        logging.critical(
            "An unhandled error occurred in the application!", exc_info=True
        )
        # Show error message in a messagebox too, since GUI might be gone
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Critical Error",
            f"An unexpected error terminated the program.\n\n"
            f"Details: {e}\n\n"
            "Please check the log file for more information."
        )
        root.destroy()
    finally:
        # This runs after the app's mainloop is finished
        logging.info("Writing error summary to log file...")
        write_log_summary()

def main():
    """Main function to set up the log window and start the application."""
    # Main window for the logger
    log_window = tk.Tk()
    log_window.title("Live Log - PowerPoint Merger")
    log_window.geometry("900x600")

    # Set up a frame for better layout
    main_frame = tk.Frame(log_window, padx=10, pady=10)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    info_label = tk.Label(
        main_frame, 
        text="This window shows a live log of the script. Close this window to exit.",
        pady=5
    )
    info_label.pack(fill=tk.X)

    # Log widget with scrollbar
    log_frame = tk.Frame(main_frame)
    log_frame.pack(fill=tk.BOTH, expand=True)

    log_text = tk.Text(log_frame, state='disabled', wrap='word', font=("Courier New", 10), bg="#f0f0f0", fg="black")
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(log_frame, command=log_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    log_text['yscrollcommand'] = scrollbar.set

    # Configure logging to point to the Text widget
    setup_logging(log_text)

    # Override standard excepthook to log all unexpected errors
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logging.critical("Unhandled exception caught by sys.excepthook:", exc_info=(exc_type, exc_value, exc_traceback))

    sys.excepthook = handle_exception
    
    # Run main application in a separate thread to avoid GUI freezing
    app_thread = threading.Thread(target=run_main_application, daemon=True)
    app_thread.start()

    # Main loop for the log window
    log_window.mainloop()

if __name__ == "__main__":
    main()
