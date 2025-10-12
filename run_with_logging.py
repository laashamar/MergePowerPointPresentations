import customtkinter as ctk
from tkinter import messagebox
import logging
import threading
import sys

# Import the new logger configuration and application starter
from logger import setup_logging, write_log_summary
from app import start_app

# PowerPoint-inspired Dark Mode Color Palette
COLORS = {
    'primary_accent': '#d35230',
    'accent_hover': '#ba3416',
    'window_bg': '#242424',
    'frame_bg': '#2b2b2b',
    'primary_text': '#e5e5e5',
    'secondary_text': '#a0a0a0',
    'button_text': '#ffffff',
}

# Font settings
FONT_FAMILY = "Helvetica"
FONT_SIZE_MEDIUM = 12
FONT_SIZE_SMALL = 10

# Set CustomTkinter appearance mode
ctk.set_appearance_mode("dark")

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
        root = ctk.CTk()
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
    log_window = ctk.CTk()
    log_window.title("Live Log - PowerPoint Merger")
    log_window.geometry("900x600")
    log_window.configure(fg_color=COLORS['window_bg'])

    # Set up a frame for better layout
    main_frame = ctk.CTkFrame(log_window, fg_color=COLORS['window_bg'])
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    info_label = ctk.CTkLabel(
        main_frame, 
        text="This window shows a live log of the script. Close this window to exit.",
        font=(FONT_FAMILY, FONT_SIZE_SMALL),
        text_color=COLORS['primary_text']
    )
    info_label.pack(fill="x", pady=5)

    # Log widget with scrollbar - using CTkTextbox for CustomTkinter
    log_text = ctk.CTkTextbox(
        main_frame,
        font=("Courier New", 10),
        fg_color=COLORS['frame_bg'],
        text_color=COLORS['primary_text'],
        wrap="word"
    )
    log_text.pack(fill="both", expand=True)

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
