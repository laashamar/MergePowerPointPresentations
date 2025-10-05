import logging
import tkinter as tk
from logging import FileHandler, Handler
import os
from datetime import datetime

# Definer den sentrale loggfilstien
LOG_FILE_PATH = os.path.join(os.path.expanduser("~"), "Downloads", "merge_powerpoint.log")

# Global liste for å samle feilmeldinger
error_list = []

class TkinterLogHandler(Handler):
    """En egendefinert logg-handler som sender records til et tkinter Text-widget."""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        """Skriver loggmeldingen til Text-widgeten."""
        msg = self.format(record)
        self.text_widget.configure(state='normal')
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.configure(state='disabled')
        self.text_widget.see(tk.END) # Auto-scroll

class ErrorListHandler(Handler):
    """En handler som samler alle ERROR og CRITICAL meldinger i en liste."""
    def emit(self, record):
        """Legger til formaterte feilmeldinger i den globale listen."""
        if record.levelno >= logging.ERROR:
            error_list.append(self.format(record))

def setup_logging(log_widget):
    """
    Konfigurerer logging til å sende output til GUI, fil og en feilliste.
    """
    # Slett gammel loggfil hvis den eksisterer for en ren start hver kjøring
    if os.path.exists(LOG_FILE_PATH):
        os.remove(LOG_FILE_PATH)

    log_format = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(module)s.%(funcName)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Rote-logger
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)

    # Handler for GUI-vinduet
    gui_handler = TkinterLogHandler(log_widget)
    gui_handler.setFormatter(log_format)
    root_logger.addHandler(gui_handler)

    # Handler for fil
    file_handler = FileHandler(LOG_FILE_PATH, mode='w', encoding='utf-8')
    file_handler.setFormatter(log_format)
    root_logger.addHandler(file_handler)

    # Handler for å samle feil
    error_handler = ErrorListHandler()
    error_handler.setFormatter(log_format)
    root_logger.addHandler(error_handler)

    logging.info("Logging er satt opp.")
    logging.info(f"Loggfil lagres til: {LOG_FILE_PATH}")


def write_log_summary():
    """Skriver en oppsummering av feil til slutten av loggfilen."""
    try:
        with open(LOG_FILE_PATH, "a", encoding="utf-8") as f:
            summary_header = "\n\n" + "="*80 + "\n"
            summary_header += " FEILOPPSUMMERING ".center(80, "=") + "\n"
            summary_header += "="*80 + "\n\n"
            f.write(summary_header)

            if error_list:
                f.write(f"Fant {len(error_list)} feil under kjøringen:\n\n")
                for error in error_list:
                    f.write(error + "\n\n")
            else:
                f.write("Ingen feil ble logget under kjøringen.\n")
            
            f.write("\n" + "="*80 + "\n")
        logging.info("Feiloppsummering er skrevet til loggfilen.")
    except Exception as e:
        logging.error(f"Kunne ikke skrive feiloppsummering til loggfil: {e}", exc_info=True)
