import tkinter as tk
from tkinter import messagebox
import logging
import threading
import sys

# Importer den nye logger-konfigurasjonen og applikasjonsstarteren
from logger import setup_logging, write_log_summary
from app import start_app

def run_main_application():
    """Kjører hovedapplikasjonen og fanger eventuelle ukjente feil."""
    try:
        logging.info("Starter PowerPoint Merger-applikasjonens arbeidsflyt...")
        start_app()
        logging.info("Applikasjonens arbeidsflyt er fullført (GUI-vinduer lukket).")
    except Exception as e:
        # Fanger opp uventede feil under oppstart eller kjøring
        logging.critical(
            "En ubehandlet feil oppstod i applikasjonen!", exc_info=True
        )
        # Vis feilmeldingen i en messagebox også, siden GUI-en kan være borte
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Kritisk Feil",
            f"En uventet feil avsluttet programmet.\n\n"
            f"Detaljer: {e}\n\n"
            "Vennligst sjekk loggfilen for mer informasjon."
        )
        root.destroy()
    finally:
        # Dette kjøres etter at appens mainloop er ferdig
        logging.info("Skriver feiloppsummering til loggfil...")
        write_log_summary()

def main():
    """Hovedfunksjon for å sette opp logg-vinduet og starte applikasjonen."""
    # Hovedvindu for loggeren
    log_window = tk.Tk()
    log_window.title("Live Logg - PowerPoint Sammenslåing")
    log_window.geometry("900x600")

    # Setter opp en ramme for bedre layout
    main_frame = tk.Frame(log_window, padx=10, pady=10)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    info_label = tk.Label(
        main_frame, 
        text="Dette vinduet viser en live logg av skriptet. Lukk dette vinduet for å avslutte.",
        pady=5
    )
    info_label.pack(fill=tk.X)

    # Logg-widget med scrollbar
    log_frame = tk.Frame(main_frame)
    log_frame.pack(fill=tk.BOTH, expand=True)

    log_text = tk.Text(log_frame, state='disabled', wrap='word', font=("Courier New", 10), bg="#f0f0f0", fg="black")
    log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(log_frame, command=log_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    log_text['yscrollcommand'] = scrollbar.set

    # Konfigurer logging til å peke til Text-widgeten
    setup_logging(log_text)

    # Overstyr standard excepthook for å logge alle uventede feil
    def handle_exception(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logging.critical("Ubehandlet unntak fanget av sys.excepthook:", exc_info=(exc_type, exc_value, exc_traceback))

    sys.excepthook = handle_exception
    
    # Kjør hovedapplikasjonen i en egen tråd for å unngå at GUI-en fryser
    app_thread = threading.Thread(target=run_main_application, daemon=True)
    app_thread.start()

    # Hoved-loopen for logg-vinduet
    log_window.mainloop()

if __name__ == "__main__":
    main()
