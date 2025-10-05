"""
Application module for PowerPoint Merger.

This module orchestrates the workflow, manages application state,
and connects GUI events with the backend merging logic.
"""
import logging
from tkinter import messagebox
import gui
import powerpoint_core  # Endret fra 'core'


class PowerPointMergerApp:
    """Main application class that manages state and workflow."""

    def __init__(self):
        """Initialize the application state."""
        self.num_files = 0
        self.selected_files = []
        self.output_filename = ""
        self.file_order = []
        logging.info("PowerPointMergerApp-instans er initialisert.")

    def run(self):
        """Start the application with the first window."""
        logging.info("Applikasjonsflyten starter.")
        gui.show_number_of_files_window(self._on_number_of_files_entered)

    def _on_number_of_files_entered(self, num_files):
        """Handle number of files input from Step 1."""
        logging.info(f"Steg 1 fullført. Antall filer satt til: {num_files}.")
        self.num_files = num_files
        gui.show_file_selection_window(
            self.num_files,
            self._on_files_selected
        )

    def _on_files_selected(self, selected_files):
        """Handle file selection from Step 2."""
        logging.info(f"Steg 2 fullført. {len(selected_files)} filer valgt.")
        logging.debug(f"Valgte filer: {selected_files}")
        self.selected_files = selected_files
        gui.show_filename_window(self._on_filename_entered)

    def _on_filename_entered(self, filename):
        """Handle filename input from Step 3."""
        logging.info(f"Steg 3 fullført. Output-filnavn satt til: '{filename}'.")
        self.output_filename = filename
        gui.show_reorder_window(
            self.selected_files,
            self._on_files_reordered
        )

    def _on_files_reordered(self, file_order):
        """Handle file reordering from Step 4."""
        logging.info("Steg 4 fullført. Filrekkefølge er satt.")
        logging.debug(f"Endelig filrekkefølge: {file_order}")
        self.file_order = file_order
        self._merge_and_launch()

    def _merge_and_launch(self):
        """Merge presentations and launch the slideshow."""
        logging.info("Starter sammenslåing og visning av presentasjoner.")
        
        # Merge presentations using COM automation
        success, output_path, error_msg = powerpoint_core.merge_presentations(
            self.file_order,
            self.output_filename
        )

        if not success:
            # Feil logges allerede i core-modulen
            messagebox.showerror(
                "Merge Failed",
                f"Klarte ikke å slå sammen presentasjonene: {error_msg}"
            )
            return

        logging.info(f"Sammenslåing var vellykket. Fil lagret til: {output_path}")
        messagebox.showinfo(
            "Success",
            f"Presentasjonen ble slått sammen!\n\nLagret i: {output_path}"
        )

        # Launch slideshow using COM automation
        logging.info(f"Starter lysbildefremvisning for: {output_path}")
        success, error_msg = powerpoint_core.launch_slideshow(output_path)

        if not success:
            # Feil logges allerede i core-modulen
            messagebox.showwarning(
                "Launch Failed",
                f"Presentasjonen ble lagret i {output_path}, "
                f"men kunne ikke startes: {error_msg}"
            )


def start_app():
    """Start the PowerPoint Merger application."""
    logging.info("Oppretter og kjører en ny instans av PowerPointMergerApp.")
    app = PowerPointMergerApp()
    app.run()

