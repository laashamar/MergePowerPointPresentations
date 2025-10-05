"""
Core module for PowerPoint presentation merging using COM automation.

This module provides the business logic for merging multiple PowerPoint
presentations using pywin32 COM automation, which ensures perfect copying
of all content, formatting, and animations.
"""
import logging
import os
import win32com.client


def merge_presentations(file_order, output_filename):
    """
    Merge multiple PowerPoint presentations into a single file using COM.

    Args:
        file_order: List of file paths in the order they should be merged
        output_filename: Name of the output file (should include .pptx extension)

    Returns:
        tuple: (success: bool, output_path: str, error_message: str or None)
    """
    logging.info("Starter 'merge_presentations'-prosessen.")
    logging.info(f"Antall filer som skal slås sammen: {len(file_order)}")
    logging.info(f"Output-filnavn: {output_filename}")

    powerpoint = None
    destination_prs = None
    source_prs = None

    try:
        # Initialize PowerPoint application
        logging.info("Initialiserer PowerPoint-applikasjonen via COM...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True
        logging.info("PowerPoint-applikasjonen er synlig.")

        # Create a new presentation
        logging.info("Oppretter ny (tom) destinasjonspresentasjon.")
        destination_prs = powerpoint.Presentations.Add()

        # Remove the default blank slide if it exists
        if destination_prs.Slides.Count > 0:
            logging.info("Fjerner standard blankt lysbilde fra destinasjonen.")
            destination_prs.Slides(1).Delete()

        # Process each source file in order
        for i, file_path in enumerate(file_order, 1):
            abs_path = os.path.abspath(file_path)
            logging.info(f"--- Behandler fil {i}/{len(file_order)}: {os.path.basename(abs_path)} ---")
            try:
                # Open source presentation
                logging.info(f"Åpner kildepresentasjon: {abs_path}")
                source_prs = powerpoint.Presentations.Open(
                    abs_path,
                    ReadOnly=True,
                    WithWindow=False
                )
                
                num_slides = source_prs.Slides.Count
                logging.info(f"Fant {num_slides} lysbilder i kildepresentasjonen.")

                # Copy all slides from source to destination in one operation
                if num_slides > 0:
                    logging.info(f"Starter kopiering av {num_slides} lysbilder...")
                    source_prs.Slides.Range().Copy()
                    destination_prs.Slides.Paste()
                    logging.info("Alle lysbilder fra kilden ble limt inn i destinasjonen.")
                else:
                    logging.warning(f"Ingen lysbilder funnet i {os.path.basename(abs_path)}. Hopper over.")

                # Close source presentation
                logging.info(f"Lukker kildepresentasjon: {os.path.basename(abs_path)}")
                source_prs.Close()
                source_prs = None

            except Exception as e:
                logging.error(
                    f"En feil oppstod under behandling av filen {os.path.basename(file_path)}",
                    exc_info=True
                )
                if source_prs:
                    source_prs.Close()
                # Re-raise to be caught by the outer try-except block
                raise Exception(
                    f"Klarte ikke å behandle filen {os.path.basename(file_path)}: {str(e)}"
                )

        # Save the merged presentation
        output_path = os.path.abspath(output_filename)
        logging.info(f"Lagrer den sammenslåtte presentasjonen til: {output_path}")
        destination_prs.SaveAs(output_path)
        logging.info("Lagring vellykket.")

        return True, output_path, None

    except Exception as e:
        error_message = f"En feil oppstod under sammenslåingen: {str(e)}"
        logging.critical(error_message, exc_info=True)
        # If any error occurs during the process, perform a full cleanup.
        try:
            logging.info("Feil oppstod. Starter opprydding av COM-objekter.")
            if destination_prs:
                destination_prs.Close()
                logging.info("Lukket destinasjonspresentasjon.")
            if source_prs:
                source_prs.Close()
                logging.info("Lukket kildepresentasjon.")
            if powerpoint:
                powerpoint.Quit()
                logging.info("Avsluttet PowerPoint-applikasjonen.")
        except Exception as cleanup_error:
            logging.error(f"En feil oppstod under opprydding: {cleanup_error}", exc_info=True)
            pass
        return False, "", str(e)


def launch_slideshow(output_path):
    """
    Launch PowerPoint slideshow using COM automation.

    Args:
        output_path: Full path to the presentation file

    Returns:
        tuple: (success: bool, error_message: str or None)
    """
    logging.info("Starter 'launch_slideshow'-prosessen.")
    powerpoint = None
    presentation = None

    try:
        # Get PowerPoint application instance
        logging.info("Henter instans av PowerPoint-applikasjonen...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        # Open the presentation
        abs_path = os.path.abspath(output_path)
        logging.info(f"Åpner presentasjon for visning: {abs_path}")
        presentation = powerpoint.Presentations.Open(
            abs_path,
            WithWindow=True
        )

        # Start the slideshow
        logging.info("Starter lysbildefremvisning...")
        presentation.SlideShowSettings.Run()
        logging.info("Lysbildefremvisning startet vellykket.")

        return True, None

    except Exception as e:
        error_message = f"Kunne ikke starte lysbildefremvisning for {output_path}: {str(e)}"
        logging.error(error_message, exc_info=True)
        return False, str(e)
