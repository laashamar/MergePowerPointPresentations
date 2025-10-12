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
    logging.info("Starting 'merge_presentations' process.")
    logging.info(f"Number of files to be merged: {len(file_order)}")
    logging.info(f"Output filename: {output_filename}")

    powerpoint = None
    destination_prs = None
    source_prs = None

    try:
        # Initialize PowerPoint application
        logging.info("Initializing PowerPoint application via COM...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True
        logging.info("PowerPoint application is visible.")

        # Create a new presentation
        logging.info("Creating new (empty) destination presentation.")
        destination_prs = powerpoint.Presentations.Add()

        # Remove the default blank slide if it exists
        if destination_prs.Slides.Count > 0:
            logging.info("Removing default blank slide from destination.")
            destination_prs.Slides(1).Delete()

        # Process each source file in order
        for i, file_path in enumerate(file_order, 1):
            abs_path = os.path.abspath(file_path)
            logging.info(
                f"--- Processing file {i}/{len(file_order)}: "
                f"{os.path.basename(abs_path)} ---"
            )
            try:
                # Open source presentation
                logging.info(f"Opening source presentation: {abs_path}")
                source_prs = powerpoint.Presentations.Open(
                    abs_path,
                    ReadOnly=True,
                    WithWindow=False
                )

                num_slides = source_prs.Slides.Count
                logging.info(f"Found {num_slides} slides in source presentation.")

                # Copy all slides from source to destination in one operation
                if num_slides > 0:
                    logging.info(f"Starting copy of {num_slides} slides...")
                    source_prs.Slides.Range().Copy()
                    destination_prs.Slides.Paste()
                    logging.info("All slides from source were pasted into destination.")
                else:
                    logging.warning(f"No slides found in {os.path.basename(abs_path)}. Skipping.")

                # Close source presentation
                logging.info(f"Closing source presentation: {os.path.basename(abs_path)}")
                source_prs.Close()
                source_prs = None

            except Exception as e:
                logging.error(
                    f"An error occurred while processing file {os.path.basename(file_path)}",
                    exc_info=True
                )
                if source_prs:
                    source_prs.Close()
                # Re-raise to be caught by the outer try-except block
                raise Exception(
                    f"Failed to process file {os.path.basename(file_path)}: {str(e)}"
                )

        # Save the merged presentation
        output_path = os.path.abspath(output_filename)
        logging.info(f"Saving merged presentation to: {output_path}")
        # The value 11 corresponds to the 'ppSaveAsDefault' format (.pptx)
        destination_prs.SaveAs(output_path, FileFormat=11)
        logging.info("Save successful.")

        return True, output_path, None

    except Exception as e:
        error_message = f"An error occurred during merging: {str(e)}"
        logging.critical(error_message, exc_info=True)
        # If any error occurs during the process, perform a full cleanup.
        try:
            logging.info("Error occurred. Starting cleanup of COM objects.")
            if destination_prs:
                destination_prs.Close()
                logging.info("Closed destination presentation.")
            if source_prs:
                source_prs.Close()
                logging.info("Closed source presentation.")
            if powerpoint:
                powerpoint.Quit()
                logging.info("Quit PowerPoint application.")
        except Exception as cleanup_error:
            logging.error(f"An error occurred during cleanup: {cleanup_error}", exc_info=True)
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
    logging.info("Starting 'launch_slideshow' process.")
    powerpoint = None
    presentation = None

    try:
        # Get PowerPoint application instance
        logging.info("Getting PowerPoint application instance...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        # Open the presentation
        abs_path = os.path.abspath(output_path)
        logging.info(f"Opening presentation for viewing: {abs_path}")
        presentation = powerpoint.Presentations.Open(
            abs_path,
            WithWindow=True
        )

        # Start the slideshow
        logging.info("Starting slideshow...")
        presentation.SlideShowSettings.Run()
        logging.info("Slideshow started successfully.")

        return True, None

    except Exception as e:
        error_message = f"Could not start slideshow for {output_path}: {str(e)}"
        logging.error(error_message, exc_info=True)
        return False, str(e)
