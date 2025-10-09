# powerpoint_core.py

"""
Core functionality for interacting with PowerPoint presentations.
"""

import logging
import os
import comtypes.client

# Set up logging
logger = logging.getLogger(__name__)

class PowerPointError(Exception):
    """Custom exception for PowerPoint-related errors."""
    pass

class PowerPointCore:
    """
    A class to encapsulate the core PowerPoint automation functionalities.
    """
    def __init__(self):
        self.powerpoint = None
        self.is_powerpoint_running = False
        try:
            self.powerpoint = comtypes.client.GetActiveObject("PowerPoint.Application")
            self.is_powerpoint_running = True
            logger.info("Connected to existing PowerPoint instance.")
        except (OSError, comtypes.COMError):
            try:
                self.powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                self.powerpoint.Visible = 1
                logger.info("Created a new PowerPoint instance.")
            except (OSError, comtypes.COMError) as e:
                logger.error("PowerPoint could not be started: %s", e)
                raise PowerPointError("PowerPoint could not be started.") from e

    def merge_presentations(self, file_paths, output_path):
        """
        Merges multiple PowerPoint presentations into a single file.
        """
        if not file_paths:
            raise ValueError("No input files provided.")

        try:
            # Create a new presentation
            base_presentation = self.powerpoint.Presentations.Add()
            logger.info("Created a new presentation for merging.")

            # Insert slides from each presentation
            for file_path in file_paths:
                if os.path.exists(file_path):
                    try:
                        # Insert all slides from the source presentation
                        base_presentation.Slides.InsertFromFile(file_path, base_presentation.Slides.Count)
                        logger.info("Inserted slides from: %s", file_path)
                    except comtypes.COMError as e:
                        logger.error("Failed to insert slides from %s: %s", file_path, e)
                        raise PowerPointError(f"Failed to insert slides from {os.path.basename(file_path)}.") from e
                else:
                    logger.warning("File not found: %s", file_path)
                    raise FileNotFoundError(f"File not found: {file_path}")

            # Save the merged presentation
            base_presentation.SaveAs(output_path)
            logger.info("Merged presentation saved to: %s", output_path)

            # Close the base presentation
            base_presentation.Close()
            logger.info("Closed the merged presentation.")

        except comtypes.COMError as e:
            logger.error("An error occurred during the merge process: %s", e)
            raise PowerPointError("An error occurred during the merge process.") from e
        except Exception as e:
            logger.error("An unexpected error occurred: %s", e, exc_info=True)
            raise

    def close(self):
        """
        Closes the PowerPoint application if it was started by this class.
        """
        if self.powerpoint and not self.is_powerpoint_running:
            self.powerpoint.Quit()
            self.powerpoint = None
            logger.info("Closed the PowerPoint instance.")
