"""
This module contains the core logic for merging PowerPoint presentations.
"""
import logging
import os
import sys

# Import comtypes - will be mocked on non-Windows platforms by conftest
import comtypes
import comtypes.client

# Define a custom exception for PowerPoint-related errors
class PowerPointError(Exception):
    """Custom exception for errors related to PowerPoint operations."""
    pass


class PowerPointCore:
    """
    Handles PowerPoint COM automation for merging presentations using comtypes.
    """

    def __init__(self):
        """
        Initialize PowerPoint COM automation.
        Attempts to connect to an existing PowerPoint instance or creates a new one.
        Raises PowerPointError if PowerPoint cannot be initialized.
        """
        if sys.platform != 'win32':
            raise PowerPointError("PowerPoint COM automation is only available on Windows.")
        
        comtypes.CoInitialize()
        self.powerpoint = None
        
        try:
            # Try to get an existing PowerPoint instance
            self.powerpoint = comtypes.client.GetActiveObject("PowerPoint.Application")
            logging.info("Connected to existing PowerPoint instance.")
        except OSError:
            # If no instance is running, create a new one
            try:
                self.powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                logging.info("Created new PowerPoint instance.")
            except OSError as e:
                logging.error("Failed to initialize PowerPoint application.")
                raise PowerPointError("Could not start PowerPoint application.") from e
        
        self.powerpoint.Visible = True

    def merge_presentations(self, file_paths, output_path):
        """
        Merge multiple PowerPoint presentations into a single file.
        
        :param file_paths: List of paths to PowerPoint files to merge.
        :param output_path: Path where the merged presentation will be saved.
        :raises FileNotFoundError: If any input file doesn't exist.
        :raises PowerPointError: If an error occurs during merging.
        """
        # Validate all files exist
        for file_path in file_paths:
            if not os.path.exists(file_path):
                logging.error(f"File not found: {file_path}")
                raise FileNotFoundError(f"File not found: {file_path}")
        
        try:
            # Create a new base presentation
            base_presentation = self.powerpoint.Presentations.Add()
            logging.info("Created base presentation for merge.")
            
            # Insert slides from each file
            for file_path in file_paths:
                abs_path = os.path.abspath(file_path)
                logging.info(f"Inserting slides from: {abs_path}")
                slide_count = base_presentation.Slides.Count
                # InsertFromFile inserts slides after the specified index
                base_presentation.Slides.InsertFromFile(abs_path, slide_count)
            
            # Save the merged presentation
            abs_output_path = os.path.abspath(output_path)
            logging.info(f"Saving merged presentation to: {abs_output_path}")
            base_presentation.SaveAs(abs_output_path)
            
            # Close the presentation
            base_presentation.Close()
            logging.info("Merge completed successfully.")
            
        except Exception as e:
            # Handle both comtypes.COMError and other exceptions
            logging.error(f"Error during merge: {e}")
            if sys.platform == 'win32' and hasattr(comtypes, 'COMError') and isinstance(e, comtypes.COMError):
                raise PowerPointError(f"Error during PowerPoint merge: {e}") from e
            raise PowerPointError(f"Unexpected error during merge: {e}") from e

    def __del__(self):
        """Cleanup COM resources."""
        try:
            if sys.platform == 'win32' and hasattr(self, 'powerpoint') and self.powerpoint:
                # Don't quit the application as it might be used by the user
                pass
            if sys.platform == 'win32':
                comtypes.CoUninitialize()
        except:
            pass


class PowerPointMerger:
    """
    Handles the core functionality of managing and merging PowerPoint files.
    """

    def __init__(self):
        """Initializes the PowerPointMerger with an empty list of files."""
        self._files = []
        logging.info("PowerPointMerger initialized.")

    def add_files(self, files):
        """
        Adds a list of files to the internal list, avoiding duplicates.
        :param files: A list of file paths to add.
        """
        for file in files:
            if file not in self._files:
                self._files.append(file)
        logging.info(f"Added files: {files}. Current list: {self._files}")

    def remove_file(self, file):
        """
        Removes a specific file from the list.
        :param file: The file path to remove.
        """
        if file in self._files:
            self._files.remove(file)
            logging.info(f"Removed file: {file}. Current list: {self._files}")

    def move_file_up(self, index):
        """
        Moves a file up in the list (to a lower index).
        :param index: The current index of the file to move.
        """
        if 0 < index < len(self._files):
            self._files[index], self._files[index - 1] = self._files[index - 1], self._files[index]
            logging.info(f"Moved file up at index {index}. New order: {self._files}")

    def move_file_down(self, index):
        """
        Moves a file down in the list (to a higher index).
        :param index: The current index of the file to move.
        """
        if 0 <= index < len(self._files) - 1:
            self._files[index], self._files[index + 1] = self._files[index + 1], self._files[index]
            logging.info(f"Moved file down at index {index}. New order: {self._files}")

    def get_files(self):
        """
        Returns the current list of files.
        :return: A list of file paths.
        """
        return self._files

    def merge(self, output_path, progress_callback=None):
        """
        Placeholder for the merge logic.
        In a real implementation, this would use COM automation or another
        library to merge the actual .pptx files.
        """
        logging.info(f"Starting merge process for output file: {output_path}")
        if not self._files:
            raise PowerPointError("No files to merge.")

        total_files = len(self._files)
        for i, file in enumerate(self._files):
            logging.info(f"Processing ({i+1}/{total_files}): {file}")
            if progress_callback:
                progress_callback(i + 1, total_files)
        
        logging.info(f"Merge successful. Output saved to {output_path}")
        # Here you would add the actual merging code.
        # For now, we'll just simulate success.
        return True
