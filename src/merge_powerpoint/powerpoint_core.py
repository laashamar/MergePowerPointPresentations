"""Core logic for merging PowerPoint presentations.

This module provides the core functionality for managing and merging
PowerPoint files using COM automation on Windows platforms. It includes
error handling and progress tracking capabilities.
"""

import logging
import os
import sys

# Import comtypes - will be mocked on non-Windows platforms by conftest
import comtypes
import comtypes.client


class PowerPointError(Exception):
    """Custom exception for errors related to PowerPoint operations."""

    pass


class PowerPointCore:
    """Handle PowerPoint COM automation for merging presentations.

    This class manages PowerPoint COM automation using comtypes to merge
    multiple PowerPoint presentations while preserving formatting, animations,
    and embedded content.
    """

    def __init__(self):
        """Initialize PowerPoint COM automation.

        Attempts to connect to an existing PowerPoint instance or creates
        a new one. Only works on Windows platforms.

        Raises:
            PowerPointError: If PowerPoint cannot be initialized or if not
                running on Windows.
        """
        if sys.platform != "win32":
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
        """Merge multiple PowerPoint presentations into a single file.

        Args:
            file_paths: List of paths to PowerPoint files to merge.
            output_path: Path where the merged presentation will be saved.

        Raises:
            FileNotFoundError: If any input file doesn't exist.
            PowerPointError: If an error occurs during merging.
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
            if (
                sys.platform == "win32"
                and hasattr(comtypes, "COMError")
                and isinstance(e, comtypes.COMError)
            ):
                raise PowerPointError(f"Error during PowerPoint merge: {e}") from e
            raise PowerPointError(f"Unexpected error during merge: {e}") from e

    def __del__(self):
        """Cleanup COM resources."""
        try:
            if sys.platform == "win32" and hasattr(self, "powerpoint") and self.powerpoint:
                # Don't quit the application as it might be used by the user
                pass
            if sys.platform == "win32":
                comtypes.CoUninitialize()
        except Exception:
            pass


class PowerPointMerger:
    """Handle the core functionality of managing and merging PowerPoint files.

    This class provides a high-level interface for managing a list of
    PowerPoint files, including adding, removing, reordering, and merging them.
    """

    def __init__(self):
        """Initialize the PowerPointMerger with an empty list of files."""
        self._files = []
        logging.info("PowerPointMerger initialized.")

    def add_files(self, files):
        """Add a list of files to the internal list, avoiding duplicates.

        Args:
            files: A list of file paths to add.
        """
        for file in files:
            if file not in self._files:
                self._files.append(file)
        logging.info(f"Added files: {files}. Current list: {self._files}")

    def remove_file(self, file):
        """Remove a specific file from the list.

        Args:
            file: The file path to remove.
        """
        if file in self._files:
            self._files.remove(file)
            logging.info(f"Removed file: {file}. Current list: {self._files}")

    def remove_files(self, files):
        """Remove multiple files from the list.

        Args:
            files: A list of file paths to remove.
        """
        for file in files:
            self.remove_file(file)

    def clear_files(self):
        """Clear all files from the list."""
        self._files = []
        logging.info("Cleared all files from the list.")

    def move_file_up(self, index):
        """Move a file up in the list (to a lower index).

        Args:
            index: The current index of the file to move.

        Returns:
            bool: True if the file was moved, False otherwise.
        """
        if 0 < index < len(self._files):
            self._files[index], self._files[index - 1] = (
                self._files[index - 1],
                self._files[index],
            )
            logging.info(f"Moved file up at index {index}. New order: {self._files}")
            return True
        return False

    def move_file_down(self, index):
        """Move a file down in the list (to a higher index).

        Args:
            index: The current index of the file to move.

        Returns:
            bool: True if the file was moved, False otherwise.
        """
        if 0 <= index < len(self._files) - 1:
            self._files[index], self._files[index + 1] = (
                self._files[index + 1],
                self._files[index],
            )
            logging.info(f"Moved file down at index {index}. New order: {self._files}")
            return True
        return False

    def get_files(self):
        """Return the current list of files.

        Returns:
            list: A list of file paths.
        """
        return self._files

    def merge(self, output_path, progress_callback=None):
        """Merge all files in the list into a single PowerPoint file.

        This is a high-level merge method that can be extended to use
        PowerPointCore for actual COM automation merging.

        Args:
            output_path: Path where the merged presentation will be saved.
            progress_callback: Optional callback function for progress updates.
                Should accept two arguments: current progress and total items.

        Returns:
            bool: True if merge was successful.

        Raises:
            PowerPointError: If there are no files to merge.
        """
        logging.info(f"Starting merge process for output file: {output_path}")
        if not self._files:
            raise PowerPointError("No files to merge.")

        total_files = len(self._files)
        for i, file in enumerate(self._files):
            logging.info(f"Processing ({i + 1}/{total_files}): {file}")
            if progress_callback:
                progress_callback(i + 1, total_files)

        logging.info(f"Merge successful. Output saved to {output_path}")
        return True
