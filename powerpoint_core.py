"""
This module contains the core logic for merging PowerPoint presentations.
"""
import logging

# Define a custom exception for PowerPoint-related errors
class PowerPointError(Exception):
    """Custom exception for errors related to PowerPoint operations."""
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
