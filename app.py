# app.py

"""
This module contains the core application logic for the PowerPoint Merger.
"""

from powerpoint_core import PowerPointCore

class PowerPointMerger:
    """
    Manages the list of PowerPoint files and orchestrates the merging process.
    """
    def __init__(self):
        self.file_paths = []
        self.core = PowerPointCore()

    def add_files(self, new_files):
        """Adds a list of files to the current list."""
        self.file_paths.extend(new_files)

    def remove_file(self, index):
        """Removes a file from the list by its index."""
        if 0 <= index < len(self.file_paths):
            del self.file_paths[index]

    def move_file_up(self, index):
        """Moves a file up one position in the list."""
        if index > 0:
            self.file_paths[index], self.file_paths[index - 1] = self.file_paths[index - 1], self.file_paths[index]

    def move_file_down(self, index):
        """Moves a file down one position in the list."""
        if 0 <= index < len(self.file_paths) - 1:
            self.file_paths[index], self.file_paths[index + 1] = self.file_paths[index + 1], self.file_paths[index]

    def get_files(self):
        """Returns the current list of file paths."""
        return self.file_paths

    def merge(self, output_path, progress_callback=None):
        """
        Merges the presentations in the list into a single file.
        """
        self.core.merge_presentations(self.file_paths, output_path)

    def close(self):
        """Closes the PowerPoint application."""
        self.core.close()

