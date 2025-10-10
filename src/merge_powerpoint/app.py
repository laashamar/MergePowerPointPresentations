"""Application controller module for the PowerPoint Merger.

This module contains the AppController, which serves as the main controller
for the application, connecting the GUI to the business logic.
"""

from merge_powerpoint.powerpoint_core import PowerPointMerger


class AppController(PowerPointMerger):
    """Application controller that manages the merging workflow.

    This controller inherits the core merging logic from PowerPointMerger
    and can be extended with additional application-specific functionality
    without altering the core logic.
    """

    def __init__(self):
        """Initialize the AppController by calling the parent constructor."""
        super().__init__()
        # Future controller-specific initializations can go here.
        # For example, loading user settings, checking for updates, etc.


if __name__ == "__main__":
    # Test code that runs when the script is executed directly
    controller = AppController()
    print("AppController created successfully.")
