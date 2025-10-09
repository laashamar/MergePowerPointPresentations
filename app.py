"""
This module contains the AppController, which serves as the main controller
for the application, connecting the GUI to the business logic.
"""
from powerpoint_core import PowerPointMerger


class AppController(PowerPointMerger):
    """
    Controller for the application. It inherits the core merging logic
    from PowerPointMerger and can be extended with additional application-
    specific functionality without altering the core logic.
    """
    def __init__(self):
        """
        Initializes the AppController by calling the parent constructor.
        """
        super().__init__()
        # Future controller-specific initializations can go here.
        # For example, loading user settings, checking for updates, etc.


# This check allows the file to be imported without running test code.
if __name__ == '__main__':
    # You can add test or demonstration code here that will only run
    # when the script is executed directly.
    # For example:
    controller = AppController()
    print("AppController created successfully.")
