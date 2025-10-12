# app.py

import logging
import gui
from powerpoint_core import merge_presentations

def start_app():
    """Initializes and runs the PowerPoint Merger application."""
    logging.info("Application started.")
    
    # The merge_presentations function from powerpoint_core will be the callback
    gui.show_modern_gui(merge_presentations)
    
    logging.info("Application closed.")

if __name__ == "__main__":
    start_app()