# main_gui.py

import os
import sys
import subprocess
import webbrowser
from tkinter import filedialog, messagebox

import customtkinter
import pywintypes

# --- Path adjustment to allow importing from the parent directory ---
# This is the crucial part that allows the new GUI to find the core logic.
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# --- Application Core Logic Imports ---
import logger
import powerpoint_core


# --- Main Application Class ---
class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("Merge PowerPoint Presentations")
        self.geometry("700x550")

if __name__ == "__main__":
    app = App()
    app.mainloop()         
