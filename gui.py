"""
Defines the main user interface for the PowerPoint Merging Tool.

This version restores all button functionality, including file selection, list management,
output path browsing, and triggering the merge process, while retaining the fix for
the drag-and-drop initialization.
"""
import logging
import os
from tkinter import filedialog
import customtkinter
from tkinterdnd2 import DND_FILES, TkinterDnD

# Assuming powerpoint_core.py is in the same directory or accessible in the path
from powerpoint_core import merge_presentations

class MainApplication(customtkinter.CTk, TkinterDnD.Tk):
    """
    The main application window, integrating CustomTkinter for appearance
    and TkinterDnD2 for drag-and-drop file handling.
    """

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("PowerPoint Presentation Merger")
        self.geometry("800x600")
        self.minsize(700, 500)

        self.file_list = []
        self.selected_index = None # To track which file is selected

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.create_widgets()
        self.configure_drag_and_drop()

    def configure_drag_and_drop(self):
        """Configures the drag-and-drop functionality for the main window."""
        try:
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self.on_drop)
            logging.info("Drag-and-drop configured successfully.")
        except Exception as e:
            logging.warning("Could not configure drag-and-drop: %s", e, exc_info=True)

    def create_widgets(self):
        """Creates and places all the widgets for the application."""
        # --- Top Frame ---
        self.top_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        self.top_frame.grid_columnconfigure(0, weight=1)

        self.label_title = customtkinter.CTkLabel(
            self.top_frame, text="Merge PowerPoint Presentations",
            font=customtkinter.CTkFont(size=20, weight="bold")
        )
        self.label_title.grid(row=0, column=0, padx=20, pady=10)

        self.theme_switch = customtkinter.CTkSwitch(
            self.top_frame, text="Dark Mode", command=self.toggle_theme
        )
        self.theme_switch.grid(row=0, column=1, padx=20, pady=10)
        self.theme_switch.select()

        # --- Main Content Frame ---
        self.main_frame = customtkinter.CTkFrame(self)
        self.main_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(0, weight=1)

        # --- File List Display ---
        self.file_display_frame = customtkinter.CTkScrollableFrame(
            self.main_frame, label_text="Files to Merge"
        )
        self.file_display_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)
        self.file_display_frame.grid_columnconfigure(0, weight=1)

        self.update_file_display() # Initial call to show placeholder

        # --- File Management Buttons Frame ---
        self.button_frame = customtkinter.CTkFrame(self.main_frame)
        self.button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.button_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)

        self.btn_add = customtkinter.CTkButton(self.button_frame, text="Add Files", command=self.add_files)
        self.btn_add.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.btn_remove = customtkinter.CTkButton(self.button_frame, text="Remove Selected", command=self.remove_selected_file)
        self.btn_remove.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        self.btn_move_up = customtkinter.CTkButton(self.button_frame, text="Move Up", command=self.move_file_up)
        self.btn_move_up.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        self.btn_move_down = customtkinter.CTkButton(self.button_frame, text="Move Down", command=self.move_file_down)
        self.btn_move_down.grid(row=0, column=3, padx=5, pady=5, sticky="ew")

        # --- Bottom Frame ---
        self.bottom_frame = customtkinter.CTkFrame(self)
        self.bottom_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        self.bottom_frame.grid_columnconfigure(1, weight=1)

        self.label_output = customtkinter.CTkLabel(self.bottom_frame, text="Output File:")
        self.label_output.grid(row=0, column=0, padx=(10, 5), pady=10)

        self.entry_output_path = customtkinter.CTkEntry(
            self.bottom_frame, placeholder_text="Select path for merged file..."
        )
        self.entry_output_path.grid(row=0, column=1, sticky="ew", padx=5, pady=10)

        self.btn_browse = customtkinter.CTkButton(self.bottom_frame, text="Browse...", command=self.browse_output_file)
        self.btn_browse.grid(row=0, column=2, padx=(5, 10), pady=10)

        self.btn_merge = customtkinter.CTkButton(
            self.bottom_frame, text="Merge Presentations", height=40, command=self.trigger_merge
        )
        self.btn_merge.grid(row=1, column=0, columnspan=3, sticky="ew", padx=10, pady=(5, 10))

        # --- Status Bar ---
        self.status_bar = customtkinter.CTkLabel(self, text="Ready", anchor="w")
        self.status_bar.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 5))

    def on_drop(self, event):
        """Handles the file drop event."""
        files = self.tk.splitlist(event.data)
        added_files = False
        for file_path in files:
            if file_path.lower().endswith('.pptx') and file_path not in self.file_list:
                self.file_list.append(file_path)
                added_files = True
            else:
                logging.warning(f"Ignored duplicate or non-PowerPoint file: {file_path}")
        if added_files:
            self.update_file_display()

    def add_files(self):
        """Opens a dialog to add multiple .pptx files."""
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if files:
            for file_path in files:
                if file_path not in self.file_list:
                    self.file_list.append(file_path)
            self.update_file_display()

    def remove_selected_file(self):
        """Removes the selected file from the list."""
        if self.selected_index is not None:
            self.file_list.pop(self.selected_index)
            self.selected_index = None
            self.update_file_display()

    def move_file_up(self):
        """Moves the selected file up in the list."""
        if self.selected_index is not None and self.selected_index > 0:
            item = self.file_list.pop(self.selected_index)
            self.file_list.insert(self.selected_index - 1, item)
            self.selected_index -= 1
            self.update_file_display()

    def move_file_down(self):
        """Moves the selected file down in the list."""
        if self.selected_index is not None and self.selected_index < len(self.file_list) - 1:
            item = self.file_list.pop(self.selected_index)
            self.file_list.insert(self.selected_index + 1, item)
            self.selected_index += 1
            self.update_file_display()

    def browse_output_file(self):
        """Opens a dialog to select the output file path."""
        file_path = filedialog.asksaveasfilename(
            title="Save Merged File As",
            defaultextension=".pptx",
            filetypes=[("PowerPoint files", "*.pptx")]
        )
        if file_path:
            self.entry_output_path.delete(0, "end")
            self.entry_output_path.insert(0, file_path)

    def trigger_merge(self):
        """Initiates the merge process."""
        output_path = self.entry_output_path.get()
        if len(self.file_list) < 2:
            self.status_bar.configure(text="Error: Please add at least two files to merge.")
            logging.error("Merge failed: Less than two files provided.")
            return
        if not output_path:
            self.status_bar.configure(text="Error: Please specify an output file path.")
            logging.error("Merge failed: No output path specified.")
            return

        self.status_bar.configure(text="Merging in progress... please wait.")
        self.update() # Force UI update

        try:
            logging.info(f"Starting merge. Files: {self.file_list}, Output: {output_path}")
            merge_presentations(self.file_list, output_path)
            self.status_bar.configure(text=f"Success! Merged file saved to {output_path}")
            logging.info("Merge completed successfully.")
        except Exception as e:
            self.status_bar.configure(text=f"Error during merge: {e}")
            logging.error(f"Merge failed with exception: {e}", exc_info=True)

    def select_file(self, index):
        """Sets the currently selected file and updates the UI."""
        self.selected_index = index
        self.update_file_display()

    def update_file_display(self):
        """Refreshes the file list in the scrollable frame."""
        for widget in self.file_display_frame.winfo_children():
            widget.destroy()

        if not self.file_list:
            # Placeholder label for the drop zone
            drop_label = customtkinter.CTkLabel(
                self.file_display_frame,
                text="Drag and drop .pptx files here\nor use 'Add Files' button",
                font=customtkinter.CTkFont(size=14), text_color="gray50"
            )
            drop_label.grid(row=0, column=0, pady=50)
            return

        for i, file_path in enumerate(self.file_list):
            file_name = os.path.basename(file_path)
            
            # Determine button appearance based on selection
            fg_color = "gray30" if i == self.selected_index else "transparent"
            
            file_btn = customtkinter.CTkButton(
                self.file_display_frame,
                text=f"{i+1}. {file_name}",
                fg_color=fg_color,
                anchor="w",
                command=lambda index=i: self.select_file(index)
            )
            file_btn.grid(row=i, column=0, sticky="ew", padx=5, pady=2)

    def toggle_theme(self):
        """Switches between light and dark themes."""
        mode = "light"
        if self.theme_switch.get() == 1:
            mode = "dark"
        customtkinter.set_appearance_mode(mode)

