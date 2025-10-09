"""Main GUI module for PowerPoint Merger application using CustomTkinter."""

import os
import sys
from tkinter import filedialog

import customtkinter

# --- Path adjustment to allow importing from the parent directory ---
# This is the crucial part that allows the new GUI to find the core logic.
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

# --- Application Core Logic Imports ---
import powerpoint_core  # noqa: E402, F401 - Will be used in future phases


# --- Main Application Class ---
class App(customtkinter.CTk):
    """Main application window for PowerPoint merger."""

    def __init__(self):
        """Initialize the main application window."""
        super().__init__()

        # Set window title and geometry
        self.title("Merge PowerPoint Presentations")
        self.geometry("800x600")

        # Initialize file list
        self.file_list = []

        # Configure grid layout (3 rows, 3 columns)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure(2, weight=0)

        # 2.1: File Display List
        self.file_list_frame = customtkinter.CTkScrollableFrame(
            self,
            label_text="Selected Presentations"
        )
        self.file_list_frame.grid(
            row=0,
            column=1,
            columnspan=2,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        # 2.2: File Management Buttons
        file_buttons_frame = customtkinter.CTkFrame(self)
        file_buttons_frame.grid(
            row=0,
            column=0,
            padx=10,
            pady=10,
            sticky="nsew"
        )

        self.add_button = customtkinter.CTkButton(
            file_buttons_frame,
            text="Add Presentation(s)",
            command=self.add_files
        )
        self.add_button.pack(padx=10, pady=10)

        self.remove_button = customtkinter.CTkButton(
            file_buttons_frame,
            text="Remove Selected",
            command=self.remove_selected_file
        )
        self.remove_button.pack(padx=10, pady=10)

        self.clear_button = customtkinter.CTkButton(
            file_buttons_frame,
            text="Clear All",
            command=self.clear_file_list
        )
        self.clear_button.pack(padx=10, pady=10)

        # 2.3: Output File Configuration
        output_frame = customtkinter.CTkFrame(self)
        output_frame.grid(
            row=1,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="ew"
        )

        output_label = customtkinter.CTkLabel(
            output_frame,
            text="Output File:"
        )
        output_label.pack(side="left", padx=10, pady=10)

        self.output_path_entry = customtkinter.CTkEntry(
            output_frame,
            placeholder_text="Select output path..."
        )
        self.output_path_entry.pack(
            side="left",
            fill="x",
            expand=True,
            padx=10,
            pady=10
        )

        self.save_as_button = customtkinter.CTkButton(
            output_frame,
            text="Save As...",
            command=self.select_output_path
        )
        self.save_as_button.pack(side="left", padx=10, pady=10)

        # 2.4: Primary Action and Status
        action_frame = customtkinter.CTkFrame(self)
        action_frame.grid(
            row=2,
            column=0,
            columnspan=3,
            padx=10,
            pady=10,
            sticky="ew"
        )

        self.merge_button = customtkinter.CTkButton(
            action_frame,
            text="Merge Presentations",
            command=self.merge_presentations,
            state="disabled"
        )
        self.merge_button.pack(pady=10)

        self.status_label = customtkinter.CTkLabel(
            action_frame,
            text="Ready"
        )
        self.status_label.pack(pady=5)

    def add_files(self):
        """Open file dialog to add PowerPoint files."""
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )

        for file in files:
            if file not in self.file_list:
                self.file_list.append(file)

        self.update_file_list_display()
        self.update_merge_button_status()

    def update_file_list_display(self):
        """Update the file list display in the scrollable frame."""
        # Clear current contents
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()

        # Display each file
        for file_path in self.file_list:
            label = customtkinter.CTkLabel(
                self.file_list_frame,
                text=os.path.basename(file_path)
            )
            label.pack(anchor="w", padx=10, pady=2)

    def remove_selected_file(self):
        """Remove the last file from the list (placeholder implementation)."""
        if self.file_list:
            self.file_list.pop()
            self.update_file_list_display()
            self.update_merge_button_status()

    def clear_file_list(self):
        """Clear all files from the list."""
        self.file_list.clear()
        self.update_file_list_display()
        self.update_merge_button_status()

    def select_output_path(self):
        """Open save-as dialog to select output file path."""
        file_path = filedialog.asksaveasfilename(
            title="Save Merged Presentation As",
            defaultextension=".pptx",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )

        if file_path:
            self.output_path_entry.delete(0, "end")
            self.output_path_entry.insert(0, file_path)
            self.update_merge_button_status()

    def update_merge_button_status(self):
        """Enable or disable merge button based on conditions."""
        # Check if at least 2 files are selected
        has_files = len(self.file_list) >= 2

        # Check if output path is specified
        has_output = bool(self.output_path_entry.get().strip())

        # Enable button if both conditions are met
        if has_files and has_output:
            self.merge_button.configure(state="normal")
        else:
            self.merge_button.configure(state="disabled")

    def merge_presentations(self):
        """Placeholder method for merging presentations."""
        self.status_label.configure(text="Merging presentations...")
        output_path = self.output_path_entry.get()
        print(f"Merging {self.file_list} into {output_path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
