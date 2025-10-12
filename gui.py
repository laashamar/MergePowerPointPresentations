"""
GUI module for PowerPoint Merger application.
This module contains the modern two-column GUI for PowerPoint file
merging, using CustomTkinter for a modern dark theme.
"""
import logging
import os
import threading
import customtkinter as ctk
import tkinter as tk
# *** FIX: Import ImageTk to handle image objects robustly ***
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox

# PowerPoint-inspired Dark Mode Color Palette
COLORS = {
    'primary_accent': '#d35230',
    'accent_hover': '#ba3416', # Corrected color
    'window_bg': '#242424',
    'frame_bg': '#2b2b2b',
    'primary_text': '#e5e5e5',
    'secondary_text': '#a0a0a0',
    'button_text': '#ffffff',
    'error_color': '#FF0000',
}

# Font settings
FONT_FAMILY = "Helvetica"
FONT_SIZE_LARGE = 14
FONT_SIZE_MEDIUM = 12
FONT_SIZE_SMALL = 10

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class PowerPointMergerGUI:
    """Modern GUI for PowerPoint Merger with two-column layout using CustomTkinter."""

    def __init__(self, merge_callback):
        self.merge_callback = merge_callback
        self.file_list = []
        self.pil_powerpoint_icon = None
        # This list is crucial to prevent garbage collection of image objects
        self.file_card_icons = []

        self.root = ctk.CTk()
        self.root.title("PowerPoint Merger")
        self.root.geometry("900x600")
        self.root.configure(fg_color=COLORS['window_bg'])

        icon_path = os.path.join(
            os.path.dirname(__file__),
            "resources",
            "MergePowerPoint.ico"
        )
        if os.path.exists(icon_path):
            try:
                self.root.iconbitmap(icon_path)
            except Exception as e:
                logging.warning(f"Could not set application icon: {e}")

        self._load_powerpoint_icon_pil()
        self._create_widgets()
        self._update_merge_queue_display()

    def _load_powerpoint_icon_pil(self):
        """Load the PowerPoint icon as a PIL Image object."""
        icon_path = os.path.join(
            os.path.dirname(__file__),
            "resources",
            "MergePowerPoint.ico"
        )
        if os.path.exists(icon_path):
            try:
                # Resize the icon for better display in the list
                pil_image = Image.open(icon_path).resize((24, 24), Image.LANCZOS)
                self.pil_powerpoint_icon = pil_image
            except Exception as e:
                logging.warning(f"Could not load powerpoint icon: {e}")
                self.pil_powerpoint_icon = None

    def _create_widgets(self):
        """Create and layout all GUI widgets."""
        main_container = ctk.CTkFrame(self.root, fg_color=COLORS['window_bg'])
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        self.queue_frame = ctk.CTkFrame(main_container, fg_color=COLORS['frame_bg'],
                                        border_width=2, border_color=COLORS['secondary_text'])
        self.queue_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        queue_label = ctk.CTkLabel(
            self.queue_frame,
            text="Merge Queue",
            font=(FONT_FAMILY, FONT_SIZE_LARGE, "bold"),
            text_color=COLORS['primary_text']
        )
        queue_label.pack(pady=(10, 5))

        self.content_frame = ctk.CTkFrame(self.queue_frame, fg_color=COLORS['frame_bg'])
        self.content_frame.pack(fill="both", expand=True, padx=10, pady=10)

        config_frame = ctk.CTkFrame(main_container, fg_color=COLORS['frame_bg'],
                                    border_width=2, border_color=COLORS['secondary_text'])
        config_frame.pack(side="right", fill="both", padx=(5, 0))
        config_frame.configure(width=300)

        config_label = ctk.CTkLabel(
            config_frame,
            text="Configuration",
            font=(FONT_FAMILY, FONT_SIZE_LARGE, "bold"),
            text_color=COLORS['primary_text']
        )
        config_label.pack(pady=(10, 5))

        folder_frame = ctk.CTkFrame(
            config_frame,
            fg_color=COLORS['frame_bg'],
            border_width=1,
            border_color=COLORS['secondary_text']
        )
        folder_frame.pack(fill="x", padx=10, pady=10)

        folder_label = ctk.CTkLabel(
            folder_frame,
            text="Output Location",
            font=(FONT_FAMILY, FONT_SIZE_SMALL, "bold"),
            text_color=COLORS['primary_text']
        )
        folder_label.pack(pady=(5, 0), padx=10, anchor="w")

        self.output_folder_var = ctk.StringVar(value=os.path.expanduser("~\\Desktop"))

        folder_entry = ctk.CTkEntry(
            folder_frame,
            textvariable=self.output_folder_var,
            font=(FONT_FAMILY, 9),
            state="readonly",
            fg_color=COLORS['frame_bg'],
            text_color=COLORS['primary_text'],
            border_color=COLORS['secondary_text']
        )
        folder_entry.pack(fill="x", pady=(5, 5), padx=10)

        browse_folder_btn = ctk.CTkButton(
            folder_frame,
            text="Browse",
            command=self._browse_output_folder,
            font=(FONT_FAMILY, FONT_SIZE_SMALL),
            fg_color=COLORS['frame_bg'],
            hover_color=COLORS['accent_hover'],
            text_color=COLORS['button_text'],
            border_width=1,
            border_color=COLORS['secondary_text']
        )
        browse_folder_btn.pack(fill="x", padx=10, pady=(0, 5))

        filename_frame = ctk.CTkFrame(
            config_frame,
            fg_color=COLORS['frame_bg'],
            border_width=1,
            border_color=COLORS['secondary_text']
        )
        filename_frame.pack(fill="x", padx=10, pady=10)

        filename_label = ctk.CTkLabel(
            filename_frame,
            text="Output Filename",
            font=(FONT_FAMILY, FONT_SIZE_SMALL, "bold"),
            text_color=COLORS['primary_text']
        )
        filename_label.pack(pady=(5, 0), padx=10, anchor="w")

        self.output_filename_var = ctk.StringVar(value="merged_presentation.pptx")

        filename_entry = ctk.CTkEntry(
            filename_frame,
            textvariable=self.output_filename_var,
            font=(FONT_FAMILY, 9),
            fg_color=COLORS['frame_bg'],
            text_color=COLORS['primary_text'],
            border_color=COLORS['primary_accent']
        )
        filename_entry.pack(fill="x", pady=(5, 5), padx=10)

        button_frame = ctk.CTkFrame(config_frame, fg_color=COLORS['frame_bg'])
        button_frame.pack(fill="x", padx=10, pady=20)

        self.merge_btn = ctk.CTkButton(
            button_frame,
            text="Merge Presentations",
            command=self._on_merge,
            font=(FONT_FAMILY, 11, "bold"),
            fg_color=COLORS['primary_accent'],
            hover_color=COLORS['accent_hover'],
            text_color=COLORS['button_text'],
            height=40
        )
        self.merge_btn.pack(fill="x", pady=(0, 10))

        clear_btn = ctk.CTkButton(
            button_frame,
            text="Clear Queue",
            command=self._clear_queue,
            font=(FONT_FAMILY, FONT_SIZE_SMALL),
            fg_color=COLORS['frame_bg'],
            hover_color=COLORS['accent_hover'],
            text_color=COLORS['button_text'],
            border_width=1,
            border_color=COLORS['secondary_text']
        )
        clear_btn.pack(fill="x")

        self.status_var = ctk.StringVar(value="Ready")
        status_label = ctk.CTkLabel(
            config_frame,
            textvariable=self.status_var,
            font=(FONT_FAMILY, 9),
            text_color=COLORS['primary_accent'],
            wraplength=280,
            anchor="w"
        )
        status_label.pack(side="bottom", fill="x", padx=10, pady=10)

    def _create_file_selector(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        selector_container = ctk.CTkFrame(self.content_frame, fg_color=COLORS['frame_bg'])
        selector_container.pack(fill="both", expand=True)

        plus_label = ctk.CTkLabel(
            selector_container,
            text="+",
            font=(FONT_FAMILY, 72, "bold"),
            text_color=COLORS['secondary_text']
        )
        plus_label.pack(expand=True, pady=(50, 10))

        instruction_label = ctk.CTkLabel(
            selector_container,
            text="Add PowerPoint files using the button below",
            font=(FONT_FAMILY, FONT_SIZE_MEDIUM),
            text_color=COLORS['secondary_text']
        )
        instruction_label.pack()

        browse_btn = ctk.CTkButton(
            selector_container,
            text="Browse for Files",
            command=self._browse_files,
            font=(FONT_FAMILY, 11),
            width=200,
            height=40,
            fg_color=COLORS['frame_bg'],
            hover_color=COLORS['accent_hover'],
            text_color=COLORS['button_text'],
            border_width=1,
            border_color=COLORS['secondary_text']
        )
        browse_btn.pack(pady=(20, 50))

    def _create_file_list(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        scrollable_frame = ctk.CTkScrollableFrame(
            self.content_frame,
            fg_color=COLORS['frame_bg']
        )
        scrollable_frame.pack(fill="both", expand=True)

        self.file_list_frame = scrollable_frame
        
        # Clear the list of icon references before repopulating the frame
        self.file_card_icons.clear()

        for i, file_path in enumerate(self.file_list):
            self._create_file_card(i, file_path)

    def _create_file_card(self, index, file_path):
        card = ctk.CTkFrame(
            self.file_list_frame,
            fg_color=COLORS['frame_bg'],
            border_width=1,
            border_color=COLORS['secondary_text']
        )
        card.pack(fill="x", padx=5, pady=2)

        info_frame = ctk.CTkFrame(card, fg_color=COLORS['frame_bg'])
        info_frame.pack(side="left", fill="both", expand=True, padx=10, pady=5)
        
        if self.pil_powerpoint_icon:
            # *** FIX: Create a standard ImageTk.PhotoImage ***
            powerpoint_icon = ImageTk.PhotoImage(self.pil_powerpoint_icon)
            icon_label = ctk.CTkLabel(info_frame, image=powerpoint_icon, text="")
            
            # *** FIX: Store the reference in the instance list ***
            self.file_card_icons.append(powerpoint_icon)
            
            icon_label.pack(side="left", padx=(0, 10))
        else:
            # Fallback: Load pp.png and use CTkImage
            try:
                pp_png_path = os.path.join(os.path.dirname(__file__), "pp.png")
                pil_image = Image.open(pp_png_path)
                icon_image = ctk.CTkImage(light_image=pil_image,
                                         dark_image=pil_image,
                                         size=(24, 24))
                icon_label = ctk.CTkLabel(info_frame, image=icon_image, text="")
                icon_label.pack(side="left", padx=(0, 10))
            except FileNotFoundError:
                # Final fallback to text emoji if pp.png is not found
                logging.warning("Icon 'pp.png' not found. Using text emoji as fallback.")
                icon_label = ctk.CTkLabel(info_frame, text="🅿️", font=(FONT_FAMILY, 16))
                icon_label.pack(side="left", padx=(0, 10))

        filename = os.path.basename(file_path)
        name_label = ctk.CTkLabel(
            info_frame,
            text=filename,
            font=(FONT_FAMILY, 12),
            text_color=COLORS['accent_hover'],
            anchor="w"
        )
        name_label.pack(side="left", fill="x", expand=True)

        self._create_tooltip(name_label, file_path)

        button_frame = ctk.CTkFrame(card, fg_color=COLORS['frame_bg'])
        button_frame.pack(side="left", padx=5)

        up_btn = ctk.CTkButton(
            button_frame,
            text="↑",
            command=lambda idx=index: self._move_file_up(idx),
            font=(FONT_FAMILY, FONT_SIZE_SMALL),
            width=30,
            fg_color=COLORS['frame_bg'],
            hover_color=COLORS['accent_hover'],
            text_color=COLORS['accent_hover'],
            border_width=1,
            border_color=COLORS['primary_accent']
        )
        up_btn.pack(side="left", padx=2)

        down_btn = ctk.CTkButton(
            button_frame,
            text="↓",
            command=lambda idx=index: self._move_file_down(idx),
            font=(FONT_FAMILY, FONT_SIZE_SMALL),
            width=30,
            fg_color=COLORS['frame_bg'],
            hover_color=COLORS['accent_hover'],
            text_color=COLORS['accent_hover'],
            border_width=1,
            border_color=COLORS['primary_accent']
        )
        down_btn.pack(side="left", padx=2)

        remove_btn = ctk.CTkButton(
            card,
            text="✕",
            command=lambda idx=index: self._remove_file(idx),
            font=(FONT_FAMILY, FONT_SIZE_SMALL),
            text_color=COLORS['accent_hover'],
            fg_color=COLORS['frame_bg'],
            hover_color=COLORS['frame_bg'],
            width=30,
            border_width=0
        )
        remove_btn.pack(side="right", padx=5)

    def _create_tooltip(self, widget, text):
        def show_tooltip(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")

            label = tk.Label(
                tooltip,
                text=text,
                background=COLORS['frame_bg'],
                foreground=COLORS['primary_text'],
                relief=tk.SOLID,
                borderwidth=1,
                font=(FONT_FAMILY, 9)
            )
            label.pack()

            widget.tooltip = tooltip

        def hide_tooltip(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                delattr(widget, 'tooltip')

        widget.bind('<Enter>', show_tooltip)
        widget.bind('<Leave>', hide_tooltip)

    def _load_icon(self, filename, size=(20, 20)):
        """Try to load and resize an icon image. Returns a PhotoImage or None."""
        try:
            icon_path = os.path.join(
                os.path.dirname(__file__),
                "Merge_PowerPoint",
                "MergePowerPointPresentations",
                filename
            )
            pil_image = Image.open(icon_path).resize(size, Image.LANCZOS)
            tk_image = ImageTk.PhotoImage(pil_image)
            return tk_image
        except Exception as e:
            logging.warning(f"Could not load icon {filename}: {e}")
            return None

    def _update_merge_queue_display(self):
        if not self.file_list:
            self._create_file_selector()
            self.merge_btn.configure(state="disabled")
        else:
            self._create_file_list()
            self.merge_btn.configure(state="normal")

    def _browse_files(self):
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[
                ("PowerPoint Files", "*.pptx *.ppsx"),
                ("All Files", "*.*")
            ]
        )

        if files:
            self._add_files(list(files))

    def _add_files(self, file_paths):
        added_count = 0

        for file_path in file_paths:
            if not file_path.lower().endswith(('.pptx', '.ppsx')):
                messagebox.showwarning(
                    "Invalid File Type",
                    f"Only PowerPoint files (.pptx and .ppsx) are supported.\n\n"
                    f"Invalid file: {os.path.basename(file_path)}"
                )
                continue

            if file_path in self.file_list:
                messagebox.showinfo(
                    "Duplicate File",
                    f"This file has already been added.\n\n{os.path.basename(file_path)}"
                )
                continue

            if not os.path.exists(file_path):
                messagebox.showerror(
                    "File Not Found",
                    f"The specified file does not exist.\n\n{file_path}"
                )
                continue

            try:
                with open(file_path, 'rb'):
                    pass

                self.file_list.append(file_path)
                added_count += 1
                logging.info(f"Added file to queue: {file_path}")

            except PermissionError:
                messagebox.showerror(
                    "Access Denied",
                    f"Access denied. Unable to open the file.\n\n"
                    f"{os.path.basename(file_path)}\n\n"
                    f"The file may be in use by another application or you may "
                    f"lack the necessary permissions."
                )
            except Exception as e:
                messagebox.showerror(
                    "File Access Error",
                    f"Could not access file: {os.path.basename(file_path)}\n\n"
                    f"Error: {str(e)}"
                )

        if added_count > 0:
            self._update_merge_queue_display()
            self.status_var.set(f"Added {added_count} file(s) to queue")

    def _move_file_up(self, index):
        if index > 0:
            self.file_list[index], self.file_list[index-1] = \
                self.file_list[index-1], self.file_list[index]
            self._update_merge_queue_display()
            logging.info(f"Moved file up: {os.path.basename(self.file_list[index])}")

    def _move_file_down(self, index):
        if index < len(self.file_list) - 1:
            self.file_list[index], self.file_list[index+1] = \
                self.file_list[index+1], self.file_list[index]
            self._update_merge_queue_display()
            logging.info(f"Moved file down: {os.path.basename(self.file_list[index])}")

    def _remove_file(self, index):
        removed_file = self.file_list.pop(index)
        logging.info(f"Removed file from queue: {removed_file}")
        self._update_merge_queue_display()
        self.status_var.set(f"Removed: {os.path.basename(removed_file)}")

    def _clear_queue(self):
        if self.file_list:
            if messagebox.askyesno(
                "Clear Queue",
                "Are you sure you want to remove all files from the queue?"
            ):
                self.file_list.clear()
                self._update_merge_queue_display()
                self.status_var.set("Queue cleared")
                logging.info("Merge queue cleared")

    def _browse_output_folder(self):
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.output_folder_var.get()
        )

        if folder:
            self.output_folder_var.set(folder)
            logging.info(f"Output folder set to: {folder}")

    def _on_merge(self):
        if not self.file_list:
            messagebox.showwarning(
                "No Files",
                "Please add at least one PowerPoint file to the merge queue."
            )
            return

        filename = self.output_filename_var.get().strip()
        if not filename:
            messagebox.showerror("Invalid Filename", "Please enter a valid output filename.")
            return

        if not filename.lower().endswith('.pptx'):
            filename += '.pptx'
            self.output_filename_var.set(filename)

        output_path = os.path.join(self.output_folder_var.get(), filename)

        if os.path.exists(output_path):
            if not messagebox.askyesno(
                "File Exists",
                f"The file '{filename}' already exists.\n\nDo you want to overwrite it?"
            ):
                return

        self.merge_btn.configure(state="disabled")
        self.status_var.set(f"Merging {len(self.file_list)} presentations...")
        self.root.update()

        def merge_task():
            success, final_path, error = self.merge_callback(self.file_list.copy(), output_path)

            if success:
                self.run_in_main_thread(self.update_status, f"Merge complete: {final_path}")
            else:
                self.run_in_main_thread(self.update_status, f"Merge failed: {error}")

            self.run_in_main_thread(self.enable_merge_button)

        threading.Thread(target=merge_task, daemon=True).start()

    def update_status(self, message):
        self.status_var.set(message)
        self.root.update()

    def enable_merge_button(self):
        if self.file_list:
            self.merge_btn.configure(state="normal")

    def run_in_main_thread(self, func, *args, **kwargs):
        self.root.after(0, lambda: func(*args, **kwargs))

    def run(self):
        logging.info("Starting PowerPoint Merger GUI")
        self.root.mainloop()


def show_modern_gui(merge_callback):
    gui = PowerPointMergerGUI(merge_callback)
    gui.run()

