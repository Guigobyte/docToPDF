import logging
import os
import tkinter as tk
import traceback
from tkinter import filedialog

import customtkinter as ctk


class DropZone(ctk.CTkFrame):
    """A drag-and-drop zone that accepts files with specified extensions."""

    def __init__(
        self,
        master,
        allowed_extensions: list[str],
        prompt_text: str = "Drop file here or click Browse",
        on_drop: callable = None,
        **kwargs,
    ):
        super().__init__(master, **kwargs)
        self.allowed_extensions = [e.lower() for e in allowed_extensions]
        self.on_drop = on_drop
        self._reject_timer = None

        self.configure(
            corner_radius=12,
            border_width=2,
            border_color=("#BBBBBB", "#555555"),
            fg_color=("gray92", "gray17"),
        )

        # Icon label
        self.icon_label = ctk.CTkLabel(
            self,
            text="\U0001F4C4",  # document emoji
            font=ctk.CTkFont(size=36),
            text_color=("gray50", "gray60"),
        )
        self.icon_label.pack(pady=(20, 4))

        # Prompt text
        self.prompt_label = ctk.CTkLabel(
            self,
            text=prompt_text,
            font=ctk.CTkFont(size=13),
            text_color=("gray40", "gray60"),
        )
        self.prompt_label.pack(pady=(0, 6))

        # Allowed types hint
        ext_text = ", ".join(self.allowed_extensions)
        self.hint_label = ctk.CTkLabel(
            self,
            text=f"Accepted: {ext_text}",
            font=ctk.CTkFont(size=11),
            text_color=("gray55", "gray50"),
        )
        self.hint_label.pack(pady=(0, 8))

        # Browse button
        self.browse_btn = ctk.CTkButton(
            self,
            text="Browse",
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            command=self._browse,
        )
        self.browse_btn.pack(pady=(0, 18))

    def _browse(self):
        filetypes = [
            (ext.upper().lstrip(".") + " files", f"*{ext}")
            for ext in self.allowed_extensions
        ]
        filetypes.append(("All files", "*.*"))
        path = filedialog.askopenfilename(filetypes=filetypes)
        if path and isinstance(path, str):
            self._handle_file(path)

    def _handle_file(self, path: str):
        try:
            ext = os.path.splitext(path)[1].lower()
            if ext in self.allowed_extensions:
                if self.on_drop:
                    self.on_drop(path)
            else:
                self._show_rejection(ext)
        except Exception:
            logging.error("Error in _handle_file:\n%s", traceback.format_exc())

    def _show_rejection(self, ext: str):
        """Flash red border and show rejection message for wrong file type."""
        # Cancel any pending reset
        if self._reject_timer is not None:
            self.after_cancel(self._reject_timer)

        self.configure(border_color="#FF4444")
        ext_text = ", ".join(self.allowed_extensions)
        self.hint_label.configure(
            text=f"'{ext}' not accepted. Use: {ext_text}",
            text_color=("#DC2626", "#F87171"),
        )
        self._reject_timer = self.after(3000, self._reset_appearance)

    def _reset_appearance(self):
        """Reset border and hint text to defaults."""
        self._reject_timer = None
        self.configure(border_color=("#BBBBBB", "#555555"))
        ext_text = ", ".join(self.allowed_extensions)
        self.hint_label.configure(
            text=f"Accepted: {ext_text}",
            text_color=("gray55", "gray50"),
        )

    def handle_drop_data(self, data: str):
        """Handle a file path from windnd or tkdnd."""
        try:
            path = data.strip().strip("{}")
            if path and os.path.isfile(path):
                self._handle_file(path)
        except Exception:
            logging.error("Error in handle_drop_data:\n%s",
                          traceback.format_exc())

    def set_highlight(self, on: bool):
        if on:
            self.configure(
                border_color=("#3B8ED0", "#3B8ED0"),
                fg_color=("gray88", "gray20"),
            )
        else:
            self.configure(
                border_color=("#BBBBBB", "#555555"),
                fg_color=("gray92", "gray17"),
            )
