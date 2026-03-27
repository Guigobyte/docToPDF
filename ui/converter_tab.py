import os
import subprocess
import threading

import customtkinter as ctk

from core.converter import convert
from ui.drop_zone import DropZone


class ConverterTab:
    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent
        self._build_ui()

    def _build_ui(self):
        # Title
        ctk.CTkLabel(
            self.parent,
            text="DOCX \u2192 PDF Converter",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=(18, 4))

        ctk.CTkLabel(
            self.parent,
            text="Drop a Word document to convert it to PDF",
            font=ctk.CTkFont(size=12),
            text_color=("gray45", "gray55"),
        ).pack(pady=(0, 12))

        # Drop zone
        self.drop_zone = DropZone(
            self.parent,
            allowed_extensions=[".docx"],
            prompt_text="Drop .docx file here or click Browse",
            on_drop=self._on_file_dropped,
            height=180,
        )
        self.drop_zone.pack(padx=30, pady=(0, 10), fill="x")

        # Status frame
        self.status_frame = ctk.CTkFrame(
            self.parent, fg_color="transparent"
        )
        self.status_frame.pack(padx=30, fill="x")

        # File name label
        self.file_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray70"),
        )
        self.file_label.pack(pady=(4, 2))

        # Progress bar
        self.progress = ctk.CTkProgressBar(self.status_frame, mode="indeterminate")
        self.progress.pack(pady=(4, 4), fill="x", padx=40)
        self.progress.pack_forget()  # hidden initially

        # Result label
        self.result_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.result_label.pack(pady=(4, 2))

        # Open folder button (hidden initially)
        self.open_btn = ctk.CTkButton(
            self.status_frame,
            text="Open Output Folder",
            width=140,
            height=30,
            font=ctk.CTkFont(size=12),
            command=self._open_folder,
        )
        self.open_btn.pack(pady=(4, 0))
        self.open_btn.pack_forget()

        self._output_dir = None

    def _on_file_dropped(self, path: str):
        filename = os.path.basename(path)
        self.file_label.configure(text=f"File: {filename}")
        self.result_label.configure(text="")
        self.open_btn.pack_forget()

        # Show progress
        self.progress.pack(pady=(4, 4), fill="x", padx=40)
        self.progress.start()

        # Run conversion in background thread
        thread = threading.Thread(target=self._convert, args=(path,), daemon=True)
        thread.start()

    def _convert(self, path: str):
        try:
            output = convert(path)
            self._output_dir = os.path.dirname(output)
            output_name = os.path.basename(output)
            self.parent.after(0, self._show_success, output_name)
        except FileNotFoundError:
            self.parent.after(
                0,
                self._show_error,
                "Microsoft Word is required for conversion.\nPlease install Word and try again.",
            )
        except PermissionError:
            self.parent.after(
                0,
                self._show_error,
                "File is open in another program.\nClose it and try again.",
            )
        except Exception as e:
            self.parent.after(0, self._show_error, f"Conversion failed:\n{e}")

    def _show_success(self, output_name: str):
        self.progress.stop()
        self.progress.pack_forget()
        self.result_label.configure(
            text=f"\u2705  Saved: {output_name}",
            text_color=("#1B8C3A", "#4ADE80"),
        )
        self.open_btn.pack(pady=(4, 0))

    def _show_error(self, message: str):
        self.progress.stop()
        self.progress.pack_forget()
        self.result_label.configure(
            text=f"\u274C  {message}",
            text_color=("#DC2626", "#F87171"),
        )

    def _open_folder(self):
        if self._output_dir and os.path.isdir(self._output_dir):
            subprocess.Popen(["explorer", os.path.normpath(self._output_dir)])
