import os
import subprocess
import threading

import customtkinter as ctk

from core.converter import convert
from ui.drop_zone import DropZone


class ConverterTab:
    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent
        self._pending_path = None
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

        # Overwrite prompt frame (hidden initially)
        self.overwrite_frame = ctk.CTkFrame(
            self.status_frame, corner_radius=10,
            fg_color=("#FEF3C7", "#2E2A1A"),
        )

        ctk.CTkLabel(
            self.overwrite_frame,
            text="\u26A0\uFE0F  A PDF already exists for this file.",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#B45309", "#FBBF24"),
        ).pack(pady=(10, 4))

        self.overwrite_detail = ctk.CTkLabel(
            self.overwrite_frame,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray55"),
            wraplength=350,
        )
        self.overwrite_detail.pack(pady=(0, 8))

        btn_row = ctk.CTkFrame(self.overwrite_frame, fg_color="transparent")
        btn_row.pack(pady=(0, 10))

        ctk.CTkButton(
            btn_row,
            text="Overwrite",
            width=100,
            height=30,
            font=ctk.CTkFont(size=12),
            fg_color=("#B45309", "#92400E"),
            hover_color=("#92400E", "#78350F"),
            command=self._confirm_overwrite,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_row,
            text="Cancel",
            width=80,
            height=30,
            font=ctk.CTkFont(size=12),
            fg_color=("gray70", "gray35"),
            hover_color=("gray60", "gray45"),
            command=self._cancel_overwrite,
        ).pack(side="left")

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
        self.overwrite_frame.pack_forget()

        # Check if output PDF already exists
        pdf_path = os.path.splitext(path)[0] + ".pdf"
        if os.path.isfile(pdf_path):
            self._pending_path = path
            pdf_name = os.path.basename(pdf_path)
            self.overwrite_detail.configure(
                text=f'"{pdf_name}" already exists in the same folder.\nOverwrite it with a new conversion?'
            )
            self.overwrite_frame.pack(pady=(4, 4), fill="x", padx=10)
            return

        self._start_conversion(path)

    def _confirm_overwrite(self):
        self.overwrite_frame.pack_forget()
        if self._pending_path:
            path = self._pending_path
            self._pending_path = None
            self._start_conversion(path)

    def _cancel_overwrite(self):
        self.overwrite_frame.pack_forget()
        self._pending_path = None
        self.file_label.configure(text="")
        self.result_label.configure(text="")

    def _start_conversion(self, path: str):
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
