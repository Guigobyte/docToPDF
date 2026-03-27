import os
import threading

import customtkinter as ctk

from core.validator import Result, validate
from ui.drop_zone import DropZone


class ValidatorTab:
    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent
        self.docx_path: str | None = None
        self.pdf_path: str | None = None
        self._build_ui()

    def _build_ui(self):
        # Title
        ctk.CTkLabel(
            self.parent,
            text="DOCX vs PDF Validator",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=(18, 4))

        ctk.CTkLabel(
            self.parent,
            text="Drop both files to verify the PDF was generated from the DOCX",
            font=ctk.CTkFont(size=12),
            text_color=("gray45", "gray55"),
        ).pack(pady=(0, 12))

        # Single drop zone for both file types
        self.drop_zone = DropZone(
            self.parent,
            allowed_extensions=[".docx", ".pdf"],
            prompt_text="Drop .docx and .pdf files here",
            on_drop=self._on_file_dropped,
            height=150,
        )
        self.drop_zone.pack(padx=30, pady=(0, 10), fill="x")

        # File status row
        self.files_frame = ctk.CTkFrame(self.parent, fg_color="transparent")
        self.files_frame.pack(padx=30, fill="x")
        self.files_frame.columnconfigure(0, weight=1)
        self.files_frame.columnconfigure(1, weight=1)

        # DOCX file indicator
        self.docx_frame = ctk.CTkFrame(self.files_frame, corner_radius=8)
        self.docx_frame.grid(row=0, column=0, padx=(0, 6), sticky="ew")

        ctk.CTkLabel(
            self.docx_frame,
            text="DOCX",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray40", "gray55"),
        ).pack(pady=(6, 0))

        self.docx_label = ctk.CTkLabel(
            self.docx_frame,
            text="Waiting...",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray50"),
            wraplength=220,
        )
        self.docx_label.pack(pady=(0, 6), padx=8)

        # PDF file indicator
        self.pdf_frame = ctk.CTkFrame(self.files_frame, corner_radius=8)
        self.pdf_frame.grid(row=0, column=1, padx=(6, 0), sticky="ew")

        ctk.CTkLabel(
            self.pdf_frame,
            text="PDF",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("gray40", "gray55"),
        ).pack(pady=(6, 0))

        self.pdf_label = ctk.CTkLabel(
            self.pdf_frame,
            text="Waiting...",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray50"),
            wraplength=220,
        )
        self.pdf_label.pack(pady=(0, 6), padx=8)

        # Result indicator
        self.result_frame = ctk.CTkFrame(
            self.parent, corner_radius=12, fg_color="transparent"
        )
        self.result_frame.pack(padx=30, pady=(12, 0), fill="x")

        self.result_icon = ctk.CTkLabel(
            self.result_frame,
            text="",
            font=ctk.CTkFont(size=40),
        )
        self.result_icon.pack(pady=(4, 0))

        self.result_label = ctk.CTkLabel(
            self.result_frame,
            text="",
            font=ctk.CTkFont(size=14, weight="bold"),
        )
        self.result_label.pack(pady=(2, 0))

        self.result_detail = ctk.CTkLabel(
            self.result_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray55"),
            wraplength=400,
        )
        self.result_detail.pack(pady=(2, 4))

        # Clear button
        self.clear_btn = ctk.CTkButton(
            self.parent,
            text="Clear",
            width=80,
            height=28,
            font=ctk.CTkFont(size=12),
            fg_color=("gray70", "gray35"),
            hover_color=("gray60", "gray45"),
            command=self._clear,
        )
        self.clear_btn.pack(pady=(10, 0))

    def _on_file_dropped(self, path: str):
        try:
            ext = os.path.splitext(path)[1].lower()
            name = os.path.basename(path)

            if ext == ".docx":
                self.docx_path = path
                self.docx_label.configure(
                    text=name, text_color=("gray15", "gray85")
                )
            elif ext == ".pdf":
                self.pdf_path = path
                self.pdf_label.configure(
                    text=name, text_color=("gray15", "gray85")
                )
            else:
                return

            # Auto-validate when both files are present
            if self.docx_path and self.pdf_path:
                self._run_validation()
        except Exception:
            pass  # Never crash from drop callback

    def _run_validation(self):
        """Run validation in a background thread to avoid freezing the UI."""
        # Show a brief "checking" state
        self.result_icon.configure(text="")
        self.result_label.configure(
            text="Checking...",
            text_color=("gray40", "gray55"),
        )
        self.result_detail.configure(text="")
        self.result_frame.configure(fg_color="transparent")

        docx = self.docx_path
        pdf = self.pdf_path

        thread = threading.Thread(
            target=self._validate_thread, args=(docx, pdf), daemon=True
        )
        thread.start()

    def _validate_thread(self, docx_path: str, pdf_path: str):
        """Perform validation off the main thread."""
        try:
            result_code, message = validate(docx_path, pdf_path)
            self.parent.after(0, self._show_result, result_code, message)
        except Exception as e:
            self.parent.after(0, self._show_result, Result.ERROR, str(e))

    def _show_result(self, result_code: str, message: str):
        """Update the UI with the validation result (called on main thread)."""
        if result_code == Result.MATCH:
            self.result_icon.configure(text="\u2705")
            self.result_label.configure(
                text="MATCH",
                text_color=("#1B8C3A", "#4ADE80"),
            )
            self.result_frame.configure(fg_color=("#E8F5E9", "#1A2E1A"))
        elif result_code == Result.MISMATCH:
            self.result_icon.configure(text="\u274C")
            self.result_label.configure(
                text="MISMATCH",
                text_color=("#DC2626", "#F87171"),
            )
            self.result_frame.configure(fg_color=("#FEE2E2", "#2E1A1A"))
        elif result_code == Result.NO_METADATA:
            self.result_icon.configure(text="\u26A0\uFE0F")
            self.result_label.configure(
                text="UNKNOWN",
                text_color=("#B45309", "#FBBF24"),
            )
            self.result_frame.configure(fg_color=("#FEF3C7", "#2E2A1A"))
        else:
            self.result_icon.configure(text="\u26A0\uFE0F")
            self.result_label.configure(
                text="ERROR",
                text_color=("#DC2626", "#F87171"),
            )
            self.result_frame.configure(fg_color=("#FEE2E2", "#2E1A1A"))

        self.result_detail.configure(text=message)

    def _clear(self):
        self.docx_path = None
        self.pdf_path = None
        self.docx_label.configure(
            text="Waiting...", text_color=("gray50", "gray50")
        )
        self.pdf_label.configure(
            text="Waiting...", text_color=("gray50", "gray50")
        )
        self.result_icon.configure(text="")
        self.result_label.configure(text="")
        self.result_detail.configure(text="")
        self.result_frame.configure(fg_color="transparent")
