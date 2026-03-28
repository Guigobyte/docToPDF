import os
import subprocess
import threading

import customtkinter as ctk

from core.converter import convert
from ui.drop_zone import DropZone

# UI states
_STATE_IDLE = "idle"
_STATE_OVERWRITE = "overwrite"
_STATE_CONVERTING = "converting"
_STATE_SUCCESS = "success"
_STATE_ERROR = "error"


class ConverterTab:
    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent
        self._pending_path = None
        self._output_dir = None
        self._state = _STATE_IDLE
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

        # Status area — uses grid so we can grid_remove instead of pack_forget.
        # grid_remove preserves config so re-showing is a clean grid() call.
        self.status_frame = ctk.CTkFrame(self.parent, fg_color="transparent")
        self.status_frame.pack(padx=30, fill="x")
        self.status_frame.columnconfigure(0, weight=1)

        row = 0

        # Row 0: File name label (always visible when a file is loaded)
        self.file_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray70"),
        )
        self.file_label.grid(row=row, column=0, pady=(4, 2))
        self.file_label.grid_remove()
        row += 1

        # Row 1: Progress bar (converting state)
        self.progress = ctk.CTkProgressBar(
            self.status_frame, mode="indeterminate"
        )
        self.progress.grid(row=row, column=0, pady=(4, 4), sticky="ew", padx=40)
        self.progress.grid_remove()
        row += 1

        # Row 2: Result label (success / error state)
        self.result_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.result_label.grid(row=row, column=0, pady=(4, 2))
        self.result_label.grid_remove()
        row += 1

        # Row 3: Overwrite prompt (overwrite state)
        self.overwrite_frame = ctk.CTkFrame(
            self.status_frame, corner_radius=10,
            fg_color=("#FEF3C7", "#2E2A1A"),
        )
        self.overwrite_frame.grid(
            row=row, column=0, pady=(4, 4), sticky="ew", padx=10
        )
        self.overwrite_frame.grid_remove()

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

        row += 1

        # Row 4: Open folder button (success state)
        self.open_btn = ctk.CTkButton(
            self.status_frame,
            text="Open Output Folder",
            width=140,
            height=30,
            font=ctk.CTkFont(size=12),
            command=self._open_folder,
        )
        self.open_btn.grid(row=row, column=0, pady=(4, 0))
        self.open_btn.grid_remove()

    # -- State transitions ------------------------------------------------

    def _set_state(self, state: str):
        """Transition to a new UI state, showing/hiding the right widgets."""
        if state == self._state:
            return
        old = self._state
        self._state = state

        # Hide widgets from old state
        if old == _STATE_OVERWRITE:
            self.overwrite_frame.grid_remove()
        elif old == _STATE_CONVERTING:
            self.progress.stop()
            self.progress.grid_remove()
        elif old == _STATE_SUCCESS:
            self.result_label.grid_remove()
            self.open_btn.grid_remove()
        elif old == _STATE_ERROR:
            self.result_label.grid_remove()

        # Show widgets for new state
        if state == _STATE_IDLE:
            self.file_label.grid_remove()
        elif state == _STATE_OVERWRITE:
            self.file_label.grid()
            self.overwrite_frame.grid()
        elif state == _STATE_CONVERTING:
            self.file_label.grid()
            self.progress.grid()
            self.progress.start()
        elif state == _STATE_SUCCESS:
            self.file_label.grid()
            self.result_label.grid()
            self.open_btn.grid()
        elif state == _STATE_ERROR:
            self.file_label.grid()
            self.result_label.grid()

    # -- Drop handling ----------------------------------------------------

    def _on_file_dropped(self, path: str):
        filename = os.path.basename(path)
        self.file_label.configure(text=f"File: {filename}")

        # Check if output PDF already exists
        pdf_path = os.path.splitext(path)[0] + ".pdf"
        if os.path.isfile(pdf_path):
            self._pending_path = path
            pdf_name = os.path.basename(pdf_path)
            self.overwrite_detail.configure(
                text=f'"{pdf_name}" already exists in the same folder.\n'
                     f'Overwrite it with a new conversion?'
            )
            self._set_state(_STATE_OVERWRITE)
            return

        self._start_conversion(path)

    def _confirm_overwrite(self):
        if self._pending_path:
            path = self._pending_path
            self._pending_path = None
            self._start_conversion(path)

    def _cancel_overwrite(self):
        self._pending_path = None
        self._set_state(_STATE_IDLE)

    # -- Conversion -------------------------------------------------------

    def _start_conversion(self, path: str):
        self._set_state(_STATE_CONVERTING)
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
                "Microsoft Word is required for conversion.\n"
                "Please install Word and try again.",
            )
        except PermissionError:
            self.parent.after(
                0,
                self._show_error,
                "File is open in another program.\n"
                "Close it and try again.",
            )
        except Exception as e:
            self.parent.after(0, self._show_error, f"Conversion failed:\n{e}")

    def _show_success(self, output_name: str):
        self.result_label.configure(
            text=f"\u2705  Saved: {output_name}",
            text_color=("#1B8C3A", "#4ADE80"),
        )
        self._set_state(_STATE_SUCCESS)

    def _show_error(self, message: str):
        self.result_label.configure(
            text=f"\u274C  {message}",
            text_color=("#DC2626", "#F87171"),
        )
        self._set_state(_STATE_ERROR)

    def _open_folder(self):
        if self._output_dir and os.path.isdir(self._output_dir):
            subprocess.Popen(["explorer", os.path.normpath(self._output_dir)])
