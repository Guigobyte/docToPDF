import logging
import os
import threading
import traceback

import customtkinter as ctk

from core.converter import convert
from ui.drop_zone import DropZone


class ConverterTab:
    def __init__(self, parent: ctk.CTkFrame):
        self.parent = parent
        self._pending_path = None
        self._output_dir = None
        self._converting = False
        self._build_ui()

    def _build_ui(self):
        # Title
        ctk.CTkLabel(
            self.parent,
            text="DOCX \u2192 PDF Converter",
            font=ctk.CTkFont(size=20, weight="bold"),
        ).pack(pady=(14, 2))

        ctk.CTkLabel(
            self.parent,
            text="Drop a Word document to convert it to PDF",
            font=ctk.CTkFont(size=12),
            text_color=("gray45", "gray55"),
        ).pack(pady=(0, 8))

        # Drop zone
        self.drop_zone = DropZone(
            self.parent,
            allowed_extensions=[".docx"],
            prompt_text="Drop .docx file here or click Browse",
            on_drop=self._on_file_dropped,
            height=150,
        )
        self.drop_zone.pack(padx=24, pady=(0, 8), fill="x")

        # --- Status area: all widgets always packed, visibility via text ---
        self.status_frame = ctk.CTkFrame(self.parent, fg_color="transparent")
        self.status_frame.pack(padx=24, fill="x")

        # File name label (always present, empty when idle)
        self.file_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=("gray30", "gray70"),
        )
        self.file_label.pack(pady=(4, 2))

        # Result label (always present, empty when idle)
        self.result_label = ctk.CTkLabel(
            self.status_frame,
            text="",
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.result_label.pack(pady=(4, 2))

        # Progress bar — this is the ONE widget we show/hide because
        # CTkProgressBar has no "empty" state. We use a wrapper frame
        # so we only ever pack_forget/pack the wrapper, not the progress
        # bar itself (avoids CTk internal state issues).
        self._progress_wrapper = ctk.CTkFrame(
            self.status_frame, fg_color="transparent", height=0
        )
        self._progress_wrapper.pack(fill="x", padx=40)
        self._progress_wrapper.pack_propagate(False)  # don't expand when empty
        self.progress = ctk.CTkProgressBar(
            self._progress_wrapper, mode="indeterminate"
        )
        # progress is NOT packed yet — we pack it when needed

        # Overwrite prompt — always present, content toggled via visibility.
        # Uses a wrapper frame that collapses when empty.
        self._overwrite_wrapper = ctk.CTkFrame(
            self.status_frame, fg_color="transparent", height=0
        )
        self._overwrite_wrapper.pack(fill="x", padx=10)
        self._overwrite_wrapper.pack_propagate(False)

        self.overwrite_inner = ctk.CTkFrame(
            self._overwrite_wrapper, corner_radius=10,
            fg_color=("#FEF3C7", "#2E2A1A"),
        )

        ctk.CTkLabel(
            self.overwrite_inner,
            text="\u26A0\uFE0F  A PDF already exists for this file.",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#B45309", "#FBBF24"),
        ).pack(pady=(10, 4))

        self.overwrite_detail = ctk.CTkLabel(
            self.overwrite_inner,
            text="",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray55"),
            wraplength=350,
        )
        self.overwrite_detail.pack(pady=(0, 8))

        btn_row = ctk.CTkFrame(self.overwrite_inner, fg_color="transparent")
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

        # Open folder button (always present, hidden via empty text trick
        # doesn't work for buttons — use wrapper)
        self._open_wrapper = ctk.CTkFrame(
            self.status_frame, fg_color="transparent", height=0
        )
        self._open_wrapper.pack(fill="x")
        self._open_wrapper.pack_propagate(False)

        self.open_btn = ctk.CTkButton(
            self._open_wrapper,
            text="Open Output Folder",
            width=140,
            height=30,
            font=ctk.CTkFont(size=12),
            command=self._open_folder,
        )

    # -- Widget show/hide helpers using wrapper frames --------------------
    # Instead of pack_forget on CTk widgets directly (which can leak state),
    # we pack/unpack the child INSIDE a fixed wrapper frame.

    def _show_progress(self):
        self._progress_wrapper.configure(height=30)
        self.progress.pack(pady=(4, 4), fill="x")
        self.progress.start()

    def _hide_progress(self):
        self.progress.stop()
        self.progress.pack_forget()
        self._progress_wrapper.configure(height=0)

    def _show_overwrite(self):
        self._overwrite_wrapper.configure(height=160)
        self.overwrite_inner.pack(pady=(4, 4), fill="x")

    def _hide_overwrite(self):
        self.overwrite_inner.pack_forget()
        self._overwrite_wrapper.configure(height=0)

    def _show_open_btn(self):
        self._open_wrapper.configure(height=36)
        self.open_btn.pack(pady=(4, 0))

    def _hide_open_btn(self):
        self.open_btn.pack_forget()
        self._open_wrapper.configure(height=0)

    # -- Reset everything to idle -----------------------------------------

    def _reset_ui(self):
        """Clear all status indicators back to idle."""
        self.file_label.configure(text="")
        self.result_label.configure(text="")
        self._hide_progress()
        self._hide_overwrite()
        self._hide_open_btn()

    # -- Drop handling ----------------------------------------------------

    def _on_file_dropped(self, path: str):
        try:
            # Ignore drops while a conversion is running
            if self._converting:
                return

            self._reset_ui()

            filename = os.path.basename(path)
            self.file_label.configure(text=f"File: {filename}")

            # Check if output PDF already exists
            pdf_path = os.path.splitext(path)[0] + ".pdf"
            if os.path.isfile(pdf_path):
                self._pending_path = path
                pdf_name = os.path.basename(pdf_path)
                self.overwrite_detail.configure(
                    text=f'"{pdf_name}" already exists in the same '
                         f'folder.\nOverwrite it with a new conversion?'
                )
                self._show_overwrite()
                return

            self._start_conversion(path)
        except Exception:
            logging.error("Error in _on_file_dropped:\n%s",
                          traceback.format_exc())

    def _confirm_overwrite(self):
        try:
            self._hide_overwrite()
            if self._pending_path:
                path = self._pending_path
                self._pending_path = None
                self._start_conversion(path)
        except Exception:
            logging.error("Error in _confirm_overwrite:\n%s",
                          traceback.format_exc())

    def _cancel_overwrite(self):
        try:
            self._pending_path = None
            self._reset_ui()
        except Exception:
            logging.error("Error in _cancel_overwrite:\n%s",
                          traceback.format_exc())

    # -- Conversion -------------------------------------------------------

    def _start_conversion(self, path: str):
        self._converting = True
        self._show_progress()
        # daemon=True is intentional: if user closes the app during conversion,
        # the thread is killed. The converter.py finally block handles COM cleanup
        # (word.Quit + CoUninitialize) so Word won't be left as a zombie.
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
            logging.error("Conversion error:\n%s", traceback.format_exc())
            self.parent.after(0, self._show_error, f"Conversion failed:\n{e}")

    def _show_success(self, output_name: str):
        try:
            self._converting = False
            self._hide_progress()
            self.result_label.configure(
                text=f"\u2705  Saved: {output_name}",
                text_color=("#1B8C3A", "#4ADE80"),
            )
            self._show_open_btn()
        except Exception:
            logging.error("Error in _show_success:\n%s",
                          traceback.format_exc())

    def _show_error(self, message: str):
        try:
            self._converting = False
            self._hide_progress()
            self.result_label.configure(
                text=f"\u274C  {message}",
                text_color=("#DC2626", "#F87171"),
            )
        except Exception:
            logging.error("Error in _show_error:\n%s",
                          traceback.format_exc())

    def _open_folder(self):
        if self._output_dir and os.path.isdir(self._output_dir):
            os.startfile(os.path.normpath(self._output_dir))
