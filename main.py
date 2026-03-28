import sys
import os
import logging
import traceback
import queue

# --- Crash logger (writes next to the exe / main.py) ----------------------
_log_dir = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                           else os.path.abspath(__file__))
_log_file = os.path.join(_log_dir, "DocxToPDF_crash.log")
try:
    logging.basicConfig(
        filename=_log_file,
        level=logging.ERROR,
        format="%(asctime)s  %(message)s",
    )
except Exception:
    # Fall back to temp dir if exe dir is read-only (e.g. Program Files)
    import tempfile
    _log_file = os.path.join(tempfile.gettempdir(), "DocxToPDF_crash.log")
    logging.basicConfig(
        filename=_log_file,
        level=logging.ERROR,
        format="%(asctime)s  %(message)s",
    )

def _global_exception_handler(exc_type, exc_value, exc_tb):
    msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    logging.error("Unhandled exception:\n%s", msg)
    # Flush immediately so the message is written before process dies
    for handler in logging.getLogger().handlers:
        handler.flush()
    sys.__excepthook__(exc_type, exc_value, exc_tb)

sys.excepthook = _global_exception_handler

# --- Bootstrap for frozen (PyInstaller) builds ----------------------------
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
    if hasattr(sys, "_MEIPASS"):
        sys.path.insert(0, sys._MEIPASS)

import customtkinter as ctk

from ui.converter_tab import ConverterTab
from ui.validator_tab import ValidatorTab


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Catch tkinter callback exceptions to the log file
        self.report_callback_exception = self._on_tk_error

        self.title("DocxToPDF")
        self.geometry("560x560")
        self.minsize(480, 480)
        self.resizable(True, True)

        # Set icon (try root-level .ico, then assets fallback)
        icon_path = self._resource_path("DocxToPDF.ico")
        if not os.path.exists(icon_path):
            icon_path = self._resource_path("assets", "icon.ico")
        if os.path.exists(icon_path):
            self.iconbitmap(icon_path)

        # Tab view
        self.tabview = ctk.CTkTabview(self, anchor="nw")
        self.tabview.pack(fill="both", expand=True, padx=12, pady=(8, 12))

        # Create tabs
        convert_tab = self.tabview.add("Convert")
        validate_tab = self.tabview.add("Validate")

        self.converter = ConverterTab(convert_tab)
        self.validator = ValidatorTab(validate_tab)

        # Queue for processing dropped files on the main thread.
        # windnd's native WM_DROPFILES handler MUST NOT call any Tkinter
        # widget methods (configure, pack, after, etc.) — doing so causes
        # reentrant Tcl interpreter access and a native crash that bypasses
        # all Python exception handlers. Instead, we queue file paths and
        # process them via a safe main-thread polling loop.
        self._drop_queue = queue.Queue()
        self._setup_dnd()
        self._poll_drop_queue()

    @staticmethod
    def _on_tk_error(exc_type, exc_value, exc_tb):
        msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logging.error("Tkinter callback error:\n%s", msg)
        for handler in logging.getLogger().handlers:
            handler.flush()

    def _resource_path(self, *parts):
        base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base, *parts)

    def _setup_dnd(self):
        """Register Windows-native drag-and-drop using windnd."""
        try:
            import windnd
            windnd.hook_dropfiles(self, func=self._on_drop)
        except Exception:
            pass  # Drag-and-drop unavailable; browse button still works

    def _poll_drop_queue(self):
        """Process queued file drops on the main thread (safe for Tk calls)."""
        try:
            while True:
                path, tab = self._drop_queue.get_nowait()
                if tab == "Convert":
                    self.converter.drop_zone.handle_drop_data(path)
                elif tab == "Validate":
                    self.validator.drop_zone.handle_drop_data(path)
        except queue.Empty:
            pass
        except Exception:
            logging.error("Error processing drop:\n%s", traceback.format_exc())
        self.after(50, self._poll_drop_queue)

    def _on_drop(self, file_list):
        """Handle files dropped onto the window via windnd.

        CRITICAL: This runs inside windnd's native WM_DROPFILES handler.
        We MUST NOT call ANY Tkinter/CTk widget methods here (configure,
        pack, after, event_generate, etc.) — that causes reentrant Tcl
        interpreter access and an instant native crash with no traceback.

        Only pure-Python operations (string ops, os.path, queue.put) are
        safe here. All widget work is deferred to _poll_drop_queue().
        """
        try:
            # tabview.get() is safe: returns a stored Python string attribute,
            # does not call into Tcl.
            active_tab = self.tabview.get()
            for raw_path in file_list:
                if isinstance(raw_path, bytes):
                    path = raw_path.decode("utf-8", errors="replace")
                else:
                    path = str(raw_path)
                path = os.path.normpath(path.strip())
                if not path or not os.path.isfile(path):
                    continue
                self._drop_queue.put((path, active_tab))
        except Exception:
            pass  # Cannot safely log from inside native handler


if __name__ == "__main__":
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
