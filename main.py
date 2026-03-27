import sys
import os

# Ensure bundled app can find its modules
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

        self.title("DocToPDF")
        self.geometry("600x580")
        self.minsize(500, 500)
        self.resizable(True, True)

        # Try to set icon
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

        # Set up drag-and-drop via windnd (Windows-native, reliable)
        self._setup_dnd()

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

    def _on_drop(self, file_list):
        """Handle files dropped onto the window."""
        active_tab = self.tabview.get()
        for raw_path in file_list:
            # windnd gives bytes on some versions
            if isinstance(raw_path, bytes):
                path = raw_path.decode("utf-8", errors="replace")
            else:
                path = str(raw_path)

            if active_tab == "Convert":
                self.converter.drop_zone.handle_drop_data(path)
            elif active_tab == "Validate":
                self.validator.drop_zone.handle_drop_data(path)


if __name__ == "__main__":
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    app = App()
    app.mainloop()
