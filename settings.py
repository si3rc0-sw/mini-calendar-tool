"""JSON-based settings persistence for the mini calendar."""

import json
import os
import sys
import winreg

_SETTINGS_PATH = os.path.join(os.path.expanduser("~"), ".mini-calendar-settings.json")

_DEFAULTS = {
    "dark_mode": False,
    "window_width": None,
    "window_height": None,
    "grid_cols": None,
    "grid_rows": None,
    "holidays": [],
    "holiday_colors": {"CH": "#FF0000", "DE": "#FFD700", "CN": "#4CAF50"},
}


def load_settings() -> dict:
    """Load settings from disk, returning defaults for missing keys."""
    settings = dict(_DEFAULTS)
    try:
        with open(_SETTINGS_PATH, "r", encoding="utf-8") as f:
            stored = json.load(f)
        if "dark_mode" in stored and isinstance(stored["dark_mode"], bool):
            settings["dark_mode"] = stored["dark_mode"]
        for key in ("window_width", "window_height", "grid_cols", "grid_rows"):
            if key in stored and isinstance(stored[key], int):
                settings[key] = stored[key]
        if "holidays" in stored and isinstance(stored["holidays"], list):
            settings["holidays"] = [k for k in stored["holidays"] if isinstance(k, str)]
        if "holiday_colors" in stored and isinstance(stored["holiday_colors"], dict):
            settings["holiday_colors"] = dict(stored["holiday_colors"])
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    return settings


def save_settings(settings: dict) -> None:
    """Persist settings to disk."""
    with open(_SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


# ------------------------------------------------------------------
# Windows autostart (registry-based)
# ------------------------------------------------------------------
_RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
_APP_NAME = "MiniCalendar"


def get_autostart() -> bool:
    """Return True if the autostart registry entry exists."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, _APP_NAME)
            return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


def set_autostart(enable: bool) -> None:
    """Create or remove the autostart registry entry."""
    if enable:
        if getattr(sys, "frozen", False):
            # Running as PyInstaller .exe
            command = f'"{sys.executable}"'
        else:
            main_script = os.path.abspath(os.path.join(os.path.dirname(__file__), "main.py"))
            pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            if not os.path.exists(pythonw):
                pythonw = sys.executable
            command = f'"{pythonw}" "{main_script}"'
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, _APP_NAME, 0, winreg.REG_SZ, command)
    else:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _RUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
                winreg.DeleteValue(key, _APP_NAME)
        except FileNotFoundError:
            pass
