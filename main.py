"""Entry point — glues pystray (daemon thread) with tkinter (main thread)."""

import ctypes
import threading

from calendar_window import CalendarWindow
from icon_gen import create_icon_image
from tray_icon import create_tray


def main() -> None:
    # DPI awareness so positions / fonts are crisp on Hi-DPI monitors
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    cal_win = CalendarWindow()

    # Callbacks marshalled onto the tkinter main thread
    def on_show() -> None:
        cal_win.root.after(0, cal_win.toggle)

    def on_exit() -> None:
        def _quit() -> None:
            tray.stop()
            cal_win.root.destroy()
        cal_win.root.after(0, _quit)

    def on_settings() -> None:
        cal_win.root.after(0, cal_win.open_settings)

    def on_about() -> None:
        cal_win.root.after(0, cal_win.open_about)

    icon_image = create_icon_image()
    tray = create_tray(icon_image, on_show, on_exit,
                       on_settings=on_settings, on_about=on_about)

    # Run pystray in a daemon thread so it doesn't block tkinter
    tray_thread = threading.Thread(target=tray.run, daemon=True)
    tray_thread.start()

    # tkinter main loop on the main thread
    cal_win.root.mainloop()


if __name__ == "__main__":
    main()
