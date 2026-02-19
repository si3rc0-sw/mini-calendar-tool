"""System-tray icon setup via pystray."""

from typing import Callable

import pystray
from PIL import Image
from pystray import MenuItem, Menu


def create_tray(
    icon_image: Image.Image,
    on_show: Callable[[], None],
    on_exit: Callable[[], None],
    on_settings: Callable[[], None] | None = None,
) -> pystray.Icon:
    """Build and return a pystray Icon (not yet started)."""
    items: list[MenuItem | Menu] = [
        MenuItem("Show Calendar", lambda _icon, _item: on_show(), default=True),
    ]
    if on_settings is not None:
        items.append(MenuItem("Settings", lambda _icon, _item: on_settings()))
    items.append(Menu.SEPARATOR)
    items.append(MenuItem("Exit", lambda _icon, _item: on_exit()))
    menu = Menu(*items)
    from datetime import date
    cw = date.today().isocalendar()[1]
    icon = pystray.Icon("mini-calendar", icon_image, f"Mini Calendar – CW {cw}", menu)
    return icon
