"""Multi-month calendar window (tkinter) positioned above the taskbar."""

import ctypes
import ctypes.wintypes
import calendar as _cal
from datetime import date
from tkinter import font as tkfont
import tkinter as tk

from calendar_logic import (
    DAY_ABBR,
    day_of_year,
    iso_week_numbers,
    month_grid,
    next_month,
    prev_month,
)
from tkinter import colorchooser

from holidays import COUNTRIES, holidays_by_country, holidays_for_year
from settings import load_settings, save_settings, get_autostart, set_autostart

VERSION = "1.5.0"

# Theme colour dictionaries
THEMES = {
    "light": {
        "accent":      "#0078D4",
        "sel_bg":      "#B3D7F2",
        "header_bg":   "#F3F3F3",
        "grid_bg":     "white",
        "wn_fg":       "#888888",
        "header_fg":   "#333333",
        "weekend_fg":  "#CC0000",
        "day_fg":      "black",
        "footer_fg":   "#555555",
        "tooltip_bg":  "#FFFFE0",
        "tooltip_fg":  "black",
        "today_fg":    "white",
        "sel_fg":      "black",
        "border":      "#E0E0E0",
        "btn_fg":      "#333333",
        "close_fg":    "#666666",
        "close_hover": "#E81123",
    },
    "dark": {
        "accent":      "#4DA6FF",
        "sel_bg":      "#2A4A6B",
        "header_bg":   "#2D2D2D",
        "grid_bg":     "#1E1E1E",
        "wn_fg":       "#777777",
        "header_fg":   "#E0E0E0",
        "weekend_fg":  "#FF6B6B",
        "day_fg":      "#E0E0E0",
        "footer_fg":   "#AAAAAA",
        "tooltip_bg":  "#3C3C3C",
        "tooltip_fg":  "#E0E0E0",
        "today_fg":    "white",
        "sel_fg":      "white",
        "border":      "#444444",
        "btn_fg":      "#E0E0E0",
        "close_fg":    "#AAAAAA",
        "close_hover": "#E81123",
    },
}


class _ToolTip:
    """Lightweight shared tooltip for holiday labels."""

    __slots__ = ("_root", "_tw", "_bg", "_fg")

    def __init__(self, root: tk.Tk, bg: str, fg: str) -> None:
        self._root = root
        self._tw: tk.Toplevel | None = None
        self._bg = bg
        self._fg = fg

    def update_colors(self, bg: str, fg: str) -> None:
        self._bg = bg
        self._fg = fg

    def show(self, widget: tk.Widget, text: str) -> None:
        self.hide()
        tw = tk.Toplevel(self._root)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)
        lbl = tk.Label(
            tw, text=text, bg=self._bg, fg=self._fg,
            relief="solid", borderwidth=1, padx=6, pady=3, justify="left",
        )
        lbl.pack()
        x = widget.winfo_rootx() + widget.winfo_width() // 2
        y = widget.winfo_rooty() + widget.winfo_height() + 2
        tw.wm_geometry(f"+{x}+{y}")
        self._tw = tw

    def hide(self) -> None:
        if self._tw:
            self._tw.destroy()
            self._tw = None


class _MonthPanel:
    """Pre-allocated widget pool for a single month (header + 6 weeks max)."""

    __slots__ = ("frame", "header", "wk_header", "day_headers",
                 "week_nums", "day_cells")

    def __init__(self, parent: tk.Frame, fonts: dict, theme: dict,
                 on_press, on_motion, on_release,
                 on_enter=None, on_leave=None,
                 on_right_click=None) -> None:
        self.frame = tk.Frame(parent, bg=theme["grid_bg"])

        self.header = tk.Label(
            self.frame, font=fonts["header"],
            bg=theme["header_bg"], fg=theme["header_fg"],
        )
        self.header.grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 2))

        self.wk_header = tk.Label(
            self.frame, text="CW", font=fonts["bold"],
            bg=theme["grid_bg"], fg=theme["wn_fg"], width=3,
        )
        self.wk_header.grid(row=1, column=0)

        self.day_headers: list[tk.Label] = []
        for col, abbr in enumerate(DAY_ABBR):
            fg = theme["weekend_fg"] if col >= 5 else theme["header_fg"]
            lbl = tk.Label(
                self.frame, text=abbr, font=fonts["bold"],
                bg=theme["grid_bg"], fg=fg, width=3,
            )
            lbl.grid(row=1, column=col + 1)
            self.day_headers.append(lbl)

        # Measure cell size to match a Label width=3
        _cell_w = fonts["cell_w"]
        _cell_h = fonts["cell_h"]

        self.week_nums: list[tk.Label] = []
        self.day_cells: list[list[tk.Canvas]] = []
        for r in range(6):  # max 6 weeks
            grid_row = r + 2
            wn = tk.Label(
                self.frame, font=fonts["wn"],
                bg=theme["grid_bg"], fg=theme["wn_fg"], width=3,
            )
            wn.grid(row=grid_row, column=0)
            self.week_nums.append(wn)

            row_cells: list[tk.Canvas] = []
            for c in range(7):
                cell = tk.Canvas(
                    self.frame, width=_cell_w, height=_cell_h,
                    bg=theme["grid_bg"], highlightthickness=0, borderwidth=0,
                )
                cell.grid(row=grid_row, column=c + 1)
                # Bind drag events once — handler checks _widget_dates
                cell.bind("<ButtonPress-1>", on_press)
                cell.bind("<B1-Motion>", on_motion)
                cell.bind("<ButtonRelease-1>", on_release)
                if on_enter:
                    cell.bind("<Enter>", on_enter)
                if on_leave:
                    cell.bind("<Leave>", on_leave)
                if on_right_click:
                    cell.bind("<Button-3>", on_right_click)
                row_cells.append(cell)
            self.day_cells.append(row_cells)


class CalendarWindow:
    """Multi-month calendar that appears above the taskbar."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(self._title())
        self.root.resizable(True, True)

        self.root.attributes("-toolwindow", True)
        self.root.attributes("-topmost", True)

        self._setup_fonts()

        settings = load_settings()
        self._theme = THEMES["dark" if settings.get("dark_mode") else "light"]
        self.root.configure(bg=self.tc("grid_bg"))
        self._set_titlebar_dark(self._theme is THEMES["dark"])

        self._saved_width: int | None = settings["window_width"]
        self._saved_height: int | None = settings["window_height"]
        self._saved_grid_cols: int | None = settings["grid_cols"]
        self._saved_grid_rows: int | None = settings["grid_rows"]

        today = date.today()
        self.center_year = today.year
        self.center_month = today.month

        self._showing = False  # suppress resize handler during show()

        # Selection state
        self.sel_start: date | None = None
        self.sel_end: date | None = None
        self._dragging = False

        # Widget-to-date mapping (filled during _build)
        self._widget_dates: dict[int, date] = {}
        # Date-to-widget mapping for highlight updates
        self._date_widgets: dict[date, tk.Label] = {}
        self._footer_label: tk.Label | None = None

        # Holiday state
        self._enabled_holidays: set[str] = set(settings.get("holidays", []))
        self._holiday_colors: dict[str, str] = dict(settings.get(
            "holiday_colors", {"CH": "#FF0000", "DE": "#FFD700", "CN": "#4CAF50"}))
        self._holiday_map: dict[date, list[tuple[str, str]]] = {}

        # Day markers (ephemeral, not persisted)
        self._day_markers: dict[date, str] = {}
        self._marker_colors: list[str] = list(settings.get(
            "marker_colors", ["#E74C3C", "#9B59B6", "#1ABC9C", "#F39C12"]))
        self._holidays_visible: bool = True

        # Auto-fit state (grid: cols x rows)
        self._month_width: int = 0
        self._month_height: int = 0
        self._grid_cols: int = 1
        self._grid_rows: int = 1
        self._current_total_months: int = 1
        self._resize_after_id: str | None = None

        # Font dict for _MonthPanel — includes cell pixel dims
        _tmp = tk.Label(self.root, text="00", font=self.font_normal, width=3)
        _tmp.update_idletasks()
        _cw = _tmp.winfo_reqwidth()
        _ch = _tmp.winfo_reqheight()
        _tmp.destroy()
        self._panel_fonts = {
            "header": self.font_header, "bold": self.font_bold,
            "normal": self.font_normal, "wn": self.font_wn,
            "cell_w": _cw, "cell_h": _ch,
        }

        # Panel pool + persistent shell
        self._panels: list[_MonthPanel] = []
        self._months_frame: tk.Frame | None = None
        self._nav_buttons: list[tk.Label] = []
        self._close_btn: tk.Label | None = None
        self._build_shell()
        self._tooltip = _ToolTip(
            self.root, self.tc("tooltip_bg"), self.tc("tooltip_fg"))
        self._rebuild_months()

        self.root.bind("<Escape>", self._on_escape)
        self.root.bind("<Configure>", self._on_configure)
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.withdraw()

    # ------------------------------------------------------------------
    # Theme helper
    # ------------------------------------------------------------------
    def tc(self, key: str) -> str:
        """Look up a colour from the active theme."""
        return self._theme[key]

    def _set_titlebar_dark(self, dark: bool) -> None:
        """Use Windows DWM API to switch the title bar between dark/light."""
        hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
        value = ctypes.c_int(1 if dark else 0)
        # DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 10 build 18985+ / Win11)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, 20, ctypes.byref(value), ctypes.sizeof(value))

    # ------------------------------------------------------------------
    # Fonts
    # ------------------------------------------------------------------
    def _setup_fonts(self) -> None:
        families = tkfont.families(self.root)
        base = "Segoe UI" if "Segoe UI" in families else "TkDefaultFont"
        self.font_normal = tkfont.Font(family=base, size=9)
        self.font_bold = tkfont.Font(family=base, size=9, weight="bold")
        self.font_header = tkfont.Font(family=base, size=10, weight="bold")
        self.font_nav = tkfont.Font(family=base, size=12, weight="bold")
        self.font_wn = tkfont.Font(family=base, size=8)
        self.font_footer = tkfont.Font(family=base, size=9)

    @staticmethod
    def _title() -> str:
        return f"Mini Calendar  Day: {day_of_year(date.today())}"

    # ------------------------------------------------------------------
    # Build shell (once) — nav bar + months placeholder + footer
    # ------------------------------------------------------------------
    def _build_shell(self) -> None:
        self._outer = tk.Frame(self.root, bg=self.tc("grid_bg"))
        self._outer.pack(padx=6, pady=4)

        # Navigation row: ◀◀◀  ◀◀  ◀  Today  ...  ▶  ▶▶  ▶▶▶  ×
        nav = tk.Frame(self._outer, bg=self.tc("grid_bg"))
        nav.pack(fill="x", pady=(0, 2))
        self._nav_frame = nav

        btn_prev_year = tk.Label(
            nav, text="\u25C0\u25C0\u25C0", font=self.font_nav,
            bg=self.tc("grid_bg"), fg=self.tc("btn_fg"), cursor="hand2",
        )
        btn_prev_year.pack(side="left", padx=6)
        btn_prev_year.bind("<Button-1>", lambda _e: self._navigate_year(-1))

        btn_prev_page = tk.Label(
            nav, text="\u25C0\u25C0", font=self.font_nav,
            bg=self.tc("grid_bg"), fg=self.tc("btn_fg"), cursor="hand2",
        )
        btn_prev_page.pack(side="left", padx=6)
        btn_prev_page.bind("<Button-1>", lambda _e: self._navigate_page(-1))

        btn_prev = tk.Label(
            nav, text="\u25C0", font=self.font_nav,
            bg=self.tc("grid_bg"), fg=self.tc("btn_fg"), cursor="hand2",
        )
        btn_prev.pack(side="left", padx=6)
        btn_prev.bind("<Button-1>", lambda _e: self._navigate(-1))

        btn_today = tk.Label(
            nav, text="Today", font=self.font_bold,
            bg=self.tc("grid_bg"), fg=self.tc("accent"), cursor="hand2",
        )
        btn_today.pack(side="left", padx=6)
        btn_today.bind("<Button-1>", lambda _e: self._go_today())

        # Custom close button (×) — packed right-side first so it's rightmost
        close_btn = tk.Label(
            nav, text="\u00D7", font=("Segoe UI", 14),
            fg=self.tc("close_fg"), bg=self.tc("grid_bg"), cursor="hand2",
        )
        close_btn.pack(side="right", padx=(0, 4))
        close_btn.bind("<Button-1>", lambda _e: self.hide())
        close_btn.bind("<Enter>",
                       lambda _e: close_btn.config(fg=self.tc("close_hover")))
        close_btn.bind("<Leave>",
                       lambda _e: close_btn.config(fg=self.tc("close_fg")))
        self._close_btn = close_btn

        btn_next_year = tk.Label(
            nav, text="\u25B6\u25B6\u25B6", font=self.font_nav,
            bg=self.tc("grid_bg"), fg=self.tc("btn_fg"), cursor="hand2",
        )
        btn_next_year.pack(side="right", padx=6)
        btn_next_year.bind("<Button-1>", lambda _e: self._navigate_year(1))

        btn_next_page = tk.Label(
            nav, text="\u25B6\u25B6", font=self.font_nav,
            bg=self.tc("grid_bg"), fg=self.tc("btn_fg"), cursor="hand2",
        )
        btn_next_page.pack(side="right", padx=6)
        btn_next_page.bind("<Button-1>", lambda _e: self._navigate_page(1))

        btn_next = tk.Label(
            nav, text="\u25B6", font=self.font_nav,
            bg=self.tc("grid_bg"), fg=self.tc("btn_fg"), cursor="hand2",
        )
        btn_next.pack(side="right", padx=6)
        btn_next.bind("<Button-1>", lambda _e: self._navigate(1))

        self._nav_buttons = [btn_prev_year, btn_prev_page, btn_prev, btn_today,
                             btn_next, btn_next_page, btn_next_year]

        # Months frame (content filled by _rebuild_months)
        self._months_frame = tk.Frame(self._outer, bg=self.tc("grid_bg"))
        self._months_frame.pack()

        # Footer row: info left, version right
        footer = tk.Frame(self._outer, bg=self.tc("grid_bg"))
        footer.pack(fill="x", pady=(4, 0))
        self._footer_frame = footer

        self._footer_label = tk.Label(
            footer, text=self._footer_text(), font=self.font_footer,
            bg=self.tc("grid_bg"), fg=self.tc("footer_fg"),
        )
        self._footer_label.pack(side="left")

        self._version_label = tk.Label(
            footer, text=f"v{VERSION}", font=self.font_wn,
            bg=self.tc("grid_bg"), fg=self.tc("wn_fg"),
        )
        self._version_label.pack(side="right")

        self._holiday_btn = tk.Label(
            footer, text="Holidays", font=self.font_wn,
            bg=self.tc("grid_bg"), fg=self.tc("accent"), cursor="hand2",
        )
        self._holiday_btn.pack(side="right", padx=(0, 8))
        self._holiday_btn.bind("<Button-1>", lambda _e: self._toggle_holidays())

        self._clear_btn = tk.Label(
            footer, text="Clear", font=self.font_wn,
            bg=self.tc("grid_bg"), fg=self.tc("footer_fg"), cursor="hand2",
        )
        self._clear_btn.pack(side="right", padx=(0, 8))
        self._clear_btn.bind("<Button-1>", lambda _e: self._clear_markers())

    # ------------------------------------------------------------------
    # Apply theme to shell widgets (called after theme switch)
    # ------------------------------------------------------------------
    def _apply_theme(self) -> None:
        """Recolour all persistent shell widgets and rebuild month panels."""
        self._set_titlebar_dark(self._theme is THEMES["dark"])
        grid_bg = self.tc("grid_bg")
        self.root.configure(bg=grid_bg)
        self._outer.configure(bg=grid_bg)
        self._nav_frame.configure(bg=grid_bg)
        for btn in self._nav_buttons:
            btn.configure(bg=grid_bg, fg=self.tc("btn_fg"))
        # "Today" button uses accent colour
        self._nav_buttons[3].configure(fg=self.tc("accent"))
        self._close_btn.configure(bg=grid_bg, fg=self.tc("close_fg"))
        self._months_frame.configure(bg=grid_bg)
        self._footer_frame.configure(bg=grid_bg)
        self._footer_label.configure(bg=grid_bg, fg=self.tc("footer_fg"))
        self._version_label.configure(bg=grid_bg, fg=self.tc("wn_fg"))
        self._holiday_btn.configure(bg=grid_bg,
            fg=self.tc("accent") if self._holidays_visible else self.tc("wn_fg"))
        self._clear_btn.configure(bg=grid_bg, fg=self.tc("footer_fg"))
        self._tooltip.update_colors(self.tc("tooltip_bg"), self.tc("tooltip_fg"))
        self._rebuild_months()

    # ------------------------------------------------------------------
    # Rebuild month grid using pooled panels (fast reconfigure)
    # ------------------------------------------------------------------
    def _rebuild_months(self) -> None:
        self._widget_dates.clear()
        self._date_widgets.clear()

        total = self._grid_cols * self._grid_rows
        self._current_total_months = total
        months_before = (total - 1) // 2

        # Grow panel pool if needed
        while len(self._panels) < total:
            self._panels.append(_MonthPanel(
                self._months_frame, self._panel_fonts, self._theme,
                self._on_press, self._on_motion, self._on_release,
                self._on_cell_enter, self._on_cell_leave,
                self._on_right_click,
            ))

        # Build month list
        month_list: list[tuple[int, int]] = []
        y, m = self.center_year, self.center_month
        for _ in range(months_before):
            y, m = prev_month(y, m)
        for _ in range(total):
            month_list.append((y, m))
            y, m = next_month(y, m)

        today = date.today()
        sel_lo, sel_hi = self._sel_range()

        # Compute holiday map for all visible years
        visible_years = set(y for y, _m in month_list)
        self._holiday_map = {}
        for y in visible_years:
            self._holiday_map.update(
                holidays_for_year(y, self._enabled_holidays))

        # Update active panels
        for i, (y, m) in enumerate(month_list):
            panel = self._panels[i]
            grid_r = i // self._grid_cols
            grid_c = i % self._grid_cols
            panel.frame.grid(row=grid_r, column=grid_c, padx=6, pady=2, sticky="n")
            self._fill_panel(panel, y, m, today, sel_lo, sel_hi)

        # Hide excess panels
        for i in range(total, len(self._panels)):
            self._panels[i].frame.grid_forget()

        if self._footer_label:
            self._footer_label.configure(text=self._footer_text())

        # Measure month dimensions once (they never change)
        if self._month_width == 0 and month_list:
            self.root.update_idletasks()
            f = self._panels[0].frame
            self._month_width = f.winfo_reqwidth() + 12   # +padx*2
            self._month_height = f.winfo_reqheight() + 4  # +pady*2

    def _fill_panel(self, panel: "_MonthPanel", year: int, month: int,
                    today: date, sel_lo: date | None, sel_hi: date | None) -> None:
        """Reconfigure an existing panel's labels — no widget creation."""
        grid_bg = self.tc("grid_bg")

        # Apply theme to structural widgets
        panel.frame.configure(bg=grid_bg)
        panel.header.configure(
            text=f"{_cal.month_name[month]} {year}",
            bg=self.tc("header_bg"), fg=self.tc("header_fg"),
        )
        panel.wk_header.configure(bg=grid_bg, fg=self.tc("wn_fg"))
        for col, lbl in enumerate(panel.day_headers):
            fg = self.tc("weekend_fg") if col >= 5 else self.tc("header_fg")
            lbl.configure(bg=grid_bg, fg=fg)

        grid = month_grid(year, month)
        weeks = iso_week_numbers(year, month)
        font_bold = self.font_bold
        font_normal = self.font_normal

        for r in range(6):
            panel.week_nums[r].configure(bg=grid_bg, fg=self.tc("wn_fg"))
            if r < len(grid):
                row_days = grid[r]
                panel.week_nums[r].configure(text=weeks[r])
                for c in range(7):
                    cell = panel.day_cells[r][c]
                    day = row_days[c]
                    if day is None:
                        cell.delete("all")
                        cell.configure(bg=grid_bg, cursor="")
                    else:
                        d = date(year, month, day)
                        is_today = d == today
                        in_sel = (sel_lo is not None and sel_lo <= d <= sel_hi)
                        h_colors = self._holiday_color_for_date(d)
                        m_color = self._day_markers.get(d)
                        bgs, fg = self._day_colors(
                            is_today, c >= 5, in_sel, h_colors)
                        self._draw_cell(
                            cell, str(day), bgs, fg,
                            font_bold if is_today else font_normal,
                            cursor="hand2",
                            marker_color=m_color,
                        )
                        self._widget_dates[id(cell)] = d
                        self._date_widgets[d] = cell
            else:
                panel.week_nums[r].configure(text="")
                for c in range(7):
                    cell = panel.day_cells[r][c]
                    cell.delete("all")
                    cell.configure(bg=grid_bg, cursor="")

    # ------------------------------------------------------------------
    # Day colour logic
    # ------------------------------------------------------------------
    def _day_colors(self, is_today: bool, is_weekend: bool, in_sel: bool,
                    holiday_colors: list[str] | None = None,
                    ) -> tuple[list[str], str]:
        if is_today:
            return [self.tc("accent")], self.tc("today_fg")
        if in_sel:
            return [self.tc("sel_bg")], self.tc("sel_fg")
        if holiday_colors:
            return holiday_colors, "white"
        if is_weekend:
            return [self.tc("grid_bg")], self.tc("weekend_fg")
        return [self.tc("grid_bg")], self.tc("day_fg")

    # ------------------------------------------------------------------
    # Selection helpers
    # ------------------------------------------------------------------
    def _sel_range(self) -> tuple[date | None, date | None]:
        if self.sel_start and self.sel_end:
            lo = min(self.sel_start, self.sel_end)
            hi = max(self.sel_start, self.sel_end)
            return lo, hi
        return None, None

    def _clear_selection(self) -> None:
        self.sel_start = None
        self.sel_end = None
        self._dragging = False

    # ------------------------------------------------------------------
    # Drag-to-select events
    # ------------------------------------------------------------------
    def _on_press(self, event: tk.Event) -> None:
        d = self._widget_dates.get(id(event.widget))
        if d:
            self.sel_start = d
            self.sel_end = d
            self._dragging = True
            self._update_highlight()

    def _on_motion(self, event: tk.Event) -> None:
        if not self._dragging:
            return
        w = event.widget.winfo_containing(event.x_root, event.y_root)
        if w and id(w) in self._widget_dates:
            new_end = self._widget_dates[id(w)]
            if new_end != self.sel_end:
                self.sel_end = new_end
                self._update_highlight()

    def _on_release(self, event: tk.Event) -> None:
        if not self._dragging:
            return
        self._dragging = False
        w = event.widget.winfo_containing(event.x_root, event.y_root)
        if w and id(w) in self._widget_dates:
            self.sel_end = self._widget_dates[id(w)]
        self._update_highlight()

    # ------------------------------------------------------------------
    # Update highlights without full rebuild
    # ------------------------------------------------------------------
    def _update_highlight(self) -> None:
        today = date.today()
        sel_lo, sel_hi = self._sel_range()

        for d, cell in self._date_widgets.items():
            is_today = d == today
            is_weekend = d.weekday() >= 5
            in_sel = sel_lo is not None and sel_lo <= d <= sel_hi
            h_colors = self._holiday_color_for_date(d)
            m_color = self._day_markers.get(d)
            bgs, fg = self._day_colors(is_today, is_weekend, in_sel, h_colors)
            self._draw_cell(
                cell, str(d.day), bgs, fg,
                self.font_bold if is_today else self.font_normal,
                cursor="hand2",
                marker_color=m_color,
            )

        if self._footer_label:
            self._footer_label.configure(text=self._footer_text())

    # ------------------------------------------------------------------
    # Holiday colour helper
    # ------------------------------------------------------------------
    def _holiday_color_for_date(self, d: date) -> list[str] | None:
        if not self._holidays_visible:
            return None
        entries = self._holiday_map.get(d)
        if not entries:
            return None
        seen: set[str] = set()
        colors: list[str] = []
        for _, country in entries:
            if country not in seen:
                seen.add(country)
                colors.append(self._holiday_colors.get(country, "#888888"))
        return colors

    # ------------------------------------------------------------------
    # Canvas cell drawing (supports multi-colour stripes)
    # ------------------------------------------------------------------
    def _draw_cell(self, cell: tk.Canvas, text: str, bg_colors: list[str],
                   fg: str, font, cursor: str = "",
                   marker_color: str | None = None) -> None:
        cell.delete("all")
        w = cell.winfo_width()
        h = cell.winfo_height()
        if w <= 1:
            w = int(cell["width"]) + 2
        if h <= 1:
            h = int(cell["height"]) + 2

        if len(bg_colors) <= 1:
            cell.configure(bg=bg_colors[0] if bg_colors else self.tc("grid_bg"))
        else:
            cell.configure(bg=bg_colors[0])
            n = len(bg_colors)
            stripe_h = h / n
            for i, c in enumerate(bg_colors):
                y1 = round(i * stripe_h)
                y2 = round((i + 1) * stripe_h)
                cell.create_rectangle(0, y1, w, y2, fill=c, outline="")

        # Marker circle (between background and text)
        if marker_color:
            r = min(w, h) // 2 - 1
            cx, cy = w // 2, h // 2
            cell.create_oval(cx - r, cy - r, cx + r, cy + r,
                             fill=marker_color, outline="")
            fg = "white"

        if text:
            cx, cy = w // 2, h // 2
            if len(bg_colors) > 1:
                # Outline for readability on multi-colour stripes
                for dx in (-1, 0, 1):
                    for dy in (-1, 0, 1):
                        if dx or dy:
                            cell.create_text(
                                cx + dx, cy + dy, text=text,
                                fill="black", font=font)
            cell.create_text(cx, cy, text=text, fill=fg, font=font)
        cell.configure(cursor=cursor)

    # ------------------------------------------------------------------
    # Tooltip on hover
    # ------------------------------------------------------------------
    def _on_cell_enter(self, event: tk.Event) -> None:
        d = self._widget_dates.get(id(event.widget))
        if d and self._holidays_visible and d in self._holiday_map:
            lines = [f"{name} ({country})" for name, country in self._holiday_map[d]]
            self._tooltip.show(event.widget, "\n".join(lines))

    def _on_cell_leave(self, _event: tk.Event) -> None:
        self._tooltip.hide()

    # ------------------------------------------------------------------
    # Right-click day marker
    # ------------------------------------------------------------------
    def _on_right_click(self, event: tk.Event) -> None:
        d = self._widget_dates.get(id(event.widget))
        if not d:
            return
        # If right-clicked inside a multi-day selection, target all selected days
        sel_lo, sel_hi = self._sel_range()
        if sel_lo and sel_hi and sel_lo != sel_hi and sel_lo <= d <= sel_hi:
            targets = sorted(dd for dd in self._date_widgets if sel_lo <= dd <= sel_hi)
        else:
            targets = [d]

        menu = tk.Menu(self.root, tearoff=0)
        for color in self._marker_colors:
            menu.add_command(
                label="    ",
                background=color,
                activebackground=color,
                command=lambda c=color: self._set_markers(targets, c))
        if any(t in self._day_markers for t in targets):
            menu.add_separator()
            menu.add_command(label="Remove",
                             command=lambda: self._remove_markers(targets))
        menu.tk_popup(event.x_root, event.y_root)

    def _set_markers(self, dates: list[date], color: str) -> None:
        for d in dates:
            self._day_markers[d] = color
        self._update_highlight()

    def _remove_markers(self, dates: list[date]) -> None:
        for d in dates:
            self._day_markers.pop(d, None)
        self._update_highlight()

    def _clear_markers(self) -> None:
        """Remove all day markers (does not affect holidays)."""
        self._day_markers.clear()
        self._update_highlight()

    def _toggle_holidays(self) -> None:
        """Show or hide holiday colours without changing enabled holidays."""
        self._holidays_visible = not self._holidays_visible
        fg = self.tc("accent") if self._holidays_visible else self.tc("wn_fg")
        self._holiday_btn.configure(fg=fg)
        self._update_highlight()

    # ------------------------------------------------------------------
    # Footer text
    # ------------------------------------------------------------------
    def _footer_text(self) -> str:
        today_str = f"Today: {date.today().strftime('%d.%m.%Y')}"
        sel_lo, sel_hi = self._sel_range()
        if sel_lo is None or sel_lo == sel_hi:
            return today_str

        total_days = (sel_hi - sel_lo).days + 1
        full_weeks, rem_days = divmod(total_days, 7)

        parts: list[str] = []
        if full_weeks:
            parts.append(f"{full_weeks} week{'s' if full_weeks != 1 else ''}")
        if rem_days:
            parts.append(f"{rem_days} day{'s' if rem_days != 1 else ''}")

        range_str = f"{sel_lo.strftime('%d.%m')} \u2192 {sel_hi.strftime('%d.%m')}"
        return f"{range_str}:  {total_days} days  ({', '.join(parts)})     {today_str}"

    # ------------------------------------------------------------------
    # ESC clears selection first, then hides
    # ------------------------------------------------------------------
    def _on_escape(self, _event: tk.Event) -> None:
        if self.sel_start is not None:
            self._clear_selection()
            self._update_highlight()
        else:
            self.hide()

    # ------------------------------------------------------------------
    # About dialog
    # ------------------------------------------------------------------
    def open_about(self) -> None:
        import webbrowser
        dlg = tk.Toplevel(self.root)
        dlg.title("About")
        dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        dlg.grab_set()

        frame = tk.Frame(dlg, padx=20, pady=16)
        frame.pack()

        tk.Label(frame, text="Mini Calendar", font=self.font_header).pack()
        tk.Label(frame, text=f"Version {VERSION}", font=self.font_normal,
                 fg="#888888").pack(pady=(2, 8))

        tk.Label(frame, text="A lightweight Windows system-tray calendar\n"
                 "with dark mode, holidays, and ISO week numbers.",
                 font=self.font_normal, justify="center").pack(pady=(0, 8))

        tk.Label(frame, text="MIT License + Commons Clause", font=self.font_wn,
                 fg="#888888").pack()

        link = tk.Label(frame, text="github.com/si3rc0-sw/mini-calendar-tool",
                        font=self.font_normal, fg=self.tc("accent"),
                        cursor="hand2")
        link.pack(pady=(4, 8))
        link.bind("<Button-1>", lambda _e: webbrowser.open(
            "https://github.com/si3rc0-sw/mini-calendar-tool"))

        tk.Button(frame, text="OK", width=8, command=dlg.destroy).pack()

    # ------------------------------------------------------------------
    # Settings dialog
    # ------------------------------------------------------------------
    def open_settings(self) -> None:
        dlg = tk.Toplevel(self.root)
        dlg.title("Settings")
        dlg.resizable(False, False)
        dlg.attributes("-topmost", True)
        dlg.grab_set()

        frame = tk.Frame(dlg, padx=12, pady=8)
        frame.pack()

        dark_var = tk.BooleanVar(value=self._theme is THEMES["dark"])
        tk.Checkbutton(
            frame, text="Dark mode", variable=dark_var,
            font=self.font_normal,
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=4)

        autostart_var = tk.BooleanVar(value=get_autostart())
        tk.Checkbutton(
            frame, text="Start with Windows", variable=autostart_var,
            font=self.font_normal,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=4)

        # --- Holiday section ---
        holiday_frame = tk.LabelFrame(
            frame, text="Holidays", font=self.font_bold, padx=8, pady=4,
        )
        holiday_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=(8, 0))

        check_vars: dict[str, tk.BooleanVar] = {}
        color_labels: dict[str, tk.Label] = {}
        color_vals: dict[str, str] = dict(self._holiday_colors)

        for col_idx, (code, country_name) in enumerate(COUNTRIES):
            col_frame = tk.Frame(holiday_frame)
            col_frame.grid(row=0, column=col_idx, padx=8, pady=2, sticky="n")

            # Country header with colour swatch
            hdr = tk.Frame(col_frame)
            hdr.pack(fill="x", pady=(0, 4))

            tk.Label(hdr, text=country_name, font=self.font_bold).pack(side="left")

            swatch = tk.Label(
                hdr, text="  ", bg=color_vals.get(code, "#888888"),
                relief="raised", borderwidth=1, cursor="hand2",
            )
            swatch.pack(side="right", padx=(4, 0))
            color_labels[code] = swatch

            def _make_picker(c=code, sw=swatch):
                def _pick(_e=None):
                    result = colorchooser.askcolor(
                        color=color_vals[c], parent=dlg, title=f"Colour for {c}")
                    if result[1]:
                        color_vals[c] = result[1]
                        sw.configure(bg=result[1])
                return _pick

            swatch.bind("<Button-1>", _make_picker())

            # Checkbuttons for each holiday
            for key, name in holidays_by_country(code):
                var = tk.BooleanVar(value=(key in self._enabled_holidays))
                check_vars[key] = var
                tk.Checkbutton(
                    col_frame, text=name, variable=var,
                    font=self.font_normal, anchor="w",
                ).pack(fill="x")

        # --- Marker Colors section ---
        marker_frame = tk.LabelFrame(
            frame, text="Marker Colors", font=self.font_bold, padx=8, pady=4,
        )
        marker_frame.grid(row=3, column=0, columnspan=2, sticky="we", pady=(8, 0))

        marker_vals = list(self._marker_colors)

        for i, color in enumerate(marker_vals):
            sw = tk.Label(
                marker_frame, text="    ", bg=color,
                relief="raised", borderwidth=1, cursor="hand2",
            )
            sw.grid(row=0, column=i, padx=4, pady=4)

            def _make_marker_picker(idx=i, swatch=sw):
                def _pick(_e=None):
                    result = colorchooser.askcolor(
                        color=marker_vals[idx], parent=dlg,
                        title=f"Marker colour {idx + 1}")
                    if result[1]:
                        marker_vals[idx] = result[1]
                        swatch.configure(bg=result[1])
                return _pick

            sw.bind("<Button-1>", _make_marker_picker())

        tk.Label(
            marker_frame, text="Right-click a day to mark it",
            font=self.font_wn, fg="#888888",
        ).grid(row=0, column=len(marker_vals), padx=(8, 0))

        # --- Buttons ---
        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=4, column=0, columnspan=2, pady=(8, 0))

        def on_ok() -> None:
            new_enabled = [k for k, v in check_vars.items() if v.get()]
            new_colors = {code: color_vals[code] for code, _ in COUNTRIES}
            new_dark = dark_var.get()

            settings = load_settings()
            settings["dark_mode"] = new_dark
            settings["holidays"] = new_enabled
            settings["holiday_colors"] = new_colors
            settings["marker_colors"] = marker_vals
            save_settings(settings)

            self._theme = THEMES["dark" if new_dark else "light"]
            self._enabled_holidays = set(new_enabled)
            self._holiday_colors = new_colors
            self._marker_colors = list(marker_vals)

            set_autostart(autostart_var.get())
            dlg.destroy()
            self._apply_theme()

        tk.Button(btn_frame, text="OK", width=8, command=on_ok).pack(
            side="left", padx=4,
        )
        tk.Button(btn_frame, text="Cancel", width=8, command=dlg.destroy).pack(
            side="left", padx=4,
        )

    # ------------------------------------------------------------------
    # Navigation
    # ------------------------------------------------------------------
    def _navigate(self, direction: int) -> None:
        if direction < 0:
            self.center_year, self.center_month = prev_month(
                self.center_year, self.center_month
            )
        else:
            self.center_year, self.center_month = next_month(
                self.center_year, self.center_month
            )
        self._rebuild_months()

    def _navigate_page(self, direction: int) -> None:
        """Skip forward/back by the number of currently displayed months."""
        steps = self._current_total_months
        y, m = self.center_year, self.center_month
        step_fn = next_month if direction > 0 else prev_month
        for _ in range(steps):
            y, m = step_fn(y, m)
        self.center_year, self.center_month = y, m
        self._rebuild_months()

    def _navigate_year(self, direction: int) -> None:
        self.center_year += direction
        self._rebuild_months()

    def _go_today(self) -> None:
        today = date.today()
        self.center_year = today.year
        self.center_month = today.month
        self._clear_selection()
        self._rebuild_months()

    # ------------------------------------------------------------------
    # Resize handling — auto-fit month count to window width
    # ------------------------------------------------------------------
    def _on_configure(self, event: tk.Event) -> None:
        if self._showing:
            return
        if event.widget is not self.root:
            return
        if self._month_width <= 0 or self._month_height <= 0:
            return
        # Track size (persisted on hide)
        self._saved_width = self.root.winfo_width()
        self._saved_height = self.root.winfo_height()
        # Debounce rebuild (30ms) to batch rapid configure events
        if self._resize_after_id is not None:
            self.root.after_cancel(self._resize_after_id)
        self._resize_after_id = self.root.after(30, self._handle_resize)

    def _handle_resize(self) -> None:
        self._resize_after_id = None
        cols = max(1, (self._saved_width - 24) // self._month_width)
        rows = max(1, (self._saved_height - 50) // self._month_height)
        if cols != self._grid_cols or rows != self._grid_rows:
            self._grid_cols = cols
            self._grid_rows = rows
            self._saved_grid_cols = cols
            self._saved_grid_rows = rows
            self._rebuild_months()
            self._persist_size()

    # ------------------------------------------------------------------
    # Persist window size
    # ------------------------------------------------------------------
    def _persist_size(self) -> None:
        settings = load_settings()
        settings["window_width"] = self._saved_width
        settings["window_height"] = self._saved_height
        settings["grid_cols"] = self._grid_cols
        settings["grid_rows"] = self._grid_rows
        save_settings(settings)

    # ------------------------------------------------------------------
    # Show / Hide / Toggle
    # ------------------------------------------------------------------
    def toggle(self) -> None:
        if self.root.state() == "withdrawn" or not self.root.winfo_viewable():
            self.show()
        else:
            self.hide()

    def show(self) -> None:
        today = date.today()
        self.center_year = today.year
        self.center_month = today.month
        self._clear_selection()
        self.root.title(self._title())

        # Restore grid layout from saved values, or default 3×1
        if self._saved_grid_cols is not None and self._saved_grid_rows is not None:
            self._grid_cols = self._saved_grid_cols
            self._grid_rows = self._saved_grid_rows
        else:
            self._grid_cols = 3
            self._grid_rows = 1

        self._showing = True  # suppress resize handler during layout
        self._rebuild_months()
        self.root.geometry("")  # clear old geometry so window fits content
        self.root.deiconify()
        self.root.update_idletasks()
        self._set_titlebar_dark(self._theme is THEMES["dark"])
        self._position_window()
        self.root.update_idletasks()
        # Delay re-enabling resize: WM configure events arrive asynchronously
        self.root.after(200, self._end_showing)

        self.root.lift()
        self.root.focus_force()

    def _end_showing(self) -> None:
        self._showing = False

    def hide(self) -> None:
        if self._saved_width is not None and self._saved_height is not None:
            self._persist_size()
        self.root.withdraw()

    # ------------------------------------------------------------------
    # Position bottom-right above taskbar
    # ------------------------------------------------------------------
    def _position_window(self, override_size: tuple[int, int] | None = None) -> None:
        self.root.update_idletasks()

        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
        work_right = rect.right
        work_bottom = rect.bottom

        if override_size:
            win_w, win_h = override_size
        else:
            win_w = self.root.winfo_reqwidth()
            win_h = self.root.winfo_reqheight()

        x = work_right - win_w - 12
        y = work_bottom - win_h - 12
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
