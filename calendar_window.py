"""Multi-month calendar window (tkinter) positioned above the taskbar."""

import ctypes
import ctypes.wintypes
import calendar as _cal
from datetime import date, timedelta
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
from settings import load_settings, save_settings, get_autostart, set_autostart

# Colours
ACCENT = "#0078D4"
SEL_BG = "#B3D7F2"
HEADER_BG = "#F3F3F3"
GRID_BG = "white"
WN_FG = "#888888"


class _MonthPanel:
    """Pre-allocated widget pool for a single month (header + 6 weeks max)."""

    __slots__ = ("frame", "header", "wk_header", "day_headers",
                 "week_nums", "day_cells")

    def __init__(self, parent: tk.Frame, fonts: dict,
                 on_press, on_motion, on_release) -> None:
        self.frame = tk.Frame(parent, bg=GRID_BG)

        self.header = tk.Label(
            self.frame, font=fonts["header"], bg=HEADER_BG, fg="#333333",
        )
        self.header.grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 2))

        self.wk_header = tk.Label(
            self.frame, text="Wk", font=fonts["bold"], bg=GRID_BG, fg=WN_FG, width=3,
        )
        self.wk_header.grid(row=1, column=0)

        self.day_headers: list[tk.Label] = []
        for col, abbr in enumerate(DAY_ABBR):
            fg = "#CC0000" if col >= 5 else "#333333"
            lbl = tk.Label(
                self.frame, text=abbr, font=fonts["bold"], bg=GRID_BG, fg=fg, width=3,
            )
            lbl.grid(row=1, column=col + 1)
            self.day_headers.append(lbl)

        self.week_nums: list[tk.Label] = []
        self.day_cells: list[list[tk.Label]] = []
        for r in range(6):  # max 6 weeks
            grid_row = r + 2
            wn = tk.Label(
                self.frame, font=fonts["wn"], bg=GRID_BG, fg=WN_FG, width=3,
            )
            wn.grid(row=grid_row, column=0)
            self.week_nums.append(wn)

            row_cells: list[tk.Label] = []
            for c in range(7):
                cell = tk.Label(self.frame, text="", bg=GRID_BG, width=3)
                cell.grid(row=grid_row, column=c + 1)
                # Bind drag events once — handler checks _widget_dates
                cell.bind("<ButtonPress-1>", on_press)
                cell.bind("<B1-Motion>", on_motion)
                cell.bind("<ButtonRelease-1>", on_release)
                row_cells.append(cell)
            self.day_cells.append(row_cells)


class CalendarWindow:
    """Multi-month calendar that appears above the taskbar."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(self._title())
        self.root.resizable(True, True)
        self.root.configure(bg=GRID_BG)

        self.root.attributes("-toolwindow", True)
        self.root.attributes("-topmost", True)

        self._setup_fonts()

        settings = load_settings()
        self.months_before: int = settings["months_before"]
        self.months_after: int = settings["months_after"]
        self._saved_width: int | None = settings["window_width"]
        self._saved_height: int | None = settings["window_height"]

        today = date.today()
        self.center_year = today.year
        self.center_month = today.month

        # Selection state
        self.sel_start: date | None = None
        self.sel_end: date | None = None
        self._dragging = False

        # Widget-to-date mapping (filled during _build)
        self._widget_dates: dict[int, date] = {}
        # Date-to-widget mapping for highlight updates
        self._date_widgets: dict[date, tk.Label] = {}
        self._footer_label: tk.Label | None = None

        # Auto-fit state (grid: cols x rows)
        self._month_width: int = 0
        self._month_height: int = 0
        self._grid_cols: int = self.months_before + 1 + self.months_after
        self._grid_rows: int = 1
        self._current_total_months: int = self._grid_cols * self._grid_rows
        self._resize_after_id: str | None = None

        # Font dict for _MonthPanel
        self._panel_fonts = {
            "header": self.font_header, "bold": self.font_bold,
            "normal": self.font_normal, "wn": self.font_wn,
        }

        # Panel pool + persistent shell
        self._panels: list[_MonthPanel] = []
        self._months_frame: tk.Frame | None = None
        self._build_shell()
        self._rebuild_months()

        self.root.bind("<Escape>", self._on_escape)
        self.root.bind("<Configure>", self._on_configure)
        self.root.protocol("WM_DELETE_WINDOW", self.hide)
        self.root.withdraw()

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
        self._outer = tk.Frame(self.root, bg=GRID_BG)
        self._outer.pack(padx=6, pady=4)

        # Navigation row: ◀◀  ◀  Today  ▶  ▶▶
        nav = tk.Frame(self._outer, bg=GRID_BG)
        nav.pack(fill="x", pady=(0, 2))

        btn_prev_year = tk.Label(
            nav, text="\u25C0\u25C0", font=self.font_nav, bg=GRID_BG, cursor="hand2"
        )
        btn_prev_year.pack(side="left", padx=6)
        btn_prev_year.bind("<Button-1>", lambda _e: self._navigate_year(-1))

        btn_prev = tk.Label(
            nav, text="\u25C0", font=self.font_nav, bg=GRID_BG, cursor="hand2"
        )
        btn_prev.pack(side="left", padx=6)
        btn_prev.bind("<Button-1>", lambda _e: self._navigate(-1))

        btn_today = tk.Label(
            nav, text="Today", font=self.font_bold, bg=GRID_BG, fg=ACCENT,
            cursor="hand2",
        )
        btn_today.pack(side="left", padx=6)
        btn_today.bind("<Button-1>", lambda _e: self._go_today())

        btn_next_year = tk.Label(
            nav, text="\u25B6\u25B6", font=self.font_nav, bg=GRID_BG, cursor="hand2"
        )
        btn_next_year.pack(side="right", padx=6)
        btn_next_year.bind("<Button-1>", lambda _e: self._navigate_year(1))

        btn_next = tk.Label(
            nav, text="\u25B6", font=self.font_nav, bg=GRID_BG, cursor="hand2"
        )
        btn_next.pack(side="right", padx=6)
        btn_next.bind("<Button-1>", lambda _e: self._navigate(1))

        # Months frame (content filled by _rebuild_months)
        self._months_frame = tk.Frame(self._outer, bg=GRID_BG)
        self._months_frame.pack()

        # Footer
        self._footer_label = tk.Label(
            self._outer, text=self._footer_text(), font=self.font_footer,
            bg=GRID_BG, fg="#555555",
        )
        self._footer_label.pack(pady=(4, 0))

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
                self._months_frame, self._panel_fonts,
                self._on_press, self._on_motion, self._on_release,
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
        panel.header.configure(text=f"{_cal.month_name[month]} {year}")

        grid = month_grid(year, month)
        weeks = iso_week_numbers(year, month)
        font_bold = self.font_bold
        font_normal = self.font_normal
        _day_colors = self._day_colors

        for r in range(6):
            if r < len(grid):
                row_days = grid[r]
                panel.week_nums[r].configure(text=weeks[r])
                for c in range(7):
                    cell = panel.day_cells[r][c]
                    day = row_days[c]
                    if day is None:
                        cell.configure(text="", bg=GRID_BG, fg=GRID_BG,
                                       font=font_normal, cursor="")
                    else:
                        d = date(year, month, day)
                        is_today = d == today
                        in_sel = (sel_lo is not None and sel_lo <= d <= sel_hi)
                        bg, fg = _day_colors(is_today, c >= 5, in_sel)
                        cell.configure(
                            text=str(day), bg=bg, fg=fg,
                            font=font_bold if is_today else font_normal,
                            cursor="hand2",
                        )
                        self._widget_dates[id(cell)] = d
                        self._date_widgets[d] = cell
            else:
                panel.week_nums[r].configure(text="")
                for c in range(7):
                    panel.day_cells[r][c].configure(
                        text="", bg=GRID_BG, fg=GRID_BG,
                        font=font_normal, cursor="")

    # ------------------------------------------------------------------
    # Day colour logic
    # ------------------------------------------------------------------
    @staticmethod
    def _day_colors(is_today: bool, is_weekend: bool, in_sel: bool
                    ) -> tuple[str, str]:
        if is_today and in_sel:
            return ACCENT, "white"
        if is_today:
            return ACCENT, "white"
        if in_sel:
            return SEL_BG, "black"
        if is_weekend:
            return GRID_BG, "#CC0000"
        return GRID_BG, "black"

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

        for d, lbl in self._date_widgets.items():
            is_today = d == today
            is_weekend = d.weekday() >= 5
            in_sel = sel_lo is not None and sel_lo <= d <= sel_hi
            bg, fg = self._day_colors(is_today, is_weekend, in_sel)
            lbl.configure(bg=bg, fg=fg)

        if self._footer_label:
            self._footer_label.configure(text=self._footer_text())

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

        tk.Label(frame, text="Months before:", font=self.font_normal).grid(
            row=0, column=0, sticky="w", pady=4,
        )
        spin_before = tk.Spinbox(
            frame, from_=0, to=6, width=4, font=self.font_normal,
        )
        spin_before.delete(0, "end")
        spin_before.insert(0, str(self.months_before))
        spin_before.grid(row=0, column=1, padx=(8, 0), pady=4)

        tk.Label(frame, text="Months after:", font=self.font_normal).grid(
            row=1, column=0, sticky="w", pady=4,
        )
        spin_after = tk.Spinbox(
            frame, from_=0, to=6, width=4, font=self.font_normal,
        )
        spin_after.delete(0, "end")
        spin_after.insert(0, str(self.months_after))
        spin_after.grid(row=1, column=1, padx=(8, 0), pady=4)

        autostart_var = tk.BooleanVar(value=get_autostart())
        tk.Checkbutton(
            frame, text="Start with Windows", variable=autostart_var,
            font=self.font_normal,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=4)

        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=(8, 0))

        def on_ok() -> None:
            try:
                self.months_before = max(0, min(6, int(spin_before.get())))
                self.months_after = max(0, min(6, int(spin_after.get())))
            except ValueError:
                return
            settings = load_settings()
            settings["months_before"] = self.months_before
            settings["months_after"] = self.months_after
            save_settings(settings)
            set_autostart(autostart_var.get())
            self._grid_cols = self.months_before + 1 + self.months_after
            self._grid_rows = 1
            self._current_total_months = self._grid_cols
            self._saved_width = None  # reset saved size so defaults apply
            self._saved_height = None
            dlg.destroy()
            self._rebuild_months()

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
            self._rebuild_months()

    # ------------------------------------------------------------------
    # Persist window size
    # ------------------------------------------------------------------
    def _persist_size(self) -> None:
        settings = load_settings()
        settings["window_width"] = self._saved_width
        settings["window_height"] = self._saved_height
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

        # Compute grid once — use saved size if available, else defaults
        has_saved = (self._saved_width is not None and self._saved_height is not None
                     and self._month_width > 0 and self._month_height > 0)
        if has_saved:
            self._grid_cols = max(1, (self._saved_width - 24) // self._month_width)
            self._grid_rows = max(1, (self._saved_height - 50) // self._month_height)
        else:
            self._grid_cols = self.months_before + 1 + self.months_after
            self._grid_rows = 1

        self._rebuild_months()
        self.root.deiconify()
        self.root.update_idletasks()

        if has_saved:
            self._position_window(override_size=(self._saved_width, self._saved_height))
        else:
            self._position_window()

        self.root.lift()
        self.root.focus_force()

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
