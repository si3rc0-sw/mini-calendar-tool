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


class CalendarWindow:
    """Multi-month calendar that appears above the taskbar."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title(self._title())
        self.root.resizable(False, False)
        self.root.configure(bg=GRID_BG)

        self.root.attributes("-toolwindow", True)
        self.root.attributes("-topmost", True)

        self._setup_fonts()

        settings = load_settings()
        self.months_before: int = settings["months_before"]
        self.months_after: int = settings["months_after"]

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

        self._build()

        self.root.bind("<Escape>", self._on_escape)
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
    # Build / rebuild
    # ------------------------------------------------------------------
    def _build(self) -> None:
        for w in self.root.winfo_children():
            w.destroy()

        self._widget_dates.clear()
        self._date_widgets.clear()

        outer = tk.Frame(self.root, bg=GRID_BG)
        outer.pack(padx=6, pady=4)

        # Navigation row
        nav = tk.Frame(outer, bg=GRID_BG)
        nav.pack(fill="x", pady=(0, 2))

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

        btn_next = tk.Label(
            nav, text="\u25B6", font=self.font_nav, bg=GRID_BG, cursor="hand2"
        )
        btn_next.pack(side="right", padx=6)
        btn_next.bind("<Button-1>", lambda _e: self._navigate(1))

        # Months: months_before + current + months_after
        months_frame = tk.Frame(outer, bg=GRID_BG)
        months_frame.pack()

        month_list: list[tuple[int, int]] = []
        y, m = self.center_year, self.center_month
        for _ in range(self.months_before):
            y, m = prev_month(y, m)
        for _ in range(self.months_before + 1 + self.months_after):
            month_list.append((y, m))
            y, m = next_month(y, m)

        for i, (y, m) in enumerate(month_list):
            f = tk.Frame(months_frame, bg=GRID_BG)
            f.grid(row=0, column=i, padx=6, pady=2, sticky="n")
            self._draw_month(f, y, m)

        # Footer
        self._footer_label = tk.Label(
            outer, text=self._footer_text(), font=self.font_footer,
            bg=GRID_BG, fg="#555555",
        )
        self._footer_label.pack(pady=(4, 0))

    # ------------------------------------------------------------------
    # Single month widget
    # ------------------------------------------------------------------
    def _draw_month(self, parent: tk.Frame, year: int, month: int) -> None:
        today = date.today()

        header_text = f"{_cal.month_name[month]} {year}"
        tk.Label(
            parent, text=header_text, font=self.font_header, bg=HEADER_BG, fg="#333333"
        ).grid(row=0, column=0, columnspan=8, sticky="we", pady=(0, 2))

        for col, abbr in enumerate(DAY_ABBR):
            fg = "#CC0000" if col >= 5 else "#333333"
            tk.Label(
                parent, text=abbr, font=self.font_bold, bg=GRID_BG, fg=fg, width=3
            ).grid(row=1, column=col + 1)

        tk.Label(
            parent, text="Wk", font=self.font_bold, bg=GRID_BG, fg=WN_FG, width=3
        ).grid(row=1, column=0)

        grid = month_grid(year, month)
        weeks = iso_week_numbers(year, month)

        sel_lo, sel_hi = self._sel_range()

        for r, (row_days, wk) in enumerate(zip(grid, weeks)):
            grid_row = r + 2
            tk.Label(
                parent, text=wk, font=self.font_wn, bg=GRID_BG, fg=WN_FG, width=3
            ).grid(row=grid_row, column=0)

            for c, day in enumerate(row_days):
                if day is None:
                    tk.Label(parent, text="", bg=GRID_BG, width=3).grid(
                        row=grid_row, column=c + 1
                    )
                    continue

                d = date(year, month, day)
                is_today = d == today
                is_weekend = c >= 5
                in_sel = sel_lo is not None and sel_lo <= d <= sel_hi

                bg, fg = self._day_colors(is_today, is_weekend, in_sel)

                lbl = tk.Label(
                    parent, text=str(day),
                    font=self.font_bold if is_today else self.font_normal,
                    bg=bg, fg=fg, width=3, cursor="hand2",
                )
                lbl.grid(row=grid_row, column=c + 1)

                # Store mappings
                self._widget_dates[id(lbl)] = d
                self._date_widgets[d] = lbl

                # Drag events
                lbl.bind("<ButtonPress-1>", self._on_press)
                lbl.bind("<B1-Motion>", self._on_motion)
                lbl.bind("<ButtonRelease-1>", self._on_release)

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
            save_settings({
                "months_before": self.months_before,
                "months_after": self.months_after,
            })
            set_autostart(autostart_var.get())
            dlg.destroy()
            self._build()

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
        self._build()

    def _go_today(self) -> None:
        today = date.today()
        self.center_year = today.year
        self.center_month = today.month
        self._clear_selection()
        self._build()

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
        self._build()
        self.root.deiconify()
        self.root.update_idletasks()
        self._position_window()
        self.root.lift()
        self.root.focus_force()

    def hide(self) -> None:
        self.root.withdraw()

    # ------------------------------------------------------------------
    # Position bottom-right above taskbar
    # ------------------------------------------------------------------
    def _position_window(self) -> None:
        self.root.update_idletasks()

        rect = ctypes.wintypes.RECT()
        ctypes.windll.user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)
        work_right = rect.right
        work_bottom = rect.bottom

        win_w = self.root.winfo_reqwidth()
        win_h = self.root.winfo_reqheight()

        x = work_right - win_w - 12
        y = work_bottom - win_h - 12
        self.root.geometry(f"+{x}+{y}")
