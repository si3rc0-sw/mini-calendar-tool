"""
Automated view persistence test for Mini Calendar.

Simulates the full lifecycle: start -> show -> resize -> exit -> restart -> verify.
Tests that grid_cols/grid_rows survive an app restart.
"""

import sys
import os
import ctypes
import time

sys.path.insert(0, "C:/dev/mini-calender-tool")
os.chdir("C:/dev/mini-calender-tool")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

from settings import load_settings, save_settings
from calendar_window import CalendarWindow


passed = 0
total = 0


def check(name, expected, actual):
    global passed, total
    total += 1
    ok = expected == actual
    if ok:
        passed += 1
    tag = "PASS" if ok else "FAIL"
    print(f"  [{tag}] {name}: expected={expected!r}, got={actual!r}")


def main():
    global passed, total
    print("=" * 55)
    print("  Mini Calendar - View Persistence Test")
    print("=" * 55)

    # --- Setup: save and clear grid values ---
    settings = load_settings()
    backup = {k: settings.get(k) for k in ("grid_cols", "grid_rows", "window_width", "window_height")}
    settings["grid_cols"] = None
    settings["grid_rows"] = None
    save_settings(settings)
    print(f"\n[SETUP] Cleared grid settings (backup: {backup})")

    # =========================================================
    # TEST 1: First launch — no saved grid -> default 3×1
    # =========================================================
    print("\n[TEST 1] First launch — default grid 3x1")
    cal = CalendarWindow()
    cal.show()
    cal.root.update()
    time.sleep(0.3)
    cal.root.update()
    check("Default grid_cols", 3, cal._grid_cols)
    check("Default grid_rows", 1, cal._grid_rows)

    # =========================================================
    # TEST 2: User resizes window to show 5 months
    # Simulate by setting grid and calling _rebuild + _persist
    # =========================================================
    print("\n[TEST 2] User resizes to 5x1")
    cal._showing = False
    cal._grid_cols = 5
    cal._grid_rows = 1
    cal._saved_grid_cols = 5
    cal._saved_grid_rows = 1
    cal._rebuild_months()
    cal.root.update_idletasks()
    cal._saved_width = cal.root.winfo_reqwidth()
    cal._saved_height = cal.root.winfo_reqheight()
    cal._persist_size()

    saved = load_settings()
    check("Persisted grid_cols", 5, saved["grid_cols"])
    check("Persisted grid_rows", 1, saved["grid_rows"])
    check("Persisted window_width is int", True, isinstance(saved["window_width"], int))

    # =========================================================
    # TEST 3: User hides window (simulates close button)
    # =========================================================
    print("\n[TEST 3] User hides window")
    cal.hide()
    saved = load_settings()
    check("Grid_cols after hide", 5, saved["grid_cols"])
    check("Grid_rows after hide", 1, saved["grid_rows"])

    # =========================================================
    # TEST 4: Simulate app exit + restart (destroy + new instance)
    # =========================================================
    print("\n[TEST 4] App exit -> restart (destroy + new CalendarWindow)")
    cal._persist_size()  # same as on_exit
    cal.root.destroy()

    cal2 = CalendarWindow()
    check("New instance _saved_grid_cols", 5, cal2._saved_grid_cols)
    check("New instance _saved_grid_rows", 1, cal2._saved_grid_rows)

    cal2.show()
    cal2.root.update()
    time.sleep(0.3)
    cal2.root.update()

    check("Restored grid_cols after restart", 5, cal2._grid_cols)
    check("Restored grid_rows after restart", 1, cal2._grid_rows)

    # =========================================================
    # TEST 5: Verify _showing flag blocks resize during show()
    # =========================================================
    print("\n[TEST 5] _showing flag blocks resize handler")
    cal2._showing = True
    cal2._saved_width = 100  # tiny — would compute cols=0->1
    fake_event = type("Event", (), {"widget": cal2.root})()
    cal2._on_configure(fake_event)
    check("Grid unchanged during _showing", 5, cal2._grid_cols)
    cal2._showing = False

    # =========================================================
    # TEST 6: Resize to 2x2, exit, restart — verify 2x2
    # =========================================================
    print("\n[TEST 6] Resize to 2x2, exit, restart")
    cal2._grid_cols = 2
    cal2._grid_rows = 2
    cal2._saved_grid_cols = 2
    cal2._saved_grid_rows = 2
    cal2._saved_width = 1000
    cal2._saved_height = 800
    cal2._persist_size()
    cal2.root.destroy()

    cal3 = CalendarWindow()
    cal3.show()
    cal3.root.update()
    time.sleep(0.3)
    cal3.root.update()

    check("Restored grid_cols=2", 2, cal3._grid_cols)
    check("Restored grid_rows=2", 2, cal3._grid_rows)
    check("Total months shown", 4, cal3._current_total_months)

    # =========================================================
    # RESULTS
    # =========================================================
    print("\n" + "=" * 55)
    print("  RESULTS")
    print("=" * 55)
    print(f"  {passed}/{total} checks passed.")
    if passed == total:
        print("  ALL CHECKS PASSED.")
    else:
        print("  SOME CHECKS FAILED.")

    # Restore original settings
    settings = load_settings()
    for k, v in backup.items():
        settings[k] = v
    save_settings(settings)
    print(f"\n[CLEANUP] Restored original settings")

    cal3.root.destroy()
    print("\nDone.")


if __name__ == "__main__":
    main()
