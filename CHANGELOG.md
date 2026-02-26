# Changelog

## v1.4.1

### Fixed
- **View persistence** — grid layout (columns/rows) is now saved and restored across app restarts; previously the month count reset to 1 after relaunch

## v1.4.0

### Added
- **About dialog** — accessible via tray menu, shows version, description, license, and clickable GitHub link
- **"Check for Updates"** tray menu entry — opens GitHub Releases page in the browser
- **Text outline** on multi-holiday days — black outline around day numbers on striped color backgrounds for better readability

### Changed
- **License** — changed from MIT to MIT + Commons Clause (free to use/modify/share, commercial sale not permitted)
- **README** — updated with light + dark screenshots side by side, new features documented

## v1.3.0

### Added
- **Dark mode** — full dark theme with switchable light/dark color dictionaries (17 theme keys); toggle via Settings checkbox; persisted across sessions
- **Dark title bar** — Windows DWM API (`DwmSetWindowAttribute`) switches the OS-drawn title bar between dark and light mode
- **Custom close button** — "×" label in the nav bar with red hover effect; replaces the hard-to-read `-toolwindow` close button
- **Page navigation** — new `◀◀` / `▶▶` buttons skip forward/back by the number of currently displayed months
- **Version display** — `v1.3.0` shown bottom-right in the footer
- **MIT License** — open-source license file added
- **GitHub Actions workflow** — automated `.exe` build and release on version tags

### Changed
- **Navigation layout** — `◀` / `▶` = 1 month, `◀◀` / `▶▶` = page (displayed month count), `◀◀◀` / `▶▶▶` = 1 year
- **Calendar week header** — renamed from "Wk" to "CW"

### Removed
- **Months before/after settings** — removed from settings dialog and defaults; the resize handler already auto-fits months to window size dynamically

## v1.2.0

### Added
- **Selectable holidays** — per-country holiday checkboxes (Switzerland, Germany, China) with colored day indicators and customizable colors

## v1.1.0

### Added
- **Resizable window** — drag horizontally and vertically to show more months in a grid layout
- **Year navigation** — `<<` / `>>` arrows to jump one year back/forward
- **Window size persistence** — window dimensions are saved and restored across sessions
- **Close to tray** — X button hides the window instead of quitting the app

### Improved
- **Performance** — widget pooling (`_MonthPanel`) reuses pre-allocated labels via `configure()` instead of destroy/create (~8-10x faster rebuilds)
- **Shell/content split** — nav bar and footer are created once; only the month grid is updated on navigation or resize
- **Debounced resize** — 30ms debounce prevents redundant rebuilds during drag

## v1.0.0

- Initial release: multi-month calendar tray app with drag-to-select, week numbers, and settings dialog
