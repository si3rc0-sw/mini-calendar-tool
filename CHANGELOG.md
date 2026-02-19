# Changelog

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
