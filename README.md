# Mini Calendar Tool

A lightweight Windows system-tray calendar that shows ISO week numbers, supports drag-to-select date ranges, and stays always on top.

![Screenshot](Printscreen/mini-calendar.png)

## Features

- **System tray icon** displaying the current ISO week number
- **Multi-month view** with configurable months before/after
- **ISO week numbers** alongside each week row
- **Drag-to-select** date ranges with day/week count in the footer
- **Weekend highlighting** (Sat/Sun in red)
- **Start with Windows** option via registry autostart
- **Keyboard:** Escape clears selection, second Escape hides the window

## Installation

### Standalone (no Python required)

Download `MiniCalendar.exe` from [Releases](https://github.com/si3rc0-sw/mini-calendar-tool/releases) and run it.

### From source

```bash
pip install pystray Pillow
python main.py
```

Use `pythonw main.py` to run without a console window.

## Build standalone .exe

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name MiniCalendar main.py
```

The output will be in `dist/MiniCalendar.exe`.

## Usage

- **Left-click** the tray icon to toggle the calendar
- **Right-click** the tray icon for Settings / Exit
- **Navigate** months with the arrow buttons or jump to today
- **Drag** across days to select a date range and see the duration

## Settings

| Option | Description |
|---|---|
| Months before | Number of months shown before the current month (0-6) |
| Months after | Number of months shown after the current month (0-6) |
| Start with Windows | Launch automatically on Windows login |

## Requirements

- Windows 10/11
- Python 3.10+ (when running from source)
