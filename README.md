# attendance_management

Python-based attendance management desktop application targeting **Windows 11**.

## Features

| Feature | Details |
|---|---|
| **Excel storage** | One file per year (`Attendance_Sheet_YYYY.xlsx`), one sheet per month (e.g. `Jan`) |
| **Columns** | Date / Start time / End time / Work time — sorted by date ascending |
| **Break deduction** | Work time > 6 h → 1-hour break automatically deducted |
| **Monthly total** | When all calendar days for a month are recorded, a total row is appended |
| **Start / End buttons** | Timestamps written to Excel immediately and displayed in the window |
| **Auto end-time** | On startup, scans all `Attendance_Sheet_*.xlsx` files across years and months to find the **latest** date with a start time but no end time, then queries Windows Event Log for a hibernate / lock / shutdown event on that date; the inferred time is written only if it falls between **12:00 and 23:59** on that date |
| **Diagnostic log** | Startup auto-fill decisions and Windows Event lookup results are written to `~/.attendance_management.log` (rotating, up to 3 backups) for troubleshooting |
| **Persistent folder** | Selected folder saved to `~/.attendance_config.json`; shown as `Empty` on first use |

## Requirements

- Python 3.10 or later
- Windows 11 (the Windows Event Log feature requires Windows; all other features work cross-platform)

## Installation

```powershell
# 1. Clone / download the repository
# 2. Install Python dependencies
pip install -r requirements.txt
```

## Usage

```powershell
python main.py
```

## Building a standalone executable

Run `build.bat` on Windows to produce a directory-based executable under `dist\attendance_management\`:

```bat
build.bat
```

The script installs all required dependencies (including [PyInstaller](https://pyinstaller.org/)) and then calls PyInstaller using the included `attendance_management.spec` configuration.

Once the build completes, distribute the **entire `dist\attendance_management\` folder** — no Python installation is required on the target machine.
The entry point is `dist\attendance_management\attendance_management.exe`.

> **Note**: The build produces **only** the `dist\attendance_management\` folder.
> No standalone `dist\attendance_management.exe` is generated in the `dist\` root.

> **Startup time**: The build uses PyInstaller's *onedir* (directory-based) mode instead of the previous *onefile* mode.
> Unlike a single-file executable, which must extract all bundled files to a temporary directory on every launch, *onedir* runs the files directly from the distribution folder.
> This significantly reduces startup time (typically from ~10 seconds down to 1–2 seconds on Windows 11 SSD).

1. Click **フォルダ選択** to choose where the Excel files will be saved.
2. Click **始業** when you start work — the time is written to Excel immediately.
3. Click **終業** when you finish work — the time and calculated work time are written to Excel.
4. Click **Excelを開く** to open the current month's file directly.

### Troubleshooting auto end-time

When auto end-time does not fill as expected, check:

- `~/.attendance_management.log`
- Latest lines around app startup (`check_previous_day`, `get_last_work_end_time`)

## File layout

```
<selected folder>/
└── Attendance_Sheet_2025.xlsx
    ├── Jan   (columns: 日付 | 始業時間 | 終業時間 | 労働時間)
    ├── Feb
    └── ...
```
