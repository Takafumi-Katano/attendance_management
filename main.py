"""
main.py – Tkinter GUI for the attendance management application.

Layout
------
  Row 0: [保存先フォルダ label] [folder path display] [フォルダ選択 button]
  Row 1: [ファイル名 label]     [filename display]    [Excelを開く button]
  ── separator ──
  Row 3: [現在時刻 label]  [live clock]
  ── separator ──
  Row 5: [始業時刻 label]  [recorded start time]  [始業 button]
  Row 6: [終業時刻 label]  [recorded end time]    [終業 button]
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox

from app_logger import get_logger
from config import Config
from attendance import AttendanceManager

logger = get_logger()

# ---------------------------------------------------------------------------
# Main application class
# ---------------------------------------------------------------------------


class AttendanceApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("勤怠管理")
        self.root.resizable(False, False)

        self.config = Config()
        self.manager = AttendanceManager(self.config)

        self._build_ui()
        self._handle_year_rollover_on_startup()
        self._load_today_times()
        self._update_clock()
        # Run previous-day check in a background thread so the window opens immediately
        threading.Thread(target=self._check_previous_day, daemon=True).start()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 5}

        # ── Row 0 : Folder path ─────────────────────────────────────────
        tk.Label(self.root, text="保存先フォルダ:", anchor="e").grid(
            row=0, column=0, sticky="e", **pad
        )
        self.folder_var = tk.StringVar(
            value=self.config.folder_path if self.config.folder_path else "未設定"
        )
        tk.Entry(
            self.root,
            textvariable=self.folder_var,
            width=45,
            state="readonly",
            readonlybackground="white",
        ).grid(row=0, column=1, columnspan=2, sticky="ew", **pad)
        tk.Button(
            self.root,
            text="フォルダ選択",
            width=12,
            command=self._choose_folder,
        ).grid(row=0, column=3, **pad)

        # ── Row 1 : File name ────────────────────────────────────────────
        tk.Label(self.root, text="ファイル名:", anchor="e").grid(
            row=1, column=0, sticky="e", **pad
        )
        self.filename_var = tk.StringVar(value=self._filename_display())
        tk.Entry(
            self.root,
            textvariable=self.filename_var,
            width=45,
            state="readonly",
            readonlybackground="white",
        ).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)
        tk.Button(
            self.root,
            text="Excelを開く",
            width=12,
            command=self._open_file,
        ).grid(row=1, column=3, **pad)

        # ── Separator ────────────────────────────────────────────────────
        tk.Frame(self.root, height=2, bd=1, relief="sunken").grid(
            row=2, column=0, columnspan=4, sticky="ew", padx=8, pady=2
        )

        # ── Row 3 : Current time ─────────────────────────────────────────
        tk.Label(self.root, text="現在時刻:", anchor="e").grid(
            row=3, column=0, sticky="e", **pad
        )
        self.current_time_var = tk.StringVar()
        tk.Label(
            self.root,
            textvariable=self.current_time_var,
            font=("", 16, "bold"),
            fg="#1565C0",
            width=12,
            anchor="w",
        ).grid(row=3, column=1, sticky="w", **pad)

        # ── Separator ────────────────────────────────────────────────────
        tk.Frame(self.root, height=2, bd=1, relief="sunken").grid(
            row=4, column=0, columnspan=4, sticky="ew", padx=8, pady=2
        )

        # ── Row 5 : Start time ───────────────────────────────────────────
        tk.Label(self.root, text="始業時刻:", anchor="e").grid(
            row=5, column=0, sticky="e", **pad
        )
        self.start_time_var = tk.StringVar(value="--:--")
        tk.Label(
            self.root,
            textvariable=self.start_time_var,
            font=("", 14),
            width=8,
            anchor="w",
        ).grid(row=5, column=1, sticky="w", **pad)
        self.start_btn = tk.Button(
            self.root,
            text="始業",
            width=10,
            bg="#388E3C",
            fg="white",
            activebackground="#2E7D32",
            font=("", 11, "bold"),
            command=self._record_start,
        )
        self.start_btn.grid(row=5, column=2, **pad)

        # ── Row 6 : End time ─────────────────────────────────────────────
        tk.Label(self.root, text="終業時刻:", anchor="e").grid(
            row=6, column=0, sticky="e", **pad
        )
        self.end_time_var = tk.StringVar(value="--:--")
        tk.Label(
            self.root,
            textvariable=self.end_time_var,
            font=("", 14),
            width=8,
            anchor="w",
        ).grid(row=6, column=1, sticky="w", **pad)
        tk.Button(
            self.root,
            text="終業",
            width=10,
            bg="#C62828",
            fg="white",
            activebackground="#B71C1C",
            font=("", 11, "bold"),
            command=self._record_end,
        ).grid(row=6, column=2, **pad)

        # Extra padding at the bottom
        tk.Label(self.root, text="").grid(row=7, column=0)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _filename_display(self) -> str:
        if not self.config.folder_path:
            return "Empty"
        return f"Attendance_Sheet_{datetime.now().year}.xlsx"

    def _update_clock(self) -> None:
        self.current_time_var.set(datetime.now().strftime("%H:%M:%S"))
        self.root.after(1000, self._update_clock)

    def _load_today_times(self) -> None:
        """Populate start/end labels from the Excel file on application start."""
        start, end = self.manager.get_today_times()
        if start:
            self.start_time_var.set(str(start))
            self.start_btn.config(state="disabled")
        if end:
            self.end_time_var.set(str(end))

    def _check_previous_day(self) -> None:
        """Silently attempt to auto-fill yesterday's missing end time."""
        try:
            self.manager.check_previous_day()
        except Exception as exc:
            print(f"[AttendanceApp] previous day check error: {exc}")
            logger.exception("[AttendanceApp] previous day check error")

    def _handle_year_rollover_on_startup(self) -> None:
        """On first launch after year change, auto-switch to new year file and notify."""
        current_year = datetime.now().year
        previous_year = self.config.last_used_year

        if previous_year is None:
            self.config.last_used_year = current_year
            self.config.save()
            return

        if previous_year == current_year:
            return

        file_path = self.manager.ensure_file_exists(current_year)
        self.config.last_used_year = current_year
        self.config.save()

        if not file_path:
            return
        self.filename_var.set(self._filename_display())

        def notify_and_open() -> None:
            messagebox.showinfo(
                "情報",
                (
                    "年が変わったため、勤怠Excelファイルを自動更新しました。\n"
                    f"新しいファイル: {os.path.basename(file_path)}"
                ),
            )
            self._open_file_path(file_path)

        self.root.after(0, notify_and_open)

    # ------------------------------------------------------------------
    # Button callbacks
    # ------------------------------------------------------------------

    def _choose_folder(self) -> None:
        current = self.folder_var.get()
        initial = current if os.path.isdir(current) else None
        path = filedialog.askdirectory(
            title="勤怠データの保存先フォルダを選択してください",
            initialdir=initial,
        )
        if not path:
            return
        self.config.folder_path = path
        self.config.save()
        self.folder_var.set(path)
        self.filename_var.set(self._filename_display())

    def _open_file(self) -> None:
        file_path = self.manager.get_file_path()
        if not file_path:
            messagebox.showinfo("情報", "保存先フォルダを選択してください。")
            return
        if not os.path.exists(file_path):
            messagebox.showinfo(
                "情報",
                f"ファイルが見つかりません:\n{file_path}\n\n"
                "始業ボタンを押すとファイルが作成されます。",
            )
            return
        self._open_file_path(file_path)

    @staticmethod
    def _open_file_path(file_path: str) -> None:
        try:
            if sys.platform == "win32":
                os.startfile(file_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", file_path], check=True)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
        except Exception as exc:
            messagebox.showerror("エラー", f"ファイルを開けませんでした:\n{exc}")

    def _record_start(self) -> None:
        if not self.config.folder_path:
            messagebox.showwarning("警告", "保存先フォルダを選択してください。")
            return
        now = datetime.now()
        result = self.manager.record_start(now)
        if result:
            self.start_time_var.set(result)
            self.start_btn.config(state="disabled")
            messagebox.showinfo("始業", f"始業時刻を記録しました: {result}")
        else:
            messagebox.showerror("エラー", "始業時刻の記録に失敗しました。")

    def _record_end(self) -> None:
        if not self.config.folder_path:
            messagebox.showwarning("警告", "保存先フォルダを選択してください。")
            return
        now = datetime.now()
        result = self.manager.record_end(now)
        if result:
            self.end_time_var.set(result)
            messagebox.showinfo("終業", f"終業時刻を記録しました: {result}")
        else:
            messagebox.showerror("エラー", "終業時刻の記録に失敗しました。")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    root = tk.Tk()
    AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
