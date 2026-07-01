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
  ── separator ──
  Row 8: [業務 label] [task dropdown] [作業開始 button]
"""

import atexit
import os
import socket
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from typing import Optional

from app_logger import get_log_path, get_logger, set_log_directory
from config import Config
from attendance import AttendanceManager
from task_config import TaskConfig
from task_session import TaskSessionManager

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
        log_path = set_log_directory(self.config.folder_path)
        logger.info("application started: log_path=%s", log_path)

        self.task_config = TaskConfig(self.config.folder_path)
        self.task_session_manager = TaskSessionManager(self.config.folder_path)
        self.manager = AttendanceManager(self.config, self.task_session_manager)

        self._build_ui()
        self._handle_year_rollover_on_startup()
        self._load_today_times()
        self._update_clock()
        # Run previous-day check in a background thread so the window opens immediately
        threading.Thread(target=self._check_previous_day, daemon=True).start()

        # System tray (best-effort; ignored if pystray/Pillow not available)
        self.tray_icon = None
        self._is_hiding_to_tray = False
        self._start_tray()

        # Hide window on close → keep residing in the system tray
        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)
        # Hide window on minimize as well (keep only tray presence)
        self.root.bind("<Unmap>", self._on_window_unmap)

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
        ).grid(row=0, column=1, columnspan=3, sticky="ew", **pad)
        tk.Button(
            self.root,
            text="フォルダ選択",
            width=12,
            command=self._choose_folder,
        ).grid(row=0, column=4, **pad)

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
        tk.Button(
            self.root,
            text="ログを開く",
            width=12,
            command=self._open_log_file,
        ).grid(row=1, column=4, **pad)

        # ── Separator ────────────────────────────────────────────────────
        tk.Frame(self.root, height=2, bd=1, relief="sunken").grid(
            row=2, column=0, columnspan=5, sticky="ew", padx=8, pady=2
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
            row=4, column=0, columnspan=5, sticky="ew", padx=8, pady=2
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

        # ── Separator ────────────────────────────────────────────────────
        tk.Frame(self.root, height=2, bd=1, relief="sunken").grid(
            row=7, column=0, columnspan=5, sticky="ew", padx=8, pady=2
        )

        # ── Row 8 : Task selection ───────────────────────────────────────
        tk.Label(self.root, text="業務:", anchor="e").grid(
            row=8, column=0, sticky="e", **pad
        )
        self.task_var = tk.StringVar()
        self.task_combo = ttk.Combobox(
            self.root,
            textvariable=self.task_var,
            state="readonly",
            width=30,
        )
        self.task_combo.grid(row=8, column=1, columnspan=2, sticky="ew", **pad)
        tk.Button(
            self.root,
            text="作業開始",
            width=10,
            bg="#1565C0",
            fg="white",
            activebackground="#0D47A1",
            font=("", 11, "bold"),
            command=self._start_task,
        ).grid(row=8, column=3, **pad)

        # ── Row 9 : Active task status label ─────────────────────────────
        self.active_task_var = tk.StringVar(value="")
        tk.Label(
            self.root,
            textvariable=self.active_task_var,
            fg="#555555",
            font=("", 9),
            anchor="w",
        ).grid(row=9, column=0, columnspan=5, sticky="w", padx=10, pady=2)

        # Extra padding at the bottom
        tk.Label(self.root, text="").grid(row=10, column=0)

        # Populate the task dropdown
        self._update_task_dropdown()

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _filename_display(self) -> str:
        if not self.config.folder_path:
            return "Empty"
        return f"Attendance_Sheet_{datetime.now().year}.xlsx"

    def _update_task_dropdown(self) -> None:
        """Reload task list from config and refresh the Combobox."""
        tasks = self.task_config.tasks
        self.task_combo["values"] = tasks
        if tasks and not self.task_var.get():
            self.task_combo.current(0)
        elif not tasks:
            self.task_var.set("")
        # Reflect currently active task in the status label
        active = self.task_session_manager.active_task
        if active:
            self.active_task_var.set(f"実施中: {active}")
        else:
            self.active_task_var.set("")

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
        logger.info("loaded today times: start=%s end=%s", start, end)

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
        log_path = set_log_directory(path)
        self.folder_var.set(path)
        self.filename_var.set(self._filename_display())
        # Reload task config and session manager for the new folder
        self.task_config.load(path)
        self.task_session_manager.set_folder_path(path)
        self._update_task_dropdown()
        # Restart the tray so its task submenu reflects the new task list
        self._restart_tray()
        logger.info("folder selected: %s log_path=%s", path, log_path)

    def _open_file(self) -> None:
        file_path = self.manager.get_file_path()
        if not file_path:
            logger.info("open excel skipped: folder not configured")
            messagebox.showinfo("情報", "保存先フォルダを選択してください。")
            return
        if not os.path.exists(file_path):
            logger.info("open excel skipped: file not found %s", file_path)
            messagebox.showinfo(
                "情報",
                f"ファイルが見つかりません:\n{file_path}\n\n"
                "始業ボタンを押すとファイルが作成されます。",
            )
            return
        logger.info("opening excel file: %s", file_path)
        self._open_file_path(file_path)

    def _open_log_file(self) -> None:
        log_path = get_log_path()
        if not os.path.exists(log_path):
            open(log_path, "a", encoding="utf-8").close()
        logger.info("opening log file: %s", log_path)
        self._open_file_path(log_path)

    @staticmethod
    def _open_file_path(file_path: str) -> None:
        try:
            logger.info("open file path requested: %s", file_path)
            if sys.platform == "win32":
                os.startfile(file_path)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.run(["open", file_path], check=True)
            else:
                subprocess.run(["xdg-open", file_path], check=True)
        except Exception as exc:
            logger.exception("failed to open file: %s", file_path)
            messagebox.showerror("エラー", f"ファイルを開けませんでした:\n{exc}")

    def _record_start(self) -> None:
        if not self.config.folder_path:
            logger.info("record start skipped: folder not configured")
            messagebox.showwarning("警告", "保存先フォルダを選択してください。")
            return
        now = datetime.now()
        logger.info("record start requested: dt=%s", now.strftime("%Y-%m-%d %H:%M:%S"))
        result = self.manager.record_start(now)
        if result:
            self.start_time_var.set(result)
            self.start_btn.config(state="disabled")
            logger.info("record start success: %s", result)
            messagebox.showinfo("始業", f"始業時刻を記録しました: {result}")
        else:
            logger.error("record start failed")
            messagebox.showerror("エラー", "始業時刻の記録に失敗しました。")

    def _record_end(self) -> None:
        if not self.config.folder_path:
            logger.info("record end skipped: folder not configured")
            messagebox.showwarning("警告", "保存先フォルダを選択してください。")
            return
        now = datetime.now()
        logger.info("record end requested: dt=%s", now.strftime("%Y-%m-%d %H:%M:%S"))
        # Close the active task session before writing end time to Excel
        self.task_session_manager.end_current_task(now)
        self.active_task_var.set("")
        result = self.manager.record_end(now)
        if result:
            self.end_time_var.set(result)
            logger.info("record end success: %s", result)
            messagebox.showinfo("終業", f"終業時刻を記録しました: {result}")
        else:
            logger.error("record end failed")
            messagebox.showerror("エラー", "終業時刻の記録に失敗しました。")

    def _start_task(self) -> None:
        """Begin the selected task (called by the 作業開始 button)."""
        if not self.config.folder_path:
            messagebox.showwarning("警告", "保存先フォルダを選択してください。")
            return
        task_name = self.task_var.get().strip()
        if not task_name:
            messagebox.showwarning("警告", "業務を選択してください。")
            return
        now = datetime.now()
        logger.info(
            "start_task: task=%s dt=%s", task_name, now.strftime("%Y-%m-%d %H:%M:%S")
        )
        self.task_session_manager.start_task(task_name, now)
        self.active_task_var.set(f"実施中: {task_name}")
        logger.info("start_task success: task=%s", task_name)

    def _start_task_by_name(self, task_name: str) -> None:
        """Begin *task_name* – called from the system tray submenu."""
        if not self.config.folder_path:
            return
        now = datetime.now()
        logger.info(
            "start_task_by_name (tray): task=%s dt=%s",
            task_name,
            now.strftime("%Y-%m-%d %H:%M:%S"),
        )
        self.task_session_manager.start_task(task_name, now)
        self.active_task_var.set(f"実施中: {task_name}")

    # ------------------------------------------------------------------
    # Window lifecycle
    # ------------------------------------------------------------------

    def _on_window_close(self) -> None:
        """Hide the main window to the system tray instead of destroying it."""
        self._hide_to_tray()

    def _on_window_unmap(self, _event) -> None:
        """When minimized, hide from taskbar and keep running in tray."""
        if self.tray_icon is None:
            return
        if self._is_hiding_to_tray:
            return
        if self.root.state() == "iconic":
            self.root.after(0, self._hide_to_tray)

    def _hide_to_tray(self) -> None:
        self._is_hiding_to_tray = True
        self.root.withdraw()
        self.root.after_idle(self._clear_hide_to_tray_flag)

    def _clear_hide_to_tray_flag(self) -> None:
        self._is_hiding_to_tray = False

    def _show_window(self) -> None:
        """Restore and focus the main window."""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _quit_app(self) -> None:
        """Fully exit the application (called from the tray menu)."""
        logger.info("quit_app called")
        if self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
        self.root.after(0, self.root.destroy)

    # ------------------------------------------------------------------
    # System tray
    # ------------------------------------------------------------------

    def _create_tray_icon_image(self):
        """Return a simple 64×64 PIL image for the system tray icon."""
        import math
        from PIL import Image, ImageDraw

        size = 64
        img = Image.new("RGB", (size, size), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        # Green circle
        draw.ellipse([4, 4, 60, 60], fill=(56, 142, 60), outline=(255, 255, 255), width=2)
        # Clock hands (decorative)
        cx, cy, r = 32, 32, 20
        hx = cx + int(r * 0.5 * math.sin(math.radians(60)))
        hy = cy - int(r * 0.5 * math.cos(math.radians(60)))
        draw.line([(cx, cy), (hx, hy)], fill="white", width=3)
        mx = cx + int(r * math.sin(math.radians(120)))
        my = cy - int(r * math.cos(math.radians(120)))
        draw.line([(cx, cy), (mx, my)], fill="white", width=2)
        return img

    def _tray_task_callback(self, task_name: str):
        """Return a pystray-compatible callback that starts *task_name*."""
        def callback(_icon, _item) -> None:
            self.root.after(0, lambda: self._start_task_by_name(task_name))
        return callback

    def _build_tray_menu(self):
        """Build a pystray Menu reflecting the current task list."""
        import pystray

        tasks = self.task_config.tasks
        if tasks:
            task_items = tuple(
                pystray.MenuItem(task, self._tray_task_callback(task))
                for task in tasks
            )
            task_submenu = pystray.Menu(*task_items)
        else:
            task_submenu = pystray.Menu(
                pystray.MenuItem("(業務未設定)", None, enabled=False)
            )

        return pystray.Menu(
            pystray.MenuItem(
                "ウィンドウを表示",
                lambda _icon, _item: self.root.after(0, self._show_window),
                default=True,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "始業",
                lambda _icon, _item: self.root.after(0, self._record_start),
            ),
            pystray.MenuItem(
                "終業",
                lambda _icon, _item: self.root.after(0, self._record_end),
            ),
            pystray.MenuItem("個別業務", task_submenu),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "アプリ終了",
                lambda _icon, _item: self._quit_app(),
            ),
        )

    def _start_tray(self) -> None:
        """Start the system tray icon in a background daemon thread (best-effort)."""
        try:
            import pystray  # presence check

            icon_image = self._create_tray_icon_image()
            menu = self._build_tray_menu()
            self.tray_icon = pystray.Icon(
                "attendance_management",
                icon_image,
                "勤怠管理",
                menu,
            )
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
            logger.info("system tray started")
        except Exception as exc:
            logger.warning("system tray unavailable: %s", exc)

    def _restart_tray(self) -> None:
        """Stop the existing tray icon and create a new one (e.g. after folder change)."""
        if self.tray_icon is not None:
            try:
                self.tray_icon.stop()
            except Exception:
                pass
            self.tray_icon = None
        self._start_tray()


# ---------------------------------------------------------------------------
# Single-instance enforcement
# ---------------------------------------------------------------------------

# Arbitrary local port used to detect and signal an already-running instance.
# Only loopback (127.0.0.1) is used, so no network exposure occurs.
# The value can be overridden via the ATTENDANCE_INSTANCE_PORT environment
# variable if port 47832 happens to be occupied by another application.
_INSTANCE_PORT_DEFAULT = 47832
try:
    _INSTANCE_PORT = int(os.environ.get("ATTENDANCE_INSTANCE_PORT", _INSTANCE_PORT_DEFAULT))
except ValueError:
    logger.warning(
        "ATTENDANCE_INSTANCE_PORT is not a valid integer; using default port %d",
        _INSTANCE_PORT_DEFAULT,
    )
    _INSTANCE_PORT = _INSTANCE_PORT_DEFAULT
# How long (seconds) to wait when probing for an already-running instance.
_INSTANCE_CHECK_TIMEOUT_SEC = 0.5
# Maximum number of pending connections accepted by the instance server socket.
_INSTANCE_SERVER_BACKLOG = 5
_INSTANCE_SIGNAL = b"attendance_management:show_window"
_INSTANCE_ACK = b"attendance_management:ok"


def _recv_exact(sock: socket.socket, size: int) -> Optional[bytes]:
    """Receive exactly *size* bytes; return None if the stream closes early."""
    chunks = []
    received = 0
    while received < size:
        chunk = sock.recv(size - received)
        if not chunk:
            return None
        chunks.append(chunk)
        received += len(chunk)
    return b"".join(chunks)


def _start_instance_server(root: tk.Tk) -> None:
    """Bind a local TCP port and listen for focus requests from duplicate launches.

    When a second instance of the application starts, it connects to this port.
    On receiving a connection the first instance brings its window to the foreground.
    The listener runs in a daemon thread and is terminated automatically on exit.
    The server socket is registered with atexit to ensure it is closed on normal exit.
    """
    server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    server.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        server.bind(("127.0.0.1", _INSTANCE_PORT))
        server.listen(_INSTANCE_SERVER_BACKLOG)
    except OSError as exc:
        # Could not bind (e.g. port occupied by another process).
        # Single-instance protection is skipped; the app still runs normally.
        logger.warning("single-instance server could not bind port %d: %s", _INSTANCE_PORT, exc)
        server.close()
        return

    # Ensure the socket is closed when the process exits normally so the port
    # is released promptly and subsequent launches are not blocked.
    atexit.register(server.close)

    def _bring_to_front() -> None:
        """Restore and focus the main window (must be called on the Tk thread)."""
        root.deiconify()
        root.lift()
        root.focus_force()

    def _listen() -> None:
        while True:
            try:
                conn, _ = server.accept()
                with conn:
                    conn.settimeout(_INSTANCE_CHECK_TIMEOUT_SEC)
                    payload = _recv_exact(conn, len(_INSTANCE_SIGNAL))
                    if payload != _INSTANCE_SIGNAL:
                        continue
                    conn.sendall(_INSTANCE_ACK)
                root.after(0, _bring_to_front)
            except OSError:
                # Server socket was closed (application exiting); stop the loop.
                break
            except Exception:
                # Stop the listener on any unexpected error to avoid a runaway loop.
                logger.exception("instance server listener encountered an unexpected error; stopping")
                break

    threading.Thread(target=_listen, daemon=True).start()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def main() -> None:
    # Single-instance guard: try to connect to a port that a running instance is
    # listening on.  If the connection succeeds the existing instance will bring
    # its window to the front; this process then exits without opening a window.
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.settimeout(_INSTANCE_CHECK_TIMEOUT_SEC)
            s.connect(("127.0.0.1", _INSTANCE_PORT))
            s.sendall(_INSTANCE_SIGNAL)
            if _recv_exact(s, len(_INSTANCE_ACK)) == _INSTANCE_ACK:
                # Connection succeeded to our existing app instance.
                return
    except OSError:
        pass  # No existing instance found; proceed with normal startup.

    root = tk.Tk()
    _start_instance_server(root)
    AttendanceApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
