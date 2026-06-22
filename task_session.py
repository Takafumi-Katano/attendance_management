"""
task_session.py – Track task work sessions and compute per-task durations.

A "session" is a (task, start_time, end_time) triple for a single calendar day.
Sessions are persisted to *task_sessions.json* in the same folder as the Excel
file so that data survives application restarts.

Lunch-break rule (requirement):
    If the session's time range [start, end] overlaps with 12:30–13:30,
    deduct exactly 1 hour from the computed duration.
"""

import json
import os
from datetime import datetime, timedelta
from typing import Dict, List, Optional

from app_logger import get_logger

SESSIONS_FILENAME = "task_sessions.json"
DATE_FMT = "%Y/%m/%d"
TIME_FMT = "%H:%M"
_LUNCH_START = "12:30"
_LUNCH_END = "13:30"

logger = get_logger()


class TaskSessionManager:
    """Manages task work sessions, persistence, and duration calculations."""

    def __init__(self, folder_path: Optional[str] = None) -> None:
        # date_str -> list of {task, start, end}
        self._sessions: Dict[str, List[Dict[str, str]]] = {}
        # {task, start, date} or None
        self._active: Optional[Dict[str, str]] = None
        self._folder_path: Optional[str] = folder_path
        if folder_path:
            self._load()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @property
    def active_task(self) -> Optional[str]:
        """Name of the currently running task, or None."""
        return self._active["task"] if self._active else None

    def set_folder_path(self, folder_path: str) -> None:
        """Switch to a new save folder and reload persisted sessions."""
        self._folder_path = folder_path
        self._sessions = {}
        self._active = None
        self._load()

    def start_task(self, task_name: str, dt: datetime) -> None:
        """Begin *task_name* at *dt*.

        Closes any currently active task first (using *dt* as its end time).
        """
        date_str = dt.strftime(DATE_FMT)
        time_str = dt.strftime(TIME_FMT)

        if self._active:
            self._close_active(time_str)

        self._active = {"task": task_name, "start": time_str, "date": date_str}
        logger.info(
            "task_session start_task: task=%s date=%s time=%s",
            task_name,
            date_str,
            time_str,
        )
        self._save()

    def end_current_task(self, dt: datetime) -> None:
        """Close the active task at *dt* (e.g. when 終業 is pressed)."""
        if not self._active:
            return
        time_str = dt.strftime(TIME_FMT)
        task_name = self._active["task"]
        self._close_active(time_str)
        logger.info(
            "task_session end_current_task: task=%s time=%s",
            task_name,
            time_str,
        )
        self._save()

    def get_task_durations_for_date(self, date_str: str) -> Dict[str, float]:
        """Return ``{task_name: excel_day_fraction}`` for *date_str*.

        The value is an Excel-compatible day fraction (``seconds / 86400``)
        that can be written directly to a cell formatted as ``[h]:mm``.
        """
        sessions = self._sessions.get(date_str, [])
        durations: Dict[str, float] = {}
        for s in sessions:
            dur = self._calc_duration_days(s["start"], s["end"])
            task = s["task"]
            durations[task] = durations.get(task, 0.0) + dur
        return durations

    def get_first_occurrence_for_month(self, year: int, month: int) -> List[str]:
        """Return tasks that occurred in *year*/*month*, ordered by first occurrence."""
        prefix = f"{year}/{month:02d}/"
        first_date: Dict[str, str] = {}
        for date_str, sessions in self._sessions.items():
            if not date_str.startswith(prefix):
                continue
            for s in sessions:
                task = s["task"]
                if task not in first_date or date_str < first_date[task]:
                    first_date[task] = date_str
        return sorted(first_date.keys(), key=lambda t: first_date[t])

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _close_active(self, end_time_str: str) -> None:
        if not self._active:
            return
        date_str = self._active["date"]
        if date_str not in self._sessions:
            self._sessions[date_str] = []
        self._sessions[date_str].append(
            {
                "task": self._active["task"],
                "start": self._active["start"],
                "end": end_time_str,
            }
        )
        self._active = None

    @staticmethod
    def _calc_duration_days(start_str: str, end_str: str) -> float:
        """Compute the duration from *start_str* to *end_str* as an Excel day
        fraction, applying the lunch-break deduction when applicable.
        """
        start = datetime.strptime(start_str, TIME_FMT)
        end = datetime.strptime(end_str, TIME_FMT)
        if end <= start:
            # Treat as zero-length session (same minute or clock goes backward)
            return 0.0

        diff = end - start

        # Deduct 1 hour if the lunch window [12:30, 13:30) overlaps [start, end)
        lunch_start = datetime.strptime(_LUNCH_START, TIME_FMT)
        lunch_end = datetime.strptime(_LUNCH_END, TIME_FMT)
        if start < lunch_end and end > lunch_start:
            diff -= timedelta(hours=1)

        total_seconds = max(0.0, diff.total_seconds())
        return total_seconds / 86400.0  # Convert to Excel day fraction

    def _load(self) -> None:
        if not self._folder_path:
            return
        path = os.path.join(self._folder_path, SESSIONS_FILENAME)
        if not os.path.exists(path):
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            self._sessions = data.get("sessions", {})
            self._active = data.get("active")
            logger.info("task_session loaded: path=%s", path)
        except (json.JSONDecodeError, OSError, TypeError):
            logger.exception("task_session load error: path=%s", path)

    def _save(self) -> None:
        if not self._folder_path:
            return
        path = os.path.join(self._folder_path, SESSIONS_FILENAME)
        try:
            with open(path, "w", encoding="utf-8") as fh:
                json.dump(
                    {"sessions": self._sessions, "active": self._active},
                    fh,
                    ensure_ascii=False,
                    indent=2,
                )
        except OSError:
            logger.exception("task_session save error: path=%s", path)
