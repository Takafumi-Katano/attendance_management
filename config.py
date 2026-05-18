"""
config.py – Loads and saves user preferences (save folder path).

The config file is stored at ~/.attendance_config.json so it persists
across sessions but is never committed to the repository.
"""

import json
import os

_CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".attendance_config.json")


class Config:
    """Manages persistent configuration for the attendance management app."""

    def __init__(self) -> None:
        self.folder_path: str | None = None
        self.last_used_year: int | None = None
        self._load()

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------

    def save(self) -> None:
        """Persist current configuration to disk."""
        with open(_CONFIG_FILE, "w", encoding="utf-8") as fh:
            json.dump(
                {
                    "folder_path": self.folder_path,
                    "last_used_year": self.last_used_year,
                },
                fh,
                ensure_ascii=False,
            )

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _load(self) -> None:
        """Load configuration from disk (silently ignore missing file)."""
        if not os.path.exists(_CONFIG_FILE):
            return
        try:
            with open(_CONFIG_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            self.folder_path = data.get("folder_path")
            year = data.get("last_used_year")
            self.last_used_year = int(year) if year is not None else None
        except (json.JSONDecodeError, OSError, ValueError, TypeError):
            pass
