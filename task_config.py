"""
task_config.py – Load and save the list of task types from a plain-text file.

File format (tasks_config.txt, stored alongside the Excel file):
    # Lines starting with '#' are comments and are ignored.
    # Add one task name per line.
    業務A
    業務B

The file is hand-editable.  A default skeleton is created automatically
the first time the save-folder is selected.
"""

import os
from typing import List, Optional

TASKS_CONFIG_FILENAME = "tasks_config.txt"

_DEFAULT_CONTENT = """\
# 業務種類の設定ファイル
# 1行に1つの業務名を記載してください。
# '#' で始まる行はコメントとして無視されます。
#
# 例:
# 開発業務
# 会議
# ドキュメント作成
"""


class TaskConfig:
    """Manages the list of task types read from *tasks_config.txt*."""

    def __init__(self, folder_path: Optional[str] = None) -> None:
        self._tasks: List[str] = []
        self._folder_path: Optional[str] = folder_path
        if folder_path:
            self.load(folder_path)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    @property
    def tasks(self) -> List[str]:
        """Return a copy of the current task list."""
        return list(self._tasks)

    def load(self, folder_path: str) -> None:
        """(Re-)load task list from *folder_path*/tasks_config.txt.

        Creates the file with default content if it does not exist.
        """
        self._folder_path = folder_path
        path = os.path.join(folder_path, TASKS_CONFIG_FILENAME)
        if not os.path.exists(path):
            try:
                with open(path, "w", encoding="utf-8") as fh:
                    fh.write(_DEFAULT_CONTENT)
            except OSError:
                pass
            self._tasks = []
            return
        self._tasks = self._parse(path)

    def get_config_path(self) -> Optional[str]:
        """Return the absolute path of the config file, or None if no folder set."""
        if self._folder_path:
            return os.path.join(self._folder_path, TASKS_CONFIG_FILENAME)
        return None

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    @staticmethod
    def _parse(path: str) -> List[str]:
        tasks: List[str] = []
        try:
            with open(path, "r", encoding="utf-8") as fh:
                for line in fh:
                    stripped = line.strip()
                    if stripped and not stripped.startswith("#"):
                        tasks.append(stripped)
        except OSError:
            pass
        return tasks
