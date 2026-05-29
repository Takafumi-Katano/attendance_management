"""
app_logger.py – Shared file logger for attendance management diagnostics.
"""

import logging
import os
from logging.handlers import RotatingFileHandler

_LOGGER_NAME = "attendance_management"
_DEFAULT_LOG_FILE = os.path.join(os.path.expanduser("~"), ".attendance_management.log")
_LOG_FILE = _DEFAULT_LOG_FILE


def _build_log_file_path(folder_path: str | None) -> str:
    if folder_path and os.path.isdir(folder_path):
        return os.path.join(folder_path, "attendance_management.log")
    return _DEFAULT_LOG_FILE


def get_logger() -> logging.Logger:
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers and getattr(logger, "_attendance_log_path", None) == _LOG_FILE:
        return logger

    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    handler = RotatingFileHandler(
        _LOG_FILE,
        maxBytes=1_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    )
    logger.addHandler(handler)
    logger.propagate = False
    logger._attendance_log_path = _LOG_FILE  # type: ignore[attr-defined]
    return logger


def set_log_directory(folder_path: str | None) -> str:
    global _LOG_FILE
    _LOG_FILE = _build_log_file_path(folder_path)
    get_logger()
    return _LOG_FILE


def get_log_path() -> str:
    return _LOG_FILE
