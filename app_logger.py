"""
app_logger.py – Shared file logger for attendance management diagnostics.
"""

import logging
import os
from logging.handlers import RotatingFileHandler

_LOGGER_NAME = "attendance_management"
_LOG_FILE = os.path.join(os.path.expanduser("~"), ".attendance_management.log")


def get_logger() -> logging.Logger:
    logger = logging.getLogger(_LOGGER_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
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
    return logger


def get_log_path() -> str:
    return _LOG_FILE
