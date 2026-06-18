from __future__ import annotations

import getpass
import logging
from datetime import datetime
from pathlib import Path
import sys


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def get_log_path() -> Path:
    log_dir = get_app_dir() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    return log_dir / f"smd_auto_appeal_request_{datetime.now():%Y-%m-%d}.log"


def setup_logger() -> logging.Logger:
    logger = logging.getLogger("smd_auto_request")
    logger.setLevel(logging.INFO)

    if logger.handlers:
        return logger

    log_path = get_log_path()

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    file_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.propagate = False

    logger.info("============================================================")
    logger.info("Application started")
    logger.info("Log file: %s", log_path)
    logger.info("User: %s", getpass.getuser())

    return logger