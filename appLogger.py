from __future__ import annotations

import getpass
import logging
from pathlib import Path
from datetime import datetime
import sys


def get_app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_log_path(log_dir: Path | None = None) -> Path:
    target_dir = log_dir if log_dir is not None else (get_app_dir() / "logs")
    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y-%m-%d_%I%M%S%p")
    return target_dir / f"smd_auto_appeal_request_{timestamp}.log"


def setup_logger(log_dir: Path | None = None, reset: bool = False) -> logging.Logger:
    logger = logging.getLogger("smd_auto_request")
    logger.setLevel(logging.INFO)

    if reset:
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
            handler.close()

    if logger.handlers:
        return logger

    log_path = get_log_path(log_dir)

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
    logger.info("Run folder: %s", log_path.parent)
    logger.info("User: %s", getpass.getuser())

    return logger