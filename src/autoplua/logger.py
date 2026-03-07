from __future__ import annotations

import logging
import os
from pathlib import Path


LOG_NAME = "autoplua"


def _log_dir() -> Path:
    appdata = os.getenv("APPDATA")
    if appdata:
        return Path(appdata) / "AutoPlua" / "logs"
    return Path.home() / ".autoplua" / "logs"


def setup_logger() -> logging.Logger:
    logger = logging.getLogger(LOG_NAME)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)
    log_dir = _log_dir()
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "app.log"

    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    return logger
