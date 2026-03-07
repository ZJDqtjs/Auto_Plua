from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

APP_NAME = "AutoPlua"


def _config_dir() -> Path:
    appdata = os.getenv("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME
    return Path.home() / f".{APP_NAME.lower()}"


def config_path() -> Path:
    return _config_dir() / "config.json"


def default_config() -> dict[str, Any]:
    return {
        "programs": [],
        "schedules": [],
        "power_enabled": False,
        "power_settings": {
            "boot_frequency": "每天",
            "boot_time": "06:30",
            "shutdown_frequency": "每天",
            "shutdown_time": "23:00",
            "shutdown_action": "关机",
            "login_user": "",
            "login_domain": "",
            "login_password": "",
            "wake_mode": "Windows任务计划",
            "wol_mac": "",
            "wol_host": "255.255.255.255",
        },
    }


def load_config() -> dict[str, Any]:
    path = config_path()
    if not path.exists():
        return default_config()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return default_config()
        return {**default_config(), **data}
    except (json.JSONDecodeError, OSError):
        return default_config()


def save_config(data: dict[str, Any]) -> None:
    path = config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
