from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

APP_NAME = "AutoPlua"
USER_CONFIG_FILENAME = "autoplua.user.json"


def _config_dir() -> Path:
    appdata = os.getenv("APPDATA")
    if appdata:
        return Path(appdata) / APP_NAME
    return Path.home() / f".{APP_NAME.lower()}"


def _workspace_root() -> Path:
    return Path(__file__).resolve().parents[2]


def _legacy_config_path() -> Path:
    return _config_dir() / "config.json"


def config_path() -> Path:
    custom_path = os.getenv("AUTOPLUA_CONFIG_PATH", "").strip()
    if custom_path:
        return Path(custom_path).expanduser().resolve()

    workspace_target = _workspace_root() / USER_CONFIG_FILENAME
    return workspace_target


def default_config() -> dict[str, Any]:
    return {
        "config_version": 2,
        "profile_name": os.getenv("USERNAME", "default"),
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
            "virtual_display_auto_prepare": False,
            "virtual_display_auto_install": False,
            "virtual_display_prepare_on_app_start": True,
            "virtual_display_strict_isolation": True,
            "virtual_display_driver_inf": "",
        },
    }


def load_config() -> dict[str, Any]:
    path = config_path()
    if not path.exists():
        legacy = _legacy_config_path()
        if legacy.exists():
            try:
                data = json.loads(legacy.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    merged = {**default_config(), **data}
                    save_config(merged)
                    return merged
            except (json.JSONDecodeError, OSError):
                pass
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
