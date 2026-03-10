from __future__ import annotations

import ctypes
import os
import subprocess
from pathlib import Path


class VirtualDisplayService:
    """Manage virtual display driver installation and activation on Windows.

    This service does not build drivers. It installs and enables an existing INF driver
    package (for example IDD/Virtual Display Driver) through Windows tooling.
    """

    _VIRTUAL_HINTS = ("virtual", "indirect", "idd", "dummy", "usb display")

    @staticmethod
    def is_admin() -> bool:
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    def install_driver_from_inf(self, inf_path: str) -> tuple[bool, str]:
        path = Path(inf_path).expanduser()
        if not path.exists() or path.suffix.lower() != ".inf":
            return False, "invalid-inf-path"

        if not self.is_admin():
            return False, "admin-required"

        result = subprocess.run(
            ["pnputil", "/add-driver", str(path), "/install"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode != 0:
            return False, (result.stderr or result.stdout or "pnputil-failed").strip()

        # Trigger device re-enumeration to let newly installed virtual display appear quickly.
        subprocess.run(["pnputil", "/scan-devices"], check=False, capture_output=True)
        return True, "ok"

    @staticmethod
    def enable_extend_mode() -> tuple[bool, str]:
        display_switch = Path(os.environ.get("WINDIR", r"C:\Windows")) / "System32" / "DisplaySwitch.exe"
        if not display_switch.exists():
            return False, "displayswitch-not-found"

        result = subprocess.run([str(display_switch), "/extend"], check=False, capture_output=True)
        if result.returncode != 0:
            return False, "displayswitch-failed"
        return True, "ok"

    def is_virtual_display_present(self) -> bool:
        # Use powershell to query active monitor names and IDs.
        command = (
            "Get-CimInstance -Namespace root\\wmi -ClassName WmiMonitorID "
            "| ForEach-Object { [System.Text.Encoding]::ASCII.GetString($_.UserFriendlyName).Trim([char]0) }"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode != 0:
            return False

        text = (result.stdout or "").lower()
        return any(hint in text for hint in self._VIRTUAL_HINTS)

    def auto_prepare(self, inf_path: str, auto_install: bool) -> tuple[bool, str]:
        if self.is_virtual_display_present():
            ok, message = self.enable_extend_mode()
            if not ok:
                return False, message
            return True, "ready"

        if not auto_install:
            return False, "virtual-display-not-present"

        ok, message = self.install_driver_from_inf(inf_path)
        if not ok:
            return False, f"install-failed-{message}"

        ok, message = self.enable_extend_mode()
        if not ok:
            return False, f"extend-failed-{message}"

        if self.is_virtual_display_present():
            return True, "installed-and-ready"
        return False, "virtual-display-not-detected"
