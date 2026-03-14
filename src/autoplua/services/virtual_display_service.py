from __future__ import annotations

import ctypes
import os
import re
import subprocess
import time
import uuid
from pathlib import Path


class _RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("rcMonitor", _RECT),
        ("rcWork", _RECT),
        ("dwFlags", ctypes.c_ulong),
    ]


class _GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]


class _SP_DEVINFO_DATA(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_ulong),
        ("ClassGuid", _GUID),
        ("DevInst", ctypes.c_ulong),
        ("Reserved", ctypes.c_void_p),
    ]


class VirtualDisplayService:
    """Manage virtual display driver installation and activation on Windows.

    This service does not build drivers. It installs and enables an existing INF driver
    package (for example IDD/Virtual Display Driver) through Windows tooling.
    """

    _VIRTUAL_HINTS = ("virtual", "indirect", "idd", "dummy", "usb display")
    _DRIVER_HINTS = ("mttvdd", "iddsample", "indirect display", "virtual display driver")
    _MONITORINFOF_PRIMARY = 0x00000001
    _DISPLAY_CLASS_GUID = "{4D36E968-E325-11CE-BFC1-08002BE10318}"

    _SPDRP_HARDWAREID = 0x00000001
    _DICD_GENERATE_ID = 0x00000001
    _DIF_REGISTERDEVICE = 0x00000019

    def _resolved_inf_name(self, inf_path: str) -> str:
        resolved, _ = self.resolve_driver_inf(inf_path)
        return resolved.name.lower() if resolved is not None else ""

    @staticmethod
    def _workspace_root() -> Path:
        return Path(__file__).resolve().parents[3]

    def embedded_driver_dir(self) -> Path:
        return self._workspace_root() / "drivers" / "virtual_display"

    def find_embedded_inf(self) -> Path | None:
        root = self.embedded_driver_dir()
        if not root.exists():
            return None

        candidates = sorted(root.rglob("*.inf"))
        if not candidates:
            return None

        # Prefer IDD-like names when multiple INFs are bundled.
        for path in candidates:
            name = path.name.lower()
            if any(key in name for key in ("idd", "virtual", "indirect", "display")):
                return path
        return candidates[0]

    def resolve_driver_inf(self, inf_path: str) -> tuple[Path | None, str]:
        custom = Path(inf_path).expanduser() if inf_path.strip() else None
        if custom:
            if custom.exists() and custom.suffix.lower() == ".inf":
                return custom, "custom"
            return None, "invalid-inf-path"

        embedded = self.find_embedded_inf()
        if embedded is None:
            return None, "embedded-driver-not-found"
        return embedded, "embedded"

    @staticmethod
    def is_admin() -> bool:
        try:
            return bool(ctypes.windll.shell32.IsUserAnAdmin())
        except Exception:
            return False

    def install_driver_from_inf(self, inf_path: str) -> tuple[bool, str]:
        path, source = self.resolve_driver_inf(inf_path)
        if path is None:
            return False, source

        if not self.is_admin():
            return False, "admin-required"

        staged_before = self.is_driver_package_staged(inf_path=str(path))

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

        # Some IDD packages are only staged by pnputil and still need explicit root-device creation.
        self._try_create_root_device_instance(path)

        staged_after = self.is_driver_package_staged(inf_path=str(path))
        if not staged_after:
            return False, "driver-package-not-staged"

        instance_present = self.is_target_display_device_present()
        if instance_present:
            return True, f"ok-{source}"

        if staged_before:
            return True, f"staged-only-{source}"
        return True, f"staged-only-{source}"

    def _extract_root_hardware_ids(self, inf_file: Path) -> list[str]:
        try:
            text = inf_file.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            return []

        ids: list[str] = []
        for line in text.splitlines():
            candidate = line.split(";", 1)[0].strip()
            if not candidate:
                continue
            match = re.search(r"(root\\[^\s,]+)", candidate, re.IGNORECASE)
            if not match:
                continue
            hwid = match.group(1)
            if hwid.lower() not in [x.lower() for x in ids]:
                ids.append(hwid)
        return ids

    @staticmethod
    def _guid_from_string(text: str) -> _GUID:
        u = uuid.UUID(text)
        data4 = (ctypes.c_ubyte * 8)(*u.bytes[8:16])
        return _GUID(
            Data1=u.time_low,
            Data2=u.time_mid,
            Data3=u.time_hi_version,
            Data4=data4,
        )

    def _create_root_device_setupapi(self, hardware_id: str, inf_file: Path) -> bool:
        setupapi = ctypes.windll.setupapi
        newdev = ctypes.windll.newdev

        class_guid = self._guid_from_string(self._DISPLAY_CLASS_GUID)

        hdevinfo = setupapi.SetupDiCreateDeviceInfoList(ctypes.byref(class_guid), None)
        if hdevinfo in (ctypes.c_void_p(-1).value, 0):
            return False

        devinfo = _SP_DEVINFO_DATA()
        devinfo.cbSize = ctypes.sizeof(_SP_DEVINFO_DATA)

        try:
            created = setupapi.SetupDiCreateDeviceInfoW(
                hdevinfo,
                "AutoPlua Virtual Display",
                ctypes.byref(class_guid),
                None,
                None,
                self._DICD_GENERATE_ID,
                ctypes.byref(devinfo),
            )
            if not created:
                return False

            # REG_MULTI_SZ for hardware IDs (utf-16le ending with double NUL)
            multi = (hardware_id + "\0\0").encode("utf-16le")
            multi_buf = ctypes.create_string_buffer(multi)
            set_prop = setupapi.SetupDiSetDeviceRegistryPropertyW(
                hdevinfo,
                ctypes.byref(devinfo),
                self._SPDRP_HARDWAREID,
                ctypes.cast(multi_buf, ctypes.c_void_p),
                len(multi),
            )
            if not set_prop:
                return False

            registered = setupapi.SetupDiCallClassInstaller(
                self._DIF_REGISTERDEVICE,
                hdevinfo,
                ctypes.byref(devinfo),
            )
            if not registered:
                return False

            reboot = ctypes.c_bool(False)
            _ = newdev.UpdateDriverForPlugAndPlayDevicesW(
                None,
                hardware_id,
                str(inf_file),
                0,
                ctypes.byref(reboot),
            )
            return True
        finally:
            setupapi.SetupDiDestroyDeviceInfoList(hdevinfo)

    def _try_create_root_device_instance(self, inf_file: Path) -> None:
        hardware_ids = self._extract_root_hardware_ids(inf_file)
        if not hardware_ids:
            return

        devcon_path = None
        where = subprocess.run(
            ["where", "devcon"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if where.returncode == 0:
            first = (where.stdout or "").splitlines()
            if first:
                devcon_path = first[0].strip()

        for hwid in hardware_ids:
            if devcon_path:
                subprocess.run(
                    [devcon_path, "install", str(inf_file), hwid],
                    check=False,
                    capture_output=True,
                    text=True,
                    encoding="utf-8",
                    errors="ignore",
                )
            else:
                self._create_root_device_setupapi(hwid, inf_file)

        # Always request re-enumeration after instance-creation attempts.
        subprocess.run(["pnputil", "/scan-devices"], check=False, capture_output=True)

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

    def is_virtual_display_driver_installed(self, inf_path: str = "") -> bool:
        command = (
            "Get-CimInstance Win32_PnPSignedDriver "
            "| ForEach-Object { \"$($_.DeviceName)|$($_.DriverName)|$($_.InfName)\" }"
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

        raw = result.stdout or ""
        text = raw.lower()

        target_inf_name = self._resolved_inf_name(inf_path)

        if target_inf_name:
            for line in raw.splitlines():
                parts = [p.strip().lower() for p in line.split("|")]
                if len(parts) < 3:
                    continue
                if parts[2] == target_inf_name:
                    return True

        return any(hint in text for hint in self._DRIVER_HINTS)

    def is_driver_package_staged(self, inf_path: str = "") -> bool:
        target_inf_name = self._resolved_inf_name(inf_path)
        if not target_inf_name:
            return False

        result = subprocess.run(
            ["pnputil", "/enum-drivers"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
        )
        if result.returncode != 0:
            return False

        return target_inf_name in (result.stdout or "").lower()

    def is_target_display_device_present(self) -> bool:
        command = (
            "Get-PnpDevice -Class Display "
            "| ForEach-Object { \"$($_.FriendlyName)|$($_.InstanceId)|$($_.Status)\" }"
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
        return (
            "root\\mttvdd" in text
            or "|virtual display driver|" in text
            or "mttvdd" in text
        )

    def has_non_primary_monitor(self) -> bool:
        user32 = ctypes.windll.user32
        monitors: list[int] = []

        @ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_void_p, ctypes.c_void_p, ctypes.POINTER(_RECT), ctypes.c_long)
        def enum_proc(hmonitor, _hdc, _rect, _lparam):
            info = _MONITORINFO()
            info.cbSize = ctypes.sizeof(_MONITORINFO)
            if user32.GetMonitorInfoW(hmonitor, ctypes.byref(info)):
                if not bool(info.dwFlags & self._MONITORINFOF_PRIMARY):
                    monitors.append(1)
            return 1

        user32.EnumDisplayMonitors(0, 0, enum_proc, 0)
        return bool(monitors)

    def ensure_automation_display_ready(
        self,
        inf_path: str,
        auto_install: bool,
        wait_seconds: float = 8.0,
    ) -> tuple[bool, str]:
        if self.has_non_primary_monitor():
            ok, message = self.enable_extend_mode()
            if not ok:
                return False, message
            return True, "ready"

        driver_installed = self.is_virtual_display_driver_installed(inf_path=inf_path) or self.is_virtual_display_present()
        if not driver_installed:
            if not auto_install:
                return False, "virtual-display-not-present"

            ok, message = self.install_driver_from_inf(inf_path)
            if not ok:
                return False, f"install-failed-{message}"

        ok, message = self.enable_extend_mode()
        if not ok:
            return False, f"extend-failed-{message}"

        deadline = time.monotonic() + max(1.0, wait_seconds)
        while time.monotonic() < deadline:
            if self.has_non_primary_monitor():
                return True, "installed-and-ready"
            time.sleep(0.4)

        if self.is_driver_package_staged(inf_path=inf_path) and not self.is_target_display_device_present():
            return False, "virtual-device-instance-missing"

        if self.is_virtual_display_driver_installed(inf_path=inf_path) or self.is_virtual_display_present():
            return False, "virtual-display-present-but-not-extended"
        return False, "virtual-display-not-detected"

    def auto_prepare(self, inf_path: str, auto_install: bool) -> tuple[bool, str]:
        return self.ensure_automation_display_ready(inf_path=inf_path, auto_install=auto_install)
