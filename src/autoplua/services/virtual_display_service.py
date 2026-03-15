from __future__ import annotations

import ctypes
import locale
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
        ("Data1", ctypes.c_uint32),
        ("Data2", ctypes.c_uint16),
        ("Data3", ctypes.c_uint16),
        ("Data4", ctypes.c_ubyte * 8),
    ]


class _SP_DEVINFO_DATA(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint32),
        ("ClassGuid", _GUID),
        ("DevInst", ctypes.c_uint32),
        ("Reserved", ctypes.c_size_t),
    ]


class VirtualDisplayService:
    """Manage virtual display driver installation and activation on Windows.

    This service does not build drivers. It installs and enables an existing INF driver
    package (for example IDD/Virtual Display Driver) through Windows tooling.
    """

    _VIRTUAL_HINTS = ("virtual", "indirect", "idd", "dummy", "usb display")
    _DEVICE_HINTS = ("mttvdd", "root\\mttvdd", "virtual display driver")
    _MONITORINFOF_PRIMARY = 0x00000001
    _SPDRP_HARDWAREID = 0x00000001
    _DICD_GENERATE_ID = 0x00000001
    _DIF_REGISTERDEVICE = 0x00000019
    _DISPLAY_CLASS_GUID = "{4D36E968-E325-11CE-BFC1-08002BE10318}"

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

        result = subprocess.run(
            ["pnputil", "/add-driver", str(path), "/install"],
            check=False,
            capture_output=True,
            text=True,
            encoding=locale.getpreferredencoding(False),
            errors="ignore",
        )
        if result.returncode != 0:
            # Repeated installs may report a non-zero code while the driver is already present.
            if self.is_virtual_driver_device_present() or self.has_non_primary_monitor():
                return True, "already-installed"

            output = f"{result.stdout or ''}\n{result.stderr or ''}".lower()
            duplicate_hints = (
                "already",
                "already installed",
                "already exists",
                "already added",
                "已安装",
                "已存在",
                "已添加",
            )
            if not any(hint in output for hint in duplicate_hints):
                return False, (result.stderr or result.stdout or "pnputil-failed").strip()

        # Trigger device re-enumeration to let newly installed virtual display appear quickly.
        subprocess.run(["pnputil", "/scan-devices"], check=False, capture_output=True)

        # Some virtual display drivers require an explicit root-enumerated device instance
        # (for example ROOT\MttVDD). pnputil may stage the package but not create the node.
        if not self.is_virtual_driver_device_present() and not self.has_non_primary_monitor():
            created, reason = self._create_root_device_and_bind_driver(
                hardware_id=r"ROOT\MttVDD",
                inf_path=path,
            )
            if not created:
                return False, f"virtual-driver-device-create-failed:{reason}"

        deadline = time.monotonic() + 6.0
        while time.monotonic() < deadline:
            if self.is_virtual_driver_device_present() or self.has_non_primary_monitor():
                return True, f"ok-{source}"
            time.sleep(0.5)

        out = (result.stdout or "").lower()
        if "driver package added successfully" in out:
            return False, "driver-package-added-but-device-not-created"
        return False, "virtual-driver-device-not-detected"

    def uninstall_driver(self, inf_path: str) -> tuple[bool, str]:
        if not self.is_admin():
            return False, "admin-required"

        preferred_inf: Path | None = None
        if inf_path.strip():
            preferred_inf = Path(inf_path).expanduser()
        else:
            preferred_inf = self.find_embedded_inf()

        removed_any = False

        for instance_id in self._list_virtual_device_instance_ids():
            result = subprocess.run(
                ["pnputil", "/remove-device", instance_id],
                check=False,
                capture_output=True,
                text=True,
                encoding=locale.getpreferredencoding(False),
                errors="ignore",
            )
            if result.returncode == 0:
                removed_any = True

        for published_name in self._list_virtual_published_driver_names(preferred_inf):
            result = subprocess.run(
                ["pnputil", "/delete-driver", published_name, "/uninstall", "/force"],
                check=False,
                capture_output=True,
                text=True,
                encoding=locale.getpreferredencoding(False),
                errors="ignore",
            )
            if result.returncode == 0:
                removed_any = True

        subprocess.run(["pnputil", "/scan-devices"], check=False, capture_output=True)

        if removed_any:
            return True, "ok"
        if not self.is_virtual_driver_device_present() and not self.has_non_primary_monitor():
            return True, "already-uninstalled"
        return False, "uninstall-failed"

    def _list_virtual_device_instance_ids(self) -> list[str]:
        command = (
            "Get-PnpDevice -Class Display -ErrorAction SilentlyContinue "
            "| Where-Object { $_.InstanceId -like 'ROOT\\MTTVDD*' -or $_.FriendlyName -like '*MttVDD*' } "
            "| Select-Object -ExpandProperty InstanceId"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            check=False,
            capture_output=True,
            text=True,
            encoding=locale.getpreferredencoding(False),
            errors="ignore",
        )
        if result.returncode != 0:
            return []
        return [line.strip() for line in (result.stdout or "").splitlines() if line.strip()]

    def _list_virtual_published_driver_names(self, preferred_inf: Path | None) -> list[str]:
        result = subprocess.run(
            ["pnputil", "/enum-drivers"],
            check=False,
            capture_output=True,
            text=True,
            encoding=locale.getpreferredencoding(False),
            errors="ignore",
        )
        if result.returncode != 0:
            return []

        text = result.stdout or ""
        blocks = re.split(r"\r?\n\s*\r?\n", text)
        matches: list[str] = []
        preferred_name = preferred_inf.name.lower() if preferred_inf else ""
        fallback_hints = ("mttvdd", "root\\mttvdd")

        for block in blocks:
            lower = block.lower()
            if preferred_name and preferred_name in lower:
                pass
            elif not any(hint in lower for hint in fallback_hints):
                continue

            for name in re.findall(r"oem\d+\.inf", lower):
                if name not in matches:
                    matches.append(name)

        return matches

    def _create_root_device_and_bind_driver(self, hardware_id: str, inf_path: Path) -> tuple[bool, str]:
        guid = self._guid_from_string(self._DISPLAY_CLASS_GUID)
        setupapi = ctypes.WinDLL("setupapi", use_last_error=True)
        newdev = ctypes.WinDLL("newdev", use_last_error=True)

        setupapi.SetupDiCreateDeviceInfoList.argtypes = [ctypes.POINTER(_GUID), ctypes.c_void_p]
        setupapi.SetupDiCreateDeviceInfoList.restype = ctypes.c_void_p

        setupapi.SetupDiCreateDeviceInfoW.argtypes = [
            ctypes.c_void_p,
            ctypes.c_wchar_p,
            ctypes.POINTER(_GUID),
            ctypes.c_wchar_p,
            ctypes.c_void_p,
            ctypes.c_uint32,
            ctypes.POINTER(_SP_DEVINFO_DATA),
        ]
        setupapi.SetupDiCreateDeviceInfoW.restype = ctypes.c_int

        setupapi.SetupDiSetDeviceRegistryPropertyW.argtypes = [
            ctypes.c_void_p,
            ctypes.POINTER(_SP_DEVINFO_DATA),
            ctypes.c_uint32,
            ctypes.POINTER(ctypes.c_ubyte),
            ctypes.c_uint32,
        ]
        setupapi.SetupDiSetDeviceRegistryPropertyW.restype = ctypes.c_int

        setupapi.SetupDiCallClassInstaller.argtypes = [ctypes.c_uint32, ctypes.c_void_p, ctypes.POINTER(_SP_DEVINFO_DATA)]
        setupapi.SetupDiCallClassInstaller.restype = ctypes.c_int

        setupapi.SetupDiDestroyDeviceInfoList.argtypes = [ctypes.c_void_p]
        setupapi.SetupDiDestroyDeviceInfoList.restype = ctypes.c_int

        newdev.UpdateDriverForPlugAndPlayDevicesW.argtypes = [
            ctypes.c_void_p,
            ctypes.c_wchar_p,
            ctypes.c_wchar_p,
            ctypes.c_uint32,
            ctypes.POINTER(ctypes.c_int),
        ]
        newdev.UpdateDriverForPlugAndPlayDevicesW.restype = ctypes.c_int

        handle = setupapi.SetupDiCreateDeviceInfoList(ctypes.byref(guid), None)
        if handle in (None, ctypes.c_void_p(-1).value):
            return False, f"SetupDiCreateDeviceInfoList-{ctypes.get_last_error()}"

        try:
            info = _SP_DEVINFO_DATA()
            info.cbSize = ctypes.sizeof(_SP_DEVINFO_DATA)

            ok_create = setupapi.SetupDiCreateDeviceInfoW(
                handle,
                "MttVDD",
                ctypes.byref(guid),
                "Virtual Display Driver",
                None,
                self._DICD_GENERATE_ID,
                ctypes.byref(info),
            )
            if not ok_create:
                return False, f"SetupDiCreateDeviceInfoW-{ctypes.get_last_error()}"

            # REG_MULTI_SZ: "ROOT\\MttVDD\0\0"
            reg_multi = f"{hardware_id}\0\0".encode("utf-16le")
            raw_buf = (ctypes.c_ubyte * len(reg_multi)).from_buffer_copy(reg_multi)
            ok_prop = setupapi.SetupDiSetDeviceRegistryPropertyW(
                handle,
                ctypes.byref(info),
                self._SPDRP_HARDWAREID,
                raw_buf,
                len(reg_multi),
            )
            if not ok_prop:
                return False, f"SetupDiSetDeviceRegistryPropertyW-{ctypes.get_last_error()}"

            ok_reg = setupapi.SetupDiCallClassInstaller(
                self._DIF_REGISTERDEVICE,
                handle,
                ctypes.byref(info),
            )
            if not ok_reg:
                return False, f"SetupDiCallClassInstaller-{ctypes.get_last_error()}"

            reboot_required = ctypes.c_int(0)
            ok_bind = newdev.UpdateDriverForPlugAndPlayDevicesW(
                None,
                hardware_id,
                str(inf_path),
                0,
                ctypes.byref(reboot_required),
            )
            if not ok_bind:
                return False, f"UpdateDriverForPlugAndPlayDevicesW-{ctypes.get_last_error()}"

            subprocess.run(["pnputil", "/scan-devices"], check=False, capture_output=True)
            return True, "ok"
        finally:
            setupapi.SetupDiDestroyDeviceInfoList(handle)

    @staticmethod
    def _guid_from_string(text: str) -> _GUID:
        parsed = uuid.UUID(text)
        raw = parsed.bytes_le
        out = _GUID()
        out.Data1 = int.from_bytes(raw[0:4], "little")
        out.Data2 = int.from_bytes(raw[4:6], "little")
        out.Data3 = int.from_bytes(raw[6:8], "little")
        out.Data4 = (ctypes.c_ubyte * 8)(*raw[8:16])
        return out

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
            encoding=locale.getpreferredencoding(False),
            errors="ignore",
        )
        if result.returncode != 0:
            return False

        text = (result.stdout or "").lower()
        return any(hint in text for hint in self._VIRTUAL_HINTS)

    def is_virtual_driver_device_present(self) -> bool:
        command = (
            "Get-PnpDevice -Class Display -ErrorAction SilentlyContinue "
            "| Where-Object { $_.InstanceId -like 'ROOT\\MTTVDD*' -or $_.FriendlyName -like '*MttVDD*' -or $_.FriendlyName -like '*Virtual Display Driver*' } "
            "| ForEach-Object { ($_.FriendlyName + '|' + $_.InstanceId) }"
        )
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            check=False,
            capture_output=True,
            text=True,
            encoding=locale.getpreferredencoding(False),
            errors="ignore",
        )
        if result.returncode != 0:
            return False

        text = (result.stdout or "").lower()
        return any(hint in text for hint in self._DEVICE_HINTS)

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

        driver_present = self.is_virtual_driver_device_present()
        monitor_present = self.is_virtual_display_present()
        if not driver_present and not monitor_present:
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

        if self.is_virtual_driver_device_present() or self.is_virtual_display_present():
            return False, "virtual-display-present-but-not-extended"
        return False, "virtual-display-not-detected"

    def auto_prepare(self, inf_path: str, auto_install: bool) -> tuple[bool, str]:
        return self.ensure_automation_display_ready(inf_path=inf_path, auto_install=auto_install)
