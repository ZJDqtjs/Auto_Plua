from __future__ import annotations

import ctypes
import os
import subprocess
from datetime import datetime, timezone


_WINDOWS_TICKS_PER_SECOND = 10_000_000
_WINDOWS_EPOCH_DIFF_SECONDS = 11_644_473_600


class PowerService:
    def __init__(self) -> None:
        self._wake_timer_handle = None

    @staticmethod
    def shutdown() -> None:
        os.system("shutdown /s /t 0")

    @staticmethod
    def restart() -> None:
        os.system("shutdown /r /t 0")

    @staticmethod
    def sleep() -> None:
        # Explicitly request sleep (S3) instead of hibernation and avoid force-critical path.
        ctypes.windll.powrprof.SetSuspendState(False, False, False)

    @staticmethod
    def lock() -> None:
        ctypes.windll.user32.LockWorkStation()

    @staticmethod
    def cancel_shutdown() -> None:
        subprocess.run(["shutdown", "/a"], check=False)

    def cancel_wake_timer(self) -> bool:
        if self._wake_timer_handle is None:
            return False

        ctypes.windll.kernel32.CancelWaitableTimer(self._wake_timer_handle)
        ctypes.windll.kernel32.CloseHandle(self._wake_timer_handle)
        self._wake_timer_handle = None
        return True

    def schedule_wake(self, wake_time: datetime) -> bool:
        # Rebuild timer whenever settings change to ensure only one wake alarm is active.
        self.cancel_wake_timer()

        handle = ctypes.windll.kernel32.CreateWaitableTimerW(None, True, None)
        if not handle:
            return False

        wake_utc = wake_time.astimezone(timezone.utc)
        unix_seconds = wake_utc.timestamp()
        filetime_ticks = int((unix_seconds + _WINDOWS_EPOCH_DIFF_SECONDS) * _WINDOWS_TICKS_PER_SECOND)
        due_time = ctypes.c_longlong(filetime_ticks)

        ok = ctypes.windll.kernel32.SetWaitableTimer(
            handle,
            ctypes.byref(due_time),
            0,
            None,
            None,
            True,
        )

        if ok:
            self._wake_timer_handle = handle
            return True

        ctypes.windll.kernel32.CloseHandle(handle)
        return False
